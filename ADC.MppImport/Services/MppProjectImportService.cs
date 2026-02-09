using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using ADC.MppImport.MppReader.Mpp;
using ADC.MppImport.MppReader.Model;

namespace ADC.MppImport.Services
{
    /// <summary>
    /// Testable business logic for importing MPP file tasks into Project Operations.
    /// No workflow/CRM framework dependencies in the constructor — only IOrganizationService + ITracingService passed in.
    /// </summary>
    public class MppProjectImportService
    {
        private readonly IOrganizationService _service;
        private readonly ITracingService _trace;

        public MppProjectImportService(IOrganizationService service, ITracingService trace)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _trace = trace;
        }

        /// <summary>
        /// Main entry point: reads MPP from the case template file field, parses it,
        /// and upserts project tasks into msdyn_projecttask for the given project.
        /// Returns an ImportResult with counts.
        /// </summary>
        public ImportResult Execute(Guid caseTemplateId, Guid projectId)
        {
            var result = new ImportResult();

            // 1. Download the MPP file bytes from the case template file field
            _trace?.Trace("Downloading MPP file from adc_adccasetemplate {0}...", caseTemplateId);
            byte[] mppBytes = DownloadFileColumn(caseTemplateId, "adc_adccasetemplate", "adc_templatemsprojectmppfile");
            if (mppBytes == null || mppBytes.Length == 0)
                throw new InvalidPluginExecutionException("No MPP file found on the case template record.");

            _trace?.Trace("MPP file downloaded: {0} bytes", mppBytes.Length);

            // 2. Parse the MPP file
            var reader = new MppFileReader();
            ProjectFile project = reader.Read(mppBytes);
            _trace?.Trace("MPP parsed: {0} tasks, {1} resources", project.Tasks.Count, project.Resources.Count);

            if (project.Tasks.Count == 0)
            {
                _trace?.Trace("No tasks found in MPP file.");
                return result;
            }

            // 3. Retrieve existing project tasks for this project (for upsert matching)
            var existingTasks = RetrieveExistingProjectTasks(projectId);
            _trace?.Trace("Existing msdyn_projecttask records: {0}", existingTasks.Count);

            // 4. Build a lookup of existing tasks by msdyn_msprojectclientid
            var existingByClientId = new Dictionary<string, Entity>(StringComparer.OrdinalIgnoreCase);
            foreach (var et in existingTasks)
            {
                string clientId = et.GetAttributeValue<string>("msdyn_msprojectclientid");
                if (!string.IsNullOrEmpty(clientId) && !existingByClientId.ContainsKey(clientId))
                    existingByClientId[clientId] = et;
            }

            // 5. First pass: create/update tasks and build a map of MPP UniqueID -> msdyn_projecttask ID
            var taskIdMap = new Dictionary<int, Guid>(); // MPP UniqueID -> CRM record ID
            var projectRef = new EntityReference("msdyn_project", projectId);

            foreach (var mppTask in project.Tasks)
            {
                if (!mppTask.UniqueID.HasValue) continue;

                string clientId = mppTask.UniqueID.Value.ToString();
                Entity existing = null;
                existingByClientId.TryGetValue(clientId, out existing);

                Entity taskEntity = MapMppTaskToEntity(mppTask, projectRef, clientId);

                if (existing != null)
                {
                    // Update
                    taskEntity.Id = existing.Id;
                    _service.Update(taskEntity);
                    taskIdMap[mppTask.UniqueID.Value] = existing.Id;
                    result.TasksUpdated++;
                }
                else
                {
                    // Create
                    Guid newId = _service.Create(taskEntity);
                    taskIdMap[mppTask.UniqueID.Value] = newId;
                    result.TasksCreated++;
                }
            }

            // 6. Second pass: set parent task references
            foreach (var mppTask in project.Tasks)
            {
                if (!mppTask.UniqueID.HasValue) continue;
                if (mppTask.ParentTask == null || !mppTask.ParentTask.UniqueID.HasValue) continue;

                Guid taskRecordId;
                Guid parentRecordId;
                if (!taskIdMap.TryGetValue(mppTask.UniqueID.Value, out taskRecordId)) continue;
                if (!taskIdMap.TryGetValue(mppTask.ParentTask.UniqueID.Value, out parentRecordId)) continue;

                var update = new Entity("msdyn_projecttask", taskRecordId);
                update["msdyn_parenttask"] = new EntityReference("msdyn_projecttask", parentRecordId);
                try
                {
                    _service.Update(update);
                }
                catch (Exception ex)
                {
                    _trace?.Trace("Warning: Could not set parent for task {0}: {1}",
                        mppTask.UniqueID.Value, ex.Message);
                }
            }

            _trace?.Trace("Import complete. Created: {0}, Updated: {1}", result.TasksCreated, result.TasksUpdated);
            return result;
        }

        /// <summary>
        /// Maps an MPP Task to a msdyn_projecttask Entity for create/update.
        /// </summary>
        private Entity MapMppTaskToEntity(Task mppTask, EntityReference projectRef, string clientId)
        {
            var entity = new Entity("msdyn_projecttask");
            entity["msdyn_project"] = projectRef;
            entity["msdyn_msprojectclientid"] = clientId;
            entity["msdyn_subject"] = mppTask.Name ?? "(Unnamed Task)";

            if (mppTask.Duration != null && mppTask.Duration.Value > 0)
            {
                // msdyn_duration expects decimal hours; convert based on units
                double durationHours = ConvertToHours(mppTask.Duration);
                // Convert to days (8h/day standard)
                entity["msdyn_duration"] = Math.Round((decimal)(durationHours / 8.0), 2);
            }

            if (mppTask.Start.HasValue)
                entity["msdyn_scheduledstart"] = mppTask.Start.Value;

            if (mppTask.Finish.HasValue)
                entity["msdyn_scheduledend"] = mppTask.Finish.Value;

            if (mppTask.OutlineLevel > 0)
                entity["msdyn_outlinelevel"] = mppTask.OutlineLevel;

            if (!string.IsNullOrEmpty(mppTask.WBS))
                entity["msdyn_wbsid"] = mppTask.WBS;

            if (mppTask.PercentComplete > 0)
                entity["msdyn_progress"] = (decimal)mppTask.PercentComplete;

            // msdyn_projecttask has a scheduledstart/end that ProjOps uses
            // Also set effort if we have work data
            if (mppTask.Work != null && mppTask.Work.Value > 0)
            {
                double workHours = ConvertToHours(mppTask.Work);
                entity["msdyn_effort"] = Math.Round((decimal)workHours, 2);
            }

            return entity;
        }

        /// <summary>
        /// Converts a Duration to hours based on its time units.
        /// </summary>
        private static double ConvertToHours(Duration duration)
        {
            double val = duration.Value;
            switch (duration.Units)
            {
                case TimeUnit.Minutes: return val / 60.0;
                case TimeUnit.Hours: return val;
                case TimeUnit.Days: return val * 8.0;
                case TimeUnit.Weeks: return val * 40.0;
                case TimeUnit.Months: return val * 160.0;
                default: return val;
            }
        }

        /// <summary>
        /// Downloads a file column's content as byte[] using InitializeFileBlocksDownload.
        /// </summary>
        private byte[] DownloadFileColumn(Guid recordId, string entityName, string fileAttributeName)
        {
            try
            {
                var initRequest = new OrganizationRequest("InitializeFileBlocksDownload");
                initRequest["Target"] = new EntityReference(entityName, recordId);
                initRequest["FileAttributeName"] = fileAttributeName;

                var initResponse = _service.Execute(initRequest);
                string fileContinuationToken = (string)initResponse["FileContinuationToken"];
                long fileSize = (long)initResponse["FileSizeInBytes"];

                if (fileSize == 0) return null;

                var allBytes = new List<byte>();
                long offset = 0;
                const long blockSize = 4 * 1024 * 1024; // 4MB blocks

                while (offset < fileSize)
                {
                    var downloadRequest = new OrganizationRequest("DownloadBlock");
                    downloadRequest["FileContinuationToken"] = fileContinuationToken;
                    downloadRequest["BlockLength"] = blockSize;
                    downloadRequest["Offset"] = offset;

                    var downloadResponse = _service.Execute(downloadRequest);
                    byte[] blockData = (byte[])downloadResponse["Data"];
                    allBytes.AddRange(blockData);
                    offset += blockData.Length;
                }

                return allBytes.ToArray();
            }
            catch (Exception ex)
            {
                _trace?.Trace("Error downloading file: {0}", ex.Message);
                throw new InvalidPluginExecutionException(
                    $"Failed to download MPP file from {entityName}.{fileAttributeName}: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Retrieves all existing msdyn_projecttask records for a given project.
        /// </summary>
        private List<Entity> RetrieveExistingProjectTasks(Guid projectId)
        {
            var query = new QueryExpression("msdyn_projecttask")
            {
                ColumnSet = new ColumnSet("msdyn_msprojectclientid", "msdyn_subject", "msdyn_wbsid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("msdyn_project", ConditionOperator.Equal, projectId),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0) // active only
                    }
                }
            };

            var results = new List<Entity>();
            query.PageInfo = new PagingInfo { PageNumber = 1, Count = 5000 };

            while (true)
            {
                var response = _service.RetrieveMultiple(query);
                results.AddRange(response.Entities);

                if (response.MoreRecords)
                {
                    query.PageInfo.PageNumber++;
                    query.PageInfo.PagingCookie = response.PagingCookie;
                }
                else
                {
                    break;
                }
            }

            return results;
        }
    }

    /// <summary>
    /// Result of an MPP import operation.
    /// </summary>
    public class ImportResult
    {
        public int TasksCreated { get; set; }
        public int TasksUpdated { get; set; }
        public int TotalProcessed => TasksCreated + TasksUpdated;
        public List<string> Warnings { get; set; } = new List<string>();
    }
}
