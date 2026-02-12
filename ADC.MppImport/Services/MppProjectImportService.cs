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
    /// Uses the Project Scheduling Service (PSS) API to create/update msdyn_projecttask records.
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
        /// and creates/updates project tasks via the PSS API for the given project.
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

            // 5. Retrieve or create the default project bucket
            var projectRef = new EntityReference("msdyn_project", projectId);
            EntityReference bucketRef = GetOrCreateDefaultBucket(projectId);
            _trace?.Trace("Using project bucket: {0}", bucketRef.Id);

            // 6. Create an OperationSet for the PSS batch
            string operationSetId = CreateOperationSet(projectId, "MPP Import");
            _trace?.Trace("OperationSet created: {0}", operationSetId);

            // 7. Pre-generate IDs for new tasks; build map of MPP UniqueID -> CRM record GUID
            var taskIdMap = new Dictionary<int, Guid>();

            // First pass: queue PssCreate or PssUpdate for each task
            foreach (var mppTask in project.Tasks)
            {
                if (!mppTask.UniqueID.HasValue) continue;

                string clientId = mppTask.UniqueID.Value.ToString();
                Entity existing = null;
                existingByClientId.TryGetValue(clientId, out existing);

                Entity taskEntity = MapMppTaskToEntity(mppTask, projectRef, bucketRef);

                if (existing != null)
                {
                    // Update via PSS
                    taskEntity.Id = existing.Id;
                    taskEntity.LogicalName = "msdyn_projecttask";
                    PssUpdate(taskEntity, operationSetId);
                    taskIdMap[mppTask.UniqueID.Value] = existing.Id;
                    result.TasksUpdated++;
                    _trace?.Trace("  Queued UPDATE for task [{0}] {1}", mppTask.UniqueID.Value, mppTask.Name);
                }
                else
                {
                    // Create via PSS — pre-generate an ID so we can reference it for parent links
                    Guid newId = Guid.NewGuid();
                    taskEntity.Id = newId;
                    taskEntity["msdyn_projecttaskid"] = newId;
                    PssCreate(taskEntity, operationSetId);
                    taskIdMap[mppTask.UniqueID.Value] = newId;
                    result.TasksCreated++;
                    _trace?.Trace("  Queued CREATE for task [{0}] {1} -> {2}", mppTask.UniqueID.Value, mppTask.Name, newId);
                }
            }

            // 8. Second pass: queue PssUpdate to set parent task references
            int parentLinksSet = 0;
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
                PssUpdate(update, operationSetId);
                parentLinksSet++;
            }

            _trace?.Trace("Parent links queued: {0}", parentLinksSet);

            // 9. Third pass: create predecessor/successor dependency records
            int dependenciesCreated = 0;
            foreach (var mppTask in project.Tasks)
            {
                if (!mppTask.UniqueID.HasValue) continue;
                if (mppTask.Predecessors == null || mppTask.Predecessors.Count == 0) continue;

                Guid successorId;
                if (!taskIdMap.TryGetValue(mppTask.UniqueID.Value, out successorId)) continue;

                foreach (var relation in mppTask.Predecessors)
                {
                    Guid predecessorId;
                    if (!taskIdMap.TryGetValue(relation.SourceTaskUniqueID, out predecessorId)) continue;

                    var dep = new Entity("msdyn_projecttaskdependency");
                    dep.Id = Guid.NewGuid();
                    dep["msdyn_projecttaskdependencyid"] = dep.Id;
                    dep["msdyn_project"] = projectRef;
                    dep["msdyn_predecessortask"] = new EntityReference("msdyn_projecttask", predecessorId);
                    dep["msdyn_successortask"] = new EntityReference("msdyn_projecttask", successorId);
                    dep["msdyn_linktype"] = new OptionSetValue(MapRelationType(relation.Type));

                    PssCreate(dep, operationSetId);
                    dependenciesCreated++;
                }
            }

            _trace?.Trace("Dependencies queued: {0}", dependenciesCreated);
            result.DependenciesCreated = dependenciesCreated;

            // 10. Execute the OperationSet to commit all changes
            _trace?.Trace("Executing OperationSet...");
            try
            {
                string executeResult = ExecuteOperationSet(operationSetId);
                _trace?.Trace("OperationSet executed successfully. Result: {0}", executeResult ?? "(none)");
            }
            catch (Exception ex)
            {
                _trace?.Trace("OperationSet execution failed: {0}", ex.Message);
                if (ex.InnerException != null)
                    _trace?.Trace("Inner: {0}", ex.InnerException.Message);

                // Query the operation set record for detailed error info
                try
                {
                    LogOperationSetStatus(operationSetId);
                }
                catch (Exception logEx)
                {
                    _trace?.Trace("Could not retrieve OperationSet status: {0}", logEx.Message);
                }

                throw;
            }

            _trace?.Trace("Import complete. Created: {0}, Updated: {1}", result.TasksCreated, result.TasksUpdated);
            return result;
        }

        #region PSS API Helpers

        /// <summary>
        /// Calls msdyn_CreateOperationSetV1 to create a new operation set for batching PSS operations.
        /// </summary>
        private string CreateOperationSet(Guid projectId, string description)
        {
            var request = new OrganizationRequest("msdyn_CreateOperationSetV1");
            request["ProjectId"] = projectId.ToString();
            request["Description"] = description;
            var response = _service.Execute(request);
            return (string)response["OperationSetId"];
        }

        /// <summary>
        /// Calls msdyn_PssCreateV1 to queue a create operation in the operation set.
        /// </summary>
        private void PssCreate(Entity entity, string operationSetId)
        {
            var request = new OrganizationRequest("msdyn_PssCreateV1");
            request["Entity"] = entity;
            request["OperationSetId"] = operationSetId;
            _service.Execute(request);
        }

        /// <summary>
        /// Calls msdyn_PssUpdateV1 to queue an update operation in the operation set.
        /// </summary>
        private void PssUpdate(Entity entity, string operationSetId)
        {
            var request = new OrganizationRequest("msdyn_PssUpdateV1");
            request["Entity"] = entity;
            request["OperationSetId"] = operationSetId;
            _service.Execute(request);
        }

        /// <summary>
        /// Calls msdyn_ExecuteOperationSetV1 to commit all queued operations.
        /// </summary>
        private string ExecuteOperationSet(string operationSetId)
        {
            var request = new OrganizationRequest("msdyn_ExecuteOperationSetV1");
            request["OperationSetId"] = operationSetId;
            var response = _service.Execute(request);

            // Safely read result — key may vary by version
            foreach (var kvp in response.Results)
            {
                _trace?.Trace("  Response key: {0} = {1}", kvp.Key, kvp.Value);
            }

            if (response.Results.ContainsKey("OperationSetDetailId"))
                return response["OperationSetDetailId"] as string;

            return operationSetId;
        }

        /// <summary>
        /// Queries the msdyn_operationset record to log detailed status/error info.
        /// </summary>
        private void LogOperationSetStatus(string operationSetId)
        {
            Guid osId;
            if (!Guid.TryParse(operationSetId, out osId)) return;

            var osRecord = _service.Retrieve("msdyn_operationset", osId,
                new ColumnSet("msdyn_status", "msdyn_statuscode", "msdyn_description"));

            foreach (var attr in osRecord.Attributes)
            {
                _trace?.Trace("  OperationSet.{0} = {1}", attr.Key, attr.Value);
            }

            // Also check for failed operation set detail records
            var detailQuery = new QueryExpression("msdyn_operationsetdetail")
            {
                ColumnSet = new ColumnSet(true),
                TopCount = 10,
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("msdyn_operationsetid", ConditionOperator.Equal, osId)
                    }
                }
            };

            var details = _service.RetrieveMultiple(detailQuery);
            _trace?.Trace("  OperationSetDetail records: {0}", details.Entities.Count);
            foreach (var detail in details.Entities)
            {
                foreach (var attr in detail.Attributes)
                {
                    _trace?.Trace("    {0} = {1}", attr.Key, attr.Value);
                }
                _trace?.Trace("    ---");
            }
        }

        #endregion

        #region Task Mapping

        /// <summary>
        /// Maps an MPP Task to a msdyn_projecttask Entity for create/update.
        /// </summary>
        private Entity MapMppTaskToEntity(Task mppTask, EntityReference projectRef, EntityReference bucketRef)
        {
            var entity = new Entity("msdyn_projecttask");
            entity["msdyn_project"] = projectRef;
            entity["msdyn_projectbucket"] = bucketRef;
            entity["msdyn_subject"] = mppTask.Name ?? "(Unnamed Task)";
            entity["msdyn_LinkStatus"] = new OptionSetValue(192350000); // Not Linked

            // Summary tasks (parents): PSS auto-calculates duration/effort from children
            // Only set duration/effort on leaf tasks
            if (!mppTask.HasChildTasks)
            {
                if (mppTask.Duration != null && mppTask.Duration.Value > 0)
                {
                    double durationHours = ConvertToHours(mppTask.Duration);
                    entity["msdyn_duration"] = Math.Round(durationHours / 8.0, 2);
                }

                if (mppTask.Work != null && mppTask.Work.Value > 0)
                {
                    double workHours = ConvertToHours(mppTask.Work);
                    entity["msdyn_effort"] = Math.Round(workHours, 2);
                }
            }

            return entity;
        }

        /// <summary>
        /// Maps MPP RelationType to Project Operations msdyn_linktype OptionSetValue.
        /// </summary>
        private static int MapRelationType(RelationType type)
        {
            switch (type)
            {
                case RelationType.FinishToStart:  return 192350000;
                case RelationType.StartToStart:   return 192350001;
                case RelationType.FinishToFinish: return 192350002;
                case RelationType.StartToFinish:  return 192350003;
                default:                          return 192350000; // default FS
            }
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

        #endregion

        #region Dataverse Helpers

        /// <summary>
        /// Retrieves the first existing msdyn_projectbucket for the project, or creates a default one.
        /// </summary>
        private EntityReference GetOrCreateDefaultBucket(Guid projectId)
        {
            // Try to find an existing bucket for this project
            var query = new QueryExpression("msdyn_projectbucket")
            {
                ColumnSet = new ColumnSet("msdyn_projectbucketid", "msdyn_name"),
                TopCount = 1,
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("msdyn_project", ConditionOperator.Equal, projectId)
                    }
                }
            };

            var results = _service.RetrieveMultiple(query);
            if (results.Entities.Count > 0)
            {
                var existing = results.Entities[0];
                _trace?.Trace("Found existing bucket: {0}", existing.GetAttributeValue<string>("msdyn_name") ?? existing.Id.ToString());
                return existing.ToEntityReference();
            }

            // No bucket exists — create a default one
            _trace?.Trace("No project bucket found, creating default bucket...");
            var bucket = new Entity("msdyn_projectbucket");
            bucket["msdyn_project"] = new EntityReference("msdyn_project", projectId);
            bucket["msdyn_name"] = "MPP Import";
            Guid bucketId = _service.Create(bucket);
            return new EntityReference("msdyn_projectbucket", bucketId);
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
                ColumnSet = new ColumnSet("msdyn_msprojectclientid", "msdyn_subject"),
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

        #endregion
    }

    /// <summary>
    /// Result of an MPP import operation.
    /// </summary>
    public class ImportResult
    {
        public int TasksCreated { get; set; }
        public int TasksUpdated { get; set; }
        public int DependenciesCreated { get; set; }
        public int TotalProcessed => TasksCreated + TasksUpdated;
        public List<string> Warnings { get; set; } = new List<string>();
    }
}
