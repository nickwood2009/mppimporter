using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ADC.MppImport.Services
{
    /// <summary>
    /// Shared business logic for importing an MPP template into a project linked to an adc_case.
    /// Called from both CaseCreatePlugin and ImportCaseActivity.
    /// </summary>
    public class CaseImportService
    {
        private readonly IOrganizationService _service;
        private readonly ITracingService _trace;

        public CaseImportService(IOrganizationService service, ITracingService trace)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _trace = trace;
        }

        /// <summary>
        /// Runs the full case import process: resolve template, download MPP, create project, start async import.
        /// </summary>
        /// <param name="caseId">The adc_case record ID.</param>
        /// <param name="templateRefOverride">Optional explicit template ref; if null, reads from case record.</param>
        /// <param name="startDateOverride">Optional start date override; if null, reads adc_originallodgementdate from case.</param>
        /// <param name="initiatingUserIdOverride">Optional initiating user; if null, reads createdby from case.</param>
        /// <returns>The created import job ID.</returns>
        public Guid RunImport(Guid caseId, EntityReference templateRefOverride = null,
            DateTime? startDateOverride = null, Guid? initiatingUserIdOverride = null)
        {
            _trace?.Trace("CaseImportService: RunImport caseId={0}", caseId);

            // 1. Resolve template
            EntityReference templateRef = templateRefOverride;
            var caseRecord = _service.Retrieve("adc_case", caseId,
                new ColumnSet("adc_adccasetemplateid", "adc_name", "adc_casenumber",
                    "createdby", "adc_originallodgementdate"));

            if (templateRef == null)
            {
                templateRef = caseRecord.GetAttributeValue<EntityReference>("adc_adccasetemplateid");
                _trace?.Trace("CaseImportService: Template from case = {0}",
                    templateRef != null ? templateRef.Id.ToString() : "NULL");
            }

            if (templateRef == null)
                throw new InvalidPluginExecutionException("No case template set on the case record.");

            // 2. Download MPP file
            byte[] mppBytes = DownloadFileColumn(templateRef.Id, "adc_adccasetemplate", "adc_templatefile");
            _trace?.Trace("CaseImportService: Downloaded MPP bytes = {0}",
                mppBytes != null ? mppBytes.Length.ToString() : "NULL");

            if (mppBytes == null || mppBytes.Length == 0)
                throw new InvalidPluginExecutionException("No MPP file found on the case template record.");

            // 3. Build project name from case
            string caseName = caseRecord.GetAttributeValue<string>("adc_name") ?? "ADC Case";
            string caseNumber = caseRecord.GetAttributeValue<string>("adc_casenumber");
            string projectName = !string.IsNullOrEmpty(caseNumber)
                ? string.Format("{0} - {1}", caseName, caseNumber)
                : caseName;

            // 4. Resolve start date
            DateTime? projectStartDate = startDateOverride
                ?? caseRecord.GetAttributeValue<DateTime?>("adc_originallodgementdate");
            _trace?.Trace("CaseImportService: Start date = {0}",
                projectStartDate.HasValue ? projectStartDate.Value.ToString("o") : "(not set)");

            // 5. Resolve initiating user
            Guid? initiatingUserId = initiatingUserIdOverride;
            if (!initiatingUserId.HasValue)
            {
                var createdBy = caseRecord.GetAttributeValue<EntityReference>("createdby");
                if (createdBy != null)
                    initiatingUserId = createdBy.Id;
            }

            // 6. Create project
            var projectEntity = new Entity("msdyn_project");
            projectEntity["msdyn_subject"] = projectName;
            projectEntity["adc_parentadccase"] = new EntityReference("adc_case", caseId);
            _trace?.Trace("CaseImportService: Creating project '{0}'...", projectName);
            Guid projectId = _service.Create(projectEntity);
            _trace?.Trace("CaseImportService: Project created: {0}", projectId);

            // 7. Link project to case + set processing status
            var caseUpdate = new Entity("adc_case", caseId);
            caseUpdate["adc_projectid"] = new EntityReference("msdyn_project", projectId);
            caseUpdate["adc_importstatus"] = new OptionSetValue(1); // Processing
            caseUpdate["adc_importmessage"] = "Creating project and starting import...";
            _service.Update(caseUpdate);

            // 8. Wait for project commit, then start import
            _trace?.Trace("CaseImportService: Sleeping 10s for project commit...");
            System.Threading.Thread.Sleep(10000);

            _trace?.Trace("CaseImportService: Calling InitializeJob...");
            var importService = new MppAsyncImportService(_service, _trace);
            Guid jobId = importService.InitializeJob(
                mppBytes, projectId, templateRef.Id, projectStartDate,
                caseId: caseId, initiatingUserId: initiatingUserId);
            _trace?.Trace("CaseImportService: Import job created: {0}", jobId);

            return jobId;
        }

        /// <summary>
        /// Overload that accepts a target Entity for plugin scenarios where
        /// attributes may be available on the target before DB commit.
        /// </summary>
        public Guid RunImportFromPlugin(Guid caseId, Entity target, Guid? fallbackUserId = null)
        {
            // Try to get template from target entity first (plugin pre-image)
            EntityReference templateRef = target?.GetAttributeValue<EntityReference>("adc_adccasetemplateid");

            // Try to get start date from target entity first
            DateTime? startDate = target?.GetAttributeValue<DateTime?>("adc_originallodgementdate");

            return RunImport(caseId, templateRef, startDate, fallbackUserId);
        }

        /// <summary>
        /// Marks the case as failed with the given error message.
        /// </summary>
        public void MarkCaseFailed(Guid caseId, string errorMessage)
        {
            try
            {
                var failUpdate = new Entity("adc_case", caseId);
                failUpdate["adc_importstatus"] = new OptionSetValue(4); // Failed
                var errMsg = "Import setup failed: " + errorMessage;
                failUpdate["adc_importmessage"] = errMsg.Length > 100 ? errMsg.Substring(0, 97) + "..." : errMsg;
                _service.Update(failUpdate);
            }
            catch (Exception updateEx)
            {
                _trace?.Trace("Could not update case to failed state: {0}", updateEx.Message);
            }
        }

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
                return null;
            }
        }
    }
}
