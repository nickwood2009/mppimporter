using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using ADC.MppImport.Services;

namespace ADC.MppImport.Workflows
{
    public class StartMppImportActivity : BaseCodeActivity
    {
        [Input("Case Template")]
        [ReferenceTarget("adc_adccasetemplate")]
        [RequiredArgument]
        public InArgument<EntityReference> CaseTemplate { get; set; }

        [Input("Target Project")]
        [ReferenceTarget("msdyn_project")]
        [RequiredArgument]
        public InArgument<EntityReference> TargetProject { get; set; }

        [Input("Case")]
        [ReferenceTarget("adc_case")]
        public InArgument<EntityReference> Case { get; set; }

        [Input("Starts On")]
        public InArgument<DateTime> StartsOn { get; set; }

        [Output("Import Job")]
        [ReferenceTarget("adc_mppimportjob")]
        public OutArgument<EntityReference> ImportJob { get; set; }

        [Output("Tasks Found")]
        public OutArgument<int> TasksFound { get; set; }

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            var templateRef = CaseTemplate.Get(executionContext);
            var projectRef = TargetProject.Get(executionContext);
            var caseRef = Case.Get(executionContext);
            var startsOn = StartsOn.Get(executionContext);

            if (templateRef == null)
                throw new InvalidPluginExecutionException("Case Template input is required.");
            if (projectRef == null)
                throw new InvalidPluginExecutionException("Target Project input is required.");

            TracingService.Trace("StartMppImport: Template={0}, Project={1}", templateRef.Id, projectRef.Id);

            DateTime? projectStartDate = (startsOn != default(DateTime)) ? (DateTime?)startsOn : null;

            Guid? caseId = caseRef?.Id;
            Guid? initiatingUserId = null;
            if (caseRef != null)
            {
                try
                {
                    var caseRecord = OrganizationService.Retrieve("adc_case", caseRef.Id,
                        new Microsoft.Xrm.Sdk.Query.ColumnSet("createdby"));
                    var createdBy = caseRecord.GetAttributeValue<EntityReference>("createdby");
                    if (createdBy != null)
                        initiatingUserId = createdBy.Id;
                }
                catch (Exception ex)
                {
                    TracingService.Trace("Could not resolve initiating user from case: {0}", ex.Message);
                }
            }
            if (!initiatingUserId.HasValue)
            {
                var context = executionContext.GetExtension<IWorkflowContext>();
                if (context != null)
                    initiatingUserId = context.InitiatingUserId;
            }

            byte[] mppBytes = DownloadFileColumn(templateRef.Id, "adc_adccasetemplate", "adc_templatefile");
            if (mppBytes == null || mppBytes.Length == 0)
                throw new InvalidPluginExecutionException("No MPP file found on the case template record.");

            var importService = new MppAsyncImportService(OrganizationService, TracingService);
            Guid jobId = importService.InitializeJob(mppBytes, projectRef.Id, templateRef.Id, projectStartDate, caseId, initiatingUserId);

            ImportJob.Set(executionContext, new EntityReference(ImportJobFields.EntityName, jobId));

            TracingService.Trace("StartMppImport: Job {0} created, import will proceed asynchronously.", jobId);
        }

        private byte[] DownloadFileColumn(Guid recordId, string entityName, string fileAttributeName)
        {
            try
            {
                var initRequest = new OrganizationRequest("InitializeFileBlocksDownload");
                initRequest["Target"] = new EntityReference(entityName, recordId);
                initRequest["FileAttributeName"] = fileAttributeName;

                var initResponse = OrganizationService.Execute(initRequest);
                string fileContinuationToken = (string)initResponse["FileContinuationToken"];
                long fileSize = (long)initResponse["FileSizeInBytes"];

                if (fileSize == 0) return null;

                var allBytes = new System.Collections.Generic.List<byte>();
                long offset = 0;
                const long blockSize = 4 * 1024 * 1024; // 4MB blocks

                while (offset < fileSize)
                {
                    var downloadRequest = new OrganizationRequest("DownloadBlock");
                    downloadRequest["FileContinuationToken"] = fileContinuationToken;
                    downloadRequest["BlockLength"] = blockSize;
                    downloadRequest["Offset"] = offset;

                    var downloadResponse = OrganizationService.Execute(downloadRequest);
                    byte[] blockData = (byte[])downloadResponse["Data"];
                    allBytes.AddRange(blockData);
                    offset += blockData.Length;
                }

                return allBytes.ToArray();
            }
            catch (Exception ex)
            {
                TracingService.Trace("Error downloading file: {0}", ex.Message);
                throw new InvalidPluginExecutionException(
                    string.Format("Failed to download MPP file from {0}.{1}: {2}", entityName, fileAttributeName, ex.Message), ex);
            }
        }
    }
}
