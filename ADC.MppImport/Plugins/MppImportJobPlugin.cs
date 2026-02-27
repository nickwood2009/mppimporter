using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using ADC.MppImport.Services;

namespace ADC.MppImport.Plugins
{
    /// <summary>
    /// Async plugin registered on:
    ///   1. adc_mppimportjob Create/Update (filtering: adc_status, adc_tick)
    ///      — processes import phases
    ///   2. adc_case Create
    ///      — auto-creates project and kicks off MPP import from linked template
    ///
    /// Each execution routes based on entity logical name.
    /// </summary>
    public class MppImportJobPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            // The plugin step must have impersonatinguserid set to a licensed Project Operations
            // user so PSS APIs succeed regardless of the triggering user.
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            Entity target = null;
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                target = (Entity)context.InputParameters["Target"];

            if (target == null)
            {
                tracingService.Trace("MppImportJobPlugin: Target is null, skipping.");
                return;
            }

            // Route based on entity
            if (target.LogicalName == "adc_case")
            {
                HandleCaseCreate(service, tracingService, context, target);
            }
            else if (target.LogicalName == ImportJobFields.EntityName)
            {
                HandleImportJob(service, tracingService, context, target);
            }
            else
            {
                tracingService.Trace("MppImportJobPlugin: Unexpected entity {0}, skipping.", target.LogicalName);
            }
        }

        #region Import Job Handler (adc_mppimportjob)

        private void HandleImportJob(IOrganizationService service, ITracingService trace,
            IPluginExecutionContext context, Entity target)
        {
            // Prevent infinite recursion. Each batch cycle adds ~2 depth levels.
            // Depth 50 supports ~24 batches = ~2400 tasks. Terminal state checks,
            // Thread.Sleep caps, and FailJob timeouts provide additional safety.
            if (context.Depth > 50)
            {
                trace.Trace("Depth {0} exceeds limit, skipping to prevent recursion.", context.Depth);
                return;
            }

            Guid jobId = target.Id;
            trace.Trace("HandleImportJob: job={0}, message={1}, depth={2}",
                jobId, context.MessageName, context.Depth);

            try
            {
                var importService = new MppAsyncImportService(service, trace);
                importService.ProcessJob(jobId);
            }
            catch (Exception ex)
            {
                trace.Trace("HandleImportJob UNHANDLED ERROR: {0}", ex.ToString());

                // Try to fail the job gracefully
                try
                {
                    var update = new Entity(ImportJobFields.EntityName, jobId);
                    update[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.Failed);
                    update[ImportJobFields.ErrorMessage] = "Unhandled plugin error: " + ex.Message;
                    service.Update(update);
                }
                catch (Exception failEx)
                {
                    trace.Trace("Could not update job to Failed state: {0}", failEx.Message);
                }

                // Don't rethrow — let the async system job succeed so it doesn't retry
            }
        }

        #endregion

        #region Case Create Handler (adc_case)

        /// <summary>
        /// When a user creates an adc_case and selects a Case Template:
        /// 1. Validates the template has an MPP file attached
        /// 2. Creates a new msdyn_project for this case
        /// 3. Links the project back to the case
        /// 4. Downloads the MPP and calls InitializeJob
        /// 5. MppImportJobPlugin then handles all async import phases
        /// </summary>
        private void HandleCaseCreate(IOrganizationService service, ITracingService trace,
            IPluginExecutionContext context, Entity target)
        {
            if (context.Depth > 5)
            {
                trace.Trace("HandleCaseCreate: Depth {0} exceeds limit, skipping.", context.Depth);
                return;
            }

            Guid caseId = target.Id;
            trace.Trace("HandleCaseCreate: case={0}, depth={1}", caseId, context.Depth);

            try
            {
                // 1. Get the case template reference
                var templateRef = target.GetAttributeValue<EntityReference>("adc_casetemplate");
                if (templateRef == null)
                {
                    trace.Trace("No case template selected, nothing to import.");
                    return;
                }

                trace.Trace("Case template: {0}", templateRef.Id);

                // Immediately set status so the user knows the import is starting
                var earlyUpdate = new Entity("adc_case", caseId);
                earlyUpdate["adc_importstatus"] = new OptionSetValue(1); // Processing
                earlyUpdate["adc_importmessage"] = "Import starting — validating template and downloading MPP file...";
                service.Update(earlyUpdate);

                // 2. Validate the template has an MPP file
                byte[] mppBytes = DownloadFileColumn(service, trace, templateRef.Id,
                    "adc_adccasetemplate", "adc_templatefile");

                if (mppBytes == null || mppBytes.Length == 0)
                {
                    trace.Trace("No MPP file on template, nothing to import.");
                    return;
                }

                trace.Trace("MPP file: {0} bytes", mppBytes.Length);

                // 3. Get the case name and lodgement date for the project
                var caseRecord = service.Retrieve("adc_case", caseId,
                    new ColumnSet("adc_name", "createdby", "adc_originallodgementdate"));
                string caseName = caseRecord.GetAttributeValue<string>("adc_name") ?? "ADC Case";

                // Read the original lodgement date — used as project start date
                DateTime? projectStartDate = caseRecord.GetAttributeValue<DateTime?>("adc_originallodgementdate");
                if (projectStartDate.HasValue)
                    trace.Trace("Original lodgement date: {0:yyyy-MM-dd}", projectStartDate.Value);
                else
                    trace.Trace("No lodgement date set, import will use MPP default dates.");

                // Resolve initiating user (the person who created the case)
                Guid? initiatingUserId = null;
                var createdBy = caseRecord.GetAttributeValue<EntityReference>("createdby");
                if (createdBy != null)
                    initiatingUserId = createdBy.Id;
                if (!initiatingUserId.HasValue)
                    initiatingUserId = context.InitiatingUserId;

                // 4. Create the msdyn_project
                trace.Trace("Creating msdyn_project...");
                var projectEntity = new Entity("msdyn_project");
                projectEntity["msdyn_subject"] = caseName;
                Guid projectId = service.Create(projectEntity);
                trace.Trace("Project created: {0}", projectId);

                // 5. Link the project back to the case and set Processing status
                var caseUpdate = new Entity("adc_case", caseId);
                caseUpdate["adc_project"] = new EntityReference("msdyn_project", projectId);
                caseUpdate["adc_importstatus"] = new OptionSetValue(1); // Processing
                caseUpdate["adc_importmessage"] = "Creating project and starting import...";
                service.Update(caseUpdate);

                // 6. Wait for PSS to initialize the project (root task creation)
                System.Threading.Thread.Sleep(10000);

                // 7. Initialize the async import job
                trace.Trace("Calling InitializeJob...");
                var importService = new MppAsyncImportService(service, trace);
                Guid jobId = importService.InitializeJob(
                    mppBytes, projectId, templateRef.Id, projectStartDate,
                    caseId: caseId, initiatingUserId: initiatingUserId);

                trace.Trace("Import job created: {0}. Async phases will follow.", jobId);
            }
            catch (Exception ex)
            {
                trace.Trace("HandleCaseCreate ERROR: {0}", ex.ToString());

                // Update case with failure status
                try
                {
                    var failUpdate = new Entity("adc_case", caseId);
                    failUpdate["adc_importstatus"] = new OptionSetValue(4); // Failed
                    failUpdate["adc_importmessage"] = "Import setup failed: " +
                        (ex.Message.Length > 200 ? ex.Message.Substring(0, 200) + "..." : ex.Message);
                    service.Update(failUpdate);
                }
                catch (Exception updateEx)
                {
                    trace.Trace("Could not update case to failed state: {0}", updateEx.Message);
                }

                // Don't rethrow — let the async system job succeed so it doesn't retry
            }
        }

        #endregion

        #region File Download Helper

        private byte[] DownloadFileColumn(IOrganizationService service, ITracingService trace,
            Guid recordId, string entityName, string fileAttributeName)
        {
            try
            {
                var initRequest = new OrganizationRequest("InitializeFileBlocksDownload");
                initRequest["Target"] = new EntityReference(entityName, recordId);
                initRequest["FileAttributeName"] = fileAttributeName;

                var initResponse = service.Execute(initRequest);
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

                    var downloadResponse = service.Execute(downloadRequest);
                    byte[] blockData = (byte[])downloadResponse["Data"];
                    allBytes.AddRange(blockData);
                    offset += blockData.Length;
                }

                return allBytes.ToArray();
            }
            catch (Exception ex)
            {
                trace.Trace("Error downloading file: {0}", ex.Message);
                return null;
            }
        }

        #endregion
    }
}
