using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using ADC.MppImport.Services;

namespace ADC.MppImport.Plugins
{
    /// <summary>
    /// Async post-operation plugin on adc_case Create.
    /// 
    /// When a user creates an ADC Case and selects a Case Template:
    /// 1. Validates the template has an MPP file attached
    /// 2. Creates a new msdyn_project for this case
    /// 3. Links the project back to the case (adc_project lookup)
    /// 4. Downloads the MPP file and calls InitializeJob
    /// 5. The MppImportJobPlugin then handles all async import phases
    ///
    /// Registration:
    ///   Entity: adc_case
    ///   Message: Create
    ///   Stage: PostOperation (40)
    ///   Mode: Async (1)
    ///   FilteringAttributes: (none — fires on all creates)
    ///   ImpersonatingUserId: licensed Project Operations user
    /// </summary>
    public class CaseCreatePlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("CaseCreatePlugin: message={0}, depth={1}", context.MessageName, context.Depth);

            if (context.Depth > 5)
            {
                tracingService.Trace("Depth {0} exceeds limit, skipping.", context.Depth);
                return;
            }

            Entity target = null;
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                target = (Entity)context.InputParameters["Target"];

            if (target == null || target.LogicalName != "adc_case")
            {
                tracingService.Trace("Target is null or wrong entity.");
                return;
            }

            Guid caseId = target.Id;
            tracingService.Trace("Processing case: {0}", caseId);

            try
            {
                // 1. Get the case template reference
                var templateRef = target.GetAttributeValue<EntityReference>("adc_casetemplate");
                if (templateRef == null)
                {
                    tracingService.Trace("No case template selected, nothing to import.");
                    return;
                }

                tracingService.Trace("Case template: {0}", templateRef.Id);

                // 2. Validate the template has an MPP file
                byte[] mppBytes = DownloadFileColumn(service, tracingService, templateRef.Id,
                    "adc_adccasetemplate", "adc_templatefile");

                if (mppBytes == null || mppBytes.Length == 0)
                {
                    tracingService.Trace("No MPP file on template, nothing to import.");
                    return;
                }

                tracingService.Trace("MPP file: {0} bytes", mppBytes.Length);

                // 3. Get the case name for the project
                var caseRecord = service.Retrieve("adc_case", caseId,
                    new ColumnSet("adc_name", "createdby"));
                string caseName = caseRecord.GetAttributeValue<string>("adc_name") ?? "ADC Case";

                // Resolve initiating user (the person who created the case)
                Guid? initiatingUserId = null;
                var createdBy = caseRecord.GetAttributeValue<EntityReference>("createdby");
                if (createdBy != null)
                    initiatingUserId = createdBy.Id;
                if (!initiatingUserId.HasValue)
                    initiatingUserId = context.InitiatingUserId;

                // 4. Create the msdyn_project
                tracingService.Trace("Creating msdyn_project...");
                var projectEntity = new Entity("msdyn_project");
                projectEntity["msdyn_subject"] = caseName;
                Guid projectId = service.Create(projectEntity);
                tracingService.Trace("Project created: {0}", projectId);

                // 5. Link the project back to the case
                var caseUpdate = new Entity("adc_case", caseId);
                caseUpdate["adc_project"] = new EntityReference("msdyn_project", projectId);
                caseUpdate["adc_importstatus"] = new OptionSetValue(1); // Processing
                caseUpdate["adc_importmessage"] = "Creating project and starting import...";
                service.Update(caseUpdate);

                // 6. Wait briefly for PSS to initialize the project (root task creation)
                System.Threading.Thread.Sleep(10000);

                // 7. Initialize the async import job
                tracingService.Trace("Calling InitializeJob...");
                var importService = new MppAsyncImportService(service, tracingService);
                Guid jobId = importService.InitializeJob(
                    mppBytes, projectId, templateRef.Id, null,
                    caseId: caseId, initiatingUserId: initiatingUserId);

                tracingService.Trace("Import job created: {0}. MppImportJobPlugin will handle phases.", jobId);
            }
            catch (Exception ex)
            {
                tracingService.Trace("CaseCreatePlugin ERROR: {0}", ex.ToString());

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
                    tracingService.Trace("Could not update case to failed state: {0}", updateEx.Message);
                }

                // Don't rethrow — let the async system job succeed so it doesn't retry
            }
        }

        /// <summary>
        /// Downloads a file column's content as byte[].
        /// </summary>
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
    }
}
