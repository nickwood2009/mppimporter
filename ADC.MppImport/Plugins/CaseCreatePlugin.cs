using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using ADC.MppImport.Services;

namespace ADC.MppImport.Plugins
{
    public class CaseCreatePlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("CaseCreatePlugin: Execute started. Message={0}, Entity={1}, Depth={2}, UserId={3}",
                context.MessageName, context.PrimaryEntityName, context.Depth, context.UserId);

            if (context.Depth > 5)
            {
                tracingService.Trace("CaseCreatePlugin: Depth {0} > 5, exiting.", context.Depth);
                return;
            }

            Entity target = null;
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                target = (Entity)context.InputParameters["Target"];

            if (target == null)
            {
                tracingService.Trace("CaseCreatePlugin: Target is null, exiting.");
                return;
            }
            tracingService.Trace("CaseCreatePlugin: Target entity={0}, id={1}", target.LogicalName, target.Id);
            if (target.LogicalName != "adc_case")
            {
                tracingService.Trace("CaseCreatePlugin: Wrong entity '{0}', expected 'adc_case'. Exiting.", target.LogicalName);
                return;
            }

            Guid caseId = target.Id;

            try
            {
                var templateRef = target.GetAttributeValue<EntityReference>("adc_casetemplate");
                tracingService.Trace("CaseCreatePlugin: adc_casetemplate = {0}",
                    templateRef != null ? templateRef.Id.ToString() : "NULL");
                if (templateRef == null)
                {
                    tracingService.Trace("CaseCreatePlugin: No template set, exiting.");
                    return;
                }

                byte[] mppBytes = DownloadFileColumn(service, tracingService, templateRef.Id,
                    "adc_adccasetemplate", "adc_templatefile");

                tracingService.Trace("CaseCreatePlugin: Downloaded MPP bytes = {0}",
                    mppBytes != null ? mppBytes.Length.ToString() : "NULL");
                if (mppBytes == null || mppBytes.Length == 0)
                {
                    tracingService.Trace("CaseCreatePlugin: No MPP file data, exiting.");
                    return;
                }

                var caseRecord = service.Retrieve("adc_case", caseId,
                    new ColumnSet("adc_name", "createdby"));
                string caseName = caseRecord.GetAttributeValue<string>("adc_name") ?? "ADC Case";

                Guid? initiatingUserId = null;
                var createdBy = caseRecord.GetAttributeValue<EntityReference>("createdby");
                if (createdBy != null)
                    initiatingUserId = createdBy.Id;
                if (!initiatingUserId.HasValue)
                    initiatingUserId = context.InitiatingUserId;

                var projectEntity = new Entity("msdyn_project");
                projectEntity["msdyn_subject"] = caseName;
                tracingService.Trace("CaseCreatePlugin: Creating project '{0}'...", caseName);
                Guid projectId = service.Create(projectEntity);
                tracingService.Trace("CaseCreatePlugin: Project created: {0}", projectId);

                var caseUpdate = new Entity("adc_case", caseId);
                caseUpdate["adc_project"] = new EntityReference("msdyn_project", projectId);
                caseUpdate["adc_importstatus"] = new OptionSetValue(1); // Processing
                caseUpdate["adc_importmessage"] = "Creating project and starting import...";
                service.Update(caseUpdate);

                tracingService.Trace("CaseCreatePlugin: Sleeping 10s for project commit...");
                System.Threading.Thread.Sleep(10000);
                tracingService.Trace("CaseCreatePlugin: Calling InitializeJob...");
                var importService = new MppAsyncImportService(service, tracingService);
                Guid jobId = importService.InitializeJob(
                    mppBytes, projectId, templateRef.Id, null,
                    caseId: caseId, initiatingUserId: initiatingUserId);
                tracingService.Trace("CaseCreatePlugin: Import job created: {0}", jobId);

            }
            catch (Exception ex)
            {
                tracingService.Trace("CaseCreatePlugin: EXCEPTION: {0}\n{1}", ex.Message, ex.StackTrace);
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
            }
        }

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
