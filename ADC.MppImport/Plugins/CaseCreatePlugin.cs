using System;
using Microsoft.Xrm.Sdk;
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
            var caseImportService = new CaseImportService(service, tracingService);

            try
            {
                // Check if template is set before delegating
                var templateRef = target.GetAttributeValue<EntityReference>("adc_adccasetemplateid");
                if (templateRef == null)
                {
                    tracingService.Trace("CaseCreatePlugin: No template in Target, will check DB via service...");
                }

                Guid jobId = caseImportService.RunImportFromPlugin(caseId, target, context.InitiatingUserId);
                tracingService.Trace("CaseCreatePlugin: Import job created: {0}", jobId);
            }
            catch (InvalidPluginExecutionException ex) when (ex.Message.Contains("No case template"))
            {
                tracingService.Trace("CaseCreatePlugin: No template set, exiting gracefully.");
                return;
            }
            catch (Exception ex)
            {
                tracingService.Trace("CaseCreatePlugin: EXCEPTION: {0}\n{1}", ex.Message, ex.StackTrace);
                caseImportService.MarkCaseFailed(caseId, ex.Message);
            }
        }
    }
}
