using System;
using Microsoft.Xrm.Sdk;
using ADC.MppImport.Services;

namespace ADC.MppImport.Plugins
{
    /// <summary>
    /// Plugin that recalculates "Day Count" on project tasks.
    ///
    /// Register TWO steps:
    ///
    /// Step A — msdyn_projecttask / Update / Post-Operation / Async
    ///   Filtering attributes: msdyn_scheduledend
    ///   When a task's scheduled finish date changes, recalculate its day count.
    ///
    /// Step B — adc_case / Update / Post-Operation / Async
    ///   Filtering attributes: adc_dateofinitiation, adc_originallodgementdate
    ///   When either reference date changes on the case, recalculate all tasks
    ///   in the linked project.
    /// </summary>
    public class DayCountPlugin : IPlugin
    {
        private const string CASE_INITIATION_DATE = "adc_dateofinitiation";
        private const string CASE_ORIGINAL_LODGEMENT_DATE = "adc_originallodgementdate";

        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            // Use SYSTEM context — PSS service account may lack privileges on custom entities (adc_case)
            var service = serviceFactory.CreateOrganizationService(null);

            Entity target = null;
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                target = (Entity)context.InputParameters["Target"];

            if (target == null)
            {
                tracingService.Trace("DayCountPlugin: Target is null, skipping.");
                return;
            }

            if (context.Depth > 50)
            {
                tracingService.Trace("DayCountPlugin: Depth {0} exceeds limit, skipping.", context.Depth);
                return;
            }

            try
            {
                if (target.LogicalName == "msdyn_projecttask")
                {
                    HandleTaskUpdate(service, tracingService, target);
                }
                else if (target.LogicalName == "adc_case")
                {
                    HandleCaseUpdate(service, tracingService, target);
                }
                else
                {
                    tracingService.Trace("DayCountPlugin: Unexpected entity '{0}', skipping.", target.LogicalName);
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("DayCountPlugin: EXCEPTION: {0}\n{1}", ex.Message, ex.StackTrace);
            }
        }

        /// <summary>
        /// Step A: msdyn_projecttask Update — recalculate day count for the changed task.
        /// </summary>
        private void HandleTaskUpdate(IOrganizationService service, ITracingService trace, Entity target)
        {
            if (!target.Contains("msdyn_scheduledend"))
            {
                trace.Trace("DayCountPlugin: msdyn_scheduledend not in target, skipping.");
                return;
            }

            trace.Trace("DayCountPlugin: Task {0} — msdyn_scheduledend changed, recalculating.", target.Id);

            var svc = new DayCountService(service, trace);
            svc.RecalcSingleTask(target.Id);
        }

        /// <summary>
        /// Step B: adc_case Update — reference date changed (initiation or original lodgement).
        /// Recalculates all tasks in the linked project.
        /// </summary>
        private void HandleCaseUpdate(IOrganizationService service, ITracingService trace, Entity target)
        {
            bool initiationChanged = target.Contains(CASE_INITIATION_DATE);
            bool lodgementChanged = target.Contains(CASE_ORIGINAL_LODGEMENT_DATE);

            if (!initiationChanged && !lodgementChanged)
            {
                trace.Trace("DayCountPlugin: Neither {0} nor {1} in target, skipping.",
                    CASE_INITIATION_DATE, CASE_ORIGINAL_LODGEMENT_DATE);
                return;
            }

            trace.Trace("DayCountPlugin: Case {0} — date changed (initiation={1}, lodgement={2}). Recalculating all tasks.",
                target.Id, initiationChanged, lodgementChanged);

            var svc = new DayCountService(service, trace);
            svc.OnCaseDateChanged(target.Id);
        }
    }
}
