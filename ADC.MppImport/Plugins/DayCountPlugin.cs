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
    ///   Filtering attributes: adc_initiationdate (PLACEHOLDER — update to actual schema name)
    ///   When the initiation date is set/changed on a case, update the "Initiation date"
    ///   milestone's finish date and recalculate all tasks in the linked project.
    /// </summary>
    public class DayCountPlugin : IPlugin
    {
        // PLACEHOLDER — update to match your environment
        private const string CASE_INITIATION_DATE = "adc_initiationdate";

        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            Entity target = null;
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                target = (Entity)context.InputParameters["Target"];

            if (target == null)
            {
                tracingService.Trace("DayCountPlugin: Target is null, skipping.");
                return;
            }

            // Guard against deep recursion — updating day count on tasks will fire
            // this plugin again, but the second pass won't have msdyn_scheduledend
            // in the target so it will exit early. This guard is a safety net.
            if (context.Depth > 10)
            {
                tracingService.Trace("DayCountPlugin: Depth {0} exceeds limit, skipping.", context.Depth);
                return;
            }

            try
            {
                if (target.LogicalName == "msdyn_projecttask")
                {
                    HandleTaskUpdate(service, tracingService, context, target);
                }
                else if (target.LogicalName == "adc_case")
                {
                    HandleCaseUpdate(service, tracingService, context, target);
                }
                else
                {
                    tracingService.Trace("DayCountPlugin: Unexpected entity '{0}', skipping.", target.LogicalName);
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("DayCountPlugin: EXCEPTION: {0}\n{1}", ex.Message, ex.StackTrace);
                // Don't rethrow — day count calc failure should not block the triggering operation
            }
        }

        /// <summary>
        /// Step A: msdyn_projecttask Update — recalculate day count for the changed task.
        /// Only fires when msdyn_scheduledend is in the update target (filtering attribute).
        /// </summary>
        private void HandleTaskUpdate(IOrganizationService service, ITracingService trace,
            IPluginExecutionContext context, Entity target)
        {
            // Only proceed if the scheduled end date actually changed
            if (!target.Contains("msdyn_scheduledend"))
            {
                trace.Trace("DayCountPlugin: msdyn_scheduledend not in target, skipping.");
                return;
            }

            trace.Trace("DayCountPlugin: Task {0} — msdyn_scheduledend changed, recalculating day count.", target.Id);

            var dayCountService = new DayCountService(service, trace);
            dayCountService.RecalcSingleTask(target.Id);
        }

        /// <summary>
        /// Step B: adc_case Update — initiation date changed.
        /// Updates the "Initiation date" milestone and recalculates all tasks.
        /// </summary>
        private void HandleCaseUpdate(IOrganizationService service, ITracingService trace,
            IPluginExecutionContext context, Entity target)
        {
            if (!target.Contains(CASE_INITIATION_DATE))
            {
                trace.Trace("DayCountPlugin: {0} not in target, skipping.", CASE_INITIATION_DATE);
                return;
            }

            DateTime? initiationDate = target.GetAttributeValue<DateTime?>(CASE_INITIATION_DATE);
            trace.Trace("DayCountPlugin: Case {0} — initiation date changed to {1}.",
                target.Id, initiationDate.HasValue ? initiationDate.Value.ToString("yyyy-MM-dd") : "(cleared)");

            var dayCountService = new DayCountService(service, trace);
            dayCountService.OnInitiationDateChanged(target.Id, initiationDate);
        }
    }
}
