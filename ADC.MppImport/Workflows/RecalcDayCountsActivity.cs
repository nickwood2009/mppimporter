using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using ADC.MppImport.Services;

namespace ADC.MppImport.Workflows
{
    /// <summary>
    /// Custom workflow activity that recalculates day counts for ALL tasks
    /// in a project. Use after import completes, or as a manual "Recalculate" action.
    ///
    /// Input: Project (msdyn_project EntityReference)
    /// Outputs: Success (bool), ResultMessage (string)
    /// </summary>
    public class RecalcDayCountsActivity : BaseCodeActivity
    {
        [Input("Project")]
        [ReferenceTarget("msdyn_project")]
        [RequiredArgument]
        public InArgument<EntityReference> Project { get; set; }

        [Output("Success")]
        public OutArgument<bool> Success { get; set; }

        [Output("Result Message")]
        public OutArgument<string> ResultMessage { get; set; }

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            var projectRef = Project.Get(executionContext);

            if (projectRef == null)
                throw new InvalidPluginExecutionException("Project input is required.");

            TracingService.Trace("RecalcDayCountsActivity: Project={0}", projectRef.Id);

            try
            {
                var dayCountService = new DayCountService(OrganizationService, TracingService);
                dayCountService.RecalcAllTasks(projectRef.Id);

                Success.Set(executionContext, true);
                ResultMessage.Set(executionContext, "Day counts recalculated successfully.");
                TracingService.Trace("RecalcDayCountsActivity: Completed successfully.");
            }
            catch (Exception ex)
            {
                TracingService.Trace("RecalcDayCountsActivity: EXCEPTION: {0}", ex.Message);
                Success.Set(executionContext, false);
                ResultMessage.Set(executionContext, ex.Message);
            }
        }
    }
}
