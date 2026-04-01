using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using ADC.MppImport.Services;

namespace ADC.MppImport.Workflows
{
    /// <summary>
    /// Custom workflow activity that recalculates day counts for ALL tasks in a project.
    /// Use after import completes, or as a manual on-demand "Recalculate" action.
    ///
    /// Provide EITHER:
    ///   - Project (msdyn_project) — recalcs tasks directly, OR
    ///   - Case (adc_case) — resolves the linked project first, then recalcs.
    /// If both are provided, Project takes priority.
    ///
    /// To run on-demand:
    ///   1. Create an on-demand workflow on adc_case (or msdyn_project).
    ///   2. Add this step and map the primary entity to the Case (or Project) input.
    /// </summary>
    public class RecalcDayCountsActivity : BaseCodeActivity
    {
        [Input("Project")]
        [ReferenceTarget("msdyn_project")]
        public InArgument<EntityReference> Project { get; set; }

        [Input("Case")]
        [ReferenceTarget("adc_case")]
        public InArgument<EntityReference> Case { get; set; }

        [Output("Success")]
        public OutArgument<bool> Success { get; set; }

        [Output("Result Message")]
        public OutArgument<string> ResultMessage { get; set; }

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            var projectRef = Project.Get(executionContext);
            var caseRef = Case.Get(executionContext);

            Guid projectId;

            if (projectRef != null)
            {
                projectId = projectRef.Id;
                TracingService.Trace("RecalcDayCountsActivity: Project input provided: {0}", projectId);
            }
            else if (caseRef != null)
            {
                TracingService.Trace("RecalcDayCountsActivity: Case input provided: {0} — resolving linked project...", caseRef.Id);
                projectId = ResolveProjectFromCase(caseRef.Id);
                if (projectId == Guid.Empty)
                {
                    string msg = string.Format("No project linked to case {0} (via adc_projectid lookup).", caseRef.Id);
                    TracingService.Trace("RecalcDayCountsActivity: {0}", msg);
                    Success.Set(executionContext, false);
                    ResultMessage.Set(executionContext, msg);
                    return;
                }
                TracingService.Trace("RecalcDayCountsActivity: Resolved project {0} from case.", projectId);
            }
            else
            {
                throw new InvalidPluginExecutionException("Either Project or Case input is required.");
            }

            try
            {
                var dayCountService = new DayCountService(OrganizationService, TracingService);
                dayCountService.RecalcAllTasks(projectId);

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

        /// <summary>
        /// Looks up the project linked to a case via the adc_projectid lookup on adc_case.
        /// </summary>
        private Guid ResolveProjectFromCase(Guid caseId)
        {
            try
            {
                var caseRecord = OrganizationService.Retrieve("adc_case", caseId,
                    new ColumnSet("adc_projectid"));
                var projectRef = caseRecord.GetAttributeValue<EntityReference>("adc_projectid");
                if (projectRef != null)
                {
                    TracingService.Trace("RecalcDayCountsActivity: Case {0} has adc_projectid = {1}", caseId, projectRef.Id);
                    return projectRef.Id;
                }
            }
            catch (Exception ex)
            {
                TracingService.Trace("RecalcDayCountsActivity: Error resolving project from case: {0}", ex.Message);
            }
            return Guid.Empty;
        }
    }
}
