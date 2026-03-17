using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using ADC.MppImport.Services;

namespace ADC.MppImport.Workflows
{
    /// <summary>
    /// Custom workflow activity that triggers the same MPP import process as CaseCreatePlugin.
    /// Useful for re-running imports or triggering from flows/workflows outside of case creation.
    /// </summary>
    public class ImportCaseActivity : BaseCodeActivity
    {
        [Input("Case")]
        [ReferenceTarget("adc_case")]
        [RequiredArgument]
        public InArgument<EntityReference> Case { get; set; }

        [Input("Case Template")]
        [ReferenceTarget("adc_adccasetemplate")]
        public InArgument<EntityReference> CaseTemplate { get; set; }

        [Input("Starts On")]
        public InArgument<DateTime> StartsOn { get; set; }

        [Output("Import Job")]
        [ReferenceTarget("adc_mppimportjob")]
        public OutArgument<EntityReference> ImportJob { get; set; }

        [Output("Success")]
        public OutArgument<bool> Success { get; set; }

        [Output("Result Message")]
        public OutArgument<string> ResultMessage { get; set; }

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            var caseRef = Case.Get(executionContext);
            var templateRef = CaseTemplate.Get(executionContext);
            var startsOn = StartsOn.Get(executionContext);

            if (caseRef == null)
                throw new InvalidPluginExecutionException("Case input is required.");

            TracingService.Trace("ImportCaseActivity: Case={0}, Template={1}, StartsOn={2}",
                caseRef.Id,
                templateRef != null ? templateRef.Id.ToString() : "(from case)",
                startsOn != default(DateTime) ? startsOn.ToString("yyyy-MM-dd") : "(from case)");

            DateTime? startDateOverride = (startsOn != default(DateTime)) ? (DateTime?)startsOn : null;

            // Resolve initiating user
            Guid? initiatingUserId = null;
            var context = executionContext.GetExtension<IWorkflowContext>();
            if (context != null)
                initiatingUserId = context.InitiatingUserId;

            var caseImportService = new CaseImportService(OrganizationService, TracingService);

            try
            {
                Guid jobId = caseImportService.RunImport(
                    caseRef.Id, templateRef, startDateOverride, initiatingUserId);

                ImportJob.Set(executionContext, new EntityReference(ImportJobFields.EntityName, jobId));
                Success.Set(executionContext, true);
                ResultMessage.Set(executionContext, "Import job created successfully.");
                TracingService.Trace("ImportCaseActivity: Job {0} created.", jobId);
            }
            catch (Exception ex)
            {
                TracingService.Trace("ImportCaseActivity: EXCEPTION: {0}", ex.Message);
                caseImportService.MarkCaseFailed(caseRef.Id, ex.Message);
                Success.Set(executionContext, false);
                ResultMessage.Set(executionContext, ex.Message);
            }
        }
    }
}
