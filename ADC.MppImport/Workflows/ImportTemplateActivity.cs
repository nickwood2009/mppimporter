using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using ADC.MppImport.Services;

namespace ADC.MppImport.Workflows
{
    /// <summary>
    /// Custom workflow activity that imports an MPP file from an ADC Case Template
    /// into a new template project (adc_istemplate = true).
    /// This is Phase A of the template-then-clone architecture:
    /// the template project is created once, then cloned per case via CloneProjectActivity.
    /// </summary>
    public class ImportTemplateActivity : BaseCodeActivity
    {
        [Input("Case Template")]
        [ReferenceTarget("adc_adccasetemplate")]
        [RequiredArgument]
        public InArgument<EntityReference> CaseTemplate { get; set; }

        [Output("Import Job")]
        [ReferenceTarget("adc_mppimportjob")]
        public OutArgument<EntityReference> ImportJob { get; set; }

        [Output("Template Project")]
        [ReferenceTarget("msdyn_project")]
        public OutArgument<EntityReference> TemplateProject { get; set; }

        [Output("Success")]
        public OutArgument<bool> Success { get; set; }

        [Output("Result Message")]
        public OutArgument<string> ResultMessage { get; set; }

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            var templateRef = CaseTemplate.Get(executionContext);

            if (templateRef == null)
                throw new InvalidPluginExecutionException("Case Template input is required.");

            TracingService.Trace("ImportTemplateActivity: Template={0}", templateRef.Id);

            // Resolve initiating user
            Guid? initiatingUserId = null;
            var context = executionContext.GetExtension<IWorkflowContext>();
            if (context != null)
                initiatingUserId = context.InitiatingUserId;

            var caseImportService = new CaseImportService(OrganizationService, TracingService);

            try
            {
                Guid jobId = caseImportService.RunTemplateImport(templateRef.Id, initiatingUserId);

                ImportJob.Set(executionContext, new EntityReference(ImportJobFields.EntityName, jobId));
                Success.Set(executionContext, true);
                ResultMessage.Set(executionContext, "Template import job created successfully.");
                TracingService.Trace("ImportTemplateActivity: Job {0} created.", jobId);

                // Read back the project linked to the template
                try
                {
                    var templateRecord = OrganizationService.Retrieve(
                        "adc_adccasetemplate", templateRef.Id,
                        new Microsoft.Xrm.Sdk.Query.ColumnSet("adc_templateproject"));
                    var projectRef = templateRecord.GetAttributeValue<EntityReference>("adc_templateproject");
                    if (projectRef != null)
                        TemplateProject.Set(executionContext, projectRef);
                }
                catch (Exception ex)
                {
                    TracingService.Trace("ImportTemplateActivity: Could not read back template project (non-fatal): {0}", ex.Message);
                }
            }
            catch (Exception ex)
            {
                TracingService.Trace("ImportTemplateActivity: EXCEPTION: {0}", ex.Message);
                caseImportService.MarkTemplateFailed(templateRef.Id, ex.Message);
                Success.Set(executionContext, false);
                ResultMessage.Set(executionContext, ex.Message);
            }
        }
    }
}
