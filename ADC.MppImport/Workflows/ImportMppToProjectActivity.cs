using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using ADC.MppImport.Services;

namespace ADC.MppImport.Workflows
{
    /// <summary>
    /// CRM Workflow Activity: Imports tasks from an MPP file (on adc_adccasetemplate)
    /// into msdyn_projecttask records under a target msdyn_project.
    /// This is a thin shell — all business logic lives in MppProjectImportService.
    /// </summary>
    public class ImportMppToProjectActivity : BaseCodeActivity
    {
        [Input("Case Template")]
        [ReferenceTarget("adc_adccasetemplate")]
        [RequiredArgument]
        public InArgument<EntityReference> CaseTemplate { get; set; }

        [Input("Target Project")]
        [ReferenceTarget("msdyn_project")]
        [RequiredArgument]
        public InArgument<EntityReference> TargetProject { get; set; }

        [Output("Tasks Created")]
        public OutArgument<int> TasksCreated { get; set; }

        [Output("Tasks Updated")]
        public OutArgument<int> TasksUpdated { get; set; }

        [Output("Total Processed")]
        public OutArgument<int> TotalProcessed { get; set; }

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            var templateRef = CaseTemplate.Get(executionContext);
            var projectRef = TargetProject.Get(executionContext);

            if (templateRef == null)
                throw new InvalidPluginExecutionException("Case Template input is required.");
            if (projectRef == null)
                throw new InvalidPluginExecutionException("Target Project input is required.");

            TracingService.Trace("ImportMppToProject: Template={0}, Project={1}",
                templateRef.Id, projectRef.Id);

            var importService = new MppProjectImportService(OrganizationService, TracingService);
            ImportResult result = importService.Execute(templateRef.Id, projectRef.Id);

            TasksCreated.Set(executionContext, result.TasksCreated);
            TasksUpdated.Set(executionContext, result.TasksUpdated);
            TotalProcessed.Set(executionContext, result.TotalProcessed);

            TracingService.Trace("ImportMppToProject complete: Created={0}, Updated={1}",
                result.TasksCreated, result.TasksUpdated);
        }
    }
}
