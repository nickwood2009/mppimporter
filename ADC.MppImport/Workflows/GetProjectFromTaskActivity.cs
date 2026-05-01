using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace ADC.MppImport.Workflows
{
    /// <summary>
    /// Takes a project task and returns the parent project entity reference.
    /// </summary>
    public class GetProjectFromTaskActivity : BaseCodeActivity
    {
        [Input("Project Task")]
        [ReferenceTarget("msdyn_projecttask")]
        [RequiredArgument]
        public InArgument<EntityReference> ProjectTask { get; set; }

        [Output("Project")]
        [ReferenceTarget("msdyn_project")]
        public OutArgument<EntityReference> Project { get; set; }

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            var taskRef = ProjectTask.Get(executionContext);

            TracingService.Trace("GetProjectFromTask: Task={0}", taskRef.Id);

            var task = OrganizationService.Retrieve("msdyn_projecttask", taskRef.Id,
                new ColumnSet("msdyn_project"));

            var projectRef = task.GetAttributeValue<EntityReference>("msdyn_project");
            if (projectRef != null)
            {
                TracingService.Trace("GetProjectFromTask: Project={0}", projectRef.Id);
                Project.Set(executionContext, projectRef);
            }
            else
            {
                TracingService.Trace("GetProjectFromTask: No project on task.");
                Project.Set(executionContext, null);
            }
        }
    }
}
