using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using ADC.MppImport.Services;

namespace ADC.MppImport.Workflows
{
    /// <summary>
    /// Resolves a user from project task assignments by role type.
    /// Finds project for case, queries tasks+assignments, picks active allocation with earliest start.
    /// </summary>
    public class ResolveRoleAllocationActivity : BaseCodeActivity
    {
        [Input("Case")]
        [ReferenceTarget("adc_case")]
        [RequiredArgument]
        public InArgument<EntityReference> Case { get; set; }

        [Input("Role Type")]
        [AttributeTarget("msdyn_projecttask", "adc_roletype")]
        [RequiredArgument]
        public InArgument<OptionSetValue> RoleType { get; set; }

        [Input("Reference Date")]
        public InArgument<DateTime> ReferenceDate { get; set; }

        [Output("Resolved User")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> ResolvedUser { get; set; }

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            var caseRef = Case.Get(executionContext);
            var roleOsv = RoleType.Get(executionContext);
            int roleType = roleOsv != null ? roleOsv.Value : 0;
            DateTime refDate = ReferenceDate.Get(executionContext);
            if (refDate == DateTime.MinValue)
                refDate = DateTime.UtcNow;

            TracingService.Trace("ResolveRoleAllocation: Case={0}, RoleType={1}, Date={2:yyyy-MM-dd}",
                caseRef.Id, roleType, refDate);

            if (roleType == 0)
            {
                TracingService.Trace("ResolveRoleAllocation: RoleType is required.");
                ResolvedUser.Set(executionContext, null);
                return;
            }

            var service = new RoleAllocationService(OrganizationService, TracingService);

            Guid? projectId = service.FindProjectForCase(caseRef.Id);
            if (!projectId.HasValue)
            {
                TracingService.Trace("ResolveRoleAllocation: No project found for case.");
                ResolvedUser.Set(executionContext, null);
                return;
            }

            var tasks = service.GetTaskAllocationsForProject(projectId.Value);
            var resolvedUser = RoleAllocationService.ResolveAllocation(roleType, tasks, refDate);
            ResolvedUser.Set(executionContext, resolvedUser);

            TracingService.Trace("ResolveRoleAllocation: Result = {0}",
                resolvedUser != null ? resolvedUser.Id.ToString() : "(null)");
        }
    }
}
