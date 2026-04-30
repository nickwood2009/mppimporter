using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using ADC.MppImport.Services;

namespace ADC.MppImport.Workflows
{
    /// <summary>
    /// Resolves a user from project task assignments by role type and updates a case lookup field.
    /// Finds project for case, queries tasks+assignments, picks active allocation with earliest start.
    /// </summary>
    public class ResolveRoleAllocationActivity : BaseCodeActivity
    {
        [Input("Case")]
        [ReferenceTarget("adc_case")]
        [RequiredArgument]
        public InArgument<EntityReference> Case { get; set; }

        [Input("Role Type Value")]
        [RequiredArgument]
        public InArgument<int> RoleTypeValue { get; set; }

        [Input("Case Field Name")]
        [RequiredArgument]
        public InArgument<string> CaseFieldName { get; set; }

        [Output("Resolved User")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> ResolvedUser { get; set; }

        [Output("Was Updated")]
        public OutArgument<bool> WasUpdated { get; set; }

        protected override void ExecuteActivity(CodeActivityContext executionContext)
        {
            var caseRef = Case.Get(executionContext);
            int roleType = RoleTypeValue.Get(executionContext);
            string caseFieldName = CaseFieldName.Get(executionContext);

            TracingService.Trace("ResolveRoleAllocation: Case={0}, RoleType={1}, CaseField={2}",
                caseRef.Id, roleType, caseFieldName ?? "(null)");

            if (string.IsNullOrEmpty(caseFieldName))
            {
                TracingService.Trace("ResolveRoleAllocation: CaseFieldName is required.");
                ResolvedUser.Set(executionContext, null);
                WasUpdated.Set(executionContext, false);
                return;
            }

            var service = new RoleAllocationService(OrganizationService, TracingService);

            Guid? projectId = service.FindProjectForCase(caseRef.Id);
            if (!projectId.HasValue)
            {
                TracingService.Trace("ResolveRoleAllocation: No project found for case. Skipping.");
                ResolvedUser.Set(executionContext, null);
                WasUpdated.Set(executionContext, false);
                return;
            }
            TracingService.Trace("ResolveRoleAllocation: Project = {0}", projectId.Value);

            var tasks = service.GetTaskAllocationsForProject(projectId.Value);
            var resolvedUser = RoleAllocationService.ResolveAllocation(roleType, tasks, DateTime.UtcNow);
            ResolvedUser.Set(executionContext, resolvedUser);

            if (resolvedUser != null)
            {
                TracingService.Trace("ResolveRoleAllocation: Resolved user = {0} ({1})",
                    resolvedUser.Id, resolvedUser.LogicalName);

                var caseUpdate = new Entity(RoleAllocationService.CASE_ENTITY, caseRef.Id);
                caseUpdate[caseFieldName] = resolvedUser;
                OrganizationService.Update(caseUpdate);
                WasUpdated.Set(executionContext, true);
                TracingService.Trace("ResolveRoleAllocation: Updated case.{0} = {1}", caseFieldName, resolvedUser.Id);
            }
            else
            {
                TracingService.Trace("ResolveRoleAllocation: No active allocation found. Retaining existing value.");
                WasUpdated.Set(executionContext, false);
            }
        }
    }
}
