using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using ADC.MppImport.Services;

namespace ADC.MppImport.Workflows
{
    /// <summary>
    /// Resolves a Case Director or Case Manager from project task resource assignments
    /// and updates the corresponding lookup field on the adc_case record.
    ///
    /// Inputs:
    ///   - Case: the adc_case record to evaluate and update.
    ///   - RoleTypeValue: the adc_assigneerole option set integer value to resolve (global optionset).
    ///   - CaseFieldName: the schema name of the lookup field on adc_case to update (e.g. "adc_casedirector").
    ///
    /// Outputs:
    ///   - ResolvedUser: the systemuser EntityReference resolved (null if no active allocation).
    ///   - WasUpdated: true if the case field was updated, false if retained/skipped.
    ///
    /// Business rules:
    ///   1. Find the project linked to the case.
    ///   2. Query all tasks + resource assignments for that project.
    ///   3. Filter to tasks with the specified role type, active dates (Start &lt;= today, End &gt;= today),
    ///      and a valid assigned user.
    ///   4. Pick the task with the earliest Start Date; return its assigned user.
    ///   5. If no active allocation exists, retain the existing case value (no update).
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

            // 1. Find the project linked to this case
            Guid? projectId = service.FindProjectForCase(caseRef.Id);
            if (!projectId.HasValue)
            {
                TracingService.Trace("ResolveRoleAllocation: No project found for case. Skipping.");
                ResolvedUser.Set(executionContext, null);
                WasUpdated.Set(executionContext, false);
                return;
            }
            TracingService.Trace("ResolveRoleAllocation: Project = {0}", projectId.Value);

            // 2. Query task allocations (tasks + resource assignments + users)
            var tasks = service.GetTaskAllocationsForProject(projectId.Value);

            // 3. Resolve allocation using pure logic (static, unit-testable)
            var resolvedUser = RoleAllocationService.ResolveAllocation(roleType, tasks, DateTime.UtcNow);
            ResolvedUser.Set(executionContext, resolvedUser);

            if (resolvedUser != null)
            {
                TracingService.Trace("ResolveRoleAllocation: Resolved user = {0} ({1})",
                    resolvedUser.Id, resolvedUser.LogicalName);

                // 4. Update the specified case field
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
