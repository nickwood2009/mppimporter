using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ADC.MppImport.Services
{
    /// <summary>
    /// Resolves Case Director and Case Manager allocations from project task assignments.
    /// The core ResolveAllocation method is static and pure (no Dataverse calls) for unit testing.
    /// Dataverse query helpers are instance methods that require IOrganizationService.
    /// </summary>
    public class RoleAllocationService
    {
        #region Schema constants — update these to match your Dataverse schema

        // adc_case entity
        public const string CASE_ENTITY = "adc_case";

        // msdyn_projecttask fields
        public const string TASK_ENTITY = "msdyn_projecttask";
        public const string TASK_PROJECT_FIELD = "msdyn_project";
        public const string TASK_NAME_FIELD = "msdyn_subject";
        public const string TASK_START_FIELD = "msdyn_scheduledstart";
        public const string TASK_END_FIELD = "msdyn_scheduledend";
        public const string TASK_ROLE_FIELD = "adc_assigneerole";       // TODO: verify schema name

        // msdyn_resourceassignment fields (Assigned To is on this related entity, not on task)
        public const string ASSIGNMENT_ENTITY = "msdyn_resourceassignment";
        public const string ASSIGNMENT_TASK_FIELD = "msdyn_taskid";
        public const string ASSIGNMENT_RESOURCE_FIELD = "msdyn_bookableresourceid";

        // bookableresource → systemuser resolution
        public const string RESOURCE_ENTITY = "bookableresource";
        public const string RESOURCE_USER_FIELD = "userid";

        // Project entity
        public const string PROJECT_ENTITY = "msdyn_project";
        public const string PROJECT_CASE_FIELD = "adc_parentadccase";


        #endregion

        private readonly IOrganizationService _service;
        private readonly ITracingService _trace;

        public RoleAllocationService(IOrganizationService service, ITracingService trace)
        {
            _service = service;
            _trace = trace;
        }

        #region Pure logic — unit-testable, no Dataverse dependency

        /// <summary>
        /// Resolves which user should be allocated for a given role type based on project task data.
        /// Returns null if no active allocation found (caller should retain existing case value).
        ///
        /// Business rules (identical for Case Director and Case Manager):
        ///   1. Filter tasks to those matching the target role type that have a valid assignee.
        ///   2. Among those, find "active" tasks where StartDate &lt;= today AND EndDate &gt;= today.
        ///   3. If one or more active tasks exist, return the assignee of the one with the earliest StartDate.
        ///   4. If no active tasks (all ended or none found), return null — caller retains existing value.
        ///
        /// If multiple assignees exist for the same task, the caller is expected to have already
        /// picked the first one when building the TaskAllocationInfo list.
        /// </summary>
        /// <param name="targetRoleType">The option set value for the role (e.g. ROLE_CASE_DIRECTOR).</param>
        /// <param name="tasks">All project tasks with their role types, dates, and resolved assignees.</param>
        /// <param name="today">Reference date for active/inactive determination.</param>
        /// <returns>EntityReference to the resolved systemuser, or null if no active allocation.</returns>
        public static EntityReference ResolveAllocation(
            int targetRoleType,
            IList<TaskAllocationInfo> tasks,
            DateTime today)
        {
            if (tasks == null || tasks.Count == 0)
                return null;

            // Step 1: Filter to tasks matching the target role type with a valid assignee
            var roleTasks = new List<TaskAllocationInfo>();
            for (int i = 0; i < tasks.Count; i++)
            {
                var t = tasks[i];
                if (t.RoleType.HasValue
                    && t.RoleType.Value == targetRoleType
                    && t.AssignedTo != null)
                {
                    roleTasks.Add(t);
                }
            }

            if (roleTasks.Count == 0)
                return null;

            // Step 2: Filter to "active" allocations: StartDate <= today AND EndDate >= today
            var activeTasks = new List<TaskAllocationInfo>();
            for (int i = 0; i < roleTasks.Count; i++)
            {
                var t = roleTasks[i];
                if (t.StartDate.HasValue && t.StartDate.Value.Date <= today.Date
                    && t.EndDate.HasValue && t.EndDate.Value.Date >= today.Date)
                {
                    activeTasks.Add(t);
                }
            }

            // Step 3: If active allocations exist, pick the one with the earliest start date
            if (activeTasks.Count > 0)
            {
                activeTasks.Sort((a, b) => a.StartDate.Value.CompareTo(b.StartDate.Value));
                return activeTasks[0].AssignedTo;
            }

            // Step 4: No active allocation — return null so caller retains existing case value
            return null;
        }


        #endregion

        #region Dataverse queries

        /// <summary>
        /// Finds the project linked to a case via adc_parentadccase on msdyn_project.
        /// Returns null if no project is linked.
        /// </summary>
        public Guid? FindProjectForCase(Guid caseId)
        {
            var query = new QueryExpression(PROJECT_ENTITY)
            {
                ColumnSet = new ColumnSet(false),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression(PROJECT_CASE_FIELD, ConditionOperator.Equal, caseId)
                    }
                },
                TopCount = 1
            };

            var results = _service.RetrieveMultiple(query);
            if (results.Entities.Count > 0)
            {
                _trace?.Trace("RoleAllocationService: Case {0} → project {1}", caseId, results.Entities[0].Id);
                return results.Entities[0].Id;
            }

            _trace?.Trace("RoleAllocationService: Case {0} → no project found.", caseId);
            return null;
        }

        /// <summary>
        /// Queries all project tasks for a given project, joins resource assignments and
        /// bookable resources to resolve the assigned systemuser for each task.
        /// Returns one TaskAllocationInfo per task (first resource assignment wins if multiple).
        /// </summary>
        public List<TaskAllocationInfo> GetTaskAllocationsForProject(Guid projectId)
        {
            _trace?.Trace("RoleAllocationService: Querying task allocations for project {0}", projectId);

            var query = new QueryExpression(TASK_ENTITY)
            {
                ColumnSet = new ColumnSet(TASK_NAME_FIELD, TASK_START_FIELD, TASK_END_FIELD, TASK_ROLE_FIELD),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression(TASK_PROJECT_FIELD, ConditionOperator.Equal, projectId)
                    }
                }
            };

            // LEFT JOIN msdyn_resourceassignment on task PK → assignment FK
            var raLink = query.AddLink(
                ASSIGNMENT_ENTITY,
                "msdyn_projecttaskid",
                ASSIGNMENT_TASK_FIELD,
                JoinOperator.LeftOuter);
            raLink.EntityAlias = "ra";
            raLink.Columns.AddColumn(ASSIGNMENT_RESOURCE_FIELD);

            // LEFT JOIN bookableresource to resolve the user behind the resource
            var brLink = raLink.AddLink(
                RESOURCE_ENTITY,
                ASSIGNMENT_RESOURCE_FIELD,
                "bookableresourceid",
                JoinOperator.LeftOuter);
            brLink.EntityAlias = "br";
            brLink.Columns.AddColumn(RESOURCE_USER_FIELD);

            var results = _service.RetrieveMultiple(query);
            _trace?.Trace("RoleAllocationService: Retrieved {0} rows (tasks x assignments)", results.Entities.Count);

            // Build one DTO per task — first resource assignment per task wins
            var seen = new Dictionary<Guid, TaskAllocationInfo>();
            foreach (var entity in results.Entities)
            {
                Guid taskId = entity.Id;

                // Skip if we already recorded this task (first assignment wins)
                if (seen.ContainsKey(taskId))
                    continue;

                var roleOsv = entity.GetAttributeValue<OptionSetValue>(TASK_ROLE_FIELD);
                int? roleType = roleOsv != null ? (int?)roleOsv.Value : null;

                // Resolve the systemuser from the joined bookableresource.userid alias
                EntityReference assignedTo = null;
                var userAlias = entity.GetAttributeValue<AliasedValue>("br." + RESOURCE_USER_FIELD);
                if (userAlias != null && userAlias.Value is EntityReference)
                {
                    assignedTo = (EntityReference)userAlias.Value;
                }

                var info = new TaskAllocationInfo
                {
                    TaskId = taskId,
                    TaskName = entity.GetAttributeValue<string>(TASK_NAME_FIELD) ?? "",
                    RoleType = roleType,
                    StartDate = entity.GetAttributeValue<DateTime?>(TASK_START_FIELD),
                    EndDate = entity.GetAttributeValue<DateTime?>(TASK_END_FIELD),
                    AssignedTo = assignedTo
                };

                seen[taskId] = info;
            }

            var list = new List<TaskAllocationInfo>(seen.Values);

            int withRole = 0;
            int withAssignee = 0;
            foreach (var t in list)
            {
                if (t.RoleType.HasValue) withRole++;
                if (t.AssignedTo != null) withAssignee++;
            }
            _trace?.Trace("RoleAllocationService: {0} tasks total, {1} with role type, {2} with assignee",
                list.Count, withRole, withAssignee);

            return list;
        }

        #endregion
    }

    /// <summary>
    /// DTO representing a project task's role allocation data.
    /// Used as input to the pure RoleAllocationService.ResolveAllocation function.
    /// Designed for easy construction in unit tests.
    /// </summary>
    public class TaskAllocationInfo
    {
        public Guid TaskId { get; set; }
        public string TaskName { get; set; }

        /// <summary>adc_assigneerole option set value (e.g. ROLE_CASE_DIRECTOR, ROLE_CASE_MANAGER).</summary>
        public int? RoleType { get; set; }

        /// <summary>msdyn_scheduledstart on the project task.</summary>
        public DateTime? StartDate { get; set; }

        /// <summary>msdyn_scheduledend on the project task.</summary>
        public DateTime? EndDate { get; set; }

        /// <summary>
        /// The resolved systemuser from the task's first resource assignment
        /// (bookableresource → userid). Null if no assignment or resource has no linked user.
        /// </summary>
        public EntityReference AssignedTo { get; set; }
    }
}
