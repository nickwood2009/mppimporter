using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ADC.MppImport.Services
{
    /// <summary>
    /// Resolves user allocations from project task assignments based on role type.
    /// ResolveAllocation is static/pure for unit testing; Dataverse helpers are instance methods.
    /// </summary>
    public class RoleAllocationService
    {
        #region Schema constants

        public const string CASE_ENTITY = "adc_case";

        public const string TASK_ENTITY = "msdyn_projecttask";
        public const string TASK_PROJECT_FIELD = "msdyn_project";
        public const string TASK_NAME_FIELD = "msdyn_subject";
        public const string TASK_START_FIELD = "msdyn_scheduledstart";
        public const string TASK_END_FIELD = "msdyn_scheduledend";
        public const string TASK_ROLE_FIELD = "adc_roletype";

        public const string ASSIGNMENT_ENTITY = "msdyn_resourceassignment";
        public const string ASSIGNMENT_TASK_FIELD = "msdyn_taskid";
        public const string ASSIGNMENT_TEAM_FIELD = "msdyn_projectteamid";

        public const string TEAM_ENTITY = "msdyn_projectteam";
        public const string TEAM_RESOURCE_FIELD = "msdyn_bookableresourceid";

        public const string RESOURCE_ENTITY = "bookableresource";
        public const string RESOURCE_USER_FIELD = "userid";

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

        #region Pure logic

        // targetRoleType: option set int value for the role
        // tasks: all project tasks with role types, dates, and resolved assignees
        // today: reference date for active/inactive determination
        public static EntityReference ResolveAllocation(
            int targetRoleType,
            IList<TaskAllocationInfo> tasks,
            DateTime today)
        {
            if (tasks == null || tasks.Count == 0)
                return null;

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

            if (activeTasks.Count > 0)
            {
                activeTasks.Sort((a, b) => a.StartDate.Value.CompareTo(b.StartDate.Value));
                return activeTasks[0].AssignedTo;
            }

            return null;
        }


        #endregion

        #region Dataverse queries

        // caseId: the adc_case record to find the linked project for
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

        // projectId: queries tasks + resource assignments + bookableresource for this project
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

            var raLink = query.AddLink(
                ASSIGNMENT_ENTITY,
                "msdyn_projecttaskid",
                ASSIGNMENT_TASK_FIELD,
                JoinOperator.LeftOuter);
            raLink.EntityAlias = "ra";
            raLink.Columns.AddColumn(ASSIGNMENT_TEAM_FIELD);

            var tmLink = raLink.AddLink(
                TEAM_ENTITY,
                ASSIGNMENT_TEAM_FIELD,
                "msdyn_projectteamid",
                JoinOperator.LeftOuter);
            tmLink.EntityAlias = "tm";
            tmLink.Columns.AddColumn(TEAM_RESOURCE_FIELD);

            var results = _service.RetrieveMultiple(query);
            _trace?.Trace("RoleAllocationService: Retrieved {0} rows (tasks x assignments)", results.Entities.Count);

            // collect unique bookable resource IDs from team members
            var brIds = new HashSet<Guid>();
            foreach (var entity in results.Entities)
            {
                var brAlias = entity.GetAttributeValue<AliasedValue>("tm." + TEAM_RESOURCE_FIELD);
                if (brAlias != null && brAlias.Value is EntityReference brRef)
                    brIds.Add(brRef.Id);
            }

            // batch resolve bookable resource → systemuser
            var brToUser = new Dictionary<Guid, EntityReference>();
            if (brIds.Count > 0)
            {
                var brQuery = new QueryExpression(RESOURCE_ENTITY)
                {
                    ColumnSet = new ColumnSet(RESOURCE_USER_FIELD),
                    Criteria = new FilterExpression()
                };
                var guids = new object[brIds.Count];
                brIds.CopyTo(new Guid[brIds.Count], 0);
                var guidList = new List<Guid>(brIds);
                brQuery.Criteria.AddCondition("bookableresourceid", ConditionOperator.In, guidList.ConvertAll(g => (object)g).ToArray());

                var brResults = _service.RetrieveMultiple(brQuery);
                foreach (var br in brResults.Entities)
                {
                    var userRef = br.GetAttributeValue<EntityReference>(RESOURCE_USER_FIELD);
                    if (userRef != null)
                        brToUser[br.Id] = userRef;
                }
                _trace?.Trace("RoleAllocationService: Resolved {0} bookable resources → {1} users", brIds.Count, brToUser.Count);
            }

            var seen = new Dictionary<Guid, TaskAllocationInfo>();
            foreach (var entity in results.Entities)
            {
                Guid taskId = entity.Id;

                if (seen.ContainsKey(taskId) && seen[taskId].AssignedTo != null)
                    continue;

                var roleOsv = entity.GetAttributeValue<OptionSetValue>(TASK_ROLE_FIELD);
                int? roleType = roleOsv != null ? (int?)roleOsv.Value : null;

                EntityReference assignedTo = null;
                var brAlias = entity.GetAttributeValue<AliasedValue>("tm." + TEAM_RESOURCE_FIELD);
                if (brAlias != null && brAlias.Value is EntityReference brRef && brToUser.ContainsKey(brRef.Id))
                {
                    assignedTo = brToUser[brRef.Id];
                }

                _trace?.Trace("  Task {0} ({1}): br={2}, user={3}",
                    taskId,
                    entity.GetAttributeValue<string>(TASK_NAME_FIELD) ?? "",
                    brAlias != null && brAlias.Value is EntityReference r ? r.Id.ToString() : "null",
                    assignedTo != null ? assignedTo.Id.ToString() : "null");

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

    public class TaskAllocationInfo
    {
        public Guid TaskId { get; set; }
        public string TaskName { get; set; }
        public int? RoleType { get; set; }        // adc_roletype option set value
        public DateTime? StartDate { get; set; }  // msdyn_scheduledstart
        public DateTime? EndDate { get; set; }    // msdyn_scheduledend
        public EntityReference AssignedTo { get; set; } // resolved systemuser from bookableresource
    }
}
