using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ADC.MppImport.Services
{
    /// <summary>
    /// Shared business logic for calculating "Day Count" on project tasks.
    /// Day Count = calendar days between a task's scheduled finish date and a reference date.
    ///
    /// Two scenarios based on case type:
    ///
    /// Scenario 1 — No Initiation Date (Duty Assessment, Accelerated Review, Exemption):
    ///   - "Application date" milestone → day count = 0 (reference point)
    ///   - All other tasks → datediff(task Finish, Application date Finish)
    ///
    /// Scenario 2 — Has Initiation Date (Dumping Inv, Subsidy Inv, D&amp;S Inv,
    ///              Review of Measures, Continuation, Anti-circumvention):
    ///   - Initially: "Application date" milestone → day count = -20
    ///   - All other tasks → datediff(task Finish, Application date Finish)
    ///   - When Initiation date is set on case:
    ///       • "Initiation date" milestone finish = case initiation date, day count = 0
    ///       • All tasks recalculated relative to Initiation date milestone finish
    ///
    /// Placeholder field names are marked with PLACEHOLDER comments — update as needed.
    /// </summary>
    public class DayCountService
    {
        private readonly IOrganizationService _service;
        private readonly ITracingService _trace;

        // =====================================================================
        // PLACEHOLDER SCHEMA NAMES — update these to match your environment
        // =====================================================================

        // msdyn_projecttask fields
        private const string TASK_ENTITY = "msdyn_projecttask";
        private const string TASK_FINISH = "msdyn_scheduledend";
        private const string TASK_NAME = "msdyn_subject";
        private const string TASK_PROJECT = "msdyn_project";
        private const string TASK_DAY_COUNT = "adc_daycount1";

        // Milestone task names (matched by msdyn_subject)
        private const string MILESTONE_APPLICATION_DATE = "Application date";
        private const string MILESTONE_INITIATION_DATE = "Initiation date";

        // adc_case fields
        private const string CASE_ENTITY = "adc_case";
        private const string CASE_TYPE = "adc_casetype";                    // PLACEHOLDER — optionset on case
        private const string CASE_INITIATION_DATE = "adc_dateofinitiation";

        // msdyn_project fields
        private const string PROJECT_ENTITY = "msdyn_project";
        private const string PROJECT_CASE_LINK = "adc_parentadccase";

        // Case type option set values — Scenario 1 (no initiation date)
        // PLACEHOLDER — replace with actual optionset int values
        private static readonly HashSet<int> Scenario1CaseTypes = new HashSet<int>
        {
            756360004, // Duty Assessment
            756360005, // Accelerated Review
            756360008  // Exemption
        };

        // Case type option set values — Scenario 2 (has initiation date)
        // PLACEHOLDER — replace with actual optionset int values
        private static readonly HashSet<int> Scenario2CaseTypes = new HashSet<int>
        {
            756360000, // Dumping Investigation
            756360001, // Subsidy Investigation
            756360002, // Dumping and Subsidy Investigation
            756360003, // Review of Measures
            756360006, // Continuation Inquiry
            756360007  // Anti-Circumvention
        };

        // The fixed day count for the Application date milestone in Scenario 2
        // before an initiation date is set
        private const decimal APPLICATION_DATE_INITIAL_DAY_COUNT = -20m;

        public DayCountService(IOrganizationService service, ITracingService trace)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _trace = trace;
        }

        // =====================================================================
        // PUBLIC API
        // =====================================================================

        /// <summary>
        /// Recalculates day count for a single project task.
        /// Called by the plugin when msdyn_scheduledend changes on a task.
        /// </summary>
        public void RecalcSingleTask(Guid taskId)
        {
            _trace?.Trace("DayCountService.RecalcSingleTask: taskId={0}", taskId);

            var task = _service.Retrieve(TASK_ENTITY, taskId,
                new ColumnSet(TASK_FINISH, TASK_NAME, TASK_PROJECT, TASK_DAY_COUNT));

            var taskName = task.GetAttributeValue<string>(TASK_NAME) ?? "(no name)";
            _trace?.Trace("DayCountService: Retrieved task '{0}' (id={1})", taskName, taskId);

            var projectRef = task.GetAttributeValue<EntityReference>(TASK_PROJECT);
            if (projectRef == null)
            {
                _trace?.Trace("DayCountService: Task '{0}' has no project reference, skipping.", taskName);
                return;
            }
            _trace?.Trace("DayCountService: Task belongs to project {0}", projectRef.Id);

            var context = BuildContext(projectRef.Id);
            if (context == null)
            {
                _trace?.Trace("DayCountService: BuildContext returned null for project {0}, aborting.", projectRef.Id);
                return;
            }

            decimal? dayCount = CalculateDayCount(task, context);
            _trace?.Trace("DayCountService: CalculateDayCount for '{0}' returned {1}",
                taskName, dayCount.HasValue ? dayCount.Value.ToString() : "(null)");

            if (dayCount.HasValue)
            {
                decimal existingDayCount = task.GetAttributeValue<decimal>(TASK_DAY_COUNT);
                if (existingDayCount != dayCount.Value)
                {
                    var update = new Entity(TASK_ENTITY, taskId);
                    update[TASK_DAY_COUNT] = dayCount.Value;
                    _service.Update(update);
                    _trace?.Trace("DayCountService: Updated task '{0}' day count: {1} → {2}",
                        taskName, existingDayCount, dayCount.Value);
                }
                else
                {
                    _trace?.Trace("DayCountService: Task '{0}' day count unchanged at {1}, no update needed.",
                        taskName, existingDayCount);
                }
            }
        }

        /// <summary>
        /// Recalculates day count for ALL project tasks in a project.
        /// Called after import completes, or when initiation date changes on the case.
        /// </summary>
        public void RecalcAllTasks(Guid projectId)
        {
            _trace?.Trace("DayCountService.RecalcAllTasks: projectId={0}", projectId);

            var context = BuildContext(projectId);
            if (context == null)
            {
                _trace?.Trace("DayCountService: BuildContext returned null for project {0}, aborting RecalcAllTasks.", projectId);
                return;
            }

            var tasks = RetrieveAllTasks(projectId);
            _trace?.Trace("DayCountService: Found {0} tasks to recalculate.", tasks.Count);

            if (tasks.Count == 0)
            {
                _trace?.Trace("DayCountService: No tasks found for project {0}, nothing to recalculate.", projectId);
                return;
            }

            int updated = 0;
            int skipped = 0;
            int unchanged = 0;
            foreach (var task in tasks)
            {
                string tName = task.GetAttributeValue<string>(TASK_NAME) ?? "(no name)";
                decimal? dayCount = CalculateDayCount(task, context);

                if (!dayCount.HasValue)
                {
                    skipped++;
                    continue;
                }

                decimal existingDayCount = task.GetAttributeValue<decimal>(TASK_DAY_COUNT);
                if (existingDayCount != dayCount.Value)
                {
                    var update = new Entity(TASK_ENTITY, task.Id);
                    update[TASK_DAY_COUNT] = dayCount.Value;
                    _service.Update(update);
                    _trace?.Trace("DayCountService:   [{0}] '{1}' — {2} → {3}",
                        updated + 1, tName, existingDayCount, dayCount.Value);
                    updated++;
                }
                else
                {
                    unchanged++;
                }
            }

            _trace?.Trace("DayCountService: RecalcAllTasks complete — {0} updated, {1} unchanged, {2} skipped (no finish/ref date).",
                updated, unchanged, skipped);
        }

        /// <summary>
        /// Called when initiation date is set/changed on an adc_case.
        /// Updates the "Initiation date" milestone's finish date and recalculates all tasks.
        /// </summary>
        public void OnInitiationDateChanged(Guid caseId, DateTime? initiationDate)
        {
            _trace?.Trace("DayCountService.OnInitiationDateChanged: caseId={0}, date={1}",
                caseId, initiationDate.HasValue ? initiationDate.Value.ToString("yyyy-MM-dd") : "(null)");

            // Find the linked project
            Guid? projectId = FindProjectForCase(caseId);
            if (!projectId.HasValue)
            {
                _trace?.Trace("DayCountService: No project linked to case {0}, skipping.", caseId);
                return;
            }

            if (initiationDate.HasValue)
            {
                // Find the "Initiation date" milestone and update its finish date
                var milestone = FindMilestoneByName(projectId.Value, MILESTONE_INITIATION_DATE);
                if (milestone != null)
                {
                    var milestoneUpdate = new Entity(TASK_ENTITY, milestone.Id);
                    milestoneUpdate[TASK_FINISH] = initiationDate.Value;
                    milestoneUpdate[TASK_DAY_COUNT] = 0m;
                    _service.Update(milestoneUpdate);
                    _trace?.Trace("DayCountService: Updated Initiation date milestone finish to {0:yyyy-MM-dd}",
                        initiationDate.Value);
                }
                else
                {
                    _trace?.Trace("DayCountService: WARNING — Could not find '{0}' milestone in project {1}.",
                        MILESTONE_INITIATION_DATE, projectId.Value);
                }
            }

            // Recalculate all tasks in the project
            RecalcAllTasks(projectId.Value);
        }

        // =====================================================================
        // CORE CALCULATION
        // =====================================================================

        /// <summary>
        /// Holds pre-fetched context needed for day count calculations.
        /// </summary>
        private class CalcContext
        {
            public Guid ProjectId;
            public int Scenario;                    // 1 or 2
            public DateTime? ApplicationDateFinish; // Finish date of the "Application date" milestone
            public DateTime? InitiationDateFinish;  // Finish date of the "Initiation date" milestone (if set)
            public bool HasInitiationDate;          // Whether the case has an initiation date set
        }

        /// <summary>
        /// Builds the calculation context for a project: determines scenario,
        /// finds milestone dates, checks case initiation date.
        /// </summary>
        private CalcContext BuildContext(Guid projectId)
        {
            _trace?.Trace("DayCountService.BuildContext: projectId={0}", projectId);

            // Get the linked case
            var project = _service.Retrieve(PROJECT_ENTITY, projectId,
                new ColumnSet(PROJECT_CASE_LINK));

            var caseRef = project.GetAttributeValue<EntityReference>(PROJECT_CASE_LINK);
            _trace?.Trace("DayCountService.BuildContext: {0} = {1}",
                PROJECT_CASE_LINK, caseRef != null ? caseRef.Id.ToString() : "(null)");

            if (caseRef == null)
            {
                _trace?.Trace("DayCountService: Project {0} has no linked case (field '{1}' is null) — skipping day count calc.",
                    projectId, PROJECT_CASE_LINK);
                return null;
            }

            // Determine scenario from case type
            var caseRecord = _service.Retrieve(CASE_ENTITY, caseRef.Id,
                new ColumnSet(CASE_TYPE, CASE_INITIATION_DATE));

            var caseTypeOsv = caseRecord.GetAttributeValue<OptionSetValue>(CASE_TYPE);
            int caseTypeValue = caseTypeOsv != null ? caseTypeOsv.Value : -1;
            _trace?.Trace("DayCountService.BuildContext: Case {0}, {1} = {2}",
                caseRef.Id, CASE_TYPE, caseTypeValue);

            int scenario;
            if (Scenario1CaseTypes.Contains(caseTypeValue))
                scenario = 1;
            else if (Scenario2CaseTypes.Contains(caseTypeValue))
                scenario = 2;
            else
            {
                _trace?.Trace("DayCountService: Unrecognised case type {0}, defaulting to Scenario 1.", caseTypeValue);
                scenario = 1;
            }

            DateTime? caseInitiationDate = caseRecord.GetAttributeValue<DateTime?>(CASE_INITIATION_DATE);

            _trace?.Trace("DayCountService.BuildContext: Scenario={0}, CaseType={1}, {2}={3}",
                scenario, caseTypeValue, CASE_INITIATION_DATE,
                caseInitiationDate.HasValue ? caseInitiationDate.Value.ToString("yyyy-MM-dd HH:mm:ss") : "(not set)");

            // Find milestone tasks
            var appMilestone = FindMilestoneByName(projectId, MILESTONE_APPLICATION_DATE);
            DateTime? appDateFinish = appMilestone?.GetAttributeValue<DateTime?>(TASK_FINISH);
            _trace?.Trace("DayCountService.BuildContext: '{0}' milestone {1}, finish={2}",
                MILESTONE_APPLICATION_DATE,
                appMilestone != null ? "FOUND (id=" + appMilestone.Id + ")" : "NOT FOUND",
                appDateFinish.HasValue ? appDateFinish.Value.ToString("yyyy-MM-dd") : "(null)");

            DateTime? initDateFinish = null;
            if (scenario == 2)
            {
                var initMilestone = FindMilestoneByName(projectId, MILESTONE_INITIATION_DATE);
                initDateFinish = initMilestone?.GetAttributeValue<DateTime?>(TASK_FINISH);
                _trace?.Trace("DayCountService.BuildContext: '{0}' milestone {1}, finish={2}",
                    MILESTONE_INITIATION_DATE,
                    initMilestone != null ? "FOUND (id=" + initMilestone.Id + ")" : "NOT FOUND",
                    initDateFinish.HasValue ? initDateFinish.Value.ToString("yyyy-MM-dd") : "(null)");
            }

            var ctx = new CalcContext
            {
                ProjectId = projectId,
                Scenario = scenario,
                ApplicationDateFinish = appDateFinish?.Date,
                InitiationDateFinish = initDateFinish?.Date,
                HasInitiationDate = caseInitiationDate.HasValue
            };

            _trace?.Trace("DayCountService.BuildContext: RESULT — Scenario={0}, AppDateFinish={1}, InitDateFinish={2}, HasInitDate={3}",
                ctx.Scenario,
                ctx.ApplicationDateFinish.HasValue ? ctx.ApplicationDateFinish.Value.ToString("yyyy-MM-dd") : "(null)",
                ctx.InitiationDateFinish.HasValue ? ctx.InitiationDateFinish.Value.ToString("yyyy-MM-dd") : "(null)",
                ctx.HasInitiationDate);

            return ctx;
        }

        /// <summary>
        /// Calculates the day count for a single task given the pre-built context.
        /// Returns null if calculation is not possible (e.g. no reference date).
        /// </summary>
        private decimal? CalculateDayCount(Entity task, CalcContext ctx)
        {
            string taskName = task.GetAttributeValue<string>(TASK_NAME) ?? "";
            DateTime? taskFinish = task.GetAttributeValue<DateTime?>(TASK_FINISH);

            // --- Handle special milestone tasks ---

            // "Application date" milestone
            if (IsApplicationDateMilestone(taskName))
            {
                if (ctx.Scenario == 1)
                {
                    _trace?.Trace("DayCountService.Calc: '{0}' is Application date milestone, Scenario 1 → 0", taskName);
                    return 0m;
                }

                // Scenario 2: -20 initially, or recalculate relative to initiation date if set
                if (ctx.HasInitiationDate && ctx.InitiationDateFinish.HasValue && taskFinish.HasValue)
                {
                    decimal val = (decimal)(taskFinish.Value.Date - ctx.InitiationDateFinish.Value).TotalDays;
                    _trace?.Trace("DayCountService.Calc: '{0}' is Application date milestone, Scenario 2 with initiation → {1} (finish={2}, initRef={3})",
                        taskName, val, taskFinish.Value.Date.ToString("yyyy-MM-dd"), ctx.InitiationDateFinish.Value.ToString("yyyy-MM-dd"));
                    return val;
                }

                _trace?.Trace("DayCountService.Calc: '{0}' is Application date milestone, Scenario 2 no initiation → {1}",
                    taskName, APPLICATION_DATE_INITIAL_DAY_COUNT);
                return APPLICATION_DATE_INITIAL_DAY_COUNT;
            }

            // "Initiation date" milestone
            if (IsInitiationDateMilestone(taskName))
            {
                _trace?.Trace("DayCountService.Calc: '{0}' is Initiation date milestone → 0", taskName);
                return 0m;
            }

            // --- Regular tasks ---

            if (!taskFinish.HasValue)
            {
                _trace?.Trace("DayCountService.Calc: Task '{0}' has no finish date, skipping.", taskName);
                return null;
            }

            // Determine the reference date:
            // If Scenario 2 and initiation date is set → use Initiation date milestone finish
            // Otherwise → use Application date milestone finish
            DateTime? referenceDate;
            string refSource;
            if (ctx.Scenario == 2 && ctx.HasInitiationDate && ctx.InitiationDateFinish.HasValue)
            {
                referenceDate = ctx.InitiationDateFinish;
                refSource = "Initiation date milestone";
            }
            else
            {
                referenceDate = ctx.ApplicationDateFinish;
                refSource = "Application date milestone";
            }

            if (!referenceDate.HasValue)
            {
                _trace?.Trace("DayCountService.Calc: No reference date ({0}) available for task '{1}', skipping.", refSource, taskName);
                return null;
            }

            decimal result = (decimal)(taskFinish.Value.Date - referenceDate.Value).TotalDays;
            _trace?.Trace("DayCountService.Calc: '{0}' finish={1} - ref={2} ({3}) = {4}",
                taskName, taskFinish.Value.Date.ToString("yyyy-MM-dd"), referenceDate.Value.ToString("yyyy-MM-dd"),
                refSource, result);
            return result;
        }

        // =====================================================================
        // HELPERS
        // =====================================================================

        /// <summary>
        /// Finds a milestone task in a project by matching msdyn_subject (case-insensitive).
        /// </summary>
        private Entity FindMilestoneByName(Guid projectId, string milestoneName)
        {
            var query = new QueryExpression(TASK_ENTITY)
            {
                ColumnSet = new ColumnSet(TASK_NAME, TASK_FINISH, TASK_DAY_COUNT),
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression(TASK_PROJECT, ConditionOperator.Equal, projectId),
                        new ConditionExpression(TASK_NAME, ConditionOperator.Equal, milestoneName)
                    }
                },
                TopCount = 1
            };

            var results = _service.RetrieveMultiple(query);
            var found = results.Entities.FirstOrDefault();
            _trace?.Trace("DayCountService.FindMilestoneByName: '{0}' in project {1} → {2}",
                milestoneName, projectId, found != null ? "FOUND (id=" + found.Id + ")" : "NOT FOUND");
            return found;
        }

        /// <summary>
        /// Retrieves all tasks for a project.
        /// </summary>
        private List<Entity> RetrieveAllTasks(Guid projectId)
        {
            var query = new QueryExpression(TASK_ENTITY)
            {
                ColumnSet = new ColumnSet(TASK_NAME, TASK_FINISH, TASK_DAY_COUNT),
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression(TASK_PROJECT, ConditionOperator.Equal, projectId)
                    }
                }
            };

            var results = _service.RetrieveMultiple(query);
            _trace?.Trace("DayCountService.RetrieveAllTasks: project {0} → {1} tasks returned.", projectId, results.Entities.Count);
            return results.Entities.ToList();
        }

        /// <summary>
        /// Finds the project linked to a case via adc_parentadccase.
        /// </summary>
        private Guid? FindProjectForCase(Guid caseId)
        {
            var query = new QueryExpression(PROJECT_ENTITY)
            {
                ColumnSet = new ColumnSet(PROJECT_CASE_LINK),
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression(PROJECT_CASE_LINK, ConditionOperator.Equal, caseId)
                    }
                },
                TopCount = 1
            };

            var results = _service.RetrieveMultiple(query);
            if (results.Entities.Count > 0)
            {
                _trace?.Trace("DayCountService.FindProjectForCase: case {0} → project {1}", caseId, results.Entities[0].Id);
                return results.Entities[0].Id;
            }

            _trace?.Trace("DayCountService.FindProjectForCase: case {0} → no project found.", caseId);
            return null;
        }

        private static bool IsApplicationDateMilestone(string taskName)
        {
            return string.Equals(taskName, MILESTONE_APPLICATION_DATE, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsInitiationDateMilestone(string taskName)
        {
            return string.Equals(taskName, MILESTONE_INITIATION_DATE, StringComparison.OrdinalIgnoreCase);
        }
    }
}
