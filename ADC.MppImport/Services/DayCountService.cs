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
        private const string TASK_DAY_COUNT = "adc_daycount";               // PLACEHOLDER — int field on task

        // Milestone task names (matched by msdyn_subject)
        private const string MILESTONE_APPLICATION_DATE = "Application date";
        private const string MILESTONE_INITIATION_DATE = "Initiation date";

        // adc_case fields
        private const string CASE_ENTITY = "adc_case";
        private const string CASE_TYPE = "adc_casetype";                    // PLACEHOLDER — optionset on case
        private const string CASE_INITIATION_DATE = "adc_dateofinitiation";

        // msdyn_project fields
        private const string PROJECT_ENTITY = "msdyn_project";
        private const string PROJECT_CASE_LINK = "adc_projectid";

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
        private const int APPLICATION_DATE_INITIAL_DAY_COUNT = -20;

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

            var projectRef = task.GetAttributeValue<EntityReference>(TASK_PROJECT);
            if (projectRef == null)
            {
                _trace?.Trace("DayCountService: Task has no project reference, skipping.");
                return;
            }

            var context = BuildContext(projectRef.Id);
            if (context == null) return;

            int? dayCount = CalculateDayCount(task, context);
            if (dayCount.HasValue)
            {
                int existingDayCount = task.GetAttributeValue<int>(TASK_DAY_COUNT);
                if (existingDayCount != dayCount.Value)
                {
                    var update = new Entity(TASK_ENTITY, taskId);
                    update[TASK_DAY_COUNT] = dayCount.Value;
                    _service.Update(update);
                    _trace?.Trace("DayCountService: Updated task {0} day count: {1} → {2}",
                        taskId, existingDayCount, dayCount.Value);
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
            if (context == null) return;

            var tasks = RetrieveAllTasks(projectId);
            _trace?.Trace("DayCountService: Found {0} tasks to recalculate.", tasks.Count);

            int updated = 0;
            foreach (var task in tasks)
            {
                int? dayCount = CalculateDayCount(task, context);
                if (dayCount.HasValue)
                {
                    int existingDayCount = task.GetAttributeValue<int>(TASK_DAY_COUNT);
                    if (existingDayCount != dayCount.Value)
                    {
                        var update = new Entity(TASK_ENTITY, task.Id);
                        update[TASK_DAY_COUNT] = dayCount.Value;
                        _service.Update(update);
                        updated++;
                    }
                }
            }

            _trace?.Trace("DayCountService: Recalculated {0}/{1} tasks.", updated, tasks.Count);
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
                    milestoneUpdate[TASK_DAY_COUNT] = 0;
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
            // Get the linked case
            var project = _service.Retrieve(PROJECT_ENTITY, projectId,
                new ColumnSet(PROJECT_CASE_LINK));

            var caseRef = project.GetAttributeValue<EntityReference>(PROJECT_CASE_LINK);
            if (caseRef == null)
            {
                _trace?.Trace("DayCountService: Project {0} has no linked case — skipping day count calc.", projectId);
                return null;
            }

            // Determine scenario from case type
            var caseRecord = _service.Retrieve(CASE_ENTITY, caseRef.Id,
                new ColumnSet(CASE_TYPE, CASE_INITIATION_DATE));

            var caseTypeOsv = caseRecord.GetAttributeValue<OptionSetValue>(CASE_TYPE);
            int caseTypeValue = caseTypeOsv != null ? caseTypeOsv.Value : -1;

            int scenario;
            if (Scenario1CaseTypes.Contains(caseTypeValue))
                scenario = 1;
            else if (Scenario2CaseTypes.Contains(caseTypeValue))
                scenario = 2;
            else
            {
                // Default to Scenario 1 if case type is unrecognised
                _trace?.Trace("DayCountService: Unrecognised case type {0}, defaulting to Scenario 1.", caseTypeValue);
                scenario = 1;
            }

            DateTime? caseInitiationDate = caseRecord.GetAttributeValue<DateTime?>(CASE_INITIATION_DATE);

            _trace?.Trace("DayCountService: Scenario={0}, CaseType={1}, InitiationDate={2}",
                scenario, caseTypeValue,
                caseInitiationDate.HasValue ? caseInitiationDate.Value.ToString("yyyy-MM-dd") : "(not set)");

            // Find milestone tasks
            var appMilestone = FindMilestoneByName(projectId, MILESTONE_APPLICATION_DATE);
            DateTime? appDateFinish = appMilestone?.GetAttributeValue<DateTime?>(TASK_FINISH);

            DateTime? initDateFinish = null;
            if (scenario == 2)
            {
                var initMilestone = FindMilestoneByName(projectId, MILESTONE_INITIATION_DATE);
                initDateFinish = initMilestone?.GetAttributeValue<DateTime?>(TASK_FINISH);
            }

            if (appDateFinish == null)
            {
                _trace?.Trace("DayCountService: WARNING — '{0}' milestone not found or has no finish date.", MILESTONE_APPLICATION_DATE);
            }

            return new CalcContext
            {
                ProjectId = projectId,
                Scenario = scenario,
                ApplicationDateFinish = appDateFinish?.Date,
                InitiationDateFinish = initDateFinish?.Date,
                HasInitiationDate = caseInitiationDate.HasValue
            };
        }

        /// <summary>
        /// Calculates the day count for a single task given the pre-built context.
        /// Returns null if calculation is not possible (e.g. no reference date).
        /// </summary>
        private int? CalculateDayCount(Entity task, CalcContext ctx)
        {
            string taskName = task.GetAttributeValue<string>(TASK_NAME) ?? "";
            DateTime? taskFinish = task.GetAttributeValue<DateTime?>(TASK_FINISH);

            // --- Handle special milestone tasks ---

            // "Application date" milestone
            if (IsApplicationDateMilestone(taskName))
            {
                if (ctx.Scenario == 1)
                    return 0; // Reference point for Scenario 1

                // Scenario 2: -20 initially, or recalculate relative to initiation date if set
                if (ctx.HasInitiationDate && ctx.InitiationDateFinish.HasValue && taskFinish.HasValue)
                    return (int)(taskFinish.Value.Date - ctx.InitiationDateFinish.Value).TotalDays;

                return APPLICATION_DATE_INITIAL_DAY_COUNT;
            }

            // "Initiation date" milestone
            if (IsInitiationDateMilestone(taskName))
            {
                return 0; // Always 0 — it IS the reference point when present
            }

            // --- Regular tasks ---

            if (!taskFinish.HasValue)
            {
                _trace?.Trace("DayCountService: Task '{0}' has no finish date, skipping.", taskName);
                return null;
            }

            // Determine the reference date:
            // If Scenario 2 and initiation date is set → use Initiation date milestone finish
            // Otherwise → use Application date milestone finish
            DateTime? referenceDate;
            if (ctx.Scenario == 2 && ctx.HasInitiationDate && ctx.InitiationDateFinish.HasValue)
                referenceDate = ctx.InitiationDateFinish;
            else
                referenceDate = ctx.ApplicationDateFinish;

            if (!referenceDate.HasValue)
            {
                _trace?.Trace("DayCountService: No reference date available for task '{0}', skipping.", taskName);
                return null;
            }

            return (int)(taskFinish.Value.Date - referenceDate.Value).TotalDays;
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
            return results.Entities.FirstOrDefault();
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
                return results.Entities[0].Id;

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
