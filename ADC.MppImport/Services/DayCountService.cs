using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ADC.MppImport.Services
{
    /// <summary>
    /// Shared business logic for calculating "Day Count" on project tasks.
    ///
    /// Day Count = datediff(task Finish date, reference date from linked ADC case).
    ///
    /// Reference date rules:
    ///   - Scenario 1 (Duty Assessment, Accelerated Review, Exemption):
    ///       → Always uses Original Lodgement Date from the case.
    ///   - Scenario 2 (Dumping Inv, Subsidy Inv, D&amp;S Inv, Review of Measures,
    ///                 Continuation Inquiry, Anti-circumvention):
    ///       → Uses Initiation Date from the case when set.
    ///       → Falls back to Original Lodgement Date if Initiation Date not yet set.
    ///
    /// Recalculation triggers:
    ///   1. Task finish date changes → recalc that single task.
    ///   2. Initiation Date or Original Lodgement Date changes on the case → recalc all tasks.
    /// </summary>
    public class DayCountService
    {
        private readonly IOrganizationService _service;
        private readonly ITracingService _trace;

        // =====================================================================
        // SCHEMA NAMES
        // =====================================================================

        // msdyn_projecttask fields
        private const string TASK_ENTITY = "msdyn_projecttask";
        private const string TASK_FINISH = "msdyn_scheduledend";
        private const string TASK_NAME = "msdyn_subject";
        private const string TASK_PROJECT = "msdyn_project";
        private const string TASK_DAY_COUNT = "adc_daycount1";

        // adc_case fields
        private const string CASE_ENTITY = "adc_case";
        private const string CASE_TYPE = "adc_casetype";
        private const string CASE_INITIATION_DATE = "adc_dateofinitiation";
        private const string CASE_ORIGINAL_LODGEMENT_DATE = "adc_originallodgementdate";

        // msdyn_project fields
        private const string PROJECT_ENTITY = "msdyn_project";
        private const string PROJECT_CASE_LINK = "adc_parentadccase";

        // Case type option set values — Scenario 1 (no initiation date — use Original Lodgement Date)
        private static readonly HashSet<int> Scenario1CaseTypes = new HashSet<int>
        {
            756360004, // Duty Assessment
            756360005, // Accelerated Review
            756360008  // Exemption
        };

        // Case type option set values — Scenario 2 (has initiation date when set)
        private static readonly HashSet<int> Scenario2CaseTypes = new HashSet<int>
        {
            756360000, // Dumping Investigation
            756360001, // Subsidy Investigation
            756360002, // Dumping and Subsidy Investigation
            756360003, // Review of Measures
            756360006, // Continuation Inquiry
            756360007  // Anti-Circumvention
        };

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

            double? dayCount = CalculateDayCount(task, context);
            double dayCountValue = dayCount ?? 0.0;
            _trace?.Trace("DayCountService: CalculateDayCount for '{0}' returned {1} (storing {2})",
                taskName, dayCount.HasValue ? dayCount.Value.ToString() : "(null)", dayCountValue);

            double existingDayCount = task.GetAttributeValue<double>(TASK_DAY_COUNT);
            if (existingDayCount != dayCountValue)
            {
                var update = new Entity(TASK_ENTITY, taskId);
                update[TASK_DAY_COUNT] = dayCountValue;
                _service.Update(update);
                _trace?.Trace("DayCountService: Updated '{0}' day count: {1} → {2}",
                    taskName, existingDayCount, dayCountValue);
            }
            else
            {
                _trace?.Trace("DayCountService: Task '{0}' day count unchanged at {1}.",
                    taskName, existingDayCount);
            }
        }

        /// <summary>
        /// Recalculates day count for ALL project tasks in a project.
        /// Called after clone completes, or when a case date changes.
        /// </summary>
        public void RecalcAllTasks(Guid projectId)
        {
            _trace?.Trace("DayCountService.RecalcAllTasks: projectId={0}", projectId);

            var context = BuildContext(projectId);
            if (context == null)
            {
                _trace?.Trace("DayCountService: BuildContext returned null for project {0}, aborting.", projectId);
                return;
            }

            var tasks = RetrieveAllTasks(projectId);
            _trace?.Trace("DayCountService: Found {0} tasks to recalculate.", tasks.Count);

            if (tasks.Count == 0)
            {
                _trace?.Trace("DayCountService: No tasks found for project {0}.", projectId);
                return;
            }

            int updated = 0, skipped = 0, unchanged = 0, errors = 0;
            foreach (var task in tasks)
            {
                string tName = task.GetAttributeValue<string>(TASK_NAME) ?? "(no name)";
                double? dayCount = CalculateDayCount(task, context);

                double dayCountValue = dayCount ?? 0.0;
                double existingDayCount = task.GetAttributeValue<double>(TASK_DAY_COUNT);

                if (!dayCount.HasValue)
                    skipped++;

                if (existingDayCount != dayCountValue)
                {
                    try
                    {
                        var update = new Entity(TASK_ENTITY, task.Id);
                        update[TASK_DAY_COUNT] = dayCountValue;
                        _service.Update(update);
                        _trace?.Trace("DayCountService:   [{0}] '{1}' — {2} → {3}",
                            updated + 1, tName, existingDayCount, dayCountValue);
                        updated++;
                    }
                    catch (Exception ex)
                    {
                        errors++;
                        _trace?.Trace("DayCountService:   ERROR updating '{0}' (id={1}): {2}",
                            tName, task.Id, ex.Message);
                    }
                }
                else
                {
                    unchanged++;
                }
            }

            _trace?.Trace("DayCountService: RecalcAllTasks complete — {0} updated, {1} unchanged, {2} skipped (no finish date), {3} errors.",
                updated, unchanged, skipped, errors);
        }

        /// <summary>
        /// Called when Initiation Date or Original Lodgement Date changes on an adc_case.
        /// Finds linked project and recalculates all tasks.
        /// </summary>
        public void OnCaseDateChanged(Guid caseId)
        {
            _trace?.Trace("DayCountService.OnCaseDateChanged: caseId={0}", caseId);

            Guid? projectId = FindProjectForCase(caseId);
            if (!projectId.HasValue)
            {
                _trace?.Trace("DayCountService: No project linked to case {0}, skipping.", caseId);
                return;
            }

            RecalcAllTasks(projectId.Value);
        }

        // =====================================================================
        // CORE CALCULATION
        // =====================================================================

        /// <summary>
        /// Pre-fetched context: the reference date from the linked case.
        /// </summary>
        private class CalcContext
        {
            public Guid ProjectId;
            public int Scenario;               // 1 or 2
            public DateTime? ReferenceDate;    // The date to diff against
            public string ReferenceDateSource; // For tracing: which field was used
        }

        /// <summary>
        /// Builds the calculation context for a project:
        ///   1. Finds linked case
        ///   2. Determines scenario from case type
        ///   3. Picks the correct reference date from the case
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
                _trace?.Trace("DayCountService: Project {0} has no linked case — skipping.", projectId);
                return null;
            }

            // Get case fields
            var caseRecord = _service.Retrieve(CASE_ENTITY, caseRef.Id,
                new ColumnSet(CASE_TYPE, CASE_INITIATION_DATE, CASE_ORIGINAL_LODGEMENT_DATE));

            var caseTypeOsv = caseRecord.GetAttributeValue<OptionSetValue>(CASE_TYPE);
            int caseTypeValue = caseTypeOsv != null ? caseTypeOsv.Value : -1;

            DateTime? initiationDate = caseRecord.GetAttributeValue<DateTime?>(CASE_INITIATION_DATE);
            DateTime? originalLodgementDate = caseRecord.GetAttributeValue<DateTime?>(CASE_ORIGINAL_LODGEMENT_DATE);

            _trace?.Trace("DayCountService.BuildContext: Case={0}, CaseType={1}, InitiationDate={2}, OriginalLodgementDate={3}",
                caseRef.Id, caseTypeValue,
                initiationDate.HasValue ? initiationDate.Value.ToString("yyyy-MM-dd") : "(null)",
                originalLodgementDate.HasValue ? originalLodgementDate.Value.ToString("yyyy-MM-dd") : "(null)");

            // Determine scenario
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

            // Determine reference date:
            //   Scenario 2 with initiation date set → Initiation Date
            //   Everything else → Original Lodgement Date
            DateTime? referenceDate;
            string refSource;

            if (scenario == 2 && initiationDate.HasValue)
            {
                referenceDate = initiationDate.Value.Date;
                refSource = CASE_INITIATION_DATE;
            }
            else
            {
                referenceDate = originalLodgementDate.HasValue ? originalLodgementDate.Value.Date : (DateTime?)null;
                refSource = CASE_ORIGINAL_LODGEMENT_DATE;
            }

            _trace?.Trace("DayCountService.BuildContext: RESULT — Scenario={0}, ReferenceDate={1} (from {2})",
                scenario,
                referenceDate.HasValue ? referenceDate.Value.ToString("yyyy-MM-dd") : "(null)",
                refSource);

            if (!referenceDate.HasValue)
            {
                _trace?.Trace("DayCountService: WARNING — reference date is null ({0}), day counts cannot be calculated.", refSource);
            }

            return new CalcContext
            {
                ProjectId = projectId,
                Scenario = scenario,
                ReferenceDate = referenceDate,
                ReferenceDateSource = refSource
            };
        }

        /// <summary>
        /// Calculates day count for a single task: datediff(task finish, reference date).
        /// Returns null if task has no finish date or no reference date is available.
        /// </summary>
        private double? CalculateDayCount(Entity task, CalcContext ctx)
        {
            string taskName = task.GetAttributeValue<string>(TASK_NAME) ?? "";
            DateTime? taskFinish = task.GetAttributeValue<DateTime?>(TASK_FINISH);

            if (!taskFinish.HasValue)
            {
                _trace?.Trace("DayCountService.Calc: '{0}' has no finish date, skipping.", taskName);
                return null;
            }

            if (!ctx.ReferenceDate.HasValue)
            {
                _trace?.Trace("DayCountService.Calc: No reference date for '{0}', skipping.", taskName);
                return null;
            }

            double result = (taskFinish.Value.Date - ctx.ReferenceDate.Value).TotalDays;
            _trace?.Trace("DayCountService.Calc: '{0}' = {1:yyyy-MM-dd} - {2:yyyy-MM-dd} ({3}) = {4}",
                taskName, taskFinish.Value.Date, ctx.ReferenceDate.Value, ctx.ReferenceDateSource, result);
            return result;
        }

        // =====================================================================
        // HELPERS
        // =====================================================================

        /// <summary>
        /// Retrieves all tasks for a project, handling paging so no tasks are missed.
        /// </summary>
        private List<Entity> RetrieveAllTasks(Guid projectId)
        {
            var allTasks = new List<Entity>();
            var query = new QueryExpression(TASK_ENTITY)
            {
                ColumnSet = new ColumnSet(TASK_NAME, TASK_FINISH, TASK_DAY_COUNT),
                Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression(TASK_PROJECT, ConditionOperator.Equal, projectId)
                    }
                },
                PageInfo = new PagingInfo
                {
                    PageNumber = 1,
                    Count = 5000
                }
            };

            while (true)
            {
                var results = _service.RetrieveMultiple(query);
                allTasks.AddRange(results.Entities);

                if (results.MoreRecords)
                {
                    query.PageInfo.PageNumber++;
                    query.PageInfo.PagingCookie = results.PagingCookie;
                    _trace?.Trace("DayCountService.RetrieveAllTasks: page {0} returned {1} tasks, more records exist — fetching next page...",
                        query.PageInfo.PageNumber - 1, results.Entities.Count);
                }
                else
                {
                    break;
                }
            }

            _trace?.Trace("DayCountService.RetrieveAllTasks: project {0} → {1} total tasks returned.", projectId, allTasks.Count);
            return allTasks;
        }

        /// <summary>
        /// Finds the project linked to a case via adc_parentadccase.
        /// </summary>
        public Guid? FindProjectForCase(Guid caseId)
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
    }
}
