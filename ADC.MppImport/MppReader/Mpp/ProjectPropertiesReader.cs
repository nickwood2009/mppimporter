using System;
using ADC.MppImport.MppReader.Model;

namespace ADC.MppImport.MppReader.Mpp
{
    /// <summary>
    /// Reads project properties from Props data.
    /// Ported from org.mpxj.mpp.ProjectPropertiesReader
    /// </summary>
    internal static class ProjectPropertiesReader
    {
        public static void Process(ProjectFile file, Props projectProps)
        {
            var properties = file.ProjectProperties;

            properties.StartDate = projectProps.GetTimestamp(Props.PROJECT_START_DATE);
            properties.FinishDate = projectProps.GetTimestamp(Props.PROJECT_FINISH_DATE);
            properties.ScheduleFrom = ScheduleFromHelper.GetInstance(projectProps.GetShort(Props.SCHEDULE_FROM));
            properties.DefaultCalendarName = projectProps.GetUnicodeString(Props.DEFAULT_CALENDAR_NAME);
            properties.DefaultStartTime = projectProps.GetTime(Props.START_TIME);
            properties.DefaultEndTime = projectProps.GetTime(Props.END_TIME);
            properties.StatusDate = projectProps.GetTimestamp(Props.STATUS_DATE);

            properties.CurrencySymbol = projectProps.GetUnicodeString(Props.CURRENCY_SYMBOL);
            properties.SymbolPosition = MppUtility.GetSymbolPosition(projectProps.GetShort(Props.CURRENCY_PLACEMENT));
            properties.CurrencyDigits = projectProps.GetShort(Props.CURRENCY_DIGITS);
            properties.CurrencyCode = projectProps.GetUnicodeString(Props.CURRENCY_CODE);

            properties.DefaultDurationUnits = MppUtility.GetDurationTimeUnits(projectProps.GetShort(Props.DURATION_UNITS));
            properties.DefaultWorkUnits = MppUtility.GetWorkTimeUnits(projectProps.GetShort(Props.WORK_UNITS));

            properties.MinutesPerDay = projectProps.GetInt(Props.MINUTES_PER_DAY);
            properties.MinutesPerWeek = projectProps.GetInt(Props.MINUTES_PER_WEEK);
            properties.DaysPerMonth = projectProps.GetInt(Props.DAYS_PER_MONTH) / 4800.0;

            properties.DefaultStandardRate = new Rate(projectProps.GetDouble(Props.STANDARD_RATE) / 100, TimeUnit.Hours);
            properties.DefaultOvertimeRate = new Rate(projectProps.GetDouble(Props.OVERTIME_RATE) / 100, TimeUnit.Hours);

            properties.UpdatingTaskStatusUpdatesResourceStatus = projectProps.GetBoolean(Props.TASK_UPDATES_RESOURCE);
            properties.SplitInProgressTasks = projectProps.GetBoolean(Props.SPLIT_TASKS);

            int criticalSlackLimit = projectProps.GetInt(Props.CRITICAL_SLACK_LIMIT);
            properties.CriticalSlackLimit = Duration.GetInstance(criticalSlackLimit, TimeUnit.Days);

            properties.DefaultTaskType = TaskTypeHelper.GetInstance(projectProps.GetShort(Props.DEFAULT_TASK_TYPE));

            int weekStartDay = projectProps.GetInt(Props.WEEK_START_DAY);
            if (weekStartDay >= 1 && weekStartDay <= 7)
                properties.WeekStartDay = (DayOfWeek)(weekStartDay - 1);

            properties.FiscalYearStartMonth = projectProps.GetShort(Props.FISCAL_YEAR_START_MONTH);
            properties.FiscalYearStart = projectProps.GetBoolean(Props.FISCAL_YEAR_START);
            properties.HonorConstraints = projectProps.GetBoolean(Props.HONOR_CONSTRAINTS);
            properties.EditableActualCosts = projectProps.GetBoolean(Props.EDITABLE_ACTUAL_COSTS);

            properties.GUID = projectProps.GetUUID(Props.GUID);
            properties.HyperlinkBase = projectProps.GetUnicodeString(Props.HYPERLINK_BASE);

            properties.NewTasksAreManual = projectProps.GetBoolean(Props.NEW_TASKS_ARE_MANUAL);
            properties.MultipleCriticalPaths = projectProps.GetBoolean(Props.MULTIPLE_CRITICAL_PATHS);

            // Baseline dates
            properties.BaselineDate = projectProps.GetTimestamp(Props.BASELINE_DATE);
            properties.Baseline1Date = projectProps.GetTimestamp(Props.BASELINE1_DATE);
            properties.Baseline2Date = projectProps.GetTimestamp(Props.BASELINE2_DATE);
            properties.Baseline3Date = projectProps.GetTimestamp(Props.BASELINE3_DATE);
            properties.Baseline4Date = projectProps.GetTimestamp(Props.BASELINE4_DATE);
            properties.Baseline5Date = projectProps.GetTimestamp(Props.BASELINE5_DATE);
            properties.Baseline6Date = projectProps.GetTimestamp(Props.BASELINE6_DATE);
            properties.Baseline7Date = projectProps.GetTimestamp(Props.BASELINE7_DATE);
            properties.Baseline8Date = projectProps.GetTimestamp(Props.BASELINE8_DATE);
            properties.Baseline9Date = projectProps.GetTimestamp(Props.BASELINE9_DATE);
            properties.Baseline10Date = projectProps.GetTimestamp(Props.BASELINE10_DATE);
        }
    }
}
