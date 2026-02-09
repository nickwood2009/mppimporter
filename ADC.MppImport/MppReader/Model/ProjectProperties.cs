using System;

namespace ADC.MppImport.MppReader.Model
{
    /// <summary>
    /// Container for project-level properties.
    /// Ported from org.mpxj.ProjectProperties
    /// </summary>
    public class ProjectProperties
    {
        public Guid? GUID { get; set; }
        public string ProjectTitle { get; set; }
        public string Subject { get; set; }
        public string Author { get; set; }
        public string Keywords { get; set; }
        public string Comments { get; set; }
        public string Manager { get; set; }
        public string Company { get; set; }
        public string Category { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? FinishDate { get; set; }
        public DateTime? StatusDate { get; set; }
        public DateTime? CurrentDate { get; set; }
        public DateTime? CreationDate { get; set; }
        public DateTime? LastSaved { get; set; }
        public ScheduleFrom ScheduleFrom { get; set; }
        public TimeSpan? DefaultStartTime { get; set; }
        public TimeSpan? DefaultEndTime { get; set; }
        public TimeUnit DefaultDurationUnits { get; set; } = TimeUnit.Days;
        public TimeUnit DefaultWorkUnits { get; set; } = TimeUnit.Hours;
        public int? MinutesPerDay { get; set; } = 480;
        public int? MinutesPerWeek { get; set; } = 2400;
        public double? DaysPerMonth { get; set; } = 20.0;
        public Rate DefaultStandardRate { get; set; }
        public Rate DefaultOvertimeRate { get; set; }
        public bool SplitInProgressTasks { get; set; }
        public bool UpdatingTaskStatusUpdatesResourceStatus { get; set; }
        public Duration CriticalSlackLimit { get; set; }
        public int? CurrencyDigits { get; set; }
        public string CurrencySymbol { get; set; }
        public string CurrencyCode { get; set; }
        public CurrencySymbolPosition SymbolPosition { get; set; }
        public TaskType DefaultTaskType { get; set; }
        public DayOfWeek WeekStartDay { get; set; } = DayOfWeek.Monday;
        public int? FiscalYearStartMonth { get; set; }
        public bool FiscalYearStart { get; set; }
        public bool HonorConstraints { get; set; }
        public bool EditableActualCosts { get; set; }
        public string DefaultCalendarName { get; set; }
        public string HyperlinkBase { get; set; }
        public bool ShowProjectSummaryTask { get; set; }
        public bool NewTasksAreManual { get; set; }
        public bool MultipleCriticalPaths { get; set; }
        public bool AutoFilter { get; set; }
        public string ProjectFilePath { get; set; }
        public int? MppFileType { get; set; }
        public string FullApplicationName { get; set; }
        public int? ApplicationVersion { get; set; }
        public string FileApplication { get; set; }
        public string FileType { get; set; }

        // Baseline dates
        public DateTime? BaselineDate { get; set; }
        public DateTime? Baseline1Date { get; set; }
        public DateTime? Baseline2Date { get; set; }
        public DateTime? Baseline3Date { get; set; }
        public DateTime? Baseline4Date { get; set; }
        public DateTime? Baseline5Date { get; set; }
        public DateTime? Baseline6Date { get; set; }
        public DateTime? Baseline7Date { get; set; }
        public DateTime? Baseline8Date { get; set; }
        public DateTime? Baseline9Date { get; set; }
        public DateTime? Baseline10Date { get; set; }
    }
}
