namespace ADC.MppImport.MppReader.Model
{
    public enum TaskType
    {
        FixedUnits = 0,
        FixedDuration = 1,
        FixedWork = 2
    }

    public enum ConstraintType
    {
        AsLateAsPossible = 0,
        AsSoonAsPossible = 1,
        MustFinishOn = 2,
        MustStartOn = 3,
        StartNoEarlierThan = 4,
        StartNoLaterThan = 5,
        FinishNoEarlierThan = 6,
        FinishNoLaterThan = 7
    }

    public enum Priority
    {
        DoNotLevel = 0,
        Lowest = 100,
        VeryLow = 200,
        Lower = 300,
        Low = 400,
        Medium = 500,
        High = 600,
        Higher = 700,
        VeryHigh = 800,
        Highest = 900,
        DoNotLevel2 = 1000
    }

    public enum ResourceType
    {
        Work = 0,
        Material = 1,
        Cost = 2
    }

    public enum TaskMode
    {
        AutoScheduled = 0,
        ManuallyScheduled = 1
    }

    public enum AccrueType
    {
        Start = 1,
        End = 2,
        Prorated = 3
    }

    public enum CurrencySymbolPosition
    {
        Before = 0,
        After = 1,
        BeforeWithSpace = 2,
        AfterWithSpace = 3
    }

    public enum ScheduleFrom
    {
        Start = 0,
        Finish = 1
    }

    public enum RelationType
    {
        FinishToStart = 0,
        FinishToFinish = 1,
        StartToFinish = 2,
        StartToStart = 3
    }

    public enum CalendarType
    {
        Global = 0,
        Resource = 1,
        Task = 2
    }

    public enum EarnedValueMethod
    {
        PercentComplete = 0,
        PhysicalPercentComplete = 1
    }

    public static class TaskTypeHelper
    {
        public static TaskType GetInstance(int value)
        {
            if (value >= 0 && value <= 2)
                return (TaskType)value;
            return TaskType.FixedUnits;
        }
    }

    public static class ScheduleFromHelper
    {
        public static ScheduleFrom GetInstance(int value)
        {
            return value == 0 ? ScheduleFrom.Start : ScheduleFrom.Finish;
        }
    }
}
