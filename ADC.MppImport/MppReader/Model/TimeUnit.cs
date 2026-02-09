namespace ADC.MppImport.MppReader.Model
{
    /// <summary>
    /// Time unit types used for durations.
    /// Ported from org.mpxj.TimeUnit
    /// </summary>
    public enum TimeUnit
    {
        Minutes = 0,
        Hours = 1,
        Days = 2,
        Weeks = 3,
        Months = 4,
        Percent = 5,
        Years = 6,
        ElapsedMinutes = 7,
        ElapsedHours = 8,
        ElapsedDays = 9,
        ElapsedWeeks = 10,
        ElapsedMonths = 11,
        ElapsedPercent = 12,
        ElapsedYears = 13
    }

    public static class TimeUnitHelper
    {
        public static TimeUnit GetInstance(int value)
        {
            if (value >= 0 && value <= 13)
                return (TimeUnit)value;
            return TimeUnit.Days;
        }
    }
}
