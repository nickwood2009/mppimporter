using System;
using System.Collections.Generic;

namespace ADC.MppImport.MppReader.Model
{
    /// <summary>
    /// Represents a project calendar.
    /// Ported from org.mpxj.ProjectCalendar
    /// </summary>
    public class ProjectCalendar
    {
        public int? UniqueID { get; set; }
        public string Name { get; set; }
        public CalendarType Type { get; set; }
        public bool Personal { get; set; }
        public int? ParentCalendarUniqueID { get; set; }
        public ProjectCalendar ParentCalendar { get; set; }
        public Guid? GUID { get; set; }

        /// <summary>
        /// Working hours for each day of the week (Sunday=0 through Saturday=6).
        /// </summary>
        public CalendarDay[] Days { get; } = new CalendarDay[7];

        /// <summary>
        /// Calendar exceptions (holidays, special working days, etc.)
        /// </summary>
        public List<CalendarException> Exceptions { get; } = new List<CalendarException>();

        /// <summary>
        /// Derived calendars inherit from a parent but override specific settings.
        /// </summary>
        public bool IsDerived => ParentCalendarUniqueID != null;

        public ProjectCalendar()
        {
            for (int i = 0; i < 7; i++)
            {
                Days[i] = new CalendarDay { DayOfWeek = (DayOfWeek)i };
            }
        }

        public override string ToString()
        {
            return $"Calendar[UniqueID={UniqueID}, Name={Name}]";
        }
    }

    public class CalendarDay
    {
        public DayOfWeek DayOfWeek { get; set; }
        public DayType Type { get; set; } = DayType.Default;
        public List<CalendarHours> Hours { get; } = new List<CalendarHours>();
    }

    public class CalendarHours
    {
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
    }

    public class CalendarException
    {
        public string Name { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public bool Working { get; set; }
        public List<CalendarHours> Hours { get; } = new List<CalendarHours>();

        // Recurrence fields
        public RecurrenceType? RecurrenceType { get; set; }
        public int? RecurrenceFrequency { get; set; }
        public int? Occurrences { get; set; }
    }

    public enum DayType
    {
        Default = 0,
        Working = 1,
        NonWorking = 2
    }

    public enum RecurrenceType
    {
        Daily = 1,
        Weekly = 2,
        Monthly = 3,
        Yearly = 4
    }
}
