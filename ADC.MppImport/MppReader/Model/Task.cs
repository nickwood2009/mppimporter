using System;
using System.Collections.Generic;

namespace ADC.MppImport.MppReader.Model
{
    /// <summary>
    /// Represents a project task.
    /// Ported from org.mpxj.Task
    /// </summary>
    public class Task
    {
        public int? UniqueID { get; set; }
        public int? ID { get; set; }
        public string Name { get; set; }
        public string WBS { get; set; }
        public int? OutlineLevel { get; set; }
        public string OutlineNumber { get; set; }
        public DateTime? Start { get; set; }
        public DateTime? Finish { get; set; }
        public DateTime? ActualStart { get; set; }
        public DateTime? ActualFinish { get; set; }
        public Duration Duration { get; set; }
        public Duration ActualDuration { get; set; }
        public Duration RemainingDuration { get; set; }
        public double? PercentComplete { get; set; }
        public double? PercentWorkComplete { get; set; }
        public double? PhysicalPercentComplete { get; set; }
        public Duration Work { get; set; }
        public Duration ActualWork { get; set; }
        public Duration RemainingWork { get; set; }
        public Duration OvertimeWork { get; set; }
        public double? Cost { get; set; }
        public double? ActualCost { get; set; }
        public double? RemainingCost { get; set; }
        public double? FixedCost { get; set; }
        public ConstraintType? ConstraintType { get; set; }
        public DateTime? ConstraintDate { get; set; }
        public DateTime? Deadline { get; set; }
        public DateTime? EarlyStart { get; set; }
        public DateTime? EarlyFinish { get; set; }
        public DateTime? LateStart { get; set; }
        public DateTime? LateFinish { get; set; }
        public Duration FreeSlack { get; set; }
        public Duration TotalSlack { get; set; }
        public int? Priority { get; set; }
        public bool? Critical { get; set; }
        public bool? Milestone { get; set; }
        public bool? Summary { get; set; }
        public bool? ExternalProject { get; set; }
        public string ExternalTaskProject { get; set; }
        public TaskType? Type { get; set; }
        public TaskMode? TaskMode { get; set; }
        public bool? Active { get; set; }
        public string Notes { get; set; }
        public string HyperlinkAddress { get; set; }
        public string HyperlinkSubAddress { get; set; }
        public string Hyperlink { get; set; }
        public int? CalendarUniqueID { get; set; }
        public Guid? GUID { get; set; }
        public DateTime? CreateDate { get; set; }

        // Baseline fields
        public Duration BaselineDuration { get; set; }
        public DateTime? BaselineStart { get; set; }
        public DateTime? BaselineFinish { get; set; }
        public double? BaselineCost { get; set; }
        public Duration BaselineWork { get; set; }

        public Duration Baseline1Duration { get; set; }
        public DateTime? Baseline1Start { get; set; }
        public DateTime? Baseline1Finish { get; set; }
        public double? Baseline1Cost { get; set; }
        public Duration Baseline1Work { get; set; }

        public Duration Baseline2Duration { get; set; }
        public DateTime? Baseline2Start { get; set; }
        public DateTime? Baseline2Finish { get; set; }
        public double? Baseline2Cost { get; set; }
        public Duration Baseline2Work { get; set; }

        public Duration Baseline3Duration { get; set; }
        public DateTime? Baseline3Start { get; set; }
        public DateTime? Baseline3Finish { get; set; }
        public double? Baseline3Cost { get; set; }
        public Duration Baseline3Work { get; set; }

        public Duration Baseline4Duration { get; set; }
        public DateTime? Baseline4Start { get; set; }
        public DateTime? Baseline4Finish { get; set; }
        public double? Baseline4Cost { get; set; }
        public Duration Baseline4Work { get; set; }

        public Duration Baseline5Duration { get; set; }
        public DateTime? Baseline5Start { get; set; }
        public DateTime? Baseline5Finish { get; set; }
        public double? Baseline5Cost { get; set; }
        public Duration Baseline5Work { get; set; }

        public Duration Baseline6Duration { get; set; }
        public DateTime? Baseline6Start { get; set; }
        public DateTime? Baseline6Finish { get; set; }
        public double? Baseline6Cost { get; set; }
        public Duration Baseline6Work { get; set; }

        public Duration Baseline7Duration { get; set; }
        public DateTime? Baseline7Start { get; set; }
        public DateTime? Baseline7Finish { get; set; }
        public double? Baseline7Cost { get; set; }
        public Duration Baseline7Work { get; set; }

        public Duration Baseline8Duration { get; set; }
        public DateTime? Baseline8Start { get; set; }
        public DateTime? Baseline8Finish { get; set; }
        public double? Baseline8Cost { get; set; }
        public Duration Baseline8Work { get; set; }

        public Duration Baseline9Duration { get; set; }
        public DateTime? Baseline9Start { get; set; }
        public DateTime? Baseline9Finish { get; set; }
        public double? Baseline9Cost { get; set; }
        public Duration Baseline9Work { get; set; }

        public Duration Baseline10Duration { get; set; }
        public DateTime? Baseline10Start { get; set; }
        public DateTime? Baseline10Finish { get; set; }
        public double? Baseline10Cost { get; set; }
        public Duration Baseline10Work { get; set; }

        // Custom fields (Text1-30, Number1-20, Cost1-10, etc.)
        public Dictionary<string, object> CustomFields { get; } = new Dictionary<string, object>();

        // Hierarchy
        public int? ParentTaskUniqueID { get; set; }
        public Task ParentTask { get; set; }
        public List<Task> ChildTasks { get; } = new List<Task>();
        public List<Relation> Predecessors { get; } = new List<Relation>();
        public List<Relation> Successors { get; } = new List<Relation>();
        public List<ResourceAssignment> Assignments { get; } = new List<ResourceAssignment>();

        public bool HasChildTasks => ChildTasks.Count > 0;

        public override string ToString()
        {
            return $"Task[ID={ID}, UniqueID={UniqueID}, Name={Name}]";
        }
    }
}
