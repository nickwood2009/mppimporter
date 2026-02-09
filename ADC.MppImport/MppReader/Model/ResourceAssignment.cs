using System;

namespace ADC.MppImport.MppReader.Model
{
    /// <summary>
    /// Represents a resource assignment to a task.
    /// Ported from org.mpxj.ResourceAssignment
    /// </summary>
    public class ResourceAssignment
    {
        public int? UniqueID { get; set; }
        public int? TaskUniqueID { get; set; }
        public int? ResourceUniqueID { get; set; }
        public string TaskName { get; set; }
        public string ResourceName { get; set; }
        public Duration Work { get; set; }
        public Duration ActualWork { get; set; }
        public Duration RemainingWork { get; set; }
        public Duration OvertimeWork { get; set; }
        public double? Units { get; set; }
        public double? Cost { get; set; }
        public double? ActualCost { get; set; }
        public double? RemainingCost { get; set; }
        public DateTime? Start { get; set; }
        public DateTime? Finish { get; set; }
        public DateTime? ActualStart { get; set; }
        public DateTime? ActualFinish { get; set; }
        public Duration Delay { get; set; }
        public double? PercentWorkComplete { get; set; }
        public Guid? GUID { get; set; }

        // Baseline fields
        public Duration BaselineWork { get; set; }
        public double? BaselineCost { get; set; }
        public DateTime? BaselineStart { get; set; }
        public DateTime? BaselineFinish { get; set; }

        // Resolved references
        public Task Task { get; set; }
        public Resource Resource { get; set; }

        public override string ToString()
        {
            return $"Assignment: Task={TaskUniqueID}, Resource={ResourceUniqueID}, Work={Work}";
        }
    }
}
