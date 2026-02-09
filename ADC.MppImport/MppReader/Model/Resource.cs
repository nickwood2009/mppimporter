using System;
using System.Collections.Generic;

namespace ADC.MppImport.MppReader.Model
{
    /// <summary>
    /// Represents a project resource.
    /// Ported from org.mpxj.Resource
    /// </summary>
    public class Resource
    {
        public int? UniqueID { get; set; }
        public int? ID { get; set; }
        public string Name { get; set; }
        public string Initials { get; set; }
        public string Group { get; set; }
        public string Code { get; set; }
        public string EmailAddress { get; set; }
        public ResourceType? Type { get; set; }
        public bool? IsNull { get; set; }
        public double? MaxUnits { get; set; }
        public Rate StandardRate { get; set; }
        public Rate OvertimeRate { get; set; }
        public double? CostPerUse { get; set; }
        public AccrueType? AccrueAt { get; set; }
        public Duration Work { get; set; }
        public Duration ActualWork { get; set; }
        public Duration RemainingWork { get; set; }
        public Duration OvertimeWork { get; set; }
        public double? Cost { get; set; }
        public double? ActualCost { get; set; }
        public double? RemainingCost { get; set; }
        public double? OvertimeCost { get; set; }
        public double? PercentWorkComplete { get; set; }
        public double? PeakUnits { get; set; }
        public string MaterialLabel { get; set; }
        public string NtAccount { get; set; }
        public string ActiveDirectoryGUID { get; set; }
        public int? CalendarUniqueID { get; set; }
        public string Notes { get; set; }
        public Guid? GUID { get; set; }
        public DateTime? CreationDate { get; set; }
        public bool? Active { get; set; }

        // Baseline fields
        public Duration BaselineWork { get; set; }
        public double? BaselineCost { get; set; }

        // Cost rate tables
        public List<CostRateTableEntry> CostRateTable1 { get; } = new List<CostRateTableEntry>();
        public List<CostRateTableEntry> CostRateTable2 { get; } = new List<CostRateTableEntry>();
        public List<CostRateTableEntry> CostRateTable3 { get; } = new List<CostRateTableEntry>();
        public List<CostRateTableEntry> CostRateTable4 { get; } = new List<CostRateTableEntry>();
        public List<CostRateTableEntry> CostRateTable5 { get; } = new List<CostRateTableEntry>();

        // Availability
        public List<AvailabilityEntry> Availability { get; } = new List<AvailabilityEntry>();

        // Custom fields
        public Dictionary<string, object> CustomFields { get; } = new Dictionary<string, object>();

        // Assignments
        public List<ResourceAssignment> Assignments { get; } = new List<ResourceAssignment>();

        // Calendar
        public ProjectCalendar Calendar { get; set; }

        public override string ToString()
        {
            return $"Resource[ID={ID}, UniqueID={UniqueID}, Name={Name}]";
        }
    }

    public class CostRateTableEntry
    {
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public Rate StandardRate { get; set; }
        public double? StandardRateAmount { get; set; }
        public Rate OvertimeRate { get; set; }
        public double? OvertimeRateAmount { get; set; }
        public double? CostPerUse { get; set; }
    }

    public class AvailabilityEntry
    {
        public DateTime? Start { get; set; }
        public DateTime? End { get; set; }
        public double? Units { get; set; }
    }
}
