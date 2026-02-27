using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace ADC.MppImport.Services
{
    /// <summary>
    /// Status values for adc_mppimportjob.adc_status option set.
    /// </summary>
    public static class ImportJobStatus
    {
        public const int Queued = 0;
        public const int CreatingTasks = 1;
        public const int WaitingForTasks = 2;
        public const int PollingGUIDs = 3;
        public const int CreatingDeps = 4;
        public const int WaitingForDeps = 5;
        public const int Completed = 6;
        public const int Failed = 7;

        public static string Label(int status)
        {
            switch (status)
            {
                case Queued: return "Queued";
                case CreatingTasks: return "CreatingTasks";
                case WaitingForTasks: return "WaitingForTasks";
                case PollingGUIDs: return "PollingGUIDs";
                case CreatingDeps: return "CreatingDeps";
                case WaitingForDeps: return "WaitingForDeps";
                case Completed: return "Completed";
                case Failed: return "Failed";
                default: return "Unknown(" + status + ")";
            }
        }
    }

    /// <summary>
    /// Schema names for the adc_mppimportjob custom table fields.
    /// </summary>
    public static class ImportJobFields
    {
        public const string EntityName = "adc_mppimportjob";
        public const string PrimaryId = "adc_mppimportjobid";
        public const string Name = "adc_name";
        public const string Project = "adc_project";
        public const string CaseTemplate = "adc_casetemplate";
        public const string Status = "adc_status";
        public const string Phase = "adc_phase";
        public const string CurrentBatch = "adc_currentbatch";
        public const string TotalBatches = "adc_totalbatches";
        public const string TotalTasks = "adc_totaltasks";
        public const string CreatedCount = "adc_createdcount";
        public const string DepsCount = "adc_depscount";
        public const string Tick = "adc_tick";
        public const string TaskDataJson = "adc_taskdatajson";
        public const string TaskIdMapJson = "adc_taskidmapjson";
        public const string BatchesJson = "adc_batchesjson";
        public const string OperationSetId = "adc_operationsetid";
        public const string ProjectStartDate = "adc_projectstartdate";
        public const string ErrorMessage = "adc_errormessage";
        public const string Case = "adc_case";
        public const string InitiatingUser = "adc_initiatinguser";
    }

    /// <summary>
    /// Icon type values for Dataverse in-app notifications (appnotification.icontype).
    /// </summary>
    public static class NotificationIconType
    {
        public const int Info = 100000000;
        public const int Success = 100000001;
        public const int Failure = 100000002;
        public const int Warning = 100000003;
    }

    /// <summary>
    /// Serializable representation of an MPP task for storage in the job record.
    /// Uses DataContract for JSON serialization compatible with .NET 4.7.2.
    /// </summary>
    [DataContract]
    public class TaskDto
    {
        [DataMember(Name = "id")]
        public int UniqueID { get; set; }

        [DataMember(Name = "name")]
        public string Name { get; set; }

        [DataMember(Name = "parentId")]
        public int? ParentUniqueID { get; set; }

        [DataMember(Name = "depth")]
        public int Depth { get; set; }

        [DataMember(Name = "isSummary")]
        public bool IsSummary { get; set; }

        [DataMember(Name = "durationHours")]
        public double? DurationHours { get; set; }

        [DataMember(Name = "effortHours")]
        public double? EffortHours { get; set; }

        [DataMember(Name = "preGenGuid")]
        public string PreGenGuid { get; set; }
    }

    /// <summary>
    /// Serializable predecessor relationship for storage in the job record.
    /// </summary>
    [DataContract]
    public class DependencyDto
    {
        [DataMember(Name = "predId")]
        public int PredecessorUniqueID { get; set; }

        [DataMember(Name = "succId")]
        public int SuccessorUniqueID { get; set; }

        /// <summary>
        /// D365 link type option set value (192350000=FS, 192350001=SS, 192350002=FF, 192350003=SF)
        /// </summary>
        [DataMember(Name = "linkType")]
        public int LinkType { get; set; }
    }

    /// <summary>
    /// A batch of task UniqueIDs to be created in a single PSS operation set.
    /// Each batch contains complete subtrees so parent links stay within the batch.
    /// </summary>
    [DataContract]
    public class TaskBatch
    {
        [DataMember(Name = "index")]
        public int Index { get; set; }

        [DataMember(Name = "taskIds")]
        public List<int> TaskUniqueIDs { get; set; }

        public TaskBatch()
        {
            TaskUniqueIDs = new List<int>();
        }
    }

    /// <summary>
    /// Root object serialized to adc_taskdatajson. Contains all parsed MPP data
    /// needed to process the import across multiple plugin executions.
    /// </summary>
    [DataContract]
    public class ImportJobPayload
    {
        [DataMember(Name = "tasks")]
        public List<TaskDto> Tasks { get; set; }

        [DataMember(Name = "deps")]
        public List<DependencyDto> Dependencies { get; set; }

        [DataMember(Name = "batches")]
        public List<TaskBatch> Batches { get; set; }

        /// <summary>
        /// Pre-generated GUID map: UniqueID -> GUID string.
        /// Populated during initialization, used for parent link correlation within batches.
        /// </summary>
        [DataMember(Name = "taskIdMap")]
        public Dictionary<int, string> TaskIdMap { get; set; }

        /// <summary>
        /// Actual CRM GUID map: UniqueID -> CRM GUID string.
        /// Populated after PollingGUIDs phase, used for dependency creation.
        /// </summary>
        [DataMember(Name = "actualIdMap")]
        public Dictionary<int, string> ActualIdMap { get; set; }

        [DataMember(Name = "bucketId")]
        public string BucketId { get; set; }

        public ImportJobPayload()
        {
            Tasks = new List<TaskDto>();
            Dependencies = new List<DependencyDto>();
            Batches = new List<TaskBatch>();
            TaskIdMap = new Dictionary<int, string>();
            ActualIdMap = new Dictionary<int, string>();
        }
    }
}
