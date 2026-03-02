using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace ADC.MppImport.Services
{
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

    public static class NotificationIconType
    {
        public const int Info = 100000000;
        public const int Success = 100000001;
        public const int Failure = 100000002;
        public const int Warning = 100000003;
    }

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

    [DataContract]
    public class DependencyDto
    {
        [DataMember(Name = "predId")]
        public int PredecessorUniqueID { get; set; }

        [DataMember(Name = "succId")]
        public int SuccessorUniqueID { get; set; }

        [DataMember(Name = "linkType")]
        public int LinkType { get; set; }
    }

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

    [DataContract]
    public class ImportJobPayload
    {
        [DataMember(Name = "tasks")]
        public List<TaskDto> Tasks { get; set; }

        [DataMember(Name = "deps")]
        public List<DependencyDto> Dependencies { get; set; }

        [DataMember(Name = "batches")]
        public List<TaskBatch> Batches { get; set; }

        [DataMember(Name = "taskIdMap")]
        public Dictionary<int, string> TaskIdMap { get; set; }

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
