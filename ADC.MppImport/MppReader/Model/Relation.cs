namespace ADC.MppImport.MppReader.Model
{
    /// <summary>
    /// Represents a task dependency/relation.
    /// Ported from org.mpxj.Relation
    /// </summary>
    public class Relation
    {
        public int? UniqueID { get; set; }
        public int SourceTaskUniqueID { get; set; }
        public int TargetTaskUniqueID { get; set; }
        public RelationType Type { get; set; }
        public Duration Lag { get; set; }

        // Resolved references (populated after all tasks are read)
        public Task SourceTask { get; set; }
        public Task TargetTask { get; set; }

        public override string ToString()
        {
            return $"Relation: {SourceTaskUniqueID} -> {TargetTaskUniqueID} ({Type}, Lag={Lag})";
        }
    }
}
