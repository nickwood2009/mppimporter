using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace ADC.MppImport.TestDataGen
{
    /// <summary>
    /// Generates large, complex MSPDI XML project files for testing.
    /// Output can be opened in MS Project and saved as .mpp.
    ///
    /// Features:
    ///   - Configurable task count (default 500)
    ///   - Deep hierarchy (up to 8 levels)
    ///   - Mixed durations (1-30 days)
    ///   - All 4 dependency types (FS, SS, FF, SF)
    ///   - Milestones
    ///   - Realistic phase/task naming
    /// </summary>
    public class MspdiGenerator
    {
        private readonly Random _rng;
        private int _nextUid = 1;
        private int _nextId = 1;
        private DateTime _projectStart;

        // Dependency types: 0=FF, 1=FS, 2=SF, 3=SS
        private const int LinkFF = 0;
        private const int LinkFS = 1;
        private const int LinkSF = 2;
        private const int LinkSS = 3;

        private static readonly string[] PhaseNames = {
            "Initiation", "Planning", "Requirements", "Analysis", "Design",
            "Development", "Integration", "Testing", "UAT", "Deployment",
            "Migration", "Training", "Compliance", "Review", "Closeout",
            "Risk Assessment", "Stakeholder Engagement", "Documentation",
            "Infrastructure", "Security", "Performance", "Data Management",
            "Quality Assurance", "Change Management", "Governance"
        };

        private static readonly string[] TaskVerbs = {
            "Prepare", "Review", "Develop", "Conduct", "Finalise",
            "Create", "Implement", "Configure", "Validate", "Approve",
            "Assess", "Document", "Analyse", "Test", "Deploy",
            "Migrate", "Integrate", "Audit", "Evaluate", "Coordinate",
            "Schedule", "Monitor", "Resolve", "Update", "Deliver"
        };

        private static readonly string[] TaskNouns = {
            "requirements", "specifications", "architecture", "design documents",
            "test cases", "test plan", "risk register", "stakeholder matrix",
            "project charter", "work breakdown structure", "scope statement",
            "quality plan", "communication plan", "resource plan", "budget",
            "milestone report", "status report", "change request", "issue log",
            "meeting minutes", "training materials", "user guide", "data model",
            "interface design", "security policy", "compliance checklist",
            "performance baseline", "acceptance criteria", "deployment plan",
            "rollback plan", "cutover plan", "operations manual"
        };

        public MspdiGenerator(int seed = 42)
        {
            _rng = new Random(seed);
        }

        /// <summary>
        /// Generates a complex project XML file.
        /// </summary>
        /// <param name="totalTasks">Approximate number of leaf + summary tasks (excludes root task 0)</param>
        /// <param name="maxDepth">Maximum outline depth (1-based, root=0)</param>
        /// <param name="depDensity">Fraction of tasks that get a dependency (0.0 - 1.0)</param>
        /// <param name="projectStart">Project start date</param>
        public void Generate(string outputPath, int totalTasks = 500, int maxDepth = 8,
            double depDensity = 0.4, DateTime? projectStart = null)
        {
            // Default to next Monday in the future
            _projectStart = projectStart ?? GetNextMonday(DateTime.Today);
            _nextUid = 0;
            _nextId = 0;

            var tasks = BuildTaskTree(totalTasks, maxDepth);
            var deps = BuildDependencies(tasks, depDensity);
            ScheduleFromDeps(tasks);

            WriteXml(tasks, deps, outputPath);
        }

        #region Task Tree Builder

        private class TaskNode
        {
            public int Uid;
            public int Id;
            public string Name;
            public int OutlineLevel;
            public string OutlineNumber;
            public bool IsSummary;
            public bool IsMilestone;
            public int DurationDays;
            public DateTime Start;
            public DateTime Finish;
            public string Wbs;
            public List<TaskNode> Children = new List<TaskNode>();
            public List<PredLink> Predecessors = new List<PredLink>();
        }

        private class PredLink
        {
            public int PredecessorUid;
            public int Type; // 0=FF, 1=FS, 2=SF, 3=SS
        }

        private List<TaskNode> BuildTaskTree(int targetCount, int maxDepth)
        {
            var allTasks = new List<TaskNode>();

            // Task 0 = project summary (root)
            var root = new TaskNode
            {
                Uid = _nextUid++,
                Id = _nextId++,
                Name = "Large Test Project",
                OutlineLevel = 0,
                OutlineNumber = "0",
                IsSummary = true,
                Wbs = "0"
            };
            allTasks.Add(root);

            // Shuffle phase names for variety
            var phases = PhaseNames.OrderBy(_ => _rng.Next()).ToList();
            int phaseIndex = 0;
            int tasksCreated = 0;
            int topLevelCounter = 0;

            while (tasksCreated < targetCount)
            {
                topLevelCounter++;
                string phaseName = phases[phaseIndex % phases.Count];
                phaseIndex++;

                // Create a phase (summary at level 1)
                int phaseTaskCount = Math.Min(
                    _rng.Next(15, 50), // 15-50 tasks per phase
                    targetCount - tasksCreated);

                var phaseTasks = BuildPhase(phaseName, topLevelCounter, 1, maxDepth, phaseTaskCount);
                allTasks.AddRange(phaseTasks);
                tasksCreated += phaseTasks.Count;
            }

            return allTasks;
        }

        private List<TaskNode> BuildPhase(string phaseName, int siblingIndex, int level, int maxDepth, int budget)
        {
            var result = new List<TaskNode>();
            if (budget <= 0) return result;

            bool hasBudgetForChildren = budget > 1 && level < maxDepth;

            // Create the summary/phase node
            var node = new TaskNode
            {
                Uid = _nextUid++,
                Id = _nextId++,
                Name = phaseName,
                OutlineLevel = level,
                OutlineNumber = siblingIndex.ToString(),
                IsSummary = hasBudgetForChildren,
                IsMilestone = !hasBudgetForChildren && _rng.NextDouble() < 0.08,
                DurationDays = hasBudgetForChildren ? 0 : _rng.Next(1, 31)
            };

            if (node.IsMilestone) node.DurationDays = 0;

            result.Add(node);

            if (!hasBudgetForChildren) return result;

            // Distribute remaining budget among children
            int remaining = budget - 1;
            int childIndex = 0;

            while (remaining > 0)
            {
                childIndex++;
                bool goDeeper = level + 1 < maxDepth && remaining > 3 && _rng.NextDouble() < 0.35;

                if (goDeeper)
                {
                    // Create a sub-summary with children
                    int subBudget = Math.Min(_rng.Next(3, 15), remaining);
                    string subName = string.Format("{0} - {1}", phaseName, GetTaskName());
                    var subTasks = BuildPhase(subName, childIndex, level + 1, maxDepth, subBudget);
                    result.AddRange(subTasks);
                    remaining -= subTasks.Count;
                }
                else
                {
                    // Create a leaf task
                    var leaf = new TaskNode
                    {
                        Uid = _nextUid++,
                        Id = _nextId++,
                        Name = GetTaskName(),
                        OutlineLevel = level + 1,
                        OutlineNumber = string.Format("{0}.{1}", siblingIndex, childIndex),
                        IsSummary = false,
                        IsMilestone = _rng.NextDouble() < 0.05,
                        DurationDays = _rng.Next(1, 21)
                    };
                    if (leaf.IsMilestone) leaf.DurationDays = 0;
                    result.Add(leaf);
                    remaining--;
                }
            }

            return result;
        }

        private string GetTaskName()
        {
            string verb = TaskVerbs[_rng.Next(TaskVerbs.Length)];
            string noun = TaskNouns[_rng.Next(TaskNouns.Length)];
            return string.Format("{0} {1}", verb, noun);
        }

        #endregion

        #region Scheduling

        /// <summary>
        /// Forward-pass schedules all leaf tasks based on their dependency constraints.
        /// Tasks with no predecessors start at the project start date.
        /// Summary task dates are computed from their children.
        /// </summary>
        private void ScheduleFromDeps(List<TaskNode> tasks)
        {
            var uidLookup = tasks.ToDictionary(t => t.Uid);
            var leafTasks = tasks.Where(t => !t.IsSummary && t.OutlineLevel > 0).ToList();
            var leafUids = new HashSet<int>(leafTasks.Select(t => t.Uid));

            // Build dependency graph (leaf-to-leaf only; skip summary deps for scheduling)
            var inDegree = new Dictionary<int, int>();
            var successorMap = new Dictionary<int, List<int>>();

            foreach (var t in leafTasks)
            {
                if (!inDegree.ContainsKey(t.Uid))
                    inDegree[t.Uid] = 0;

                foreach (var pred in t.Predecessors)
                {
                    if (!leafUids.Contains(pred.PredecessorUid)) continue;
                    inDegree[t.Uid]++;
                    if (!successorMap.ContainsKey(pred.PredecessorUid))
                        successorMap[pred.PredecessorUid] = new List<int>();
                    successorMap[pred.PredecessorUid].Add(t.Uid);
                }
            }

            // Topological sort (Kahn's algorithm)
            var queue = new Queue<int>();
            foreach (var kvp in inDegree)
                if (kvp.Value == 0) queue.Enqueue(kvp.Key);

            var sorted = new List<int>();
            while (queue.Count > 0)
            {
                int uid = queue.Dequeue();
                sorted.Add(uid);
                List<int> succs;
                if (successorMap.TryGetValue(uid, out succs))
                {
                    foreach (int succUid in succs)
                    {
                        inDegree[succUid]--;
                        if (inDegree[succUid] == 0)
                            queue.Enqueue(succUid);
                    }
                }
            }

            // Safety: add any tasks not reached by topo sort (cycle fallback)
            var sortedSet = new HashSet<int>(sorted);
            foreach (var t in leafTasks)
            {
                if (!sortedSet.Contains(t.Uid))
                    sorted.Add(t.Uid);
            }

            // Forward pass: compute earliest start/finish based on dep constraints
            foreach (int uid in sorted)
            {
                var t = uidLookup[uid];
                DateTime earliest = _projectStart;

                foreach (var pred in t.Predecessors)
                {
                    if (!leafUids.Contains(pred.PredecessorUid)) continue;
                    TaskNode predTask;
                    if (!uidLookup.TryGetValue(pred.PredecessorUid, out predTask)) continue;
                    if (predTask.Start == default(DateTime)) continue;

                    DateTime constraint;
                    switch (pred.Type)
                    {
                        case LinkFS: // Successor starts after predecessor finishes
                            constraint = predTask.Finish;
                            break;
                        case LinkSS: // Successor starts when predecessor starts
                            constraint = predTask.Start;
                            break;
                        case LinkFF: // Successor finishes when predecessor finishes
                            constraint = predTask.Start; // approximate: run in parallel
                            break;
                        default:
                            constraint = predTask.Finish;
                            break;
                    }

                    if (constraint > earliest)
                        earliest = constraint;
                }

                // Ensure working day
                while (earliest.DayOfWeek == DayOfWeek.Saturday || earliest.DayOfWeek == DayOfWeek.Sunday)
                    earliest = earliest.AddDays(1);

                t.Start = earliest;
                t.Finish = t.DurationDays == 0 ? earliest : AddWorkdays(earliest, t.DurationDays);
            }

            ComputeSummaryDates(tasks);
        }

        /// <summary>
        /// Groups tasks by their level-1 phase ancestor.
        /// Returns a list of phases, each containing all tasks (summary + leaf) in outline order.
        /// </summary>
        private List<List<TaskNode>> ExtractPhases(List<TaskNode> tasks)
        {
            var phases = new List<List<TaskNode>>();
            List<TaskNode> current = null;

            for (int i = 1; i < tasks.Count; i++)
            {
                if (tasks[i].OutlineLevel == 1)
                {
                    current = new List<TaskNode> { tasks[i] };
                    phases.Add(current);
                }
                else if (current != null)
                {
                    current.Add(tasks[i]);
                }
            }

            return phases;
        }

        private void ComputeSummaryDates(List<TaskNode> tasks)
        {
            // Work backwards through outline levels
            int maxLevel = tasks.Max(t => t.OutlineLevel);

            for (int level = maxLevel; level >= 0; level--)
            {
                foreach (var summary in tasks.Where(t => t.OutlineLevel == level && t.IsSummary))
                {
                    // Find children at level+1 that appear after this summary and before next same-level task
                    int idx = tasks.IndexOf(summary);
                    var children = new List<TaskNode>();
                    for (int i = idx + 1; i < tasks.Count; i++)
                    {
                        if (tasks[i].OutlineLevel <= level) break;
                        children.Add(tasks[i]);
                    }

                    if (children.Count > 0)
                    {
                        var dated = children.Where(c => c.Start != default(DateTime)).ToList();
                        if (dated.Count > 0)
                        {
                            summary.Start = dated.Min(c => c.Start);
                            summary.Finish = dated.Max(c => c.Finish);
                            summary.DurationDays = CountWorkdays(summary.Start, summary.Finish);
                        }
                    }
                }
            }

            // Root task
            var root = tasks[0];
            var allDated = tasks.Where(t => t.OutlineLevel > 0 && t.Start != default(DateTime)).ToList();
            if (allDated.Count > 0)
            {
                root.Start = allDated.Min(t => t.Start);
                root.Finish = allDated.Max(t => t.Finish);
                root.DurationDays = CountWorkdays(root.Start, root.Finish);
            }
        }

        private static DateTime GetNextMonday(DateTime from)
        {
            DateTime d = from.Date.AddHours(8); // 8 AM
            int daysUntilMonday = ((int)DayOfWeek.Monday - (int)d.DayOfWeek + 7) % 7;
            if (daysUntilMonday == 0) daysUntilMonday = 7; // always next week
            return d.AddDays(daysUntilMonday);
        }

        private DateTime AddWorkdays(DateTime start, int days)
        {
            DateTime d = start;
            int added = 0;
            while (added < days)
            {
                d = d.AddDays(1);
                if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday)
                    added++;
            }
            return d;
        }

        private int CountWorkdays(DateTime start, DateTime end)
        {
            int count = 0;
            DateTime d = start;
            while (d < end)
            {
                d = d.AddDays(1);
                if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday)
                    count++;
            }
            return Math.Max(count, 1);
        }

        #endregion

        #region Dependencies

        /// <summary>
        /// Builds realistic, structured dependencies:
        ///   1. Within each phase: chain ~85% of consecutive leaf tasks (FS/SS)
        ///   2. Between phases: link last leaf of phase N to first leaf of N+1 (FS)
        ///   3. Cross-phase extras: ~8% of leaves get a dep from a random earlier phase
        ///   4. Summary-task deps: ~15% of summaries get a dep to exercise the importer's
        ///      entry/exit leaf transformation logic
        /// </summary>
        private List<PredLink> BuildDependencies(List<TaskNode> tasks, double density)
        {
            var allDeps = new List<PredLink>();
            var usedPairs = new HashSet<string>();
            var phases = ExtractPhases(tasks);
            var summaries = tasks.Where(t => t.IsSummary && t.OutlineLevel >= 1).ToList();

            // 1. Within each phase: chain consecutive leaf tasks
            foreach (var phase in phases)
            {
                var leaves = phase.Where(t => !t.IsSummary).ToList();
                for (int i = 1; i < leaves.Count; i++)
                {
                    // density controls how many get chained (default 0.4 → ~85%)
                    if (_rng.NextDouble() > Math.Min(0.85, 0.5 + density)) continue;

                    // Mostly FS, some SS for parallel work
                    int linkType;
                    double roll = _rng.NextDouble();
                    if (roll < 0.70) linkType = LinkFS;
                    else if (roll < 0.90) linkType = LinkSS;
                    else linkType = LinkFF;

                    AddDep(leaves[i], leaves[i - 1].Uid, linkType, usedPairs, allDeps);
                }
            }

            // 2. Cross-phase sequential links (last leaf → first leaf of next phase)
            for (int p = 0; p < phases.Count - 1; p++)
            {
                var prevLeaves = phases[p].Where(t => !t.IsSummary).ToList();
                var nextLeaves = phases[p + 1].Where(t => !t.IsSummary).ToList();
                if (prevLeaves.Count > 0 && nextLeaves.Count > 0)
                    AddDep(nextLeaves[0], prevLeaves[prevLeaves.Count - 1].Uid,
                        LinkFS, usedPairs, allDeps);
            }

            // 3. Extra cross-phase deps for variety
            for (int p = 1; p < phases.Count; p++)
            {
                var currentLeaves = phases[p].Where(t => !t.IsSummary).ToList();
                var earlierLeaves = new List<TaskNode>();
                for (int ep = 0; ep < p; ep++)
                    earlierLeaves.AddRange(phases[ep].Where(t => !t.IsSummary));

                if (earlierLeaves.Count == 0) continue;

                foreach (var leaf in currentLeaves)
                {
                    if (_rng.NextDouble() > 0.08) continue; // ~8% get a cross-phase dep

                    var pred = earlierLeaves[_rng.Next(earlierLeaves.Count)];
                    int linkType = _rng.NextDouble() < 0.80 ? LinkFS : LinkFF;
                    AddDep(leaf, pred.Uid, linkType, usedPairs, allDeps);
                }
            }

            // 4. Summary-task deps (exercises the importer's entry/exit leaf transformation)
            if (summaries.Count > 1 && phases.Count > 2)
            {
                int summaryDepTarget = Math.Max(2, (int)(summaries.Count * 0.15));

                // Successor is a summary (importer must push dep to entry leaves)
                for (int s = 0; s < summaryDepTarget; s++)
                {
                    int succPhaseIdx = _rng.Next(1, phases.Count);
                    var succSummaries = phases[succPhaseIdx].Where(t => t.IsSummary).ToList();
                    if (succSummaries.Count == 0) continue;

                    var earlierLeaves = new List<TaskNode>();
                    for (int ep = 0; ep < succPhaseIdx; ep++)
                        earlierLeaves.AddRange(phases[ep].Where(t => !t.IsSummary));
                    if (earlierLeaves.Count == 0) continue;

                    var predLeaf = earlierLeaves[_rng.Next(earlierLeaves.Count)];
                    var succSummary = succSummaries[_rng.Next(succSummaries.Count)];
                    AddDep(succSummary, predLeaf.Uid, LinkFS, usedPairs, allDeps);
                }

                // Predecessor is a summary (importer must push dep from exit leaves)
                for (int s = 0; s < summaryDepTarget; s++)
                {
                    int predPhaseIdx = _rng.Next(0, Math.Max(1, phases.Count - 1));
                    var predSummaries = phases[predPhaseIdx].Where(t => t.IsSummary).ToList();
                    if (predSummaries.Count == 0) continue;

                    var laterLeaves = new List<TaskNode>();
                    for (int lp = predPhaseIdx + 1; lp < phases.Count; lp++)
                        laterLeaves.AddRange(phases[lp].Where(t => !t.IsSummary));
                    if (laterLeaves.Count == 0) continue;

                    var succLeaf = laterLeaves[_rng.Next(laterLeaves.Count)];
                    var predSummary = predSummaries[_rng.Next(predSummaries.Count)];
                    AddDep(succLeaf, predSummary.Uid, LinkFS, usedPairs, allDeps);
                }
            }

            return allDeps;
        }

        private bool AddDep(TaskNode successor, int predecessorUid, int linkType,
            HashSet<string> usedPairs, List<PredLink> allDeps)
        {
            string key = string.Format("{0}|{1}", predecessorUid, successor.Uid);
            if (usedPairs.Contains(key)) return false;
            if (predecessorUid == successor.Uid) return false;
            usedPairs.Add(key);

            var dep = new PredLink { PredecessorUid = predecessorUid, Type = linkType };
            successor.Predecessors.Add(dep);
            allDeps.Add(dep);
            return true;
        }

        #endregion

        #region XML Writer

        private void WriteXml(List<TaskNode> tasks, List<PredLink> deps, string outputPath)
        {
            var settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  ",
                Encoding = Encoding.UTF8,
                OmitXmlDeclaration = false
            };

            using (var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            using (var writer = XmlWriter.Create(stream, settings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Project", "http://schemas.microsoft.com/project");

                // SaveVersion is required by MS Project to identify the schema version
                writer.WriteElementString("SaveVersion", "14");
                writer.WriteElementString("Name", "Large Test Project");
                writer.WriteElementString("Title", "Generated Test Project");
                writer.WriteElementString("ScheduleFromStart", "1");
                writer.WriteElementString("StartDate", FormatDate(_projectStart));
                writer.WriteElementString("FinishDate", FormatDate(tasks[0].Finish));
                writer.WriteElementString("CalendarUID", "1");
                writer.WriteElementString("MinutesPerDay", "480");
                writer.WriteElementString("MinutesPerWeek", "2400");
                writer.WriteElementString("DaysPerMonth", "20");

                // Standard calendar
                WriteCalendar(writer);

                // Tasks
                writer.WriteStartElement("Tasks");
                foreach (var t in tasks)
                {
                    WriteTask(writer, t);
                }
                writer.WriteEndElement(); // Tasks

                writer.WriteEndElement(); // Project
                writer.WriteEndDocument();
            }
        }

        private void WriteCalendar(XmlWriter w)
        {
            w.WriteStartElement("Calendars");
            w.WriteStartElement("Calendar");
            w.WriteElementString("UID", "1");
            w.WriteElementString("Name", "Standard");
            w.WriteElementString("IsBaseCalendar", "1");

            w.WriteStartElement("WeekDays");
            // Sunday (non-working)
            WriteWeekDay(w, 1, false);
            // Monday-Friday (working 8:00-12:00, 13:00-17:00)
            for (int day = 2; day <= 6; day++)
                WriteWeekDay(w, day, true);
            // Saturday (non-working)
            WriteWeekDay(w, 7, false);
            w.WriteEndElement(); // WeekDays

            w.WriteEndElement(); // Calendar
            w.WriteEndElement(); // Calendars
        }

        private void WriteWeekDay(XmlWriter w, int dayType, bool working)
        {
            w.WriteStartElement("WeekDay");
            w.WriteElementString("DayType", dayType.ToString());
            w.WriteElementString("DayWorking", working ? "1" : "0");
            if (working)
            {
                w.WriteStartElement("WorkingTimes");
                w.WriteStartElement("WorkingTime");
                w.WriteElementString("FromTime", "08:00:00");
                w.WriteElementString("ToTime", "12:00:00");
                w.WriteEndElement();
                w.WriteStartElement("WorkingTime");
                w.WriteElementString("FromTime", "13:00:00");
                w.WriteElementString("ToTime", "17:00:00");
                w.WriteEndElement();
                w.WriteEndElement(); // WorkingTimes
            }
            w.WriteEndElement(); // WeekDay
        }

        private void WriteTask(XmlWriter w, TaskNode t)
        {
            w.WriteStartElement("Task");
            w.WriteElementString("UID", t.Uid.ToString());
            w.WriteElementString("ID", t.Id.ToString());
            w.WriteElementString("Name", t.Name);
            w.WriteElementString("OutlineLevel", t.OutlineLevel.ToString());
            w.WriteElementString("OutlineNumber", t.OutlineNumber);
            w.WriteElementString("WBS", t.OutlineNumber);
            w.WriteElementString("Summary", t.IsSummary ? "1" : "0");
            w.WriteElementString("Milestone", t.IsMilestone ? "1" : "0");
            w.WriteElementString("Start", FormatDate(t.Start));
            w.WriteElementString("Finish", FormatDate(t.Finish));
            w.WriteElementString("Duration", FormatDuration(t.DurationDays));
            w.WriteElementString("DurationFormat", "7"); // 7 = days

            if (!t.IsSummary && !t.IsMilestone)
            {
                // Work = duration * 8 hours
                w.WriteElementString("Work", FormatDuration(t.DurationDays, asWork: true));
            }

            // Predecessors
            foreach (var pred in t.Predecessors)
            {
                w.WriteStartElement("PredecessorLink");
                w.WriteElementString("PredecessorUID", pred.PredecessorUid.ToString());
                w.WriteElementString("Type", pred.Type.ToString());
                w.WriteElementString("CrossProject", "0");
                w.WriteEndElement();
            }

            w.WriteEndElement(); // Task
        }

        private string FormatDate(DateTime dt)
        {
            if (dt == default(DateTime))
                return _projectStart.ToString("yyyy-MM-ddT08:00:00", CultureInfo.InvariantCulture);
            return dt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
        }

        private string FormatDuration(int days, bool asWork = false)
        {
            if (asWork)
            {
                // Work in ISO 8601 duration: PT<hours>H0M0S
                int hours = days * 8;
                return string.Format("PT{0}H0M0S", hours);
            }
            // Duration in ISO 8601: PT<hours>H0M0S
            int durationHours = days * 8;
            return string.Format("PT{0}H0M0S", durationHours);
        }

        #endregion
    }

    class Program
    {
        static void Main(string[] args)
        {
            int taskCount = 500;
            string outputPath = null;

            // Parse args
            for (int i = 0; i < args.Length; i++)
            {
                if ((args[i] == "-n" || args[i] == "--tasks") && i + 1 < args.Length)
                    int.TryParse(args[++i], out taskCount);
                else if ((args[i] == "-o" || args[i] == "--output") && i + 1 < args.Length)
                    outputPath = args[++i];
            }

            if (string.IsNullOrEmpty(outputPath))
                outputPath = string.Format("TestProject_{0}tasks.xml", taskCount);

            Console.WriteLine("Generating MSPDI XML with ~{0} tasks...", taskCount);

            string fullPath = Path.GetFullPath(outputPath);
            var dir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            var gen = new MspdiGenerator();
            gen.Generate(fullPath, totalTasks: taskCount, maxDepth: 8, depDensity: 0.4);

            // Count actual tasks and deps from file
            string xml = File.ReadAllText(fullPath);
            int taskLines = xml.Split(new[] { "<Task>" }, StringSplitOptions.None).Length - 1;
            int depLines = xml.Split(new[] { "<PredecessorLink>" }, StringSplitOptions.None).Length - 1;

            Console.WriteLine("Written: {0}", fullPath);
            Console.WriteLine("  Tasks: {0}", taskLines);
            Console.WriteLine("  Dependencies: {0}", depLines);
            Console.WriteLine("  File size: {0:N0} bytes", new FileInfo(fullPath).Length);
            Console.WriteLine();
            Console.WriteLine("Open this file in MS Project, then Save As .mpp");
        }
    }
}
