using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using ADC.MppImport.MppReader.Mpp;
using ADC.MppImport.MppReader.Model;

namespace ADC.MppImport.Services
{
    /// <summary>
    /// Testable business logic for importing MPP file tasks into Project Operations.
    /// Uses the Project Scheduling Service (PSS) API to create/update msdyn_projecttask records.
    /// </summary>
    public class MppProjectImportService
    {
        private readonly IOrganizationService _service;
        private readonly ITracingService _trace;

        public MppProjectImportService(IOrganizationService service, ITracingService trace)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _trace = trace;
        }

        /// <summary>
        /// Main entry point: reads MPP from the case template file field, parses it,
        /// and creates/updates project tasks via the PSS API for the given project.
        /// </summary>
        public ImportResult Execute(Guid caseTemplateId, Guid projectId, DateTime? projectStartDate = null)
        {
            // 1. Download the MPP file bytes from the case template file field
            _trace?.Trace("Downloading MPP file from adc_adccasetemplate {0}...", caseTemplateId);
            byte[] mppBytes = DownloadFileColumn(caseTemplateId, "adc_adccasetemplate", "adc_templatefile");
            if (mppBytes == null || mppBytes.Length == 0)
                throw new InvalidPluginExecutionException("No MPP file found on the case template record.");

            _trace?.Trace("MPP file downloaded: {0} bytes", mppBytes.Length);

            return ExecuteFromBytes(mppBytes, projectId, projectStartDate);
        }

        /// <summary>
        /// Overload that accepts raw MPP file bytes directly (for testing without a case template).
        /// </summary>
        public ImportResult ExecuteFromBytes(byte[] mppBytes, Guid projectId, DateTime? projectStartDate = null)
        {
            if (mppBytes == null || mppBytes.Length == 0)
                throw new ArgumentException("MPP file bytes are empty.");

            var result = new ImportResult();

            // Parse the MPP file
            var reader = new MppFileReader();
            ProjectFile project = reader.Read(mppBytes);
            _trace?.Trace("MPP parsed: {0} tasks, {1} resources", project.Tasks.Count, project.Resources.Count);

            if (project.Tasks.Count == 0)
            {
                _trace?.Trace("No tasks found in MPP file.");
                return result;
            }

            // 3. Retrieve existing project tasks for this project (for upsert matching)
            var existingTasks = RetrieveExistingProjectTasks(projectId);
            _trace?.Trace("Existing msdyn_projecttask records: {0}", existingTasks.Count);

            // If no existing tasks, the project may have just been created (e.g. by a workflow).
            // PSS needs time to initialise the project (create root task etc.) before accepting operations.
            if (existingTasks.Count == 0)
            {
                _trace?.Trace("New project detected — polling for PSS initialisation (root task)...");
                WaitForProjectReady(projectId);
            }

            // 4. Build a lookup of existing tasks by msdyn_msprojectclientid
            var existingByClientId = new Dictionary<string, Entity>(StringComparer.OrdinalIgnoreCase);
            foreach (var et in existingTasks)
            {
                string clientId = et.GetAttributeValue<string>("msdyn_msprojectclientid");
                if (!string.IsNullOrEmpty(clientId) && !existingByClientId.ContainsKey(clientId))
                    existingByClientId[clientId] = et;
            }

            // 5. Retrieve or create the default project bucket
            var projectRef = new EntityReference("msdyn_project", projectId);
            EntityReference bucketRef = GetOrCreateDefaultBucket(projectId);
            _trace?.Trace("Using project bucket: {0}", bucketRef.Id);


            // 6. Derive parent from OutlineLevel for tasks where ParentTask was not set by the MPP reader.
            //    Must run before task entity creation so HasChildTasks is correct for summary detection.
            // Filter out the MPP project summary task (UniqueID 0, outline level 0) — it represents the
            // project itself, not a real task. PSS already creates its own root task for the project.
            var sortedByOrder = project.Tasks
                .Where(t => t.UniqueID.HasValue && t.UniqueID.Value != 0)
                .ToList(); // preserve MPP file order (= outline order)

            // Dump reader diagnostics
            _trace?.Trace("=== MPP READER DIAGNOSTICS ===");
            _trace?.Trace("  MppFileType={0}", project.ProjectProperties.MppFileType ?? -1);
            foreach (var msg in project.DiagnosticMessages)
                _trace?.Trace("  {0}", msg);

            // (before-parent-derivation dump removed to save trace space)
            var lastAtLevel = new Dictionary<int, Task>();
            int parentsDerived = 0;
            foreach (var mppTask in sortedByOrder)
            {
                int level = mppTask.OutlineLevel ?? 0;
                if (level > 0 && mppTask.ParentTask == null)
                {
                    Task derivedParent;
                    if (lastAtLevel.TryGetValue(level - 1, out derivedParent))
                    {
                        mppTask.ParentTask = derivedParent;
                        if (!derivedParent.ChildTasks.Contains(mppTask))
                            derivedParent.ChildTasks.Add(mppTask);
                        parentsDerived++;
                        _trace?.Trace("  Derived parent for [{0}] '{1}' (level {2}) -> [{3}] '{4}'",
                            mppTask.UniqueID.Value, mppTask.Name, level,
                            derivedParent.UniqueID.Value, derivedParent.Name);
                    }
                }
                lastAtLevel[level] = mppTask;
            }
            if (parentsDerived > 0)
                _trace?.Trace("Derived {0} parent relationships from outline levels", parentsDerived);

            // D365 Project Operations supports max 10 outline levels (1-10 for user tasks,
            // 0 is the PSS root). Compute actual depth from parent chain (OutlineLevel can be null)
            // and reparent any task deeper than 10 to cap at level 10.
            const int MAX_DEPTH = 10;
            var depthCache = new Dictionary<int, int>(); // UniqueID -> depth
            int tasksCapped = 0;

            // Helper: compute depth by walking parent chain (depth 1 = top-level, no parent)
            foreach (var mppTask in sortedByOrder)
            {
                if (!mppTask.UniqueID.HasValue) continue;
                int depth = 1;
                var p = mppTask.ParentTask;
                while (p != null && p.UniqueID.HasValue && p.UniqueID.Value != 0)
                {
                    depth++;
                    p = p.ParentTask;
                }
                depthCache[mppTask.UniqueID.Value] = depth;
            }

            // Log max depth found for diagnostics
            int maxDepthFound = depthCache.Count > 0 ? depthCache.Values.Max() : 0;
            _trace?.Trace("Max parent chain depth: {0} (limit={1})", maxDepthFound, MAX_DEPTH);
            if (maxDepthFound > MAX_DEPTH)
            {
                foreach (var kvp in depthCache.Where(kv => kv.Value > MAX_DEPTH).Take(5))
                {
                    var t = sortedByOrder.FirstOrDefault(x => x.UniqueID.HasValue && x.UniqueID.Value == kvp.Key);
                    _trace?.Trace("  Deep task [{0}] '{1}' depth={2}", kvp.Key, t?.Name ?? "?", kvp.Value);
                }
            }

            foreach (var mppTask in sortedByOrder)
            {
                if (!mppTask.UniqueID.HasValue) continue;
                int depth;
                if (!depthCache.TryGetValue(mppTask.UniqueID.Value, out depth)) continue;
                if (depth <= MAX_DEPTH || mppTask.ParentTask == null) continue;

                // Walk up to find ancestor at depth MAX_DEPTH - 1 (so this task becomes depth MAX_DEPTH)
                var ancestor = mppTask.ParentTask;
                while (ancestor != null && ancestor.UniqueID.HasValue)
                {
                    int ancestorDepth;
                    if (depthCache.TryGetValue(ancestor.UniqueID.Value, out ancestorDepth) && ancestorDepth < MAX_DEPTH)
                        break;
                    ancestor = ancestor.ParentTask;
                }
                if (ancestor != null)
                {
                    _trace?.Trace("  Capping [{0}] '{1}' from depth {2} -> {3} (parent [{4}] '{5}')",
                        mppTask.UniqueID.Value, mppTask.Name, depth, MAX_DEPTH,
                        ancestor.UniqueID.Value, ancestor.Name);
                    mppTask.ParentTask = ancestor;
                    depthCache[mppTask.UniqueID.Value] = MAX_DEPTH;
                    tasksCapped++;
                }
            }
            if (tasksCapped > 0)
                _trace?.Trace("Capped {0} tasks from exceeding max depth {1}", tasksCapped, MAX_DEPTH);

            // Compact task dump (first 5 only)
            _trace?.Trace("=== TASK SAMPLE (first 5 of {0}) ===", sortedByOrder.Count);
            for (int ti = 0; ti < Math.Min(5, sortedByOrder.Count); ti++)
            {
                var t = sortedByOrder[ti];
                string durStr = t.Duration != null ? string.Format("{0} {1}", t.Duration.Value, t.Duration.Units) : "null";
                _trace?.Trace("  [{0}] L{1} '{2}' | Dur={3}", t.UniqueID.Value, t.OutlineLevel ?? 0, t.Name, durStr);
            }

            // 7. Build reliable summary-task set from ParentTask references.
            //    HasChildTasks depends on ChildTasks being populated, which doesn't happen
            //    when parent derivation is skipped (e.g. OutlineLevel = 0 for all tasks).
            var summaryTaskIds = new HashSet<int>();
            foreach (var mppTask in sortedByOrder)
            {
                if (mppTask.ParentTask != null && mppTask.ParentTask.UniqueID.HasValue
                    && mppTask.ParentTask.UniqueID.Value != 0)
                    summaryTaskIds.Add(mppTask.ParentTask.UniqueID.Value);
            }
            _trace?.Trace("Summary tasks (from parent refs): {0}", summaryTaskIds.Count);

            // Pre-generate IDs for new tasks; build map of MPP UniqueID -> CRM record GUID
            var taskIdMap = new Dictionary<int, Guid>();

            // Collect all task operations first to allow batching
            var taskOps = new List<Action<string>>();

            // First pass: queue PssCreate or PssUpdate for each task (sorted by MPP ID for correct display order).
            // Parent links are set here using pre-generated GUIDs as correlation IDs
            // (PSS honours these within the SAME operation set; PssUpdate in a separate op set is silently ignored).
            int parentLinksSet = 0;
            foreach (var mppTask in sortedByOrder)
            {
                if (!mppTask.UniqueID.HasValue) continue;

                string clientId = mppTask.UniqueID.Value.ToString();
                Entity existing = null;
                existingByClientId.TryGetValue(clientId, out existing);

                bool isSummary = summaryTaskIds.Contains(mppTask.UniqueID.Value);
                Entity taskEntity = MapMppTaskToEntity(mppTask, projectRef, bucketRef, isSummary);

                // Set explicit outline level from computed depth — helps PSS when MPP
                // OutlineLevels are all 0 (reader didn't populate them)
                int taskDepth;
                if (depthCache.TryGetValue(mppTask.UniqueID.Value, out taskDepth))
                    taskEntity["msdyn_outlinelevel"] = taskDepth;

                // Set parent link using pre-generated GUID (parents are already in taskIdMap
                // because sortedByOrder is in outline order = parents before children)
                if (mppTask.ParentTask != null && mppTask.ParentTask.UniqueID.HasValue)
                {
                    Guid parentGuid;
                    if (taskIdMap.TryGetValue(mppTask.ParentTask.UniqueID.Value, out parentGuid))
                    {
                        taskEntity["msdyn_parenttask"] = new EntityReference("msdyn_projecttask", parentGuid);
                        parentLinksSet++;
                    }
                }

                if (existing != null)
                {
                    taskEntity.Id = existing.Id;
                    taskEntity.LogicalName = "msdyn_projecttask";
                    var capturedEntity = taskEntity;
                    taskOps.Add(osId => PssUpdate(capturedEntity, osId));
                    taskIdMap[mppTask.UniqueID.Value] = existing.Id;
                    result.TasksUpdated++;
                    _trace?.Trace("  Queued UPDATE for task [{0}] {1}", mppTask.UniqueID.Value, mppTask.Name);
                }
                else
                {
                    Guid newId = Guid.NewGuid();
                    taskEntity.Id = newId;
                    taskEntity["msdyn_projecttaskid"] = newId;
                    var capturedEntity = taskEntity;
                    taskOps.Add(osId => PssCreate(capturedEntity, osId));
                    taskIdMap[mppTask.UniqueID.Value] = newId;
                    result.TasksCreated++;
                    _trace?.Trace("  Queued CREATE for task [{0}] {1} -> {2}", mppTask.UniqueID.Value, mppTask.Name, newId);
                }
            }
            _trace?.Trace("Parent links set in Phase 1: {0}", parentLinksSet);

            // 7. Execute in two phases.
            //    Phase 1: task creates WITH parent links (pre-generated GUIDs work as
            //             correlation IDs within the same operation set).
            //    Phase 2: poll CRM for actual GUIDs, then dependencies only
            //             (deps reference two tasks and must use real CRM GUIDs).
            const int PSS_BATCH_LIMIT = 200;
            int batchNum = 0;

            // Phase 1: project start date + task creates
            var phase1Ops = new List<Action<string>>();

            if (projectStartDate.HasValue)
            {
                _trace?.Trace("Setting project start date to {0:yyyy-MM-dd}", projectStartDate.Value);
                var projectUpdate = new Entity("msdyn_project", projectId);
                projectUpdate["msdyn_scheduledstart"] = projectStartDate.Value;
                phase1Ops.Add(osId => PssUpdate(projectUpdate, osId));
            }

            phase1Ops.AddRange(taskOps);

            _trace?.Trace("Phase 1: {0} ops (task creates)", phase1Ops.Count);
            batchNum = ExecuteOpsBatched(phase1Ops, projectId, PSS_BATCH_LIMIT, "Tasks", batchNum);

            // Phase 2: poll CRM for actual task GUIDs (by name match), then build
            //          parent links + dependencies using real CRM record IDs.
            int expectedCount = taskIdMap.Count;
            List<Entity> crmTasks = null;
            // Scale poll attempts with task count: PSS takes longer for larger batches
            // ~40s for 23 tasks, so budget ~2s per task, min 15 polls, max 24 (120s)
            int maxPolls = Math.Max(15, Math.Min(24, expectedCount / 2));
            _trace?.Trace("Phase 2: polling for {0} tasks (max {1} polls, {2}s)", expectedCount, maxPolls, maxPolls * 5);
            for (int poll = 0; poll < maxPolls; poll++)
            {
                System.Threading.Thread.Sleep(5000);
                crmTasks = RetrieveExistingProjectTasks(projectId);
                _trace?.Trace("Phase 2 poll {0}: {1} CRM tasks found (expecting ~{2})", poll + 1, crmTasks.Count, expectedCount);
                if (crmTasks.Count >= expectedCount)
                    break;
            }

            // Build name -> actual GUID lookup from CRM tasks
            var crmTasksByName = new Dictionary<string, List<Guid>>(StringComparer.OrdinalIgnoreCase);
            foreach (var crmTask in crmTasks)
            {
                string name = crmTask.GetAttributeValue<string>("msdyn_subject");
                if (string.IsNullOrEmpty(name)) continue;
                List<Guid> ids;
                if (!crmTasksByName.TryGetValue(name, out ids))
                {
                    ids = new List<Guid>();
                    crmTasksByName[name] = ids;
                }
                ids.Add(crmTask.Id);
            }

            // Map MPP UniqueID -> actual CRM GUID (consume matches in order for duplicate names)
            var actualTaskIdMap = new Dictionary<int, Guid>();
            foreach (var mppTask in sortedByOrder)
            {
                if (!mppTask.UniqueID.HasValue) continue;
                string name = mppTask.Name ?? "(Unnamed Task)";
                List<Guid> ids;
                if (crmTasksByName.TryGetValue(name, out ids) && ids.Count > 0)
                {
                    actualTaskIdMap[mppTask.UniqueID.Value] = ids[0];
                    ids.RemoveAt(0);
                }
                else
                {
                    _trace?.Trace("  WARNING: No CRM match for MPP task [{0}] '{1}'", mppTask.UniqueID.Value, name);
                }
            }
            _trace?.Trace("Task matching: {0} matched out of {1}", actualTaskIdMap.Count, taskIdMap.Count);

            // Build dependency ops using actual CRM GUIDs
            var existingDeps = RetrieveExistingDependencies(projectId);
            var seenLinks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var ed in existingDeps)
            {
                var predRef = ed.GetAttributeValue<EntityReference>("msdyn_predecessortask");
                var succRef = ed.GetAttributeValue<EntityReference>("msdyn_successortask");
                if (predRef != null && succRef != null)
                    seenLinks.Add(string.Format("{0}|{1}", predRef.Id, succRef.Id));
            }

            // PSS rejects deps on summary tasks — transform them to leaf-task equivalents
            _trace?.Trace("Summary tasks (have children): {0}", summaryTaskIds.Count);

            // Build parent → children map for resolving summary deps
            var childrenOfMap = new Dictionary<int, List<int>>();
            foreach (var t in project.Tasks)
            {
                if (!t.UniqueID.HasValue || t.UniqueID.Value == 0) continue;
                if (t.ParentTask != null && t.ParentTask.UniqueID.HasValue && t.ParentTask.UniqueID.Value != 0)
                {
                    int parentId = t.ParentTask.UniqueID.Value;
                    List<int> list;
                    if (!childrenOfMap.TryGetValue(parentId, out list))
                    {
                        list = new List<int>();
                        childrenOfMap[parentId] = list;
                    }
                    list.Add(t.UniqueID.Value);
                }
            }

            // Build predecessor/successor maps from leaf-to-leaf deps for entry/exit analysis
            var predecessorMap = new Dictionary<int, List<int>>();
            var successorMap = new Dictionary<int, List<int>>();
            foreach (var t in project.Tasks)
            {
                if (!t.UniqueID.HasValue || t.UniqueID.Value == 0) continue;
                if (t.Predecessors == null) continue;
                foreach (var rel in t.Predecessors)
                {
                    if (rel.SourceTaskUniqueID == 0) continue;
                    if (summaryTaskIds.Contains(rel.SourceTaskUniqueID) || summaryTaskIds.Contains(t.UniqueID.Value))
                        continue; // only leaf-to-leaf
                    List<int> list;
                    if (!predecessorMap.TryGetValue(t.UniqueID.Value, out list))
                    {
                        list = new List<int>();
                        predecessorMap[t.UniqueID.Value] = list;
                    }
                    list.Add(rel.SourceTaskUniqueID);
                    if (!successorMap.TryGetValue(rel.SourceTaskUniqueID, out list))
                    {
                        list = new List<int>();
                        successorMap[rel.SourceTaskUniqueID] = list;
                    }
                    list.Add(t.UniqueID.Value);
                }
            }

            int totalMppPredecessors = 0;
            foreach (var t in project.Tasks)
                totalMppPredecessors += (t.Predecessors != null ? t.Predecessors.Count : 0);

            var depOps = new List<Action<string>>();
            int duplicatesSkipped = 0;
            int summaryDepsTransformed = 0;

            if (totalMppPredecessors > 0)
            {
                _trace?.Trace("MPP has {0} explicit predecessor relationships, using those.", totalMppPredecessors);
                foreach (var mppTask in project.Tasks)
                {
                    if (!mppTask.UniqueID.HasValue) continue;
                    if (mppTask.Predecessors == null || mppTask.Predecessors.Count == 0) continue;

                    foreach (var relation in mppTask.Predecessors)
                    {
                        if (relation.SourceTaskUniqueID == 0) continue;

                        int predUid = relation.SourceTaskUniqueID;
                        int succUid = mppTask.UniqueID.Value;
                        bool predIsSummary = summaryTaskIds.Contains(predUid);
                        bool succIsSummary = summaryTaskIds.Contains(succUid);
                        int linkType = MapRelationType(relation.Type);

                        // Resolve summary endpoints to entry/exit leaf tasks
                        List<int> predIds = predIsSummary
                            ? GetExitLeaves(predUid, childrenOfMap, summaryTaskIds, successorMap)
                            : new List<int> { predUid };
                        List<int> succIds = succIsSummary
                            ? GetEntryLeaves(succUid, childrenOfMap, summaryTaskIds, predecessorMap)
                            : new List<int> { succUid };

                        if (predIsSummary || succIsSummary)
                        {
                            summaryDepsTransformed++;
                            _trace?.Trace("  Summary dep [{0}]->[{1}] → {2} pred x {3} succ leaf deps",
                                predUid, succUid, predIds.Count, succIds.Count);
                        }

                        foreach (var pId in predIds)
                        {
                            foreach (var sId in succIds)
                            {
                                if (pId == sId) continue;
                                Guid predecessorGuid, successorGuid;
                                if (!actualTaskIdMap.TryGetValue(pId, out predecessorGuid)) continue;
                                if (!actualTaskIdMap.TryGetValue(sId, out successorGuid)) continue;

                                string linkKey = string.Format("{0}|{1}", predecessorGuid, successorGuid);
                                if (!seenLinks.Add(linkKey))
                                {
                                    duplicatesSkipped++;
                                    continue;
                                }

                                var dep = new Entity("msdyn_projecttaskdependency");
                                dep.Id = Guid.NewGuid();
                                dep["msdyn_projecttaskdependencyid"] = dep.Id;
                                dep["msdyn_project"] = projectRef;
                                dep["msdyn_predecessortask"] = new EntityReference("msdyn_projecttask", predecessorGuid);
                                dep["msdyn_successortask"] = new EntityReference("msdyn_projecttask", successorGuid);
                                dep["msdyn_linktype"] = new OptionSetValue(linkType);
                                var capturedDep = dep;
                                depOps.Add(osId => PssCreate(capturedDep, osId));
                            }
                        }
                    }
                }
                if (duplicatesSkipped > 0)
                    _trace?.Trace("Skipped {0} duplicate dependency links", duplicatesSkipped);
                if (summaryDepsTransformed > 0)
                    _trace?.Trace("Transformed {0} summary deps to leaf-task deps", summaryDepsTransformed);
            }
            else
            {
                _trace?.Trace("MPP has no predecessor relationships. Auto-creating sequential FS dependencies.");
                depOps = BuildAutoSequentialDependencyOps(project, projectRef, actualTaskIdMap);
            }

            // Execute Phase 2: dependencies (fault-tolerant — log and continue on failure)
            // Parent links were already set in Phase 1 PssCreate.
            result.DependenciesCreated = depOps.Count;
            _trace?.Trace("Phase 2b: {0} dependency ops", depOps.Count);
            if (depOps.Count > 0)
            {
                try
                {
                    batchNum = ExecuteOpsBatched(depOps, projectId, PSS_BATCH_LIMIT, "Deps", batchNum);
                }
                catch (Exception depEx)
                {
                    _trace?.Trace("WARNING: Dependency batch failed — tasks and hierarchy are intact. Error: {0}", depEx.Message);
                    result.DependenciesCreated = 0;
                }
            }

            _trace?.Trace("Import complete. {0} batches, Created: {1}, Updated: {2}, Dependencies: {3}",
                batchNum, result.TasksCreated, result.TasksUpdated, result.DependenciesCreated);
            return result;
        }

        #region PSS API Helpers

        /// <summary>
        /// Executes a list of deferred PSS operations in batches of up to <paramref name="batchLimit"/>.
        /// Returns the updated batch number.
        /// </summary>
        private int ExecuteOpsBatched(List<Action<string>> ops, Guid projectId, int batchLimit, string phase, int batchNum)
        {
            int start = 0;
            while (start < ops.Count)
            {
                int size = Math.Min(batchLimit, ops.Count - start);
                batchNum++;

                string operationSetId = CreateOperationSet(projectId, string.Format("MPP Import {0} Batch {1}", phase, batchNum));
                _trace?.Trace("OperationSet {0} batch {1} created: {2} ({3} ops)", phase, batchNum, operationSetId, size);

                for (int i = start; i < start + size; i++)
                    ops[i](operationSetId);

                try
                {
                    string executeResult = ExecuteOperationSet(operationSetId);
                    _trace?.Trace("Batch {0} executed successfully. Result: {1}", batchNum, executeResult ?? "(none)");

                    // Wait for the operation set to fully complete before moving on
                    WaitForOperationSetCompletion(operationSetId);
                }
                catch (Exception ex)
                {
                    _trace?.Trace("Batch {0} ({1}) execution failed: {2}", batchNum, phase, ex.Message);
                    if (ex.InnerException != null)
                        _trace?.Trace("Inner: {0}", ex.InnerException.Message);

                    try { LogOperationSetStatus(operationSetId); }
                    catch (Exception logEx) { _trace?.Trace("Could not retrieve OperationSet status: {0}", logEx.Message); }

                    throw;
                }

                start += size;
            }
            return batchNum;
        }

        /// <summary>
        /// Waits for PSS to finish initialising a newly created project by polling for the root task.
        /// PSS creates a root msdyn_projecttask automatically; until it exists, operations will fail with ProjectNotFound.
        /// </summary>
        private void WaitForProjectReady(Guid projectId)
        {
            const int MAX_POLLS = 5;
            const int POLL_INTERVAL_MS = 2000; // 2s × 5 = 10s max (async workflow wait step handles most of the delay)

            for (int attempt = 0; attempt < MAX_POLLS; attempt++)
            {
                System.Threading.Thread.Sleep(POLL_INTERVAL_MS);

                try
                {
                    var query = new QueryExpression("msdyn_projecttask")
                    {
                        ColumnSet = new ColumnSet("msdyn_subject"),
                        TopCount = 1,
                        Criteria = new FilterExpression
                        {
                            Conditions =
                            {
                                new ConditionExpression("msdyn_project", ConditionOperator.Equal, projectId)
                            }
                        }
                    };
                    var result = _service.RetrieveMultiple(query);
                    if (result.Entities.Count > 0)
                    {
                        _trace?.Trace("PSS project ready (root task found) after {0}s", (attempt + 1) * POLL_INTERVAL_MS / 1000);
                        return;
                    }
                }
                catch (Exception ex)
                {
                    _trace?.Trace("  Project ready check error: {0}", ex.Message);
                }

                if (attempt % 5 == 4)
                    _trace?.Trace("  Still waiting for PSS project init... ({0}s)", (attempt + 1) * POLL_INTERVAL_MS / 1000);
            }

            _trace?.Trace("WARNING: PSS project root task not found after {0}s — proceeding anyway", MAX_POLLS * POLL_INTERVAL_MS / 1000);
        }

        /// <summary>
        /// Polls the msdyn_operationset record until its status indicates completion or failure.
        /// PSS ExecuteOperationSetV1 is async — records may not be committed immediately.
        /// </summary>
        private void WaitForOperationSetCompletion(string operationSetId)
        {
            Guid osId;
            if (!Guid.TryParse(operationSetId, out osId)) return;

            // msdyn_operationset statuses: 0=Open, 1=Completed, 2=Failed, 3=Cancelled (may vary)
            const int MAX_POLLS = 20;
            const int POLL_INTERVAL_MS = 2000; // 2s × 20 = 40s max

            for (int attempt = 0; attempt < MAX_POLLS; attempt++)
            {
                System.Threading.Thread.Sleep(POLL_INTERVAL_MS);

                try
                {
                    var osRecord = _service.Retrieve("msdyn_operationset", osId,
                        new ColumnSet("msdyn_status", "statuscode", "msdyn_description"));

                    // Try multiple possible status field names for compatibility
                    int status = -1;
                    foreach (string fieldName in new[] { "msdyn_status", "statuscode" })
                    {
                        var statusValue = osRecord.GetAttributeValue<OptionSetValue>(fieldName);
                        if (statusValue != null)
                        {
                            status = statusValue.Value;
                            break;
                        }
                    }

                    // 192350001 = Completed, 192350002 = Failed, 192350003 = Cancelled
                    // Some environments use 1/2/3 instead
                    if (status == 192350001 || status == 1)
                    {
                        _trace?.Trace("OperationSet {0} completed (status={1}) after {2}s", operationSetId, status, (attempt + 1) * POLL_INTERVAL_MS / 1000);
                        return;
                    }
                    else if (status == 192350002 || status == 2 || status == 192350003 || status == 3)
                    {
                        _trace?.Trace("OperationSet {0} failed/cancelled (status={1})", operationSetId, status);
                        throw new InvalidPluginExecutionException(string.Format("OperationSet {0} ended with status {1}", operationSetId, status));
                    }

                    // Still processing — continue polling
                    if (attempt % 5 == 4)
                        _trace?.Trace("  Waiting for OperationSet completion... ({0}s)", (attempt + 1) * POLL_INTERVAL_MS / 1000);
                }
                catch (InvalidPluginExecutionException) { throw; }
                catch (Exception ex)
                {
                    _trace?.Trace("  Poll error: {0}", ex.Message);
                }
            }

            _trace?.Trace("WARNING: OperationSet {0} did not complete within {1}s — proceeding anyway", operationSetId, MAX_POLLS * POLL_INTERVAL_MS / 1000);
        }

        /// <summary>
        /// Calls msdyn_CreateOperationSetV1 to create a new operation set for batching PSS operations.
        /// </summary>
        private string CreateOperationSet(Guid projectId, string description)
        {
            var request = new OrganizationRequest("msdyn_CreateOperationSetV1");
            request["ProjectId"] = projectId.ToString();
            request["Description"] = description;
            var response = _service.Execute(request);
            return (string)response["OperationSetId"];
        }

        /// <summary>
        /// Calls msdyn_PssCreateV1 to queue a create operation in the operation set.
        /// </summary>
        private void PssCreate(Entity entity, string operationSetId)
        {
            var request = new OrganizationRequest("msdyn_PssCreateV1");
            request["Entity"] = entity;
            request["OperationSetId"] = operationSetId;
            _service.Execute(request);
        }

        /// <summary>
        /// Calls msdyn_PssUpdateV1 to queue an update operation in the operation set.
        /// </summary>
        private void PssUpdate(Entity entity, string operationSetId)
        {
            var request = new OrganizationRequest("msdyn_PssUpdateV1");
            request["Entity"] = entity;
            request["OperationSetId"] = operationSetId;
            _service.Execute(request);
        }

        /// <summary>
        /// Calls msdyn_ExecuteOperationSetV1 to commit all queued operations.
        /// </summary>
        private string ExecuteOperationSet(string operationSetId)
        {
            var request = new OrganizationRequest("msdyn_ExecuteOperationSetV1");
            request["OperationSetId"] = operationSetId;
            var response = _service.Execute(request);

            // Safely read result — key may vary by version
            foreach (var kvp in response.Results)
            {
                _trace?.Trace("  Response key: {0} = {1}", kvp.Key, kvp.Value);
            }

            if (response.Results.ContainsKey("OperationSetDetailId"))
                return response["OperationSetDetailId"] as string;

            return operationSetId;
        }

        /// <summary>
        /// Queries the msdyn_operationset record to log detailed status/error info.
        /// </summary>
        private void LogOperationSetStatus(string operationSetId)
        {
            Guid osId;
            if (!Guid.TryParse(operationSetId, out osId)) return;

            var osRecord = _service.Retrieve("msdyn_operationset", osId,
                new ColumnSet("msdyn_status", "statuscode", "msdyn_description"));

            foreach (var attr in osRecord.Attributes)
            {
                string val = attr.Value is OptionSetValue osv ? osv.Value.ToString() : (attr.Value?.ToString() ?? "(null)");
                _trace?.Trace("  OperationSet.{0} = {1}", attr.Key, val);
            }
        }

        #endregion

        #region Task Mapping

        /// <summary>
        /// Maps an MPP Task to a msdyn_projecttask Entity for create/update.
        /// </summary>
        private Entity MapMppTaskToEntity(Task mppTask, EntityReference projectRef, EntityReference bucketRef, bool isSummary = false)
        {
            var entity = new Entity("msdyn_projecttask");
            entity["msdyn_project"] = projectRef;
            entity["msdyn_projectbucket"] = bucketRef;
            entity["msdyn_subject"] = mppTask.Name ?? "(Unnamed Task)";
            entity["msdyn_LinkStatus"] = new OptionSetValue(192350000); // Not Linked

            // Summary tasks (parents): PSS auto-calculates duration/effort/dates from children
            // Only set duration/effort/dates on leaf tasks.
            // Use the isSummary flag (built from ParentTask refs) rather than HasChildTasks,
            // because ChildTasks may not be populated when OutlineLevel is 0.
            if (!isSummary)
            {
                if (mppTask.Duration != null && mppTask.Duration.Value >= 0)
                {
                    double durationHours = ConvertToHours(mppTask.Duration);
                    entity["msdyn_duration"] = Math.Round(durationHours / 8.0, 2);
                }

                if (mppTask.Work != null && mppTask.Work.Value > 0)
                {
                    double workHours = ConvertToHours(mppTask.Work);
                    entity["msdyn_effort"] = Math.Round(workHours, 2);
                }
            }

            return entity;
        }

        /// <summary>
        /// Builds deferred dependency-create operations for auto-sequential FS dependencies.
        /// Groups tasks by their parent and chains each task to the next within that group.
        /// </summary>
        private List<Action<string>> BuildAutoSequentialDependencyOps(ProjectFile project, EntityReference projectRef,
            Dictionary<int, Guid> taskIdMap)
        {
            var ops = new List<Action<string>>();

            // Group tasks by parent UniqueID (null parent = top-level)
            var childrenByParent = new Dictionary<int, List<Task>>();
            foreach (var mppTask in project.Tasks)
            {
                if (!mppTask.UniqueID.HasValue) continue;
                if (!taskIdMap.ContainsKey(mppTask.UniqueID.Value)) continue;

                int parentKey = (mppTask.ParentTask != null && mppTask.ParentTask.UniqueID.HasValue)
                    ? mppTask.ParentTask.UniqueID.Value
                    : -1; // top-level

                if (!childrenByParent.ContainsKey(parentKey))
                    childrenByParent[parentKey] = new List<Task>();
                childrenByParent[parentKey].Add(mppTask);
            }

            // For each group of siblings, create FS dependencies between consecutive tasks
            foreach (var kvp in childrenByParent)
            {
                var siblings = kvp.Value;
                for (int i = 1; i < siblings.Count; i++)
                {
                    var prevTask = siblings[i - 1];
                    var currTask = siblings[i];

                    if (!prevTask.UniqueID.HasValue || !currTask.UniqueID.HasValue) continue;

                    Guid predecessorId, successorId;
                    if (!taskIdMap.TryGetValue(prevTask.UniqueID.Value, out predecessorId)) continue;
                    if (!taskIdMap.TryGetValue(currTask.UniqueID.Value, out successorId)) continue;

                    var dep = new Entity("msdyn_projecttaskdependency");
                    dep.Id = Guid.NewGuid();
                    dep["msdyn_projecttaskdependencyid"] = dep.Id;
                    dep["msdyn_project"] = projectRef;
                    dep["msdyn_predecessortask"] = new EntityReference("msdyn_projecttask", predecessorId);
                    dep["msdyn_successortask"] = new EntityReference("msdyn_projecttask", successorId);
                    dep["msdyn_linktype"] = new OptionSetValue(192350000); // Finish-to-Start

                    var capturedDep = dep;
                    ops.Add(osId => PssCreate(capturedDep, osId));

                    _trace?.Trace("    FS: [{0}] {1} -> [{2}] {3}",
                        prevTask.UniqueID, prevTask.Name, currTask.UniqueID, currTask.Name);
                }
            }

            return ops;
        }

        /// <summary>
        /// Maps MPP RelationType to Project Operations msdyn_linktype OptionSetValue.
        /// </summary>
        private static int MapRelationType(RelationType type)
        {
            switch (type)
            {
                case RelationType.FinishToStart:  return 192350000;
                case RelationType.StartToStart:   return 192350001;
                case RelationType.FinishToFinish: return 192350002;
                case RelationType.StartToFinish:  return 192350003;
                default:                          return 192350000; // default FS
            }
        }

        /// <summary>
        /// Converts a Duration to hours based on its time units.
        /// </summary>
        private static double ConvertToHours(Duration duration)
        {
            double val = duration.Value;
            switch (duration.Units)
            {
                case TimeUnit.Minutes:
                case TimeUnit.ElapsedMinutes:
                    return val / 60.0;
                case TimeUnit.Hours:
                case TimeUnit.ElapsedHours:
                    return val;
                case TimeUnit.Days:
                case TimeUnit.ElapsedDays:
                    return val * 8.0;
                case TimeUnit.Weeks:
                case TimeUnit.ElapsedWeeks:
                    return val * 40.0;
                case TimeUnit.Months:
                case TimeUnit.ElapsedMonths:
                    return val * 160.0;
                default: return val;
            }
        }

        private HashSet<int> GetAllDescendants(int summaryId, Dictionary<int, List<int>> childrenOf)
        {
            var result = new HashSet<int>();
            var stack = new Stack<int>();
            List<int> children;
            if (childrenOf.TryGetValue(summaryId, out children))
            {
                foreach (var c in children) stack.Push(c);
            }
            while (stack.Count > 0)
            {
                int id = stack.Pop();
                result.Add(id);
                if (childrenOf.TryGetValue(id, out children))
                {
                    foreach (var c in children) stack.Push(c);
                }
            }
            return result;
        }

        private List<int> GetEntryLeaves(int summaryId, Dictionary<int, List<int>> childrenOf,
            HashSet<int> summaryTaskIds, Dictionary<int, List<int>> predecessorMap)
        {
            var descendants = GetAllDescendants(summaryId, childrenOf);
            var entries = new List<int>();
            foreach (var id in descendants)
            {
                if (summaryTaskIds.Contains(id)) continue;
                bool hasInternalPred = false;
                List<int> preds;
                if (predecessorMap.TryGetValue(id, out preds))
                {
                    foreach (var predId in preds)
                    {
                        if (descendants.Contains(predId)) { hasInternalPred = true; break; }
                    }
                }
                if (!hasInternalPred) entries.Add(id);
            }
            if (entries.Count == 0)
            {
                foreach (var id in descendants)
                    if (!summaryTaskIds.Contains(id)) entries.Add(id);
            }
            return entries;
        }

        private List<int> GetExitLeaves(int summaryId, Dictionary<int, List<int>> childrenOf,
            HashSet<int> summaryTaskIds, Dictionary<int, List<int>> successorMap)
        {
            var descendants = GetAllDescendants(summaryId, childrenOf);
            var exits = new List<int>();
            foreach (var id in descendants)
            {
                if (summaryTaskIds.Contains(id)) continue;
                bool hasInternalSucc = false;
                List<int> succs;
                if (successorMap.TryGetValue(id, out succs))
                {
                    foreach (var succId in succs)
                    {
                        if (descendants.Contains(succId)) { hasInternalSucc = true; break; }
                    }
                }
                if (!hasInternalSucc) exits.Add(id);
            }
            if (exits.Count == 0)
            {
                foreach (var id in descendants)
                    if (!summaryTaskIds.Contains(id)) exits.Add(id);
            }
            return exits;
        }

        #endregion

        #region Dataverse Helpers

        /// <summary>
        /// Retrieves the first existing msdyn_projectbucket for the project, or creates a default one.
        /// </summary>
        private EntityReference GetOrCreateDefaultBucket(Guid projectId)
        {
            // Try to find an existing bucket for this project
            var query = new QueryExpression("msdyn_projectbucket")
            {
                ColumnSet = new ColumnSet("msdyn_projectbucketid", "msdyn_name"),
                TopCount = 1,
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("msdyn_project", ConditionOperator.Equal, projectId)
                    }
                }
            };

            var results = _service.RetrieveMultiple(query);
            if (results.Entities.Count > 0)
            {
                var existing = results.Entities[0];
                _trace?.Trace("Found existing bucket: {0}", existing.GetAttributeValue<string>("msdyn_name") ?? existing.Id.ToString());
                return existing.ToEntityReference();
            }

            // No bucket exists — create a default one
            _trace?.Trace("No project bucket found, creating default bucket...");
            var bucket = new Entity("msdyn_projectbucket");
            bucket["msdyn_project"] = new EntityReference("msdyn_project", projectId);
            bucket["msdyn_name"] = "MPP Import";
            Guid bucketId = _service.Create(bucket);
            return new EntityReference("msdyn_projectbucket", bucketId);
        }

        /// <summary>
        /// Downloads a file column's content as byte[] using InitializeFileBlocksDownload.
        /// </summary>
        private byte[] DownloadFileColumn(Guid recordId, string entityName, string fileAttributeName)
        {
            try
            {
                var initRequest = new OrganizationRequest("InitializeFileBlocksDownload");
                initRequest["Target"] = new EntityReference(entityName, recordId);
                initRequest["FileAttributeName"] = fileAttributeName;

                var initResponse = _service.Execute(initRequest);
                string fileContinuationToken = (string)initResponse["FileContinuationToken"];
                long fileSize = (long)initResponse["FileSizeInBytes"];

                if (fileSize == 0) return null;

                var allBytes = new List<byte>();
                long offset = 0;
                const long blockSize = 4 * 1024 * 1024; // 4MB blocks

                while (offset < fileSize)
                {
                    var downloadRequest = new OrganizationRequest("DownloadBlock");
                    downloadRequest["FileContinuationToken"] = fileContinuationToken;
                    downloadRequest["BlockLength"] = blockSize;
                    downloadRequest["Offset"] = offset;

                    var downloadResponse = _service.Execute(downloadRequest);
                    byte[] blockData = (byte[])downloadResponse["Data"];
                    allBytes.AddRange(blockData);
                    offset += blockData.Length;
                }

                return allBytes.ToArray();
            }
            catch (Exception ex)
            {
                _trace?.Trace("Error downloading file: {0}", ex.Message);
                throw new InvalidPluginExecutionException(
                    $"Failed to download MPP file from {entityName}.{fileAttributeName}: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Retrieves all existing msdyn_projecttask records for a given project.
        /// </summary>
        private List<Entity> RetrieveExistingProjectTasks(Guid projectId)
        {
            var query = new QueryExpression("msdyn_projecttask")
            {
                ColumnSet = new ColumnSet("msdyn_msprojectclientid", "msdyn_subject"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("msdyn_project", ConditionOperator.Equal, projectId)
                    }
                }
            };

            var results = new List<Entity>();
            query.PageInfo = new PagingInfo { PageNumber = 1, Count = 5000 };

            while (true)
            {
                var response = _service.RetrieveMultiple(query);
                results.AddRange(response.Entities);

                if (response.MoreRecords)
                {
                    query.PageInfo.PageNumber++;
                    query.PageInfo.PagingCookie = response.PagingCookie;
                }
                else
                {
                    break;
                }
            }

            return results;
        }

        /// <summary>
        /// Retrieves all existing msdyn_projecttaskdependency records for a given project.
        /// </summary>
        private List<Entity> RetrieveExistingDependencies(Guid projectId)
        {
            var query = new QueryExpression("msdyn_projecttaskdependency")
            {
                ColumnSet = new ColumnSet("msdyn_predecessortask", "msdyn_successortask"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("msdyn_project", ConditionOperator.Equal, projectId)
                    }
                }
            };

            var results = new List<Entity>();
            query.PageInfo = new PagingInfo { PageNumber = 1, Count = 5000 };

            while (true)
            {
                var response = _service.RetrieveMultiple(query);
                results.AddRange(response.Entities);

                if (response.MoreRecords)
                {
                    query.PageInfo.PageNumber++;
                    query.PageInfo.PagingCookie = response.PagingCookie;
                }
                else
                {
                    break;
                }
            }

            return results;
        }

        #endregion
    }

    /// <summary>
    /// Result of an MPP import operation.
    /// </summary>
    public class ImportResult
    {
        public int TasksCreated { get; set; }
        public int TasksUpdated { get; set; }
        public int DependenciesCreated { get; set; }
        public int TotalProcessed => TasksCreated + TasksUpdated;
        public List<string> Warnings { get; set; } = new List<string>();
    }
}
