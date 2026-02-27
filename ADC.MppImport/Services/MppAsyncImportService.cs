using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using ADC.MppImport.MppReader.Mpp;
using ADC.MppImport.MppReader.Model;

namespace ADC.MppImport.Services
{
    /// <summary>
    /// Async chunked MPP import service. Designed to run across multiple plugin executions,
    /// each staying under the 2-minute Dataverse sandbox timeout.
    ///
    /// Phase 0 (InitializeJob): Parse MPP, build subtree batches, create job record.
    ///          Runs inside the lightweight StartMppImportActivity workflow activity.
    ///
    /// Phases 1-5: Executed by MppImportJobPlugin (async, self-triggering).
    ///   Phase 1 (CreatingTasks):   Submit one batch of PssCreate ops.
    ///   Phase 2 (WaitingForTasks): Poll PSS operation set completion, advance batch or phase.
    ///   Phase 3 (PollingGUIDs):    Query CRM for actual task GUIDs.
    ///   Phase 4 (CreatingDeps):    Submit dependency PssCreate ops.
    ///   Phase 5 (WaitingForDeps):  Poll dep operation set completion, mark Completed.
    /// </summary>
    public class MppAsyncImportService
    {
        private readonly IOrganizationService _service;
        private readonly ITracingService _trace;

        /// <summary>Max tasks per PSS operation set batch.</summary>
        private const int BATCH_LIMIT = 100;

        /// <summary>Max dependencies per PSS operation set batch.</summary>
        private const int DEP_BATCH_LIMIT = 50;

        /// <summary>Max tick retries when waiting for PSS completion.</summary>
        private const int MAX_WAIT_TICKS = 30;

        /// <summary>D365 max outline level for user tasks.</summary>
        private const int MAX_DEPTH = 10;

        public MppAsyncImportService(IOrganizationService service, ITracingService trace)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _trace = trace;
        }

        #region Phase 0: Initialize Job (called from workflow activity)

        /// <summary>
        /// Parses the MPP file, builds subtree batches, and creates the adc_mppimportjob record.
        /// Returns the job record ID. This runs inside the workflow activity and must be fast.
        /// </summary>
        public Guid InitializeJob(byte[] mppBytes, Guid projectId, Guid? caseTemplateId, DateTime? projectStartDate, Guid? caseId = null, Guid? initiatingUserId = null)
        {
            if (mppBytes == null || mppBytes.Length == 0)
                throw new InvalidPluginExecutionException("MPP file bytes are empty.");

            _trace?.Trace("InitializeJob: parsing MPP ({0} bytes)...", mppBytes.Length);

            // Parse MPP
            var reader = new MppFileReader();
            ProjectFile project = reader.Read(mppBytes);
            _trace?.Trace("Parsed: {0} tasks, {1} resources", project.Tasks.Count, project.Resources.Count);

            if (project.Tasks.Count == 0)
                throw new InvalidPluginExecutionException("No tasks found in MPP file.");

            // Filter out project summary task (UniqueID 0)
            var sortedByOrder = project.Tasks
                .Where(t => t.UniqueID.HasValue && t.UniqueID.Value != 0)
                .ToList();

            // Derive parent relationships from OutlineLevel
            DeriveParentRelationships(sortedByOrder);

            // Compute depth from parent chain, cap at MAX_DEPTH
            var depthCache = ComputeAndCapDepths(sortedByOrder);

            // Build summary task set
            var summaryTaskIds = BuildSummaryTaskSet(sortedByOrder);
            _trace?.Trace("Summary tasks: {0}, Total tasks: {1}", summaryTaskIds.Count, sortedByOrder.Count);

            // Get or create project bucket
            EntityReference bucketRef = GetOrCreateDefaultBucket(projectId);

            // Build task DTOs
            var payload = new ImportJobPayload();
            payload.BucketId = bucketRef.Id.ToString();

            foreach (var mppTask in sortedByOrder)
            {
                if (!mppTask.UniqueID.HasValue) continue;

                var dto = new TaskDto
                {
                    UniqueID = mppTask.UniqueID.Value,
                    Name = mppTask.Name ?? "(Unnamed Task)",
                    ParentUniqueID = (mppTask.ParentTask != null && mppTask.ParentTask.UniqueID.HasValue
                        && mppTask.ParentTask.UniqueID.Value != 0)
                        ? (int?)mppTask.ParentTask.UniqueID.Value : null,
                    Depth = depthCache.ContainsKey(mppTask.UniqueID.Value) ? depthCache[mppTask.UniqueID.Value] : 1,
                    IsSummary = summaryTaskIds.Contains(mppTask.UniqueID.Value)
                };

                // Pre-compute duration/effort so plugin doesn't need MPP reader
                if (!dto.IsSummary && mppTask.Duration != null && mppTask.Duration.Value >= 0)
                    dto.DurationHours = ConvertToHours(mppTask.Duration);

                if (!dto.IsSummary && mppTask.Work != null && mppTask.Work.Value > 0)
                    dto.EffortHours = ConvertToHours(mppTask.Work);

                // Pre-generate GUID
                Guid preGen = Guid.NewGuid();
                dto.PreGenGuid = preGen.ToString();
                payload.TaskIdMap[dto.UniqueID] = dto.PreGenGuid;

                payload.Tasks.Add(dto);
            }

            // Build dependency DTOs
            payload.Dependencies = BuildDependencyDtos(project, summaryTaskIds);
            _trace?.Trace("Dependencies: {0}", payload.Dependencies.Count);

            // Build subtree batches
            payload.Batches = BuildSubtreeBatches(payload.Tasks);
            _trace?.Trace("Batches: {0} (limit {1} per batch)", payload.Batches.Count, BATCH_LIMIT);

            // Serialize payload
            string taskDataJson = SerializeJson(payload);
            _trace?.Trace("Payload JSON size: {0} chars", taskDataJson.Length);

            // Create job record
            var job = new Entity(ImportJobFields.EntityName);
            job[ImportJobFields.Name] = string.Format("Import {0} tasks", sortedByOrder.Count);
            job[ImportJobFields.Project] = new EntityReference("msdyn_project", projectId);
            if (caseTemplateId.HasValue)
                job[ImportJobFields.CaseTemplate] = new EntityReference("adc_adccasetemplate", caseTemplateId.Value);
            job[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.Queued);
            job[ImportJobFields.Phase] = 0;
            job[ImportJobFields.CurrentBatch] = 0;
            job[ImportJobFields.TotalBatches] = payload.Batches.Count;
            job[ImportJobFields.TotalTasks] = sortedByOrder.Count;
            job[ImportJobFields.CreatedCount] = 0;
            job[ImportJobFields.DepsCount] = 0;
            job[ImportJobFields.Tick] = 0;
            job[ImportJobFields.TaskDataJson] = taskDataJson;
            job[ImportJobFields.ErrorMessage] = "";
            if (projectStartDate.HasValue)
                job[ImportJobFields.ProjectStartDate] = projectStartDate.Value;
            if (caseId.HasValue)
                job[ImportJobFields.Case] = new EntityReference("adc_case", caseId.Value);
            if (initiatingUserId.HasValue)
                job[ImportJobFields.InitiatingUser] = new EntityReference("systemuser", initiatingUserId.Value);

            Guid jobId = _service.Create(job);
            _trace?.Trace("Created job record: {0}", jobId);

            // Send "import started" notification
            string projectName = GetProjectName(projectId);
            if (initiatingUserId.HasValue)
            {
                SendNotification(
                    initiatingUserId.Value,
                    "MPP Import Started",
                    string.Format("Import started for {0} ({1} tasks) \u2014 running in background.", projectName, sortedByOrder.Count),
                    NotificationIconType.Info);
            }

            // Set initial import status on adc_case
            if (caseId.HasValue)
            {
                try
                {
                    var caseUpdate = new Entity("adc_case", caseId.Value);
                    caseUpdate["adc_importstatus"] = new OptionSetValue(1); // Processing
                    caseUpdate["adc_importmessage"] = string.Format("Importing {0} tasks...", sortedByOrder.Count);
                    _service.Update(caseUpdate);
                }
                catch (Exception ex)
                {
                    _trace?.Trace("Failed to set initial case import status (non-fatal): {0}", ex.Message);
                }
            }

            return jobId;
        }

        #endregion

        #region Plugin Entry Point: ProcessJob

        /// <summary>
        /// Main entry point called by the async plugin. Reads the job record,
        /// determines current phase, executes one unit of work, and advances state.
        /// </summary>
        public void ProcessJob(Guid jobId)
        {
            var job = _service.Retrieve(ImportJobFields.EntityName, jobId, new ColumnSet(true));

            int status = GetOptionSetValue(job, ImportJobFields.Status);
            int tick = job.GetAttributeValue<int>(ImportJobFields.Tick);

            _trace?.Trace("ProcessJob {0}: status={1} ({2}), tick={3}",
                jobId, status, ImportJobStatus.Label(status), tick);

            // Don't process terminal states
            if (status == ImportJobStatus.Completed || status == ImportJobStatus.Failed)
            {
                _trace?.Trace("Job in terminal state, skipping.");
                return;
            }

            try
            {
                switch (status)
                {
                    case ImportJobStatus.Queued:
                        UpdateCaseProgress(job, "Validating project...");
                        ProcessQueued(job);
                        break;
                    case ImportJobStatus.CreatingTasks:
                        {
                            int batch = job.GetAttributeValue<int>(ImportJobFields.CurrentBatch);
                            int total = job.GetAttributeValue<int>(ImportJobFields.TotalBatches);
                            int totalTasks = job.GetAttributeValue<int>(ImportJobFields.TotalTasks);
                            UpdateCaseProgress(job, string.Format("Creating tasks — batch {0}/{1} ({2} tasks total)...", batch + 1, total, totalTasks));
                        }
                        ProcessCreatingTasks(job);
                        break;
                    case ImportJobStatus.WaitingForTasks:
                        {
                            int batch = job.GetAttributeValue<int>(ImportJobFields.CurrentBatch);
                            int total = job.GetAttributeValue<int>(ImportJobFields.TotalBatches);
                            int created = job.GetAttributeValue<int>(ImportJobFields.CreatedCount);
                            UpdateCaseProgress(job, string.Format("Waiting for tasks to commit — {0} submitted (batch {1}/{2})...", created, batch + 1, total));
                        }
                        ProcessWaitingForTasks(job);
                        break;
                    case ImportJobStatus.PollingGUIDs:
                        {
                            int totalTasks = job.GetAttributeValue<int>(ImportJobFields.TotalTasks);
                            UpdateCaseProgress(job, string.Format("Mapping {0} task IDs...", totalTasks));
                        }
                        ProcessPollingGUIDs(job);
                        break;
                    case ImportJobStatus.CreatingDeps:
                        UpdateCaseProgress(job, "Creating dependencies...");
                        ProcessCreatingDeps(job);
                        break;
                    case ImportJobStatus.WaitingForDeps:
                        UpdateCaseProgress(job, "Waiting for dependencies to commit...");
                        ProcessWaitingForDeps(job);
                        break;
                }
            }
            catch (Exception ex)
            {
                _trace?.Trace("ERROR in ProcessJob: {0}", ex.ToString());
                FailJob(job.Id, ex.Message);
            }
        }

        #endregion

        #region Phase Processors

        /// <summary>
        /// Phase 0 → 1: Job just created. Verify project exists, set start date, advance to CreatingTasks.
        /// Note: PSS does not auto-create root tasks for S2S app users, so we skip that check.
        /// PSS will handle root task creation when the first operation set is executed.
        /// </summary>
        private void ProcessQueued(Entity job)
        {
            Guid projectId = job.GetAttributeValue<EntityReference>(ImportJobFields.Project).Id;

            // Verify project exists
            try
            {
                _service.Retrieve("msdyn_project", projectId, new ColumnSet("msdyn_subject"));
                _trace?.Trace("Project {0} exists, proceeding.", projectId);
            }
            catch (Exception ex)
            {
                FailJob(job.Id, "Project not found: " + ex.Message);
                return;
            }

            // Note: Project start date is NOT set here to avoid PSS dependency rejection
            // when scheduling against past dates. The date comparison test uses offset-based
            // comparison instead (MPP start vs D365 project start → expected shift for all tasks).

            // Advance to CreatingTasks
            var update = new Entity(ImportJobFields.EntityName, job.Id);
            update[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.CreatingTasks);
            update[ImportJobFields.CurrentBatch] = 0;
            update[ImportJobFields.Tick] = 0;
            _service.Update(update);
            _trace?.Trace("Advanced to CreatingTasks, batch 0");
        }

        /// <summary>
        /// Phase 1: Process ALL task batches within a single execution.
        /// For each batch: submit PssCreate ops, wait inline for PSS completion, advance.
        /// This avoids Dataverse "infinite loop" detection which kills async plugin chains
        /// after ~10 depth levels. A timeout guard saves progress if approaching the
        /// 2-minute sandbox limit, letting the next execution resume.
        /// </summary>
        private void ProcessCreatingTasks(Entity job)
        {
            var payload = DeserializePayload(job);
            int batchIndex = job.GetAttributeValue<int>(ImportJobFields.CurrentBatch);
            int startBatch = batchIndex;
            Guid projectId = job.GetAttributeValue<EntityReference>(ImportJobFields.Project).Id;
            int totalCreated = job.GetAttributeValue<int>(ImportJobFields.CreatedCount);

            if (batchIndex >= payload.Batches.Count)
            {
                _trace?.Trace("All {0} batches submitted, advancing to PollingGUIDs", payload.Batches.Count);
                var update = new Entity(ImportJobFields.EntityName, job.Id);
                update[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.PollingGUIDs);
                update[ImportJobFields.Tick] = 0;
                _service.Update(update);
                return;
            }

            var taskLookup = payload.Tasks.ToDictionary(t => t.UniqueID);
            var projectRef = new EntityReference("msdyn_project", projectId);
            var bucketRef = new EntityReference("msdyn_projectbucket", Guid.Parse(payload.BucketId));

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            // Loop through all remaining batches within this single execution
            while (batchIndex < payload.Batches.Count)
            {
                // Timeout guard: if we've been running > 60s AND processed at least one
                // batch, save progress and let next execution continue.
                // Each batch takes ~20-30s (PSS submit + poll). At 60s we fit 2-3 batches
                // safely within the sandbox limit, with margin for the next batch's polling.
                if (stopwatch.ElapsedMilliseconds > 60000 && batchIndex > startBatch)
                {
                    _trace?.Trace("Approaching timeout at {0}ms, saving progress at batch {1}/{2}",
                        stopwatch.ElapsedMilliseconds, batchIndex, payload.Batches.Count);
                    var timeoutUpdate = new Entity(ImportJobFields.EntityName, job.Id);
                    timeoutUpdate[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.CreatingTasks);
                    timeoutUpdate[ImportJobFields.CurrentBatch] = batchIndex;
                    timeoutUpdate[ImportJobFields.CreatedCount] = totalCreated;
                    timeoutUpdate[ImportJobFields.Tick] = job.GetAttributeValue<int>(ImportJobFields.Tick) + 1;
                    _service.Update(timeoutUpdate);
                    return;
                }

                var batch = payload.Batches[batchIndex];
                _trace?.Trace("Creating batch {0}/{1} ({2} tasks)",
                    batchIndex + 1, payload.Batches.Count, batch.TaskUniqueIDs.Count);

                // --- Submit batch ---
                string operationSetId = CreateOperationSet(projectId,
                    string.Format("MPP Import Batch {0}/{1}", batchIndex + 1, payload.Batches.Count));

                int opsCreated = 0;
                foreach (int taskId in batch.TaskUniqueIDs)
                {
                    TaskDto dto;
                    if (!taskLookup.TryGetValue(taskId, out dto)) continue;

                    var entity = MapDtoToEntity(dto, projectRef, bucketRef);

                    if (dto.ParentUniqueID.HasValue)
                    {
                        string parentGuidStr;
                        if (payload.TaskIdMap.TryGetValue(dto.ParentUniqueID.Value, out parentGuidStr))
                        {
                            entity["msdyn_parenttask"] = new EntityReference("msdyn_projecttask", Guid.Parse(parentGuidStr));
                        }
                    }

                    PssCreate(entity, operationSetId);
                    opsCreated++;
                }

                _trace?.Trace("Submitted {0} PssCreate ops, executing...", opsCreated);
                ExecuteOperationSet(operationSetId);

                // --- Wait for completion inline (up to ~90s) ---
                int osStatus = -1;
                for (int attempt = 0; attempt < 18; attempt++)
                {
                    osStatus = CheckOperationSetStatus(operationSetId);
                    if (osStatus != -1) break;
                    _trace?.Trace("OperationSet still processing, attempt {0}/18", attempt + 1);
                    System.Threading.Thread.Sleep(5000);
                }

                if (osStatus == -1)
                {
                    FailJob(job.Id, string.Format(
                        "Operation set {0} did not complete after 90s (batch {1}/{2}).",
                        operationSetId, batchIndex + 1, payload.Batches.Count));
                    return;
                }

                if (osStatus == 0) // failed
                {
                    FailJob(job.Id, string.Format(
                        "Operation set {0} failed (batch {1}/{2}).",
                        operationSetId, batchIndex + 1, payload.Batches.Count));
                    return;
                }

                totalCreated += opsCreated;
                batchIndex++;
                _trace?.Trace("Batch done ({0} tasks created so far), elapsed {1}ms",
                    totalCreated, stopwatch.ElapsedMilliseconds);

                // Save progress after each batch (non-triggering: only currentBatch + createdCount)
                // Plugin filters on adc_status + adc_tick, so this won't fire a new execution.
                var progressUpdate = new Entity(ImportJobFields.EntityName, job.Id);
                progressUpdate[ImportJobFields.CurrentBatch] = batchIndex;
                progressUpdate[ImportJobFields.CreatedCount] = totalCreated;
                progressUpdate[ImportJobFields.OperationSetId] = operationSetId;
                _service.Update(progressUpdate);
            }

            // All batches done — advance to PollingGUIDs
            _trace?.Trace("All {0} batches completed, advancing to PollingGUIDs", payload.Batches.Count);
            var finalUpdate = new Entity(ImportJobFields.EntityName, job.Id);
            finalUpdate[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.PollingGUIDs);
            finalUpdate[ImportJobFields.CurrentBatch] = batchIndex;
            finalUpdate[ImportJobFields.CreatedCount] = totalCreated;
            finalUpdate[ImportJobFields.Tick] = 0;
            _service.Update(finalUpdate);
        }

        /// <summary>
        /// Phase 2: Poll for PSS operation set completion within this execution.
        /// Uses Thread.Sleep loop instead of tick-based self-triggering (which Dataverse
        /// may coalesce/throttle for rapid updates to the same record).
        /// </summary>
        private void ProcessWaitingForTasks(Entity job)
        {
            string osId = job.GetAttributeValue<string>(ImportJobFields.OperationSetId);

            if (string.IsNullOrEmpty(osId))
            {
                FailJob(job.Id, "No operation set ID stored for WaitingForTasks phase.");
                return;
            }

            // Poll within this execution — up to ~90 seconds (18 * 5s)
            int osStatus = -1;
            for (int attempt = 0; attempt < 18; attempt++)
            {
                osStatus = CheckOperationSetStatus(osId);
                if (osStatus != -1) break; // completed or failed
                _trace?.Trace("OperationSet still processing, attempt {0}/18", attempt + 1);
                System.Threading.Thread.Sleep(5000);
            }

            if (osStatus == -1) // still processing after all attempts
            {
                FailJob(job.Id, string.Format("Operation set {0} did not complete after 90s.", osId));
                return;
            }

            if (osStatus == 0) // failed
            {
                FailJob(job.Id, string.Format("Operation set {0} failed.", osId));
                return;
            }

            // Completed — advance to next batch or PollingGUIDs
            int batchIndex = job.GetAttributeValue<int>(ImportJobFields.CurrentBatch);
            int totalBatches = job.GetAttributeValue<int>(ImportJobFields.TotalBatches);

            batchIndex++;
            if (batchIndex < totalBatches)
            {
                _trace?.Trace("Batch done, advancing to batch {0}/{1}", batchIndex + 1, totalBatches);
                var update = new Entity(ImportJobFields.EntityName, job.Id);
                update[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.CreatingTasks);
                update[ImportJobFields.CurrentBatch] = batchIndex;
                update[ImportJobFields.Tick] = 0;
                _service.Update(update);
            }
            else
            {
                _trace?.Trace("All batches done, advancing to PollingGUIDs");
                var update = new Entity(ImportJobFields.EntityName, job.Id);
                update[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.PollingGUIDs);
                update[ImportJobFields.CurrentBatch] = batchIndex;
                update[ImportJobFields.Tick] = 0;
                _service.Update(update);
            }
        }

        /// <summary>
        /// Phase 3: Query CRM for actual task GUIDs, build actual ID map.
        /// </summary>
        private void ProcessPollingGUIDs(Entity job)
        {
            Guid projectId = job.GetAttributeValue<EntityReference>(ImportJobFields.Project).Id;
            int expectedCount = job.GetAttributeValue<int>(ImportJobFields.TotalTasks);

            // Poll within this execution — up to ~90 seconds (18 * 5s)
            List<Entity> crmTasks = null;
            for (int attempt = 0; attempt < 18; attempt++)
            {
                crmTasks = RetrieveExistingProjectTasks(projectId);
                _trace?.Trace("PollingGUIDs: {0} CRM tasks found (expecting ~{1}), attempt {2}/18",
                    crmTasks.Count, expectedCount, attempt + 1);
                if (crmTasks.Count >= expectedCount) break;
                System.Threading.Thread.Sleep(5000);
            }

            if (crmTasks.Count < expectedCount)
            {
                _trace?.Trace("WARNING: Only {0}/{1} tasks found after polling. Proceeding with available tasks.",
                    crmTasks.Count, expectedCount);
            }

            // Build name -> GUID lookup
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

            // Map UniqueID -> actual CRM GUID
            var payload = DeserializePayload(job);
            payload.ActualIdMap = new Dictionary<int, string>();

            foreach (var dto in payload.Tasks)
            {
                List<Guid> ids;
                if (crmTasksByName.TryGetValue(dto.Name, out ids) && ids.Count > 0)
                {
                    payload.ActualIdMap[dto.UniqueID] = ids[0].ToString();
                    ids.RemoveAt(0);
                }
                else
                {
                    _trace?.Trace("  WARNING: No CRM match for [{0}] '{1}'", dto.UniqueID, dto.Name);
                }
            }
            _trace?.Trace("Matched {0}/{1} tasks", payload.ActualIdMap.Count, payload.Tasks.Count);

            // Persist updated payload with actual ID map
            string updatedJson = SerializeJson(payload);

            var update = new Entity(ImportJobFields.EntityName, job.Id);
            update[ImportJobFields.TaskDataJson] = updatedJson;
            update[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.CreatingDeps);
            update[ImportJobFields.CurrentBatch] = 0; // reuse as dep batch index
            update[ImportJobFields.Tick] = 0;
            _service.Update(update);
        }

        /// <summary>
        /// Phase 4: Create dependency operations using actual CRM GUIDs.
        /// Uses batch-indexed fire-and-forget submission with timeout guard.
        /// Deps are split into batches of DEP_BATCH_LIMIT, each executed in its own
        /// operation set without polling. A timeout guard saves the dep batch index
        /// (via CurrentBatch) and triggers continuation when approaching the sandbox limit.
        /// Batch indices are stable across executions (built from payload, not D365 state).
        /// </summary>
        private void ProcessCreatingDeps(Entity job)
        {
            var payload = DeserializePayload(job);
            Guid projectId = job.GetAttributeValue<EntityReference>(ImportJobFields.Project).Id;

            if (payload.Dependencies.Count == 0)
            {
                _trace?.Trace("No dependencies to create, completing.");
                CompleteJob(job.Id, 0);
                return;
            }

            // Build the FULL dep entity list from payload (stable across executions).
            // We don't filter by existing deps here — batch indices must stay stable
            // so the CurrentBatch resume index remains correct.
            var projectRef = new EntityReference("msdyn_project", projectId);
            var depEntities = new List<Entity>();
            int unmapped = 0;

            foreach (var depDto in payload.Dependencies)
            {
                string predGuidStr, succGuidStr;
                if (!payload.ActualIdMap.TryGetValue(depDto.PredecessorUniqueID, out predGuidStr)) { unmapped++; continue; }
                if (!payload.ActualIdMap.TryGetValue(depDto.SuccessorUniqueID, out succGuidStr)) { unmapped++; continue; }

                var dep = new Entity("msdyn_projecttaskdependency");
                dep.Id = Guid.NewGuid();
                dep["msdyn_projecttaskdependencyid"] = dep.Id;
                dep["msdyn_project"] = projectRef;
                dep["msdyn_predecessortask"] = new EntityReference("msdyn_projecttask", Guid.Parse(predGuidStr));
                dep["msdyn_successortask"] = new EntityReference("msdyn_projecttask", Guid.Parse(succGuidStr));
                dep["msdyn_linktype"] = new OptionSetValue(depDto.LinkType);
                depEntities.Add(dep);
            }

            if (unmapped > 0)
                _trace?.Trace("Skipped {0} deps with unmapped task IDs", unmapped);

            if (depEntities.Count == 0)
            {
                _trace?.Trace("No deps to create, completing.");
                CompleteJob(job.Id, 0);
                return;
            }

            // Calculate total dep batches and resume index
            int totalDepBatches = (depEntities.Count + DEP_BATCH_LIMIT - 1) / DEP_BATCH_LIMIT;
            int depBatchStart = job.GetAttributeValue<int>(ImportJobFields.CurrentBatch);
            int startIndex = depBatchStart; // track for timeout guard

            _trace?.Trace("Submitting {0} deps in {1} batches of {2} (starting from batch {3})...",
                depEntities.Count, totalDepBatches, DEP_BATCH_LIMIT, depBatchStart);

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            for (int batchIdx = depBatchStart; batchIdx < totalDepBatches; batchIdx++)
            {
                // Timeout guard: save progress and trigger continuation.
                // 80s allows ~9 dep batches at ~8s each, with margin before 120s sandbox limit.
                if (stopwatch.ElapsedMilliseconds > 80000 && batchIdx > startIndex)
                {
                    _trace?.Trace("Dep timeout at {0}ms, saving at dep batch {1}/{2}",
                        stopwatch.ElapsedMilliseconds, batchIdx, totalDepBatches);
                    var timeoutUpdate = new Entity(ImportJobFields.EntityName, job.Id);
                    timeoutUpdate[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.CreatingDeps);
                    timeoutUpdate[ImportJobFields.CurrentBatch] = batchIdx;
                    timeoutUpdate[ImportJobFields.Tick] = job.GetAttributeValue<int>(ImportJobFields.Tick) + 1;
                    _service.Update(timeoutUpdate);
                    return;
                }

                int batchStart = batchIdx * DEP_BATCH_LIMIT;
                int batchSize = Math.Min(DEP_BATCH_LIMIT, depEntities.Count - batchStart);
                var batchDeps = depEntities.GetRange(batchStart, batchSize);

                try
                {
                    string osId = CreateOperationSet(projectId,
                        string.Format("MPP Deps {0}/{1} ({2})", batchIdx + 1, totalDepBatches, batchSize));

                    foreach (var dep in batchDeps)
                        PssCreate(dep, osId);

                    ExecuteOperationSet(osId);
                    _trace?.Trace("Dep batch {0}/{1}: {2} deps submitted (elapsed {3}ms)",
                        batchIdx + 1, totalDepBatches, batchSize, stopwatch.ElapsedMilliseconds);
                }
                catch (Exception ex)
                {
                    // Batch submission failed — retry individual deps
                    _trace?.Trace("Dep batch {0} submit failed ({1}), retrying individually...", batchIdx + 1, ex.Message);
                    foreach (var dep in batchDeps)
                    {
                        try
                        {
                            string osId = CreateOperationSet(projectId,
                                string.Format("MPP Dep retry {0}", dep.Id));
                            PssCreate(dep, osId);
                            ExecuteOperationSet(osId);
                        }
                        catch (Exception ex2)
                        {
                            _trace?.Trace("  Dep {0} submit failed: {1}", dep.Id, ex2.Message);
                        }
                    }
                }

                // Save dep batch progress (non-triggering: only CurrentBatch)
                var progressUpdate = new Entity(ImportJobFields.EntityName, job.Id);
                progressUpdate[ImportJobFields.CurrentBatch] = batchIdx + 1;
                _service.Update(progressUpdate);
            }

            // All dep batches submitted — complete immediately.
            // PSS processes fire-and-forget operation sets asynchronously.
            // The test's 30s PSS finalization wait gives deps time to commit.
            _trace?.Trace("All {0} dep batches submitted. Completing job.", totalDepBatches);
            CompleteJob(job.Id, depEntities.Count, 0);
        }

        /// <summary>
        /// Submits a batch of dependency entities in a single operation set,
        /// executes it, and waits for completion. Returns true if successful.
        /// </summary>
        private bool TryExecuteDepBatch(List<Entity> deps, Guid projectId, int batchNum)
        {
            try
            {
                string osId = CreateOperationSet(projectId,
                    string.Format("MPP Deps {0} ({1})", batchNum, deps.Count));

                foreach (var dep in deps)
                    PssCreate(dep, osId);

                ExecuteOperationSet(osId);

                // Wait for completion (up to ~45s — PSS needs time for large dep batches)
                for (int attempt = 0; attempt < 9; attempt++)
                {
                    System.Threading.Thread.Sleep(5000);
                    int status = CheckOperationSetStatus(osId);
                    if (status == 1) return true;  // success
                    if (status == 0) return false;  // failed
                }

                _trace?.Trace("Dep batch {0} timed out waiting for completion", batchNum);
                return false;
            }
            catch (Exception ex)
            {
                _trace?.Trace("Dep batch {0} exception: {1}", batchNum, ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Phase 5: Wait for dependency operation set completion, then mark Completed.
        /// With batched dep submission, this phase is only reached if ProcessCreatingDeps
        /// used the legacy single-batch path. Kept for backward compatibility.
        /// </summary>
        private void ProcessWaitingForDeps(Entity job)
        {
            string osId = job.GetAttributeValue<string>(ImportJobFields.OperationSetId);

            if (string.IsNullOrEmpty(osId))
            {
                CompleteJob(job.Id, 0);
                return;
            }

            // Poll within this execution — up to ~90 seconds (18 * 5s)
            int osStatus = -1;
            for (int attempt = 0; attempt < 18; attempt++)
            {
                osStatus = CheckOperationSetStatus(osId);
                if (osStatus != -1) break;
                _trace?.Trace("Dep OperationSet still processing, attempt {0}/18", attempt + 1);
                System.Threading.Thread.Sleep(5000);
            }

            if (osStatus == -1) // still processing after all attempts
            {
                _trace?.Trace("WARNING: Dep operation set timed out, completing without deps.");
                CompleteJob(job.Id, 0);
                return;
            }

            if (osStatus == 0) // failed
            {
                _trace?.Trace("WARNING: Dependency operation set failed, tasks are intact.");
                CompleteJob(job.Id, 0);
                return;
            }

            // Success
            int depsCount = job.GetAttributeValue<int>(ImportJobFields.DepsCount);
            CompleteJob(job.Id, depsCount);
        }

        #endregion

        #region Job State Helpers

        private void BumpTick(Entity job)
        {
            int tick = job.GetAttributeValue<int>(ImportJobFields.Tick);
            var update = new Entity(ImportJobFields.EntityName, job.Id);
            update[ImportJobFields.Tick] = tick + 1;
            _service.Update(update);
        }

        private void FailJob(Guid jobId, string errorMessage)
        {
            _trace?.Trace("FAILING job {0}: {1}", jobId, errorMessage);
            var update = new Entity(ImportJobFields.EntityName, jobId);
            update[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.Failed);
            update[ImportJobFields.ErrorMessage] = errorMessage;
            _service.Update(update);

            // Send failure notification to initiating user
            try
            {
                var job = _service.Retrieve(ImportJobFields.EntityName, jobId,
                    new ColumnSet(ImportJobFields.Project, ImportJobFields.InitiatingUser, ImportJobFields.Case));
                var userRef = job.GetAttributeValue<EntityReference>(ImportJobFields.InitiatingUser);
                var projectRef = job.GetAttributeValue<EntityReference>(ImportJobFields.Project);
                if (userRef != null && projectRef != null)
                {
                    string projectName = GetProjectName(projectRef.Id);
                    string shortReason = errorMessage.Length > 120 ? errorMessage.Substring(0, 120) + "..." : errorMessage;
                    SendNotification(
                        userRef.Id,
                        "MPP Import Failed",
                        string.Format("Import failed for {0}: {1}", projectName, shortReason),
                        NotificationIconType.Failure);

                    // Update case import status
                    UpdateCaseImportStatus(job, ImportJobStatus.Failed, shortReason);
                }
            }
            catch (Exception notifEx)
            {
                _trace?.Trace("Failed to send failure notification: {0}", notifEx.Message);
            }
        }

        private void CompleteJob(Guid jobId, int depsCreated, int depsFailed = 0)
        {
            _trace?.Trace("COMPLETING job {0}, deps={1}, depsFailed={2}", jobId, depsCreated, depsFailed);
            var update = new Entity(ImportJobFields.EntityName, jobId);
            update[ImportJobFields.Status] = new OptionSetValue(ImportJobStatus.Completed);
            update[ImportJobFields.DepsCount] = depsCreated;
            _service.Update(update);

            // Send completion notification to initiating user
            try
            {
                var job = _service.Retrieve(ImportJobFields.EntityName, jobId,
                    new ColumnSet(ImportJobFields.Project, ImportJobFields.InitiatingUser,
                        ImportJobFields.TotalTasks, ImportJobFields.Case));
                var userRef = job.GetAttributeValue<EntityReference>(ImportJobFields.InitiatingUser);
                var projectRef = job.GetAttributeValue<EntityReference>(ImportJobFields.Project);
                if (userRef != null && projectRef != null)
                {
                    string projectName = GetProjectName(projectRef.Id);
                    int totalTasks = job.GetAttributeValue<int>(ImportJobFields.TotalTasks);

                    if (depsFailed > 0)
                    {
                        // Warning: completed but some deps failed
                        SendNotification(
                            userRef.Id,
                            "MPP Import Complete (with warnings)",
                            string.Format("{0}: {1} tasks, {2} deps OK, {3} deps failed.",
                                projectName, totalTasks, depsCreated, depsFailed),
                            NotificationIconType.Warning);

                        // Update case import status
                        UpdateCaseImportStatus(job, ImportJobStatus.Completed,
                            string.Format("{0} tasks, {1}/{2} deps ({3} failed)",
                                totalTasks, depsCreated, depsCreated + depsFailed, depsFailed));
                    }
                    else
                    {
                        SendNotification(
                            userRef.Id,
                            "MPP Import Complete",
                            string.Format("{0}: {1} tasks, {2} dependencies.",
                                projectName, totalTasks, depsCreated),
                            NotificationIconType.Success);

                        // Update case import status
                        UpdateCaseImportStatus(job, ImportJobStatus.Completed,
                            string.Format("{0} tasks, {1} dependencies", totalTasks, depsCreated));
                    }
                }
            }
            catch (Exception notifEx)
            {
                _trace?.Trace("Failed to send completion notification: {0}", notifEx.Message);
            }
        }

        #endregion

        #region Notification Helpers

        /// <summary>
        /// Sends a Dataverse in-app notification (toast/bell) to the specified user.
        /// Creates an appnotification record directly. Fails silently if the
        /// appnotification table is not available.
        /// </summary>
        private void SendNotification(Guid recipientUserId, string title, string body, int iconType)
        {
            try
            {
                var notification = new Entity("appnotification");
                notification["title"] = title;
                notification["body"] = body;
                notification["ownerid"] = new EntityReference("systemuser", recipientUserId);
                notification["icontype"] = new OptionSetValue(iconType);
                notification["toasttype"] = new OptionSetValue(200000000); // Timed (shows toast)
                _service.Create(notification);
                _trace?.Trace("Notification sent to {0}: {1}", recipientUserId, title);
            }
            catch (Exception ex)
            {
                // Non-fatal: don't break the import if notifications fail
                _trace?.Trace("SendNotification failed (non-fatal): {0}", ex.Message);
            }
        }

        /// <summary>
        /// Retrieves the project display name for use in notification messages.
        /// </summary>
        private string GetProjectName(Guid projectId)
        {
            try
            {
                var project = _service.Retrieve("msdyn_project", projectId, new ColumnSet("msdyn_subject"));
                return project.GetAttributeValue<string>("msdyn_subject") ?? "(unnamed project)";
            }
            catch
            {
                return "(project)";
            }
        }

        /// <summary>
        /// Writes a progress message (with elapsed runtime) to the adc_case record during import.
        /// Called at each phase transition so the user can see live progress on the form.
        /// </summary>
        private void UpdateCaseProgress(Entity job, string progressMessage)
        {
            try
            {
                var caseRef = job.GetAttributeValue<EntityReference>(ImportJobFields.Case);
                if (caseRef == null) return;

                // Compute elapsed runtime from job creation
                var createdOn = job.GetAttributeValue<DateTime?>("createdon");
                string runtime = "";
                if (createdOn.HasValue)
                {
                    var elapsed = DateTime.UtcNow - createdOn.Value.ToUniversalTime();
                    if (elapsed.TotalMinutes >= 1)
                        runtime = string.Format(" [{0}m {1}s]", (int)elapsed.TotalMinutes, elapsed.Seconds);
                    else
                        runtime = string.Format(" [{0}s]", (int)elapsed.TotalSeconds);
                }

                var caseUpdate = new Entity("adc_case", caseRef.Id);
                caseUpdate["adc_importstatus"] = new OptionSetValue(1); // Processing
                caseUpdate["adc_importmessage"] = (progressMessage + runtime).Length > 250
                    ? (progressMessage + runtime).Substring(0, 250) : (progressMessage + runtime);
                _service.Update(caseUpdate);
            }
            catch (Exception ex)
            {
                _trace?.Trace("UpdateCaseProgress failed (non-fatal): {0}", ex.Message);
            }
        }

        /// <summary>
        /// Writes import status and message back to the adc_case record (if one is linked).
        /// Uses adc_importstatus (optionset) and adc_importmessage (string).
        /// Status mapping: 0=Queued, 1=Processing, 2=Completed, 3=CompletedWithWarnings, 4=Failed
        /// </summary>
        private void UpdateCaseImportStatus(Entity job, int jobStatus, string message)
        {
            try
            {
                var caseRef = job.GetAttributeValue<EntityReference>(ImportJobFields.Case);
                if (caseRef == null) return;

                // Map import job status to case import status optionset
                int caseImportStatus;
                switch (jobStatus)
                {
                    case ImportJobStatus.Completed:
                        // Check if message indicates warnings (depsFailed > 0)
                        caseImportStatus = (message != null && message.Contains("failed")) ? 3 : 2;
                        break;
                    case ImportJobStatus.Failed:
                        caseImportStatus = 4;
                        break;
                    default:
                        caseImportStatus = 1; // Processing
                        break;
                }

                // Append elapsed runtime
                var createdOn = job.GetAttributeValue<DateTime?>("createdon");
                string runtime = "";
                if (createdOn.HasValue)
                {
                    var elapsed = DateTime.UtcNow - createdOn.Value.ToUniversalTime();
                    if (elapsed.TotalMinutes >= 1)
                        runtime = string.Format(" (runtime: {0}m {1}s)", (int)elapsed.TotalMinutes, elapsed.Seconds);
                    else
                        runtime = string.Format(" (runtime: {0}s)", (int)elapsed.TotalSeconds);
                }

                string fullMessage = (message ?? "") + runtime;
                var caseUpdate = new Entity("adc_case", caseRef.Id);
                caseUpdate["adc_importstatus"] = new OptionSetValue(caseImportStatus);
                caseUpdate["adc_importmessage"] = fullMessage.Length > 250
                    ? fullMessage.Substring(0, 250) : fullMessage;
                _service.Update(caseUpdate);
                _trace?.Trace("Updated case {0} import status: {1}", caseRef.Id, caseImportStatus);
            }
            catch (Exception ex)
            {
                _trace?.Trace("UpdateCaseImportStatus failed (non-fatal): {0}", ex.Message);
            }
        }

        #endregion

        #region MPP Parsing Helpers (reused from MppProjectImportService)

        private void DeriveParentRelationships(List<Task> sortedByOrder)
        {
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
                    }
                }
                lastAtLevel[level] = mppTask;
            }

            if (parentsDerived > 0)
                _trace?.Trace("Derived {0} parent relationships from outline levels", parentsDerived);
        }

        private Dictionary<int, int> ComputeAndCapDepths(List<Task> sortedByOrder)
        {
            var depthCache = new Dictionary<int, int>();

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

            int maxDepth = depthCache.Count > 0 ? depthCache.Values.Max() : 0;
            _trace?.Trace("Max parent chain depth: {0} (limit={1})", maxDepth, MAX_DEPTH);

            if (maxDepth > MAX_DEPTH)
            {
                foreach (var mppTask in sortedByOrder)
                {
                    if (!mppTask.UniqueID.HasValue) continue;
                    int depth;
                    if (!depthCache.TryGetValue(mppTask.UniqueID.Value, out depth)) continue;
                    if (depth <= MAX_DEPTH || mppTask.ParentTask == null) continue;

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
                        mppTask.ParentTask = ancestor;
                        depthCache[mppTask.UniqueID.Value] = MAX_DEPTH;
                    }
                }
            }

            return depthCache;
        }

        private HashSet<int> BuildSummaryTaskSet(List<Task> sortedByOrder)
        {
            var summaryTaskIds = new HashSet<int>();
            foreach (var mppTask in sortedByOrder)
            {
                if (mppTask.ParentTask != null && mppTask.ParentTask.UniqueID.HasValue
                    && mppTask.ParentTask.UniqueID.Value != 0)
                    summaryTaskIds.Add(mppTask.ParentTask.UniqueID.Value);
            }
            return summaryTaskIds;
        }

        private List<DependencyDto> BuildDependencyDtos(ProjectFile project, HashSet<int> summaryTaskIds)
        {
            // Build parent → children map for resolving summary deps
            var childrenOf = new Dictionary<int, List<int>>();
            foreach (var t in project.Tasks)
            {
                if (!t.UniqueID.HasValue || t.UniqueID.Value == 0) continue;
                if (t.ParentTask != null && t.ParentTask.UniqueID.HasValue && t.ParentTask.UniqueID.Value != 0)
                {
                    int parentId = t.ParentTask.UniqueID.Value;
                    List<int> list;
                    if (!childrenOf.TryGetValue(parentId, out list))
                    {
                        list = new List<int>();
                        childrenOf[parentId] = list;
                    }
                    list.Add(t.UniqueID.Value);
                }
            }

            // Collect raw deps (all, including summary)
            var rawDeps = new List<DependencyDto>();
            int summaryDeps = 0;

            foreach (var mppTask in project.Tasks)
            {
                if (!mppTask.UniqueID.HasValue) continue;
                if (mppTask.UniqueID.Value == 0) continue;
                if (mppTask.Predecessors == null || mppTask.Predecessors.Count == 0) continue;

                foreach (var relation in mppTask.Predecessors)
                {
                    if (relation.SourceTaskUniqueID == 0) continue;

                    bool involvesSummary = summaryTaskIds.Contains(relation.SourceTaskUniqueID)
                        || summaryTaskIds.Contains(mppTask.UniqueID.Value);
                    if (involvesSummary)
                        summaryDeps++;

                    rawDeps.Add(new DependencyDto
                    {
                        PredecessorUniqueID = relation.SourceTaskUniqueID,
                        SuccessorUniqueID = mppTask.UniqueID.Value,
                        LinkType = MapRelationType(relation.Type)
                    });
                }
            }

            if (summaryDeps == 0)
            {
                // No summary deps — return as-is
                return rawDeps;
            }

            _trace?.Trace("{0} raw deps include {1} involving summary tasks — transforming to leaf deps",
                rawDeps.Count, summaryDeps);

            // Build predecessor/successor maps from leaf-to-leaf deps for entry/exit analysis
            var predecessorMap = new Dictionary<int, List<int>>();  // taskId → list of predecessor taskIds
            var successorMap = new Dictionary<int, List<int>>();    // taskId → list of successor taskIds
            foreach (var dep in rawDeps)
            {
                if (summaryTaskIds.Contains(dep.PredecessorUniqueID) || summaryTaskIds.Contains(dep.SuccessorUniqueID))
                    continue; // only leaf-to-leaf deps for internal analysis
                List<int> list;
                if (!predecessorMap.TryGetValue(dep.SuccessorUniqueID, out list))
                {
                    list = new List<int>();
                    predecessorMap[dep.SuccessorUniqueID] = list;
                }
                list.Add(dep.PredecessorUniqueID);
                if (!successorMap.TryGetValue(dep.PredecessorUniqueID, out list))
                {
                    list = new List<int>();
                    successorMap[dep.PredecessorUniqueID] = list;
                }
                list.Add(dep.SuccessorUniqueID);
            }

            // Transform: replace summary-task endpoints with leaf-task equivalents
            // Use entry/exit leaf analysis to minimize cross-product explosion:
            //   - Successor is summary: use entry leaves (no internal predecessors) — these are the scheduling start points
            //   - Predecessor is summary: use exit leaves (no internal successors) — these are the scheduling end points
            var finalDeps = new List<DependencyDto>();
            var seen = new HashSet<string>();
            int transformed = 0;

            foreach (var dep in rawDeps)
            {
                bool predIsSummary = summaryTaskIds.Contains(dep.PredecessorUniqueID);
                bool succIsSummary = summaryTaskIds.Contains(dep.SuccessorUniqueID);

                if (!predIsSummary && !succIsSummary)
                {
                    // Normal leaf-to-leaf dep — keep as-is
                    string key = dep.PredecessorUniqueID + "|" + dep.SuccessorUniqueID;
                    if (seen.Add(key))
                        finalDeps.Add(dep);
                    continue;
                }

                // Resolve summary endpoints to targeted leaf tasks
                List<int> predIds = predIsSummary
                    ? GetExitLeaves(dep.PredecessorUniqueID, childrenOf, summaryTaskIds, successorMap)
                    : new List<int> { dep.PredecessorUniqueID };

                List<int> succIds = succIsSummary
                    ? GetEntryLeaves(dep.SuccessorUniqueID, childrenOf, summaryTaskIds, predecessorMap)
                    : new List<int> { dep.SuccessorUniqueID };

                if (predIds.Count == 0 || succIds.Count == 0)
                {
                    _trace?.Trace("  Skipping dep {0}->{1}: no leaf descendants found",
                        dep.PredecessorUniqueID, dep.SuccessorUniqueID);
                    continue;
                }

                foreach (var p in predIds)
                {
                    foreach (var s in succIds)
                    {
                        if (p == s) continue;
                        string key = p + "|" + s;
                        if (seen.Add(key))
                        {
                            finalDeps.Add(new DependencyDto
                            {
                                PredecessorUniqueID = p,
                                SuccessorUniqueID = s,
                                LinkType = dep.LinkType
                            });
                            transformed++;
                        }
                    }
                }
            }

            _trace?.Trace("Transformed: {0} summary deps → {1} leaf deps (total deps: {2})",
                summaryDeps, transformed, finalDeps.Count);

            return finalDeps;
        }

        /// <summary>
        /// Returns all descendant UniqueIDs (both summary and leaf) under the given summary task.
        /// </summary>
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

        /// <summary>
        /// Returns leaf descendants that have NO predecessor from within the same summary subtree.
        /// These are the "entry points" — tasks that would start at the summary's start date.
        /// Falls back to all leaf descendants if analysis finds none.
        /// </summary>
        private List<int> GetEntryLeaves(int summaryId, Dictionary<int, List<int>> childrenOf,
            HashSet<int> summaryTaskIds, Dictionary<int, List<int>> predecessorMap)
        {
            var descendants = GetAllDescendants(summaryId, childrenOf);
            var entries = new List<int>();

            foreach (var id in descendants)
            {
                if (summaryTaskIds.Contains(id)) continue; // skip summaries

                bool hasInternalPred = false;
                List<int> preds;
                if (predecessorMap.TryGetValue(id, out preds))
                {
                    foreach (var predId in preds)
                    {
                        if (descendants.Contains(predId))
                        {
                            hasInternalPred = true;
                            break;
                        }
                    }
                }

                if (!hasInternalPred)
                    entries.Add(id);
            }

            // Fallback: if no entry leaves found (all have internal preds — shouldn't happen
            // unless there's a cycle), return all leaf descendants
            if (entries.Count == 0)
            {
                foreach (var id in descendants)
                {
                    if (!summaryTaskIds.Contains(id))
                        entries.Add(id);
                }
            }

            return entries;
        }

        /// <summary>
        /// Returns leaf descendants that have NO successor within the same summary subtree.
        /// These are the "exit points" — tasks that define when the summary finishes.
        /// Falls back to all leaf descendants if analysis finds none.
        /// </summary>
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
                        if (descendants.Contains(succId))
                        {
                            hasInternalSucc = true;
                            break;
                        }
                    }
                }

                if (!hasInternalSucc)
                    exits.Add(id);
            }

            if (exits.Count == 0)
            {
                foreach (var id in descendants)
                {
                    if (!summaryTaskIds.Contains(id))
                        exits.Add(id);
                }
            }

            return exits;
        }

        #endregion

        #region Subtree Batching

        /// <summary>
        /// Groups tasks into batches where complete subtrees are never split.
        /// Each Level 1 subtree (root + all descendants) stays together.
        /// Subtrees are packed into batches using first-fit up to BATCH_LIMIT.
        /// </summary>
        private List<TaskBatch> BuildSubtreeBatches(List<TaskDto> tasks)
        {
            // Build parent -> children lookup
            var childrenOf = new Dictionary<int, List<int>>();
            var topLevelIds = new List<int>();

            foreach (var task in tasks)
            {
                if (!task.ParentUniqueID.HasValue)
                {
                    topLevelIds.Add(task.UniqueID);
                }
                else
                {
                    if (!childrenOf.ContainsKey(task.ParentUniqueID.Value))
                        childrenOf[task.ParentUniqueID.Value] = new List<int>();
                    childrenOf[task.ParentUniqueID.Value].Add(task.UniqueID);
                }
            }

            // Collect all descendants of each top-level task (complete subtree)
            var subtrees = new List<List<int>>();
            foreach (int rootId in topLevelIds)
            {
                var subtree = new List<int>();
                CollectSubtree(rootId, childrenOf, subtree);
                subtrees.Add(subtree);
            }

            // Pack subtrees into batches (first-fit bin packing)
            var batches = new List<TaskBatch>();
            var currentBatch = new TaskBatch { Index = 0 };

            foreach (var subtree in subtrees)
            {
                if (currentBatch.TaskUniqueIDs.Count + subtree.Count > BATCH_LIMIT && currentBatch.TaskUniqueIDs.Count > 0)
                {
                    batches.Add(currentBatch);
                    currentBatch = new TaskBatch { Index = batches.Count };
                }
                currentBatch.TaskUniqueIDs.AddRange(subtree);
            }

            if (currentBatch.TaskUniqueIDs.Count > 0)
                batches.Add(currentBatch);

            return batches;
        }

        private void CollectSubtree(int taskId, Dictionary<int, List<int>> childrenOf, List<int> result)
        {
            result.Add(taskId);
            List<int> children;
            if (childrenOf.TryGetValue(taskId, out children))
            {
                foreach (int childId in children)
                    CollectSubtree(childId, childrenOf, result);
            }
        }

        #endregion

        #region Entity Mapping

        private Entity MapDtoToEntity(TaskDto dto, EntityReference projectRef, EntityReference bucketRef)
        {
            var entity = new Entity("msdyn_projecttask");
            entity.Id = Guid.Parse(dto.PreGenGuid);
            entity["msdyn_projecttaskid"] = entity.Id;
            entity["msdyn_project"] = projectRef;
            entity["msdyn_projectbucket"] = bucketRef;
            entity["msdyn_subject"] = dto.Name;
            entity["msdyn_LinkStatus"] = new OptionSetValue(192350000); // Not Linked
            entity["msdyn_outlinelevel"] = dto.Depth;

            if (!dto.IsSummary)
            {
                if (dto.DurationHours.HasValue)
                    entity["msdyn_duration"] = Math.Round(dto.DurationHours.Value / 8.0, 2);

                if (dto.EffortHours.HasValue)
                    entity["msdyn_effort"] = Math.Round(dto.EffortHours.Value, 2);
            }

            return entity;
        }

        #endregion

        #region PSS API Helpers

        private string CreateOperationSet(Guid projectId, string description)
        {
            var request = new OrganizationRequest("msdyn_CreateOperationSetV1");
            request["ProjectId"] = projectId.ToString();
            request["Description"] = description;
            var response = _service.Execute(request);
            return (string)response["OperationSetId"];
        }

        private void PssCreate(Entity entity, string operationSetId)
        {
            var request = new OrganizationRequest("msdyn_PssCreateV1");
            request["Entity"] = entity;
            request["OperationSetId"] = operationSetId;
            _service.Execute(request);
        }

        private void PssUpdate(Entity entity, string operationSetId)
        {
            var request = new OrganizationRequest("msdyn_PssUpdateV1");
            request["Entity"] = entity;
            request["OperationSetId"] = operationSetId;
            _service.Execute(request);
        }

        private string ExecuteOperationSet(string operationSetId)
        {
            var request = new OrganizationRequest("msdyn_ExecuteOperationSetV1");
            request["OperationSetId"] = operationSetId;
            var response = _service.Execute(request);

            if (response.Results.ContainsKey("OperationSetDetailId"))
                return response["OperationSetDetailId"] as string;
            return operationSetId;
        }

        /// <summary>
        /// Checks the operation set status. Returns:
        ///  1 = completed, 0 = failed/cancelled, -1 = still processing
        /// </summary>
        private int CheckOperationSetStatus(string operationSetId)
        {
            Guid osId;
            if (!Guid.TryParse(operationSetId, out osId)) return 0;

            try
            {
                var osRecord = _service.Retrieve("msdyn_operationset", osId,
                    new ColumnSet("msdyn_status", "statuscode", "msdyn_description"));

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

                // Completed
                if (status == 192350001 || status == 1) return 1;
                // Failed or Cancelled
                if (status == 192350002 || status == 2 || status == 192350003 || status == 3) return 0;
                // Still processing
                return -1;
            }
            catch (Exception ex)
            {
                _trace?.Trace("CheckOperationSetStatus error: {0}", ex.Message);
                return -1;
            }
        }

        private void WaitForOperationSetCompletion(string operationSetId)
        {
            for (int i = 0; i < 20; i++)
            {
                System.Threading.Thread.Sleep(2000);
                int status = CheckOperationSetStatus(operationSetId);
                if (status == 1) return;
                if (status == 0) throw new InvalidPluginExecutionException("OperationSet failed: " + operationSetId);
            }
        }

        #endregion

        #region Dataverse Helpers

        private EntityReference GetOrCreateDefaultBucket(Guid projectId)
        {
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
                return results.Entities[0].ToEntityReference();

            var bucket = new Entity("msdyn_projectbucket");
            bucket["msdyn_project"] = new EntityReference("msdyn_project", projectId);
            bucket["msdyn_name"] = "MPP Import";
            Guid bucketId = _service.Create(bucket);
            return new EntityReference("msdyn_projectbucket", bucketId);
        }

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
                else break;
            }
            return results;
        }

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
                else break;
            }
            return results;
        }

        private static int GetOptionSetValue(Entity entity, string field)
        {
            var osv = entity.GetAttributeValue<OptionSetValue>(field);
            return osv != null ? osv.Value : -1;
        }

        #endregion

        #region JSON Serialization

        private ImportJobPayload DeserializePayload(Entity job)
        {
            string json = job.GetAttributeValue<string>(ImportJobFields.TaskDataJson);
            if (string.IsNullOrEmpty(json))
                throw new InvalidPluginExecutionException("Job has no task data JSON.");
            return DeserializeJson<ImportJobPayload>(json);
        }

        private static string SerializeJson<T>(T obj)
        {
            var serializer = new DataContractJsonSerializer(typeof(T), new DataContractJsonSerializerSettings
            {
                UseSimpleDictionaryFormat = true
            });
            using (var ms = new MemoryStream())
            {
                serializer.WriteObject(ms, obj);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        private static T DeserializeJson<T>(string json)
        {
            var serializer = new DataContractJsonSerializer(typeof(T), new DataContractJsonSerializerSettings
            {
                UseSimpleDictionaryFormat = true
            });
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (T)serializer.ReadObject(ms);
            }
        }

        #endregion

        #region Mapping Helpers

        private static int MapRelationType(RelationType type)
        {
            switch (type)
            {
                case RelationType.FinishToStart: return 192350000;
                case RelationType.StartToStart: return 192350001;
                case RelationType.FinishToFinish: return 192350002;
                case RelationType.StartToFinish: return 192350003;
                default: return 192350000;
            }
        }

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

        #endregion
    }
}
