# Async Chunked MPP Import — Architecture & Plan

## Overview

The existing `MppProjectImportService` + `ImportMppToProjectActivity` workflow activity runs
the entire MPP import (parse → create tasks → poll GUIDs → create deps) in a **single execution**.
This hits Dataverse's **hard 2-minute sandbox timeout** for projects with 100+ tasks.

This document describes a new **async chunked import** architecture that splits the work into
phases, each staying well under 2 minutes.

The existing service and workflow activity are **preserved for backward compatibility**.

---

## Architecture

### Trigger Flow

```
adc_case (created with adc_casetemplate selected)
  │
  │  Plugin / process on adc_case create:
  │    1. Creates msdyn_project
  │    2. Links adc_case ↔ msdyn_project (bidirectional)
  │    3. Creates adc_mppimportjob (Queued)
  │    4. Sends "Import started" in-app notification
  │
  ▼
adc_mppimportjob
  │  triggers async plugin chain
  ▼
MppImportJobPlugin (phases 0-5)
  │  on Completed/Failed:
  │    Sends "Import completed/failed" in-app notification
  ▼
Done
```

### Component Diagram

```
┌─────────────────────────┐
│  ImportMppToProjectActivity (existing — unchanged)
│  Runs full import in one shot. Works for ≤100 tasks.
└─────────────────────────┘

┌─────────────────────────┐         ┌─────────────────────────┐
│  StartMppImportActivity │────────>│  adc_mppimportjob       │
│  (NEW workflow activity)│ creates │  (NEW custom table)     │
│                         │         │  Status = Queued        │
│  • Parse MPP            │         │  Stores serialized data │
│  • Serialize to JSON    │         └────────────┬────────────┘
│  • Create job record    │                      │ triggers
│  • Return immediately   │                      ▼
└─────────────────────────┘         ┌─────────────────────────┐
                                    │  MppImportJobPlugin     │
                                    │  (NEW async plugin)     │
                                    │                         │
                                    │  Registered on Create & │
                                    │  Update of              │
                                    │  adc_mppimportjob       │
                                    │                         │
                                    │  Processes one phase    │
                                    │  per execution, using   │
                                    │  in-execution polling   │
                                    │  (Thread.Sleep loops)   │
                                    │  for wait phases.       │
                                    │                         │
                                    │  Self-triggers next     │
                                    │  phase by updating      │
                                    │  adc_status.            │
                                    └─────────────────────────┘
```

### Parent Entity: `adc_case`

The import is triggered in the context of an `adc_case` record:

| Relationship | Description |
|---|---|
| `adc_case` → `adc_adccasetemplate` | Case template selected on case creation (contains MPP file) |
| `adc_case` → `msdyn_project` | Project created by the import process |
| `msdyn_project` → `adc_case` | Back-link from project to originating case |
| `adc_mppimportjob` → `msdyn_project` | Job tracks which project it's importing into |
| `adc_mppimportjob` → `adc_adccasetemplate` | Job tracks which template MPP was sourced from |

The `adc_case` form is the primary UI context — users create a case, select a template,
and expect to see progress/completion of the import via **in-app notifications** on that form.

---

## Custom Table: `adc_mppimportjob`

| Field (Schema Name)       | Type             | Description                                        |
|---------------------------|------------------|----------------------------------------------------|
| `adc_mppimportjobid`     | Uniqueidentifier | Primary key                                        |
| `adc_name`               | Text (200)       | Display name (e.g., "Import Project2.mpp")         |
| `adc_project`            | Lookup           | Target `msdyn_project`                             |
| `adc_casetemplate`       | Lookup           | Source `adc_adccasetemplate` (optional)             |
| `adc_status`             | OptionSet        | See Status values below                            |
| `adc_phase`              | Int              | Current processing phase (1-4)                     |
| `adc_currentbatch`       | Int              | Current batch index (0-based)                      |
| `adc_totalbatches`       | Int              | Total number of task batches                       |
| `adc_totaltasks`         | Int              | Total task count from MPP                          |
| `adc_createdcount`       | Int              | Tasks created so far                               |
| `adc_depscount`          | Int              | Dependencies created                               |
| `adc_tick`               | Int              | Bumped to re-trigger plugin on same status          |
| `adc_taskdatajson`       | Multiline (max)  | Serialized task definitions (JSON)                  |
| `adc_taskidmapjson`      | Multiline (max)  | UniqueID → pre-gen GUID map (JSON)                  |
| `adc_batchesjson`        | Multiline (max)  | Subtree batch assignments (JSON)                    |
| `adc_operationsetid`     | Text (100)       | Current PSS operation set ID                        |
| `adc_projectstartdate`   | DateTime         | Optional project start date override                |
| `adc_errormessage`       | Multiline (max)  | Error details if failed                             |

### Status OptionSet Values (`adc_status`)

| Value | Label           | Description                                              |
|-------|-----------------|----------------------------------------------------------|
| 0     | Queued          | Job created, ready for processing                        |
| 1     | CreatingTasks   | Submitting PSS task creates for current batch             |
| 2     | WaitingForTasks | Polling for PSS operation set to complete                 |
| 3     | PollingGUIDs    | Querying CRM for actual task GUIDs                        |
| 4     | CreatingDeps    | Submitting PSS dependency creates                         |
| 5     | WaitingForDeps  | Polling for dependency operation set to complete           |
| 6     | Completed       | All tasks and dependencies imported successfully           |
| 7     | Failed          | Import failed — see adc_errormessage                      |

---

## Processing Phases (each ≤ 2 minutes)

### Phase 0: Initialization (in StartMppImportActivity)

1. Download MPP bytes from `adc_adccasetemplate` file column
2. Parse with `MppFileReader`
3. Derive parent relationships, compute depth, cap at 10 levels
4. Build summary task set, subtree batches
5. Serialize everything to JSON
6. Create `adc_mppimportjob` with status = Queued
7. Return immediately → workflow completes fast

### Phase 1: CreateTasks (per batch)

1. Read `adc_batchesjson` to get current batch's task UniqueIDs
2. Create PSS operation set
3. For each task in batch: `PssCreate` with parent links + outline level
4. `ExecuteOperationSet`
5. Store `adc_operationsetid`
6. Update status → WaitingForTasks

### Phase 2: WaitForTasks

1. Check operation set completion status
2. If not done → bump `adc_tick` to re-trigger (natural async delay ~2-5s)
3. If done → check if more batches remain
   - Yes → increment `adc_currentbatch`, status → CreatingTasks
   - No → status → PollingGUIDs

### Phase 3: PollingGUIDs

1. Query CRM for all project tasks
2. Build `actualTaskIdMap` (UniqueID → CRM GUID) by name matching
3. Store in `adc_taskidmapjson`
4. Status → CreatingDeps

### Phase 4: CreateDeps

1. Read task ID map and dependency definitions
2. Build dependency entities using actual CRM GUIDs
3. Skip summary task deps, duplicates
4. Create PSS operation set, execute
5. Status → WaitingForDeps

### Phase 5: WaitForDeps

1. Check dep operation set completion
2. If not done → bump tick
3. If done → status → Completed

---

## Subtree Batching Algorithm

Tasks are grouped into batches where **complete subtrees are never split**.
This ensures parent links (which must be in the same PSS operation set) are always
within the same batch.

```
1. Group tasks by their Level 1 ancestor (top-level subtree root)
2. For each Level 1 subtree, count total descendants
3. Pack subtrees into batches using first-fit:
   - While current batch size + next subtree size ≤ BATCH_LIMIT (e.g., 100):
     add subtree to current batch
   - Otherwise: start a new batch
4. Edge case: if a single subtree exceeds BATCH_LIMIT, it gets its own batch
   (PSS can handle up to 200 ops)
```

Dependencies are NOT affected by batching — they're always created in Phase 4
after all task batches have been committed and GUIDs retrieved.

---

## Self-Triggering & Execution Model

The async plugin chains phases by **updating `adc_status`** on the job record:

- Each status change triggers the Update plugin step → next phase runs
- **Wait phases (WaitingForTasks, PollingGUIDs, WaitingForDeps)** use **in-execution
  `Thread.Sleep` polling loops** (up to ~90 seconds, 18 × 5s) instead of tick-based
  self-triggering. Tick-only updates are unreliable — Dataverse coalesces/throttles
  rapid updates to the same record within async plugin chains.
- Plugin filters: registered on **Create** and **Update** of `adc_mppimportjob`,
  filtered on `adc_status` and `adc_tick` attributes

### Depth Escalation

Each phase transition increments the execution depth by 1 (status update within
plugin → new async execution at depth+1). Each batch cycle (CreatingTasks →
WaitingForTasks) adds ~2 depth levels. **Depth guard is set to `> 50`**, supporting
up to ~24 batches (~2,400 tasks). Terminal state checks, Thread.Sleep caps, and
FailJob timeouts provide additional infinite-loop prevention.

### Plugin Step Impersonation

Both plugin steps (Create and Update) must have `impersonatinguserid` set to a
licensed Project Operations user. S2S app users do not have PSS licenses, and
`CreateOrganizationService(null)` (SYSTEM) fails because PSS requires an AAD user ID.

### Infinite Loop Prevention

- In-execution polling has a hard cap of 18 attempts (90 seconds)
- If exceeded → status = Failed with timeout error message
- Plugin checks `adc_status` is not Completed/Failed before processing
- Depth guard at 10 prevents runaway chains

---

## New Files

| File | Purpose |
|------|---------|
| `ADC.MppImport/Services/MppAsyncImportService.cs` | Core logic: parsing, batching, per-phase processing |
| `ADC.MppImport/Services/MppImportJobData.cs` | JSON-serializable DTOs for task data, batches, ID maps |
| `ADC.MppImport/Workflows/StartMppImportActivity.cs` | New workflow activity (lightweight, creates job record) |
| `ADC.MppImport/Plugins/MppImportJobPlugin.cs` | Async plugin on adc_mppimportjob Create/Update |
| `scripts/create-import-job-table.ps1` | PowerShell script to create `adc_mppimportjob` table via pac CLI |
| `scripts/deploy-plugin.ps1` | PowerShell script to build, deploy DLL, and register plugin steps |
| `scripts/register-plugin-steps.ps1` | Plugin Registration Tool automation for step registration |

---

## Deployment Scripts

### Prerequisites

- **Power Platform CLI** (`pac`) installed and authenticated
- **Plugin Registration Tool** (from NuGet or SDK tools)
- App registration with System Administrator or Customizer role

### Table Creation (`scripts/create-import-job-table.ps1`)

Uses `pac modelbuilder` or direct Web API calls to create:
1. `adc_mppimportjob` table with all fields
2. OptionSet for `adc_status`
3. Relationships (lookups to `msdyn_project`, `adc_adccasetemplate`)

### Plugin Deployment (`scripts/deploy-plugin.ps1`)

1. `msbuild` the ADC.MppImport project in Release mode
2. Use `pac plugin push` to upload the assembly to D365
3. Register plugin steps:
   - **Step 1**: Create of `adc_mppimportjob`, Async, Post-Operation
   - **Step 2**: Update of `adc_mppimportjob`, Async, Post-Operation,
     filtering attributes: `adc_status,adc_tick`

### Redeployment

After code changes:
```powershell
.\scripts\deploy-plugin.ps1 -Environment https://orgfbe0a613.crm6.dynamics.com
```

This rebuilds, pushes the updated DLL, and re-registers steps if needed.

---

## Existing Code Preserved

| File | Status |
|------|--------|
| `Services/MppProjectImportService.cs` | **Unchanged** — existing workflow activity continues to use it |
| `Workflows/ImportMppToProjectActivity.cs` | **Unchanged** — existing workflows keep working |
| `Shared/BaseCodeActivity.cs` | **Unchanged** — shared base class |

---

## Testing Strategy

1. **Unit test**: Verify subtree batching produces correct groups
2. **Integration test** (`asynctest` CLI mode): End-to-end via `Program.cs`
   - Creates case template, uploads MPP, creates project
   - Creates `adc_mppimportjob` via `MppAsyncImportService.InitializeJob`
   - Polls job status through all phases
   - Validates tasks, hierarchy, dependencies in D365
   - Cleans up test records
3. **Large file test**: `exampleFiles/Kas Case Management Template 1.mpp`
   — one of the largest real-world templates, good for verifying batching
   and multi-batch phase transitions
4. **Failure recovery**: Test that Failed status contains useful error info
5. **Notification test**: Verify in-app notifications appear for the initiating
   user when import starts, completes, and fails

---

## In-App Notifications

Use the Dataverse **in-app notification** API to send toast/bell notifications to the
user who initiated the import (from the `adc_case` form). This provides real-time
feedback without requiring the user to manually poll or refresh.

### Notification Points

| When | Message | Icon/Type |
|------|---------|----------|
| Import starts (job record created) | "MPP import started for **{ProjectName}** — running in background" | Info |
| Import completes | "MPP import completed for **{ProjectName}** ({TaskCount} tasks, {DepCount} dependencies)" | Success |
| Import fails | "MPP import failed for **{ProjectName}**: {ShortReason} (open job for details)" | Error/Warning |

### Implementation

Notifications are created by calling `msdyn_SendAppNotification` (or creating
`appnotification` records directly via the Dataverse API).

```csharp
// Example: Send in-app notification from within the plugin
var notif = new OrganizationRequest("SendAppNotification");
notif["Title"] = "MPP Import Complete";
notif["Body"] = string.Format("Import completed for {0} ({1} tasks)", projectName, taskCount);
notif["Recipient"] = new EntityReference("systemuser", initiatingUserId);
notif["IconType"] = 100000000; // Success
// ... additional params as needed
_service.Execute(notif);
```

**Key details:**

- **Recipient**: The user who created the `adc_case` record (stored as `createdby`
  on the case or passed through to the job record). This is NOT the impersonated
  plugin user — it's the original initiating user.
- **Where to send from**:
  - "Import started" → sent during job initialization (Phase 0 / `ProcessQueued`)
  - "Import completed" → sent in `CompleteJob()`
  - "Import failed" → sent in `FailJob()`
- **Link back to case**: Include an `adc_case` record URL or entity reference in
  the notification body so the user can navigate directly to the case form.
- **No external dependencies**: In-app notifications are a native Dataverse feature
  (model-driven apps). No Azure Functions, Power Automate, or custom connectors needed.

### New Field on `adc_mppimportjob`

| Field (Schema Name) | Type | Description |
|---|---|---|
| `adc_case` | Lookup | Link to originating `adc_case` record |
| `adc_initiatinguser` | Lookup | The user who created the case (notification recipient) |

These fields allow the async plugin to know:
1. Which case the import belongs to (for notification body/links)
2. Who to send the notification to (may differ from the impersonated plugin user)

### UX on the `adc_case` Form

- User creates a case → selects template → saves
- Immediately sees a toast: "MPP import started for {Project} — running in background"
- Can continue working on other records
- When import finishes: bell notification appears with result
- Clicking the notification navigates to the case or project

---

## Open Questions

1. **Max multiline text size**: Dataverse multiline text fields support up to 1,048,576 characters.
   A 500-task project serializes to ~100KB JSON — well within limits.
2. **Concurrent imports**: Should we prevent multiple simultaneous imports to the same project?
   Could add a check in Phase 1 for existing Queued/Processing jobs.
3. **Retry logic**: If a batch fails, should we retry automatically or mark as Failed?
   Initial implementation: mark Failed. Can add retry later.
4. **Notification fallback**: If the initiating user ID is not available (e.g., S2S-only
   trigger), should we fall back to notifying a team or skipping the notification?
5. **Notification action buttons**: Should the completion notification include a button
   to "Open Project" or "View Tasks"? The `SendAppNotification` API supports custom actions.
