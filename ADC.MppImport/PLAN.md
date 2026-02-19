# ADC.MppImport — MPP to Project Operations Import Workflow

## Overview
A self-contained Dynamics 365 CRM workflow activity assembly that reads Microsoft Project (.mpp) files 
from a custom entity's file field and creates/updates project tasks in Project Operations (`msdyn_project` / `msdyn_projecttask`).

**Zero external NuGet dependencies** beyond the standard CRM SDK assemblies (`Microsoft.CrmSdk.CoreAssemblies`, `Microsoft.CrmSdk.Workflow`).

## Architecture

```
ADC.MppImport/
├── PLAN.md                          # This file
├── ADC.MppImport.csproj           # Old-style csproj, .NET 4.7.2, CRM SDK only
├── packages.config                  # CRM SDK NuGet refs only
├── Properties/
│   └── AssemblyInfo.cs
├── Shared/
│   ├── BaseCodeActivity.cs          # Copied from ADC.Xero.WFShared
│   └── SDKHelpers.cs                # Copied from ADC.Xero.WFShared (namespace updated)
├── MppReader/
│   ├── Ole2/
│   │   └── CompoundFile.cs          # Minimal self-contained OLE2/CBF reader (replaces OpenMcdf)
│   ├── Common/
│   │   └── ByteArrayHelper.cs       # Byte manipulation helpers
│   ├── Model/
│   │   ├── Duration.cs
│   │   ├── Enums.cs
│   │   ├── ProjectCalendar.cs
│   │   ├── ProjectFile.cs
│   │   ├── ProjectProperties.cs
│   │   ├── Rate.cs
│   │   ├── Relation.cs
│   │   ├── Resource.cs
│   │   ├── ResourceAssignment.cs
│   │   ├── Task.cs
│   │   └── TimeUnit.cs
│   └── Mpp/
│       ├── CompObj.cs
│       ├── DocumentInputStreamFactory.cs
│       ├── FieldMap.cs
│       ├── FixedData.cs
│       ├── FixedMeta.cs
│       ├── IMppVariantReader.cs
│       ├── Mpp8Reader.cs
│       ├── Mpp9Reader.cs
│       ├── Mpp12Reader.cs
│       ├── Mpp14Reader.cs
│       ├── MppFileReader.cs          # Modified to use Ole2.CompoundFile instead of OpenMcdf
│       ├── MppUtility.cs
│       ├── ProjectPropertiesReader.cs
│       ├── Props.cs / Props8-14.cs
│       ├── Var2Data.cs
│       └── VarMeta.cs
├── Services/
│   └── MppProjectImportService.cs   # Testable business logic (no CRM dependency in signature)
└── Workflows/
    └── ImportMppToProjectActivity.cs # Thin workflow shell — delegates to MppProjectImportService
```

## Entity Schema

### Input: `adc_adccasetemplate`
- `adc_templatemsprojectmppfile` — File column containing the .mpp binary

### Target: `msdyn_project` (Project Operations)
- Passed as EntityReference input to the workflow

### Upsert Target: `msdyn_projecttask`
Key field mapping (MPP Task → msdyn_projecttask):

| MPP Task Field        | msdyn_projecttask Field        | Notes                              |
|-----------------------|--------------------------------|------------------------------------|
| UniqueID              | msdyn_msprojectclientid        | String, used as match key          |
| Name                  | msdyn_subject                  | Task name                          |
| Duration              | msdyn_duration                 | Decimal, in days                   |
| Start                 | msdyn_scheduledstart           | DateTime                           |
| Finish                | msdyn_scheduledend             | DateTime                           |
| OutlineLevel          | msdyn_outlinelevel             | Int                                |
| WBS                   | msdyn_wbsid                    | String (e.g. "1.2.3")             |
| PercentComplete       | msdyn_progress                 | Decimal (0-100)                    |
| ParentTaskUniqueID    | msdyn_parenttask               | EntityReference (looked up)        |
| (project ref)         | msdyn_project                  | EntityReference to msdyn_project   |

## Implementation Steps

### Step 1: Build Minimal OLE2 Reader (`Ole2/CompoundFile.cs`)
Replace OpenMcdf with ~350 lines implementing:
- Parse CBF header (sector size, FAT locations, directory start, mini-stream cutoff)
- Read FAT chain, DIFAT, Mini-FAT
- Navigate directory tree (storage/stream entries, 128-byte records)
- Read stream data (regular sectors + mini-stream)
- API surface: `CompoundFile(byte[])`, `RootStorage`, `GetStorage(name)`, `GetStream(name)`, `GetData()`, `VisitEntries()`

### Step 2: Create Project & Copy Files
- Create old-style .csproj targeting .NET 4.7.2
- Copy BaseCodeActivity.cs and SDKHelpers.cs (update namespace to `ADC.MppImport`)
- Copy all MppReader .cs files, update namespaces
- Update MppFileReader.cs to reference `Ole2.CompoundFile` instead of `OpenMcdf.CompoundFile`

### Step 3: Build MppProjectImportService
Testable business logic class:
```csharp
public class MppProjectImportService
{
    public ImportResult ImportMppToProject(
        IOrganizationService service,
        ITracingService trace,
        Guid templateId,
        Guid projectId)
    { ... }
}
```
- Downloads MPP file bytes from `adc_adccasetemplate.adc_templatemsprojectmppfile`
- Parses with MppFileReader
- Queries existing `msdyn_projecttask` records for the project
- Upserts tasks based on UniqueID match

### Step 4: Build Workflow Activity Shell
Thin wrapper:
```csharp
public class ImportMppToProjectActivity : BaseCodeActivity
{
    [Input("Case Template"), ReferenceTarget("adc_adccasetemplate"), RequiredArgument]
    public InArgument<EntityReference> CaseTemplate { get; set; }

    [Input("Target Project"), ReferenceTarget("msdyn_project"), RequiredArgument]
    public InArgument<EntityReference> TargetProject { get; set; }

    [Output("Tasks Created")]
    public OutArgument<int> TasksCreated { get; set; }

    [Output("Tasks Updated")]
    public OutArgument<int> TasksUpdated { get; set; }
}
```

### Step 5: Verify Build
- 0 external NuGet dependencies beyond CRM SDK
- 0 warnings, 0 errors
- Assembly can be registered via Plugin Registration Tool

---

## Current Status (v1.0.21.0 — 2026-02-19)

### ✅ WORKS — Confirmed in testing

| Feature | Notes |
|---------|-------|
| MPP download & parsing | Downloads from `adc_templatemsprojectmppfile`, parses tasks/resources/predecessors correctly |
| Task creation via PssCreate | Tasks created with correct names, duration, effort. Batched in 200-op operation sets |
| Task update via PssUpdate | Existing tasks matched by `msdyn_msprojectclientid` and updated |
| Parent-child hierarchy | Derived from `OutlineLevel` when MPP reader doesn't set `ParentTask`. Parent links set via PssUpdate |
| Project start date | Set via PssUpdate on `msdyn_project.msdyn_scheduledstart` from workflow input |
| Operation set batching | 200-op batches, async execution, completion polling (status 192350001 = Completed) |
| Duplicate dependency detection | Queries existing `msdyn_projecttaskdependency` records and skips duplicates |
| Workflow activity shell | Thin wrapper with Case Template, Target Project, and optional Starts On inputs |
| Default bucket creation | Auto-creates project bucket if none exists |
| Assembly versioning | Incremented on each deploy to avoid CRM caching |

### ❌ DOES NOT WORK — PSS rejects these fields on PssCreate

| Field attempted | Error | Version tested |
|----------------|-------|----------------|
| `msdyn_displayorder` | "not a valid column in the table msdyn_projecttask" | 1.0.13 |
| `msdyn_scheduledstart` / `msdyn_scheduledend` | ScheduleAPI error — PSS calculates these from dependencies/duration | 1.0.18, 1.0.19 |
| `msdyn_msprojectclientid` | "Input to the column msdyn_msprojectclientid during create operation is not allowed" | 1.0.21 |

### ⚠️ CORE PROBLEM: Task GUID Mapping for Dependencies

**The fundamental issue:** After Phase 1 creates tasks via PssCreate, we need to reference their actual CRM GUIDs to create dependency records in Phase 2. PSS makes this difficult because:

1. **PSS reassigns GUIDs** — Even though we set `entity.Id` and `msdyn_projecttaskid` on PssCreate, PSS assigns its own GUIDs. Our pre-generated GUIDs do NOT match the actual CRM records. Confirmed by error: "Entity msdyn_projecttask with id ... does not exist" (v1.0.20).

2. **Cannot tag tasks with clientId** — `msdyn_msprojectclientid` is read-only on create (v1.0.21 error). We cannot write a custom identifier to match tasks back.

3. **Name matching works but is fragile** — Matching CRM tasks by `msdyn_subject` worked for DA0999 test file (v1.0.12–1.0.17). But fails when:
   - Duplicate task names exist (consumed in order, but risky)
   - Phase 2 polling doesn't find tasks in time (Accel1 file: 0 tasks found after 18s)

4. **New projects have longer propagation delay** — For brand-new projects, `WaitForProjectReady` never finds the root task even after 10–60s. Phase 2 polling found 0 tasks after 18s for the Accel1 file, even though the operation set reported "Completed". The DA0999 file worked because tasks appeared within 5s.

### Dependencies end-to-end status

| Test file | Tasks created | Dependencies created | Notes |
|-----------|--------------|---------------------|-------|
| DA0999 (23 tasks) | ✅ 23 | ✅ ~20 | Name matching worked, tasks found on first poll |
| Accel1 (28 tasks) | ✅ 28 | ❌ 0 | Phase 2 polling found 0 CRM tasks → 0 matches → 0 deps |

### What we need to solve

**Option A: Fix name matching + increase polling**
- Increase Phase 2 polling back to 12×5s = 60s (worked for DA0999)
- Risk: May still fail for new projects with slow propagation
- Risk: Duplicate task names cause wrong mappings
- Total time budget: ~10s init + ~5s batch + ~60s poll + ~5s deps = ~80s (under 120s limit)

**Option B: Include dependencies in same operation set as tasks (Phase 1)**
- PSS allows referencing tasks by the GUIDs set on PssCreate WITHIN the same operation set
- Parent links already work this way (same batch as creates)
- If dependencies are added to the SAME operation set as task creates, PSS should resolve the references internally
- This eliminates the need for Phase 2 entirely
- Risk: May exceed 200-op batch limit for large projects (tasks + parents + deps)

**Option C: Use operation set response to get actual GUIDs**
- The `ExecuteOperationSetV1` response contains a JSON payload with correlation IDs
- Could potentially extract actual GUIDs from the response
- Needs investigation of response format

### Key finding from MS docs (v1.0.23)

> "The ID property is optional. If you provide the ID property, the system tries to use it and throws an exception if it can't be used."

This means PSS DOES honour our pre-generated GUIDs. The v1.0.20 error was a **timing issue** (deps submitted before task records were committed), not a GUID mismatch.

### Implemented solution (v1.0.23.0)

Two-phase batching with pre-generated GUIDs — no polling, no name matching:

- **Phase 1**: Task creates (batched at 200). Each task gets a pre-generated GUID via `entity.Id` and `msdyn_projecttaskid`.
- **10s wait**: Ensures Phase 1 records are committed to DB.
- **Phase 2**: Parent links + dependencies (batched at 200). References the SAME pre-generated GUIDs from `taskIdMap`.

This scales to any project size. No CRM querying needed between phases.
