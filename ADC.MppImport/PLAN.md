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
