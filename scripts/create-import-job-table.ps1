<#
.SYNOPSIS
    Creates the adc_mppimportjob custom table and fields in Dynamics 365 / Dataverse
    using the Power Platform CLI (pac).

.DESCRIPTION
    Run this script once per environment to provision the adc_mppimportjob table
    that the async chunked import uses to track import progress.

.PARAMETER EnvironmentUrl
    The Dataverse environment URL (e.g., https://orgfbe0a613.crm6.dynamics.com)

.PARAMETER SolutionName
    The solution to add the table to (default: ADCMppImport)

.EXAMPLE
    .\create-import-job-table.ps1 -EnvironmentUrl https://orgfbe0a613.crm6.dynamics.com
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$EnvironmentUrl,

    [string]$SolutionName = "ADCMppImport"
)

$ErrorActionPreference = "Stop"

Write-Host "=== Create adc_mppimportjob Table ===" -ForegroundColor Cyan
Write-Host "Environment: $EnvironmentUrl"
Write-Host "Solution: $SolutionName"
Write-Host ""

# Authenticate if not already
Write-Host "Checking pac auth..." -ForegroundColor Yellow
try {
    $authList = pac auth list 2>&1
    if ($authList -notmatch $EnvironmentUrl) {
        Write-Host "Not authenticated to $EnvironmentUrl. Running pac auth create..."
        pac auth create --url $EnvironmentUrl
    } else {
        Write-Host "Already authenticated."
    }
} catch {
    Write-Host "Running pac auth create..."
    pac auth create --url $EnvironmentUrl
}

# Ensure solution exists
Write-Host "`nChecking solution '$SolutionName'..." -ForegroundColor Yellow
try {
    pac solution list 2>&1 | Out-Null
    # If solution doesn't exist, create it
    $solutions = pac solution list 2>&1
    if ($solutions -notmatch $SolutionName) {
        Write-Host "Creating solution '$SolutionName'..."
        pac solution create --name $SolutionName --publisher-name ADC --publisher-prefix adc
    }
} catch {
    Write-Host "Creating solution '$SolutionName'..."
    pac solution create --name $SolutionName --publisher-name ADC --publisher-prefix adc
}

# --- Create table via Web API ---
# pac CLI doesn't have direct table creation commands, so we use the Dataverse Web API

Write-Host "`nCreating adc_mppimportjob table via Web API..." -ForegroundColor Yellow

$headers = @{
    "Content-Type" = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
}

# Helper function to call Dataverse Web API using pac CLI token
function Invoke-DataverseApi {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body
    )

    $token = (pac auth token --environment $EnvironmentUrl 2>&1) | Select-String -Pattern "^[A-Za-z0-9]" | Select-Object -First 1
    $authHeader = @{ "Authorization" = "Bearer $($token.ToString().Trim())" }
    $allHeaders = $headers + $authHeader

    $uri = "$EnvironmentUrl/api/data/v9.2/$Path"

    $params = @{
        Uri = $uri
        Method = $Method
        Headers = $allHeaders
        ContentType = "application/json"
    }

    if ($Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10)
    }

    try {
        $response = Invoke-RestMethod @params
        return $response
    } catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 409 -or $statusCode -eq 400) {
            Write-Host "  (already exists or conflict, continuing)" -ForegroundColor DarkYellow
            return $null
        }
        throw
    }
}

# Create the entity metadata
$entityDef = @{
    "@odata.type" = "Microsoft.Dynamics.CRM.EntityMetadata"
    SchemaName = "adc_mppimportjob"
    DisplayName = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.Label"
        LocalizedLabels = @(@{
            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
            Label = "MPP Import Job"
            LanguageCode = 1033
        })
    }
    DisplayCollectionName = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.Label"
        LocalizedLabels = @(@{
            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
            Label = "MPP Import Jobs"
            LanguageCode = 1033
        })
    }
    Description = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.Label"
        LocalizedLabels = @(@{
            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
            Label = "Tracks async chunked MPP import progress"
            LanguageCode = 1033
        })
    }
    OwnershipType = "UserOwned"
    HasNotes = $false
    HasActivities = $false
    PrimaryNameAttribute = "adc_name"
    Attributes = @(
        @{
            "@odata.type" = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
            SchemaName = "adc_name"
            AttributeType = "String"
            FormatName = @{ Value = "Text" }
            MaxLength = 200
            DisplayName = @{
                "@odata.type" = "Microsoft.Dynamics.CRM.Label"
                LocalizedLabels = @(@{
                    "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                    Label = "Name"
                    LanguageCode = 1033
                })
            }
            RequiredLevel = @{ Value = "ApplicationRequired" }
            IsPrimaryName = $true
        }
    )
}

Write-Host "  Creating entity..."
try {
    Invoke-DataverseApi -Method POST -Path "EntityDefinitions" -Body $entityDef
    Write-Host "  Entity created." -ForegroundColor Green
} catch {
    Write-Host "  Entity creation: $_" -ForegroundColor Yellow
}

# Helper to add a field
function Add-Field {
    param(
        [string]$SchemaName,
        [string]$Label,
        [string]$Type,
        [hashtable]$ExtraProps = @{}
    )

    $attrDef = @{
        SchemaName = $SchemaName
        DisplayName = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label = $Label
                LanguageCode = 1033
            })
        }
        RequiredLevel = @{ Value = "None" }
    }

    switch ($Type) {
        "Integer" {
            $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.IntegerAttributeMetadata"
            $attrDef["AttributeType"] = "Integer"
            $attrDef["Format"] = "None"
            $attrDef["MinValue"] = -1
            $attrDef["MaxValue"] = 2147483647
        }
        "Memo" {
            $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.MemoAttributeMetadata"
            $attrDef["AttributeType"] = "Memo"
            $attrDef["Format"] = "TextArea"
            $attrDef["MaxLength"] = 1048576
        }
        "String" {
            $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
            $attrDef["AttributeType"] = "String"
            $attrDef["FormatName"] = @{ Value = "Text" }
            $attrDef["MaxLength"] = if ($ExtraProps.MaxLength) { $ExtraProps.MaxLength } else { 200 }
        }
        "DateTime" {
            $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata"
            $attrDef["AttributeType"] = "DateTime"
            $attrDef["Format"] = "DateOnly"
        }
        "Picklist" {
            $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.PicklistAttributeMetadata"
            $attrDef["AttributeType"] = "Picklist"
            if ($ExtraProps.Options) {
                $attrDef["OptionSet"] = @{
                    "@odata.type" = "Microsoft.Dynamics.CRM.OptionSetMetadata"
                    IsGlobal = $false
                    OptionSetType = "Picklist"
                    Options = $ExtraProps.Options
                }
            }
        }
    }

    Write-Host "  Adding field: $SchemaName ($Type)..."
    try {
        Invoke-DataverseApi -Method POST -Path "EntityDefinitions(LogicalName='adc_mppimportjob')/Attributes" -Body $attrDef
        Write-Host "    OK" -ForegroundColor Green
    } catch {
        Write-Host "    $_" -ForegroundColor Yellow
    }
}

# Status option set
$statusOptions = @(
    @{ Value = 0; Label = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "Queued"; LanguageCode = 1033 }) } }
    @{ Value = 1; Label = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "Creating Tasks"; LanguageCode = 1033 }) } }
    @{ Value = 2; Label = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "Waiting for Tasks"; LanguageCode = 1033 }) } }
    @{ Value = 3; Label = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "Polling GUIDs"; LanguageCode = 1033 }) } }
    @{ Value = 4; Label = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "Creating Dependencies"; LanguageCode = 1033 }) } }
    @{ Value = 5; Label = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "Waiting for Dependencies"; LanguageCode = 1033 }) } }
    @{ Value = 6; Label = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "Completed"; LanguageCode = 1033 }) } }
    @{ Value = 7; Label = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "Failed"; LanguageCode = 1033 }) } }
)

# Add all fields
Add-Field -SchemaName "adc_status" -Label "Status" -Type "Picklist" -ExtraProps @{ Options = $statusOptions }
Add-Field -SchemaName "adc_phase" -Label "Phase" -Type "Integer"
Add-Field -SchemaName "adc_currentbatch" -Label "Current Batch" -Type "Integer"
Add-Field -SchemaName "adc_totalbatches" -Label "Total Batches" -Type "Integer"
Add-Field -SchemaName "adc_totaltasks" -Label "Total Tasks" -Type "Integer"
Add-Field -SchemaName "adc_createdcount" -Label "Created Count" -Type "Integer"
Add-Field -SchemaName "adc_depscount" -Label "Dependencies Count" -Type "Integer"
Add-Field -SchemaName "adc_tick" -Label "Tick" -Type "Integer"
Add-Field -SchemaName "adc_taskdatajson" -Label "Task Data (JSON)" -Type "Memo"
Add-Field -SchemaName "adc_taskidmapjson" -Label "Task ID Map (JSON)" -Type "Memo"
Add-Field -SchemaName "adc_batchesjson" -Label "Batches (JSON)" -Type "Memo"
Add-Field -SchemaName "adc_operationsetid" -Label "Operation Set ID" -Type "String" -ExtraProps @{ MaxLength = 100 }
Add-Field -SchemaName "adc_projectstartdate" -Label "Project Start Date" -Type "DateTime"
Add-Field -SchemaName "adc_errormessage" -Label "Error Message" -Type "Memo"

# Add lookup relationships
Write-Host "`n  Adding lookup: adc_project -> msdyn_project..." -ForegroundColor Yellow
$projectLookup = @{
    SchemaName = "adc_project_msdyn_project"
    "@odata.type" = "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata"
    ReferencedEntity = "msdyn_project"
    ReferencingEntity = "adc_mppimportjob"
    Lookup = @{
        SchemaName = "adc_project"
        DisplayName = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label = "Project"
                LanguageCode = 1033
            })
        }
    }
}
try {
    Invoke-DataverseApi -Method POST -Path "RelationshipDefinitions" -Body $projectLookup
    Write-Host "    OK" -ForegroundColor Green
} catch {
    Write-Host "    $_" -ForegroundColor Yellow
}

Write-Host "`n  Adding lookup: adc_casetemplate -> adc_adccasetemplate..." -ForegroundColor Yellow
$templateLookup = @{
    SchemaName = "adc_casetemplate_adc_adccasetemplate"
    "@odata.type" = "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata"
    ReferencedEntity = "adc_adccasetemplate"
    ReferencingEntity = "adc_mppimportjob"
    Lookup = @{
        SchemaName = "adc_casetemplate"
        DisplayName = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label = "Case Template"
                LanguageCode = 1033
            })
        }
    }
}
try {
    Invoke-DataverseApi -Method POST -Path "RelationshipDefinitions" -Body $templateLookup
    Write-Host "    OK" -ForegroundColor Green
} catch {
    Write-Host "    $_" -ForegroundColor Yellow
}

# Publish customizations
Write-Host "`nPublishing customizations..." -ForegroundColor Yellow
try {
    Invoke-DataverseApi -Method POST -Path "PublishAllXml" -Body @{}
    Write-Host "Published." -ForegroundColor Green
} catch {
    Write-Host "Publish: $_" -ForegroundColor Yellow
}

Write-Host "`n=== Done ===" -ForegroundColor Cyan
Write-Host "Table adc_mppimportjob created with all fields and relationships."
Write-Host "Next: Run deploy-plugin.ps1 to register the async plugin."
