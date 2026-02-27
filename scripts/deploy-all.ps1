<#
.SYNOPSIS
    All-in-one: Creates the adc_mppimportjob table, deploys the plugin assembly,
    and registers plugin steps using ClientSecret (S2S) authentication.

.DESCRIPTION
    Uses the Dataverse Web API with OAuth2 client_credentials flow.
    No pac CLI required.

.PARAMETER CrmUrl
    Dataverse environment URL

.PARAMETER ClientId
    App registration client ID

.PARAMETER ClientSecret
    App registration client secret

.PARAMETER SkipTableCreation
    Skip table/field creation (if already done)

.PARAMETER SkipPluginDeploy
    Skip assembly upload (if just registering steps)

.EXAMPLE
    .\deploy-all.ps1 -CrmUrl https://orgfbe0a613.crm6.dynamics.com -ClientId "xxx" -ClientSecret "xxx"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CrmUrl,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$ClientSecret,

    [switch]$SkipTableCreation,
    [switch]$SkipPluginDeploy
)

$ErrorActionPreference = "Stop"

# Trim trailing slash
$CrmUrl = $CrmUrl.TrimEnd('/')

# Derive tenant authority from CRM URL
# For Dynamics 365, the authority endpoint can be discovered or we use the common endpoint
$resource = $CrmUrl
$authorityUrl = "https://login.microsoftonline.com"

Write-Host "=== ADC MppImport - Full Deployment ===" -ForegroundColor Cyan
Write-Host "CRM:      $CrmUrl"
Write-Host "ClientId: $ClientId"
Write-Host ""

# ============================================================
# AUTH: Get OAuth2 token via client_credentials
# ============================================================

function Get-AccessToken {
    # Discover the tenant authority from the CRM URL's WWW-Authenticate header
    $tenantAuthority = $null
    Write-Host "Discovering tenant from $CrmUrl..."
    try {
        # Make an unauthenticated request - the 401 response contains the authority URL
        Invoke-WebRequest -Uri "$CrmUrl/api/data/v9.2/" -Method GET -UseBasicParsing -ErrorAction Stop | Out-Null
    } catch {
        if ($_.Exception.Response) {
            # PowerShell 5.x: headers via .Headers property
            $wwwAuth = $_.Exception.Response.Headers["WWW-Authenticate"]
            if (-not $wwwAuth) {
                # Try as string
                $responseHeaders = $_.Exception.Response.Headers
                if ($responseHeaders) {
                    $wwwAuth = $responseHeaders.ToString()
                }
            }
            if ($wwwAuth) {
                $match = [regex]::Match($wwwAuth, 'authorization_uri="?([^"\s,]+)"?')
                if ($match.Success) {
                    $tenantAuthority = $match.Groups[1].Value
                }
            }
        }
    }

    if (-not $tenantAuthority) {
        # Fallback: try to extract tenant from CRM URL or use common
        Write-Host "  Could not discover tenant authority, trying alternate discovery..." -ForegroundColor Yellow
        # Try the OpenID configuration endpoint
        try {
            $openIdUrl = "$CrmUrl/.well-known/openid-configuration"
            $openIdConfig = Invoke-RestMethod -Uri $openIdUrl -Method GET -ErrorAction Stop
            $tenantAuthority = $openIdConfig.authorization_endpoint -replace '/oauth2/authorize$', ''
        } catch {
            Write-Host "  OpenID discovery also failed." -ForegroundColor Yellow
        }
    }

    if (-not $tenantAuthority) {
        Write-Host "  ERROR: Could not discover tenant. Please provide TenantId parameter." -ForegroundColor Red
        throw "Tenant discovery failed"
    }

    # Strip /oauth2/authorize if present (discovery returns the authorize endpoint, not the base)
    $tenantAuthority = $tenantAuthority -replace '/oauth2/authorize$', ''
    Write-Host "  Authority: $tenantAuthority"
    $tokenUrl = "$tenantAuthority/oauth2/token"

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        resource      = $CrmUrl
    }

    Write-Host "  Requesting token from $tokenUrl..."
    $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -ContentType "application/x-www-form-urlencoded"
    Write-Host "  Token acquired." -ForegroundColor Green
    return $tokenResponse.access_token
}

$global:accessToken = Get-AccessToken

function Invoke-CrmApi {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body,
        [switch]$IgnoreErrors
    )

    $headers = @{
        "Authorization"  = "Bearer $global:accessToken"
        "Content-Type"   = "application/json"
        "OData-MaxVersion" = "4.0"
        "OData-Version"  = "4.0"
        "Accept"         = "application/json"
    }

    $uri = "$CrmUrl/api/data/v9.2/$Path"
    $params = @{
        Uri         = $uri
        Method      = $Method
        Headers     = $headers
        ContentType = "application/json"
    }

    if ($Body) {
        $jsonBody = $Body | ConvertTo-Json -Depth 20 -Compress
        $params.Body = $jsonBody
    }

    try {
        $response = Invoke-RestMethod @params
        return $response
    } catch {
        $statusCode = 0
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }
        $errorBody = ""
        try {
            $reader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
            $errorBody = $reader.ReadToEnd()
        } catch { }

        if ($IgnoreErrors -or $statusCode -eq 409 -or ($statusCode -eq 400 -and $errorBody -match "already exists")) {
            Write-Host "    (already exists, continuing)" -ForegroundColor DarkYellow
            return $null
        }

        Write-Host "  API Error ($statusCode): $($_.Exception.Message)" -ForegroundColor Red
        if ($errorBody) { Write-Host "  Body: $errorBody" -ForegroundColor Red }
        throw
    }
}

# ============================================================
# STEP 1: Create adc_mppimportjob table
# ============================================================

if (-not $SkipTableCreation) {
    Write-Host "`n=== Step 1: Create adc_mppimportjob Table ===" -ForegroundColor Yellow

    # Check if entity already exists
    $existingEntity = $null
    try {
        $existingEntity = Invoke-CrmApi -Method GET -Path "EntityDefinitions(LogicalName='adc_mppimportjob')?`$select=LogicalName" -IgnoreErrors
    } catch { }

    if ($existingEntity) {
        Write-Host "  Table adc_mppimportjob already exists, skipping creation." -ForegroundColor DarkYellow
    } else {
        Write-Host "  Creating entity adc_mppimportjob..."

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

        Invoke-CrmApi -Method POST -Path "EntityDefinitions" -Body $entityDef
        Write-Host "  Entity created." -ForegroundColor Green

        # Wait a moment for entity to be available
        Start-Sleep -Seconds 3
    }

    # --- Add fields ---
    function Add-CrmField {
        param(
            [string]$SchemaName,
            [string]$Label,
            [string]$Type,
            [hashtable]$Extra = @{}
        )

        # Check if field exists
        try {
            $existing = Invoke-CrmApi -Method GET -Path "EntityDefinitions(LogicalName='adc_mppimportjob')/Attributes(LogicalName='$($SchemaName.ToLower())')?`$select=LogicalName" -IgnoreErrors
            if ($existing) {
                Write-Host "  Field $SchemaName already exists." -ForegroundColor DarkGray
                return
            }
        } catch { }

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
                $attrDef["Format"] = "None"
                $attrDef["MinValue"] = -1
                $attrDef["MaxValue"] = 2147483647
            }
            "Memo" {
                $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.MemoAttributeMetadata"
                $attrDef["Format"] = "TextArea"
                $attrDef["MaxLength"] = 1048576
            }
            "String100" {
                $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
                $attrDef["FormatName"] = @{ Value = "Text" }
                $attrDef["MaxLength"] = 100
            }
            "DateTime" {
                $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata"
                $attrDef["Format"] = "DateOnly"
            }
            "Picklist" {
                $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.PicklistAttributeMetadata"
                if ($Extra.Options) {
                    $attrDef["OptionSet"] = @{
                        "@odata.type" = "Microsoft.Dynamics.CRM.OptionSetMetadata"
                        IsGlobal = $false
                        OptionSetType = "Picklist"
                        Options = $Extra.Options
                    }
                }
            }
        }

        Write-Host "  Adding field: $SchemaName ($Type)..."
        try {
            Invoke-CrmApi -Method POST -Path "EntityDefinitions(LogicalName='adc_mppimportjob')/Attributes" -Body $attrDef
            Write-Host "    OK" -ForegroundColor Green
        } catch {
            Write-Host "    Failed (may already exist)" -ForegroundColor Yellow
        }
    }

    # Status picklist
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

    Add-CrmField -SchemaName "adc_status" -Label "Status" -Type "Picklist" -Extra @{ Options = $statusOptions }
    Add-CrmField -SchemaName "adc_phase" -Label "Phase" -Type "Integer"
    Add-CrmField -SchemaName "adc_currentbatch" -Label "Current Batch" -Type "Integer"
    Add-CrmField -SchemaName "adc_totalbatches" -Label "Total Batches" -Type "Integer"
    Add-CrmField -SchemaName "adc_totaltasks" -Label "Total Tasks" -Type "Integer"
    Add-CrmField -SchemaName "adc_createdcount" -Label "Created Count" -Type "Integer"
    Add-CrmField -SchemaName "adc_depscount" -Label "Dependencies Count" -Type "Integer"
    Add-CrmField -SchemaName "adc_tick" -Label "Tick" -Type "Integer"
    Add-CrmField -SchemaName "adc_taskdatajson" -Label "Task Data (JSON)" -Type "Memo"
    Add-CrmField -SchemaName "adc_taskidmapjson" -Label "Task ID Map (JSON)" -Type "Memo"
    Add-CrmField -SchemaName "adc_batchesjson" -Label "Batches (JSON)" -Type "Memo"
    Add-CrmField -SchemaName "adc_operationsetid" -Label "Operation Set ID" -Type "String100"
    Add-CrmField -SchemaName "adc_projectstartdate" -Label "Project Start Date" -Type "DateTime"
    Add-CrmField -SchemaName "adc_errormessage" -Label "Error Message" -Type "Memo"

    # --- Add lookup fields via relationships ---
    Write-Host "`n  Adding lookup: adc_project -> msdyn_project..."
    $projectRel = @{
        SchemaName = "adc_mppimportjob_project"
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
        Invoke-CrmApi -Method POST -Path "RelationshipDefinitions" -Body $projectRel
        Write-Host "    OK" -ForegroundColor Green
    } catch {
        Write-Host "    (may already exist)" -ForegroundColor Yellow
    }

    Write-Host "  Adding lookup: adc_casetemplate -> adc_adccasetemplate..."
    $templateRel = @{
        SchemaName = "adc_mppimportjob_casetemplate"
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
        Invoke-CrmApi -Method POST -Path "RelationshipDefinitions" -Body $templateRel
        Write-Host "    OK" -ForegroundColor Green
    } catch {
        Write-Host "    (may already exist)" -ForegroundColor Yellow
    }

    # Publish
    Write-Host "`n  Publishing customizations..."
    try {
        Invoke-CrmApi -Method POST -Path "PublishAllXml" -Body @{}
        Write-Host "  Published." -ForegroundColor Green
    } catch {
        Write-Host "  Publish warning: $_" -ForegroundColor Yellow
    }

    Write-Host "`n  Table creation complete." -ForegroundColor Green
} else {
    Write-Host "`nStep 1: SKIPPED (table creation)" -ForegroundColor DarkGray
}

# ============================================================
# STEP 2: Build and deploy plugin assembly
# ============================================================

if (-not $SkipPluginDeploy) {
    Write-Host "`n=== Step 2: Build & Deploy Plugin Assembly ===" -ForegroundColor Yellow

    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $repoRoot = Split-Path -Parent $scriptDir
    $projectFile = Join-Path $repoRoot "ADC.MppImport\ADC.MppImport.csproj"
    $outputDir = Join-Path $repoRoot "ADC.MppImport\bin\Debug"
    $assemblyPath = Join-Path $outputDir "ADC.MppImport.dll"

    # Build
    Write-Host "  Building Debug..."
    & msbuild $projectFile /p:Configuration=Debug /v:minimal /t:Build
    if ($LASTEXITCODE -ne 0) {
        Write-Host "BUILD FAILED" -ForegroundColor Red
        exit 1
    }

    if (-not (Test-Path $assemblyPath)) {
        Write-Host "ERROR: Assembly not found at $assemblyPath" -ForegroundColor Red
        exit 1
    }

    # Read assembly bytes and encode as base64
    Write-Host "  Reading assembly: $assemblyPath"
    $assemblyBytes = [System.IO.File]::ReadAllBytes($assemblyPath)
    $assemblyBase64 = [Convert]::ToBase64String($assemblyBytes)
    Write-Host "  Assembly size: $($assemblyBytes.Length) bytes"

    # Check if assembly already registered
    Write-Host "  Checking for existing plugin assembly..."
    $existingAssemblies = Invoke-CrmApi -Method GET -Path "pluginassemblies?`$filter=name eq 'ADC.MppImport'&`$select=pluginassemblyid,name,version"

    if ($existingAssemblies.value.Count -gt 0) {
        $assemblyId = $existingAssemblies.value[0].pluginassemblyid
        Write-Host "  Found existing assembly: $assemblyId - updating..." -ForegroundColor DarkYellow

        $updateBody = @{
            content = $assemblyBase64
        }
        Invoke-CrmApi -Method PATCH -Path "pluginassemblies($assemblyId)" -Body $updateBody
        Write-Host "  Assembly updated." -ForegroundColor Green
    } else {
        Write-Host "  Registering new assembly..."

        # Get assembly version info
        $assemblyInfo = [System.Reflection.AssemblyName]::GetAssemblyName($assemblyPath)

        $registerBody = @{
            name = "ADC.MppImport"
            content = $assemblyBase64
            isolationmode = 2  # Sandbox
            sourcetype = 0     # Database
            version = $assemblyInfo.Version.ToString()
            culture = "neutral"
            publickeytoken = ""
        }

        try {
            $result = Invoke-CrmApi -Method POST -Path "pluginassemblies" -Body $registerBody
            Write-Host "  Assembly registered." -ForegroundColor Green
        } catch {
            Write-Host "  Assembly registration failed: $_" -ForegroundColor Red
            Write-Host "  You may need to register it manually via Plugin Registration Tool." -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "`nStep 2: SKIPPED (plugin deploy)" -ForegroundColor DarkGray
}

# ============================================================
# STEP 3: Register plugin steps
# ============================================================

Write-Host "`n=== Step 3: Register Plugin Steps ===" -ForegroundColor Yellow

# Refresh token in case it expired
$global:accessToken = Get-AccessToken

# Find plugin type
Write-Host "  Looking up plugin type: ADC.MppImport.Plugins.MppImportJobPlugin..."
$pluginTypes = Invoke-CrmApi -Method GET -Path "plugintypes?`$filter=typename eq 'ADC.MppImport.Plugins.MppImportJobPlugin'&`$select=plugintypeid,typename"

if ($pluginTypes.value.Count -eq 0) {
    Write-Host "  Plugin type not found. Assembly may not be registered yet." -ForegroundColor Red
    Write-Host "  Register the assembly first, then re-run with -SkipTableCreation -SkipPluginDeploy" -ForegroundColor Yellow
    exit 1
}

$pluginTypeId = $pluginTypes.value[0].plugintypeid
Write-Host "  Found plugin type: $pluginTypeId" -ForegroundColor Green

# Find message IDs
$createMsg = Invoke-CrmApi -Method GET -Path "sdkmessages?`$filter=name eq 'Create'&`$select=sdkmessageid"
$updateMsg = Invoke-CrmApi -Method GET -Path "sdkmessages?`$filter=name eq 'Update'&`$select=sdkmessageid"

$createMsgId = $createMsg.value[0].sdkmessageid
$updateMsgId = $updateMsg.value[0].sdkmessageid

Write-Host "  Create message: $createMsgId"
Write-Host "  Update message: $updateMsgId"

# Find message filters for adc_mppimportjob
$createFilter = Invoke-CrmApi -Method GET -Path "sdkmessagefilters?`$filter=primaryobjecttypecode eq 'adc_mppimportjob' and _sdkmessageid_value eq $createMsgId&`$select=sdkmessagefilterid"
$updateFilter = Invoke-CrmApi -Method GET -Path "sdkmessagefilters?`$filter=primaryobjecttypecode eq 'adc_mppimportjob' and _sdkmessageid_value eq $updateMsgId&`$select=sdkmessagefilterid"

if ($createFilter.value.Count -eq 0 -or $updateFilter.value.Count -eq 0) {
    Write-Host "  Message filters not found for adc_mppimportjob." -ForegroundColor Red
    Write-Host "  The table may not be published yet. Try running PublishAllXml and re-running." -ForegroundColor Yellow

    # Try to publish and retry
    Write-Host "  Publishing and retrying..."
    Invoke-CrmApi -Method POST -Path "PublishAllXml" -Body @{}
    Start-Sleep -Seconds 5

    $createFilter = Invoke-CrmApi -Method GET -Path "sdkmessagefilters?`$filter=primaryobjecttypecode eq 'adc_mppimportjob' and _sdkmessageid_value eq $createMsgId&`$select=sdkmessagefilterid"
    $updateFilter = Invoke-CrmApi -Method GET -Path "sdkmessagefilters?`$filter=primaryobjecttypecode eq 'adc_mppimportjob' and _sdkmessageid_value eq $updateMsgId&`$select=sdkmessagefilterid"

    if ($createFilter.value.Count -eq 0 -or $updateFilter.value.Count -eq 0) {
        Write-Host "  Still not found. Register steps manually." -ForegroundColor Red
        exit 1
    }
}

$createFilterId = $createFilter.value[0].sdkmessagefilterid
$updateFilterId = $updateFilter.value[0].sdkmessagefilterid

Write-Host "  Create filter: $createFilterId"
Write-Host "  Update filter: $updateFilterId"

# Check for existing steps
$existingSteps = Invoke-CrmApi -Method GET -Path "sdkmessageprocessingsteps?`$filter=_plugintypeid_value eq $pluginTypeId&`$select=sdkmessageprocessingstepid,name"

if ($existingSteps.value.Count -gt 0) {
    Write-Host "  Found $($existingSteps.value.Count) existing step(s):" -ForegroundColor DarkYellow
    foreach ($step in $existingSteps.value) {
        Write-Host "    - $($step.name) ($($step.sdkmessageprocessingstepid))" -ForegroundColor DarkYellow
    }
    Write-Host "  Skipping step registration (already registered)." -ForegroundColor DarkYellow
} else {
    # Register Create step
    Write-Host "  Registering Create step (async, post-op)..."
    $createStep = @{
        name = "MppImportJobPlugin: Create of adc_mppimportjob"
        mode = 1  # Async
        rank = 1
        stage = 40  # PostOperation
        supporteddeployment = 0  # Server only
        asyncautodelete = $true
        "plugintypeid@odata.bind" = "/plugintypes($pluginTypeId)"
        "sdkmessageid@odata.bind" = "/sdkmessages($createMsgId)"
        "sdkmessagefilterid@odata.bind" = "/sdkmessagefilters($createFilterId)"
    }

    try {
        Invoke-CrmApi -Method POST -Path "sdkmessageprocessingsteps" -Body $createStep
        Write-Host "    Create step registered." -ForegroundColor Green
    } catch {
        Write-Host "    Create step failed: $_" -ForegroundColor Red
    }

    # Register Update step with filtering attributes
    Write-Host "  Registering Update step (async, post-op, filter: adc_status,adc_tick)..."
    $updateStep = @{
        name = "MppImportJobPlugin: Update of adc_mppimportjob"
        mode = 1  # Async
        rank = 1
        stage = 40  # PostOperation
        supporteddeployment = 0  # Server only
        filteringattributes = "adc_status,adc_tick"
        asyncautodelete = $true
        "plugintypeid@odata.bind" = "/plugintypes($pluginTypeId)"
        "sdkmessageid@odata.bind" = "/sdkmessages($updateMsgId)"
        "sdkmessagefilterid@odata.bind" = "/sdkmessagefilters($updateFilterId)"
    }

    try {
        Invoke-CrmApi -Method POST -Path "sdkmessageprocessingsteps" -Body $updateStep
        Write-Host "    Update step registered." -ForegroundColor Green
    } catch {
        Write-Host "    Update step failed: $_" -ForegroundColor Red
    }
}

# ============================================================
# DONE
# ============================================================

Write-Host "`n=== Deployment Complete ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Summary:" -ForegroundColor White
Write-Host "  Table:    adc_mppimportjob (with all fields and lookups)"
Write-Host "  Assembly: ADC.MppImport (sandbox, database)"
Write-Host "  Steps:    Create + Update of adc_mppimportjob (async, post-op)"
Write-Host ""
Write-Host "To test, trigger the StartMppImportActivity workflow or create" -ForegroundColor Gray
Write-Host "an adc_mppimportjob record manually and monitor its adc_status field." -ForegroundColor Gray
