<#
.SYNOPSIS
    Creates the adc_mppimportjob table with all fields and lookups.

.DESCRIPTION
    Uses Dataverse Web API with OAuth2 client_credentials (S2S).
    Safe to re-run — skips components that already exist.

.PARAMETER CrmUrl
    Dataverse environment URL (e.g. https://org123.crm6.dynamics.com)

.PARAMETER ClientId
    Azure AD app registration client ID

.PARAMETER ClientSecret
    Azure AD app registration client secret

.EXAMPLE
    .\create-importjob-table.ps1 -CrmUrl https://org123.crm6.dynamics.com -ClientId "xxx" -ClientSecret "yyy"
#>

param(
    [Parameter(Mandatory = $true)][string]$CrmUrl,
    [Parameter(Mandatory = $true)][string]$ClientId,
    [Parameter(Mandatory = $true)][string]$ClientSecret
)

$ErrorActionPreference = "Stop"
$CrmUrl = $CrmUrl.TrimEnd('/')

Write-Host "=== Create adc_mppimportjob Table ===" -ForegroundColor Cyan
Write-Host "CRM: $CrmUrl"
Write-Host ""

# --- Auth ---
function Get-AccessToken {
    try {
        Invoke-WebRequest -Uri "$CrmUrl/api/data/v9.2/" -Method GET -UseBasicParsing -ErrorAction Stop | Out-Null
    } catch {
        $wwwAuth = $null
        if ($_.Exception.Response) {
            $wwwAuth = $_.Exception.Response.Headers["WWW-Authenticate"]
            if (-not $wwwAuth) {
                $responseHeaders = $_.Exception.Response.Headers
                if ($responseHeaders) { $wwwAuth = $responseHeaders.ToString() }
            }
        }
        $tenantAuthority = $null
        if ($wwwAuth) {
            $match = [regex]::Match($wwwAuth, 'authorization_uri="?([^"\s,]+)"?')
            if ($match.Success) { $tenantAuthority = $match.Groups[1].Value }
        }
        if (-not $tenantAuthority) {
            try {
                $openIdConfig = Invoke-RestMethod -Uri "$CrmUrl/.well-known/openid-configuration" -Method GET -ErrorAction Stop
                $tenantAuthority = $openIdConfig.authorization_endpoint -replace '/oauth2/authorize$', ''
            } catch { }
        }
        if (-not $tenantAuthority) { throw "Could not discover tenant authority from $CrmUrl" }
        $tenantAuthority = $tenantAuthority -replace '/oauth2/authorize$', ''
        Write-Host "  Authority: $tenantAuthority"

        $body = @{
            grant_type    = "client_credentials"
            client_id     = $ClientId
            client_secret = $ClientSecret
            resource      = $CrmUrl
        }
        $tokenUrl = "$tenantAuthority/oauth2/token"
        $tokenResponse = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -ContentType "application/x-www-form-urlencoded"
        return $tokenResponse.access_token
    }
}

Write-Host "Authenticating..." -ForegroundColor Yellow
$global:token = Get-AccessToken
Write-Host "  Token acquired." -ForegroundColor Green

function Invoke-Api {
    param([string]$Method, [string]$Path, [object]$Body, [switch]$IgnoreErrors)

    $headers = @{
        Authorization      = "Bearer $global:token"
        "Content-Type"     = "application/json"
        "OData-MaxVersion" = "4.0"
        "OData-Version"    = "4.0"
        Accept             = "application/json"
    }
    $uri = "$CrmUrl/api/data/v9.2/$Path"
    $params = @{ Uri = $uri; Method = $Method; Headers = $headers; ContentType = "application/json" }
    if ($Body) { $params.Body = ($Body | ConvertTo-Json -Depth 20 -Compress) }

    try {
        return Invoke-RestMethod @params
    } catch {
        $statusCode = 0
        if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }
        if ($IgnoreErrors -or $statusCode -eq 409 -or $statusCode -eq 400) {
            Write-Host "    (already exists, skipping)" -ForegroundColor DarkYellow
            return $null
        }
        throw
    }
}

# --- 1. Create entity ---
Write-Host "`nCreating entity adc_mppimportjob..." -ForegroundColor Yellow

$existing = $null
try { $existing = Invoke-Api -Method GET -Path "EntityDefinitions(LogicalName='adc_mppimportjob')?`$select=LogicalName" -IgnoreErrors } catch { }

if ($existing) {
    Write-Host "  Table already exists, skipping entity creation." -ForegroundColor DarkYellow
} else {
    $entityDef = @{
        "@odata.type"        = "Microsoft.Dynamics.CRM.EntityMetadata"
        SchemaName           = "adc_mppimportjob"
        DisplayName          = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "MPP Import Job"; LanguageCode = 1033 }) }
        DisplayCollectionName = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "MPP Import Jobs"; LanguageCode = 1033 }) }
        Description          = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "Tracks async chunked MPP import progress"; LanguageCode = 1033 }) }
        OwnershipType        = "UserOwned"
        HasNotes             = $false
        HasActivities        = $false
        PrimaryNameAttribute = "adc_name"
        Attributes           = @(@{
            "@odata.type" = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
            SchemaName    = "adc_name"
            AttributeType = "String"
            FormatName    = @{ Value = "Text" }
            MaxLength     = 200
            DisplayName   = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = "Name"; LanguageCode = 1033 }) }
            RequiredLevel = @{ Value = "ApplicationRequired" }
            IsPrimaryName = $true
        })
    }
    Invoke-Api -Method POST -Path "EntityDefinitions" -Body $entityDef
    Write-Host "  Created." -ForegroundColor Green
    Start-Sleep -Seconds 3
}

# --- 2. Add fields ---
function Add-Field {
    param([string]$SchemaName, [string]$Label, [string]$Type, [hashtable]$Extra = @{})

    try {
        $ex = Invoke-Api -Method GET -Path "EntityDefinitions(LogicalName='adc_mppimportjob')/Attributes(LogicalName='$($SchemaName.ToLower())')?`$select=LogicalName" -IgnoreErrors
        if ($ex) { Write-Host "  $SchemaName — exists" -ForegroundColor DarkGray; return }
    } catch { }

    $attr = @{
        SchemaName    = $SchemaName
        DisplayName   = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = $Label; LanguageCode = 1033 }) }
        RequiredLevel = @{ Value = "None" }
    }
    switch ($Type) {
        "Integer"  { $attr["@odata.type"] = "Microsoft.Dynamics.CRM.IntegerAttributeMetadata"; $attr["Format"] = "None"; $attr["MinValue"] = -1; $attr["MaxValue"] = 2147483647 }
        "Memo"     { $attr["@odata.type"] = "Microsoft.Dynamics.CRM.MemoAttributeMetadata"; $attr["Format"] = "TextArea"; $attr["MaxLength"] = 1048576 }
        "String"   { $attr["@odata.type"] = "Microsoft.Dynamics.CRM.StringAttributeMetadata"; $attr["FormatName"] = @{ Value = "Text" }; $attr["MaxLength"] = if ($Extra.MaxLength) { $Extra.MaxLength } else { 200 } }
        "DateTime" { $attr["@odata.type"] = "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata"; $attr["Format"] = "DateOnly" }
        "Picklist" { $attr["@odata.type"] = "Microsoft.Dynamics.CRM.PicklistAttributeMetadata"; if ($Extra.Options) { $attr["OptionSet"] = @{ "@odata.type" = "Microsoft.Dynamics.CRM.OptionSetMetadata"; IsGlobal = $false; OptionSetType = "Picklist"; Options = $Extra.Options } } }
    }
    Write-Host "  $SchemaName ($Type)..."
    try { Invoke-Api -Method POST -Path "EntityDefinitions(LogicalName='adc_mppimportjob')/Attributes" -Body $attr; Write-Host "    OK" -ForegroundColor Green }
    catch { Write-Host "    Failed (may already exist)" -ForegroundColor Yellow }
}

Write-Host "`nAdding fields..." -ForegroundColor Yellow

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

Add-Field -SchemaName "adc_status"          -Label "Status"            -Type "Picklist" -Extra @{ Options = $statusOptions }
Add-Field -SchemaName "adc_phase"           -Label "Phase"             -Type "Integer"
Add-Field -SchemaName "adc_currentbatch"    -Label "Current Batch"     -Type "Integer"
Add-Field -SchemaName "adc_totalbatches"    -Label "Total Batches"     -Type "Integer"
Add-Field -SchemaName "adc_totaltasks"      -Label "Total Tasks"       -Type "Integer"
Add-Field -SchemaName "adc_createdcount"    -Label "Created Count"     -Type "Integer"
Add-Field -SchemaName "adc_depscount"       -Label "Dependencies Count" -Type "Integer"
Add-Field -SchemaName "adc_tick"            -Label "Tick"              -Type "Integer"
Add-Field -SchemaName "adc_taskdatajson"    -Label "Task Data (JSON)"  -Type "Memo"
Add-Field -SchemaName "adc_taskidmapjson"   -Label "Task ID Map (JSON)" -Type "Memo"
Add-Field -SchemaName "adc_batchesjson"     -Label "Batches (JSON)"    -Type "Memo"
Add-Field -SchemaName "adc_operationsetid"  -Label "Operation Set ID"  -Type "String" -Extra @{ MaxLength = 100 }
Add-Field -SchemaName "adc_projectstartdate" -Label "Project Start Date" -Type "DateTime"
Add-Field -SchemaName "adc_errormessage"    -Label "Error Message"     -Type "Memo"

# --- 3. Add lookup relationships ---
function Add-Lookup {
    param([string]$SchemaName, [string]$RelName, [string]$Label, [string]$Target)

    Write-Host "  Lookup: $SchemaName -> $Target..."
    $rel = @{
        SchemaName        = $RelName
        "@odata.type"     = "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata"
        ReferencedEntity  = $Target
        ReferencingEntity = "adc_mppimportjob"
        Lookup            = @{
            SchemaName  = $SchemaName
            DisplayName = @{ "@odata.type" = "Microsoft.Dynamics.CRM.Label"; LocalizedLabels = @(@{ "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"; Label = $Label; LanguageCode = 1033 }) }
        }
    }
    try { Invoke-Api -Method POST -Path "RelationshipDefinitions" -Body $rel; Write-Host "    OK" -ForegroundColor Green }
    catch { Write-Host "    (may already exist)" -ForegroundColor Yellow }
}

Write-Host "`nAdding lookups..." -ForegroundColor Yellow
Add-Lookup -SchemaName "adc_project"        -RelName "adc_mppimportjob_project"        -Label "Project"         -Target "msdyn_project"
Add-Lookup -SchemaName "adc_casetemplate"   -RelName "adc_mppimportjob_casetemplate"   -Label "Case Template"   -Target "adc_adccasetemplate"
Add-Lookup -SchemaName "adc_case"           -RelName "adc_mppimportjob_case"           -Label "Case"            -Target "adc_case"
Add-Lookup -SchemaName "adc_initiatinguser" -RelName "adc_mppimportjob_initiatinguser" -Label "Initiating User" -Target "systemuser"

# --- 4. Publish ---
Write-Host "`nPublishing customizations..." -ForegroundColor Yellow
try { Invoke-Api -Method POST -Path "PublishAllXml" -Body @{}; Write-Host "  Published." -ForegroundColor Green }
catch { Write-Host "  Publish: $_" -ForegroundColor Yellow }

Write-Host "`n=== Done ===" -ForegroundColor Cyan
Write-Host "Table adc_mppimportjob created with all fields and lookups."
