<#
.SYNOPSIS
    Creates the adc_adccasetemplate table with a file field (adc_templatefile)
    for uploading MPP files. Used by both the existing and async import services.

.PARAMETER CrmUrl
    Dataverse environment URL

.PARAMETER ClientId
    App registration client ID

.PARAMETER ClientSecret
    App registration client secret
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CrmUrl,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$ClientSecret
)

$ErrorActionPreference = "Stop"
$CrmUrl = $CrmUrl.TrimEnd('/')

Write-Host "=== Create adc_adccasetemplate Table ===" -ForegroundColor Cyan

# --- Auth ---
Write-Host "Discovering tenant..."
$tenantAuthority = $null
try {
    Invoke-WebRequest -Uri "$CrmUrl/api/data/v9.2/" -Method GET -UseBasicParsing -ErrorAction Stop | Out-Null
} catch {
    if ($_.Exception.Response) {
        $wwwAuth = $_.Exception.Response.Headers["WWW-Authenticate"]
        if ($wwwAuth) {
            $match = [regex]::Match($wwwAuth, 'authorization_uri="?([^"\s,]+)"?')
            if ($match.Success) {
                $tenantAuthority = $match.Groups[1].Value -replace '/oauth2/authorize$', ''
            }
        }
    }
}
if (-not $tenantAuthority) { throw "Could not discover tenant" }

$tokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    resource      = $CrmUrl
}
$token = (Invoke-RestMethod -Uri "$tenantAuthority/oauth2/token" -Method POST -Body $tokenBody -ContentType "application/x-www-form-urlencoded").access_token
Write-Host "Authenticated." -ForegroundColor Green

function Invoke-CrmApi {
    param([string]$Method, [string]$Path, [object]$Body, [switch]$IgnoreErrors, [switch]$Raw)

    $headers = @{
        "Authorization"    = "Bearer $token"
        "Content-Type"     = "application/json"
        "OData-MaxVersion" = "4.0"
        "OData-Version"    = "4.0"
        "Accept"           = "application/json"
    }

    $uri = "$CrmUrl/api/data/v9.2/$Path"
    $params = @{ Uri = $uri; Method = $Method; Headers = $headers; ContentType = "application/json" }

    if ($Body) {
        $params.Body = $Body | ConvertTo-Json -Depth 20 -Compress
    }

    try {
        if ($Raw) {
            return Invoke-WebRequest @params
        }
        return Invoke-RestMethod @params
    } catch {
        $statusCode = 0
        if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }
        $errorBody = ""
        try {
            $reader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
            $errorBody = $reader.ReadToEnd()
        } catch { }

        if ($IgnoreErrors) {
            Write-Host "    (ignored: $statusCode)" -ForegroundColor DarkYellow
            return $null
        }
        Write-Host "  API Error ($statusCode): $errorBody" -ForegroundColor Red
        throw
    }
}

# --- Check if table exists ---
Write-Host "`nChecking if adc_adccasetemplate exists..."
$existing = $null
try {
    $existing = Invoke-CrmApi -Method GET -Path "EntityDefinitions(LogicalName='adc_adccasetemplate')?`$select=LogicalName" -IgnoreErrors
} catch { }

if ($existing) {
    Write-Host "  Table already exists." -ForegroundColor DarkYellow
} else {
    Write-Host "  Creating entity adc_adccasetemplate..."

    $entityDef = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.EntityMetadata"
        SchemaName = "adc_adccasetemplate"
        DisplayName = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label = "ADC Case Template"
                LanguageCode = 1033
            })
        }
        DisplayCollectionName = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label = "ADC Case Templates"
                LanguageCode = 1033
            })
        }
        Description = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label = "Case template with MPP file for project import"
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
    Start-Sleep -Seconds 3
}

# --- Add file field ---
Write-Host "`nChecking for adc_templatefile field..."
$existingField = $null
try {
    $existingField = Invoke-CrmApi -Method GET -Path "EntityDefinitions(LogicalName='adc_adccasetemplate')/Attributes(LogicalName='adc_templatefile')?`$select=LogicalName" -IgnoreErrors
} catch { }

if ($existingField) {
    Write-Host "  File field already exists." -ForegroundColor DarkYellow
} else {
    Write-Host "  Adding adc_templatefile (File column)..."

    $fileAttr = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.FileAttributeMetadata"
        SchemaName = "adc_templatefile"
        MaxSizeInKB = 131072  # 128 MB
        DisplayName = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label = "Template File"
                LanguageCode = 1033
            })
        }
        Description = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label = "MPP file to import into project"
                LanguageCode = 1033
            })
        }
        RequiredLevel = @{ Value = "None" }
    }

    Invoke-CrmApi -Method POST -Path "EntityDefinitions(LogicalName='adc_adccasetemplate')/Attributes" -Body $fileAttr
    Write-Host "  File field created." -ForegroundColor Green
}

# --- Publish ---
Write-Host "`nPublishing customizations..."
Invoke-CrmApi -Method POST -Path "PublishAllXml" -Body @{}
Write-Host "Published." -ForegroundColor Green

Write-Host "`n=== Done ===" -ForegroundColor Cyan
Write-Host "Table: adc_adccasetemplate"
Write-Host "Fields: adc_name (text), adc_templatefile (file, 128MB max)"
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Create a record in D365 (or via API)"
Write-Host "  2. Upload an MPP file to the adc_templatefile field"
Write-Host "  3. Use the record ID as CaseTemplate input to the import"
