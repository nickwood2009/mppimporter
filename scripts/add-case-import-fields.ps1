<#
.SYNOPSIS
    Adds import status fields to the existing adc_case table:
    - adc_importstatus (optionset: Queued, Processing, Completed, CompletedWithWarnings, Failed)
    - adc_importmessage (string, 500 chars)
    Also creates and registers the JS web resource for form notifications.

.PARAMETER CrmUrl
    Dataverse environment URL

.PARAMETER ClientId
    App registration client ID

.PARAMETER ClientSecret
    App registration client secret

.EXAMPLE
    .\add-case-import-fields.ps1 -CrmUrl https://orgfbe0a613.crm6.dynamics.com -ClientId ... -ClientSecret ...
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

Write-Host "=== Add Import Status Fields to adc_case ===" -ForegroundColor Cyan

# --- Auth ---
Write-Host "Authenticating..."

# Try tenant discovery, fall back to well-known tenant
$tenantAuthority = "https://login.microsoftonline.com/2abc9b7d-4ced-403e-b3e3-b8e5a6240807"
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

$tokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    resource      = $CrmUrl
}
$token = (Invoke-RestMethod -Uri "$tenantAuthority/oauth2/token" -Method POST -Body $tokenBody -ContentType "application/x-www-form-urlencoded").access_token
Write-Host "Authenticated." -ForegroundColor Green

function Invoke-CrmApi {
    param([string]$Method, [string]$Path, [object]$Body, [switch]$IgnoreErrors)

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
            Write-Host "    (ignored: $statusCode - $errorBody)" -ForegroundColor DarkYellow
            return $null
        }
        Write-Host "  API Error ($statusCode): $errorBody" -ForegroundColor Red
        throw
    }
}

# ============================================================
# 1. Add adc_importstatus picklist (optionset) to adc_case
# ============================================================
Write-Host "`nChecking adc_importstatus field..."
$existingAttr = $null
try {
    $existingAttr = Invoke-CrmApi -Method GET -Path "EntityDefinitions(LogicalName='adc_case')/Attributes(LogicalName='adc_importstatus')?`$select=LogicalName" -IgnoreErrors
} catch { }

if ($existingAttr) {
    Write-Host "  Field already exists." -ForegroundColor DarkYellow
} else {
    Write-Host "  Creating adc_importstatus (picklist)..."

    $picklistDef = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.PicklistAttributeMetadata"
        SchemaName    = "adc_importstatus"
        DisplayName   = @{
            "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label         = "Import Status"
                LanguageCode  = 1033
            })
        }
        Description = @{
            "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label         = "MPP import status for this case"
                LanguageCode  = 1033
            })
        }
        RequiredLevel = @{ Value = "None" }
        OptionSet     = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.OptionSetMetadata"
            IsGlobal      = $false
            OptionSetType = "Picklist"
            Options       = @(
                @{
                    Value = 0
                    Label = @{
                        "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
                        LocalizedLabels = @(@{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            Label         = "Queued"
                            LanguageCode  = 1033
                        })
                    }
                },
                @{
                    Value = 1
                    Label = @{
                        "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
                        LocalizedLabels = @(@{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            Label         = "Processing"
                            LanguageCode  = 1033
                        })
                    }
                },
                @{
                    Value = 2
                    Label = @{
                        "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
                        LocalizedLabels = @(@{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            Label         = "Completed"
                            LanguageCode  = 1033
                        })
                    }
                },
                @{
                    Value = 3
                    Label = @{
                        "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
                        LocalizedLabels = @(@{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            Label         = "Completed with Warnings"
                            LanguageCode  = 1033
                        })
                    }
                },
                @{
                    Value = 4
                    Label = @{
                        "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
                        LocalizedLabels = @(@{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            Label         = "Failed"
                            LanguageCode  = 1033
                        })
                    }
                }
            )
        }
    }

    Invoke-CrmApi -Method POST -Path "EntityDefinitions(LogicalName='adc_case')/Attributes" -Body $picklistDef
    Write-Host "  Created." -ForegroundColor Green
}

# ============================================================
# 2. Add adc_importmessage string to adc_case
# ============================================================
Write-Host "`nChecking adc_importmessage field..."
$existingMsg = $null
try {
    $existingMsg = Invoke-CrmApi -Method GET -Path "EntityDefinitions(LogicalName='adc_case')/Attributes(LogicalName='adc_importmessage')?`$select=LogicalName" -IgnoreErrors
} catch { }

if ($existingMsg) {
    Write-Host "  Field already exists." -ForegroundColor DarkYellow
} else {
    Write-Host "  Creating adc_importmessage (string)..."

    $stringDef = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
        SchemaName    = "adc_importmessage"
        AttributeType = "String"
        FormatName    = @{ Value = "Text" }
        MaxLength     = 500
        DisplayName   = @{
            "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label         = "Import Message"
                LanguageCode  = 1033
            })
        }
        Description = @{
            "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label         = "Status message from the most recent MPP import"
                LanguageCode  = 1033
            })
        }
        RequiredLevel = @{ Value = "None" }
    }

    Invoke-CrmApi -Method POST -Path "EntityDefinitions(LogicalName='adc_case')/Attributes" -Body $stringDef
    Write-Host "  Created." -ForegroundColor Green
}

# ============================================================
# 3. Create JS web resource for form notifications
# ============================================================
Write-Host "`nChecking JS web resource..."
$existingWr = $null
try {
    $existingWr = Invoke-CrmApi -Method GET -Path "webresourceset?`$filter=name eq 'adc_/scripts/caseImportBanner.js'&`$select=webresourceid" -IgnoreErrors
} catch { }

# Read JS file from local path
$jsPath = Join-Path $PSScriptRoot "..\webresources\adc_caseImportBanner.js"
if (-not (Test-Path $jsPath)) {
    Write-Host "  JS file not found at $jsPath, skipping web resource creation." -ForegroundColor Yellow
} else {
    $jsContent = [System.IO.File]::ReadAllText($jsPath)
    $jsBase64  = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jsContent))

    if ($existingWr -and $existingWr.value -and $existingWr.value.Count -gt 0) {
        $wrId = $existingWr.value[0].webresourceid
        Write-Host "  Updating existing web resource ($wrId)..."
        Invoke-CrmApi -Method PATCH -Path "webresourceset($wrId)" -Body @{
            content = $jsBase64
        }
        Write-Host "  Updated." -ForegroundColor Green
    } else {
        Write-Host "  Creating web resource..."
        $wrDef = @{
            name            = "adc_/scripts/caseImportBanner.js"
            displayname     = "Case Import Banner Notifications"
            description     = "Shows form-level notifications on adc_case for MPP import status"
            webresourcetype = 3  # JScript
            content         = $jsBase64
        }
        $result = Invoke-CrmApi -Method POST -Path "webresourceset" -Body $wrDef
        Write-Host "  Created." -ForegroundColor Green
    }
}

# ============================================================
# 4. Publish
# ============================================================
Write-Host "`nPublishing customizations..."
Invoke-CrmApi -Method POST -Path "PublishAllXml" -Body @{}
Write-Host "Published." -ForegroundColor Green

Write-Host "`n=== Done ===" -ForegroundColor Cyan
Write-Host "Fields added to adc_case:"
Write-Host "  adc_importstatus  (picklist: 0=Queued, 1=Processing, 2=Completed, 3=CompletedWithWarnings, 4=Failed)"
Write-Host "  adc_importmessage (string, 500 chars)"
Write-Host "Web resource: adc_/scripts/caseImportBanner.js"
Write-Host ""
Write-Host "NEXT STEPS:" -ForegroundColor Yellow
Write-Host "  1. Add adc_importstatus and adc_importmessage to the adc_case form"
Write-Host "  2. Register caseImportBanner.js as onLoad event on the adc_case form"
Write-Host "     Function: ADC.CaseImportBanner.onLoad"
Write-Host "     Pass execution context: Yes"
