<#
.SYNOPSIS
    Creates the adc_case table with required relationships:
    - adc_case -> adc_adccasetemplate (lookup: adc_casetemplate)
    - adc_case -> msdyn_project (lookup: adc_project)
    - msdyn_project -> adc_case (lookup: adc_case, back-link)
    - adc_mppimportjob -> adc_case (lookup: adc_case)

.PARAMETER CrmUrl
    Dataverse environment URL

.PARAMETER ClientId
    App registration client ID

.PARAMETER ClientSecret
    App registration client secret

.EXAMPLE
    .\create-case-table.ps1 -CrmUrl https://orgfbe0a613.crm6.dynamics.com -ClientId ... -ClientSecret ...
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

Write-Host "=== Create adc_case Table ===" -ForegroundColor Cyan

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
# 1. Create adc_case entity
# ============================================================
Write-Host "`nChecking if adc_case exists..."
$existing = $null
try {
    $existing = Invoke-CrmApi -Method GET -Path "EntityDefinitions(LogicalName='adc_case')?`$select=LogicalName" -IgnoreErrors
} catch { }

if ($existing) {
    Write-Host "  Table already exists." -ForegroundColor DarkYellow
} else {
    Write-Host "  Creating entity adc_case..."

    $entityDef = @{
        "@odata.type" = "Microsoft.Dynamics.CRM.EntityMetadata"
        SchemaName = "adc_case"
        DisplayName = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label = "ADC Case"
                LanguageCode = 1033
            })
        }
        DisplayCollectionName = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label = "ADC Cases"
                LanguageCode = 1033
            })
        }
        Description = @{
            "@odata.type" = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label = "Case record that triggers MPP import from a case template"
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

# ============================================================
# 2. Add lookup: adc_case -> adc_adccasetemplate
# ============================================================
Write-Host "`nAdding lookup: adc_casetemplate -> adc_adccasetemplate..."
$rel1 = @{
    SchemaName        = "adc_case_casetemplate"
    "@odata.type"     = "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata"
    ReferencedEntity  = "adc_adccasetemplate"
    ReferencingEntity = "adc_case"
    Lookup            = @{
        SchemaName   = "adc_casetemplate"
        DisplayName  = @{
            "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label         = "Case Template"
                LanguageCode  = 1033
            })
        }
    }
}
try {
    Invoke-CrmApi -Method POST -Path "RelationshipDefinitions" -Body $rel1
    Write-Host "  OK" -ForegroundColor Green
} catch {
    Write-Host "  (may already exist, continuing)" -ForegroundColor DarkYellow
}

# ============================================================
# 3. Add lookup: adc_case -> msdyn_project
# ============================================================
Write-Host "`nAdding lookup: adc_project -> msdyn_project (on adc_case)..."
$rel2 = @{
    SchemaName        = "adc_case_project"
    "@odata.type"     = "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata"
    ReferencedEntity  = "msdyn_project"
    ReferencingEntity = "adc_case"
    Lookup            = @{
        SchemaName   = "adc_project"
        DisplayName  = @{
            "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label         = "Project"
                LanguageCode  = 1033
            })
        }
    }
}
try {
    Invoke-CrmApi -Method POST -Path "RelationshipDefinitions" -Body $rel2
    Write-Host "  OK" -ForegroundColor Green
} catch {
    Write-Host "  (may already exist, continuing)" -ForegroundColor DarkYellow
}

# ============================================================
# 4. Add back-link: msdyn_project -> adc_case
# ============================================================
Write-Host "`nAdding lookup: adc_case -> adc_case (on msdyn_project, back-link)..."
$rel3 = @{
    SchemaName        = "adc_project_case"
    "@odata.type"     = "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata"
    ReferencedEntity  = "adc_case"
    ReferencingEntity = "msdyn_project"
    Lookup            = @{
        SchemaName   = "adc_case"
        DisplayName  = @{
            "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label         = "ADC Case"
                LanguageCode  = 1033
            })
        }
    }
}
try {
    Invoke-CrmApi -Method POST -Path "RelationshipDefinitions" -Body $rel3
    Write-Host "  OK" -ForegroundColor Green
} catch {
    Write-Host "  (may already exist, continuing)" -ForegroundColor DarkYellow
}

# ============================================================
# 5. Add lookup: adc_mppimportjob -> adc_case
# ============================================================
Write-Host "`nAdding lookup: adc_case -> adc_case (on adc_mppimportjob)..."
$rel4 = @{
    SchemaName        = "adc_mppimportjob_case"
    "@odata.type"     = "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata"
    ReferencedEntity  = "adc_case"
    ReferencingEntity = "adc_mppimportjob"
    Lookup            = @{
        SchemaName   = "adc_case"
        DisplayName  = @{
            "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label         = "Case"
                LanguageCode  = 1033
            })
        }
    }
}
try {
    Invoke-CrmApi -Method POST -Path "RelationshipDefinitions" -Body $rel4
    Write-Host "  OK" -ForegroundColor Green
} catch {
    Write-Host "  (may already exist, continuing)" -ForegroundColor DarkYellow
}

# ============================================================
# 6. Publish
# ============================================================
Write-Host "`nPublishing customizations..."
Invoke-CrmApi -Method POST -Path "PublishAllXml" -Body @{}
Write-Host "Published." -ForegroundColor Green

Write-Host "`n=== Done ===" -ForegroundColor Cyan
Write-Host "Table: adc_case"
Write-Host "Fields: adc_name (text, primary), adc_casetemplate (lookup), adc_project (lookup)"
Write-Host "Back-link: msdyn_project.adc_case (lookup to adc_case)"
Write-Host "Import job link: adc_mppimportjob.adc_case (lookup to adc_case)"
