<#
.SYNOPSIS
    Adds adc_case and adc_initiatinguser lookup fields to the adc_mppimportjob table
    for in-app notification support.

.DESCRIPTION
    Run this script once per environment after the adc_mppimportjob table exists.
    Adds:
    - adc_case: Lookup to adc_case (originating case record)
    - adc_initiatinguser: Lookup to systemuser (notification recipient)

.EXAMPLE
    .\add-notification-fields.ps1
#>

$ErrorActionPreference = "Stop"

# Load .env from project root
$envFile = Join-Path $PSScriptRoot "..\.env"
if (Test-Path $envFile) {
    Get-Content $envFile | ForEach-Object {
        if ($_ -match '^\s*([^#][^=]+)=(.*)$') {
            [System.Environment]::SetEnvironmentVariable($matches[1].Trim(), $matches[2].Trim())
        }
    }
}

$crmUrl = $env:CRM_URL
$tenantId = $env:TENANT_ID
$clientId = $env:CLIENT_ID
$clientSecret = $env:CLIENT_SECRET

if (-not $crmUrl -or -not $tenantId -or -not $clientId -or -not $clientSecret) {
    Write-Error "Missing .env values (CRM_URL, TENANT_ID, CLIENT_ID, CLIENT_SECRET)"
    exit 1
}

Write-Host "=== Add Notification Fields to adc_mppimportjob ===" -ForegroundColor Cyan
Write-Host "Environment: $crmUrl"

# Get OAuth token
$body = @{
    grant_type    = "client_credentials"
    client_id     = $clientId
    client_secret = $clientSecret
    resource      = $crmUrl
}
$token = (Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/token" `
    -Method POST -Body $body -ContentType "application/x-www-form-urlencoded").access_token

$headers = @{
    Authorization    = "Bearer $token"
    "Content-Type"   = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version"  = "4.0"
}

$apiBase = "$crmUrl/api/data/v9.2"

# --- Add adc_case lookup (adc_mppimportjob -> adc_case) ---
Write-Host "`nAdding lookup: adc_case -> adc_case..." -ForegroundColor Yellow
$caseLookup = @{
    SchemaName         = "adc_mppimportjob_case"
    "@odata.type"      = "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata"
    ReferencedEntity   = "adc_case"
    ReferencingEntity  = "adc_mppimportjob"
    Lookup             = @{
        SchemaName    = "adc_case"
        DisplayName   = @{
            "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label         = "Case"
                LanguageCode  = 1033
            })
        }
    }
} | ConvertTo-Json -Depth 10

try {
    Invoke-RestMethod -Uri "$apiBase/RelationshipDefinitions" -Method POST -Headers $headers -Body $caseLookup -ContentType "application/json"
    Write-Host "  OK" -ForegroundColor Green
} catch {
    $code = $_.Exception.Response.StatusCode.value__
    if ($code -eq 409 -or $code -eq 400) {
        Write-Host "  (already exists or conflict, continuing)" -ForegroundColor DarkYellow
    } else {
        Write-Host "  Error: $_" -ForegroundColor Red
    }
}

# --- Add adc_initiatinguser lookup (adc_mppimportjob -> systemuser) ---
Write-Host "`nAdding lookup: adc_initiatinguser -> systemuser..." -ForegroundColor Yellow
$userLookup = @{
    SchemaName         = "adc_mppimportjob_initiatinguser"
    "@odata.type"      = "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata"
    ReferencedEntity   = "systemuser"
    ReferencingEntity  = "adc_mppimportjob"
    Lookup             = @{
        SchemaName    = "adc_initiatinguser"
        DisplayName   = @{
            "@odata.type"   = "Microsoft.Dynamics.CRM.Label"
            LocalizedLabels = @(@{
                "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                Label         = "Initiating User"
                LanguageCode  = 1033
            })
        }
    }
} | ConvertTo-Json -Depth 10

try {
    Invoke-RestMethod -Uri "$apiBase/RelationshipDefinitions" -Method POST -Headers $headers -Body $userLookup -ContentType "application/json"
    Write-Host "  OK" -ForegroundColor Green
} catch {
    $code = $_.Exception.Response.StatusCode.value__
    if ($code -eq 409 -or $code -eq 400) {
        Write-Host "  (already exists or conflict, continuing)" -ForegroundColor DarkYellow
    } else {
        Write-Host "  Error: $_" -ForegroundColor Red
    }
}

# Publish
Write-Host "`nPublishing customizations..." -ForegroundColor Yellow
try {
    Invoke-RestMethod -Uri "$apiBase/PublishAllXml" -Method POST -Headers $headers -Body "{}" -ContentType "application/json"
    Write-Host "  Published." -ForegroundColor Green
} catch {
    Write-Host "  Publish: $_" -ForegroundColor Yellow
}

Write-Host "`n=== Done ===" -ForegroundColor Cyan
Write-Host "Fields added: adc_case (lookup to adc_case), adc_initiatinguser (lookup to systemuser)"
