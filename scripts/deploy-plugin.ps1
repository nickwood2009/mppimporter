<#
.SYNOPSIS
    Builds the ADC.MppImport assembly and deploys it to Dynamics 365 / Dataverse,
    then registers plugin steps for the async chunked import.

.DESCRIPTION
    This script:
    1. Builds the ADC.MppImport project in Release mode
    2. Pushes the assembly to D365 using pac plugin push
    3. Registers (or updates) async plugin steps for MppImportJobPlugin

.PARAMETER EnvironmentUrl
    The Dataverse environment URL (e.g., https://orgfbe0a613.crm6.dynamics.com)

.PARAMETER AssemblyId
    The plugin assembly GUID in D365 (required for updates, optional for first deploy).
    After first deploy, the script outputs the assembly ID to use for subsequent deploys.

.PARAMETER BuildConfig
    Build configuration (default: Release)

.EXAMPLE
    # First deployment:
    .\deploy-plugin.ps1 -EnvironmentUrl https://orgfbe0a613.crm6.dynamics.com

    # Subsequent deployments (update existing):
    .\deploy-plugin.ps1 -EnvironmentUrl https://orgfbe0a613.crm6.dynamics.com -AssemblyId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$EnvironmentUrl,

    [string]$AssemblyId = "",

    [string]$BuildConfig = "Release"
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptDir
$projectDir = Join-Path $repoRoot "ADC.MppImport"
$projectFile = Join-Path $projectDir "ADC.MppImport.csproj"
$outputDir = Join-Path $projectDir "bin\$BuildConfig"
$assemblyPath = Join-Path $outputDir "ADC.MppImport.dll"

Write-Host "=== Deploy ADC.MppImport Plugin ===" -ForegroundColor Cyan
Write-Host "Environment:  $EnvironmentUrl"
Write-Host "Project:      $projectFile"
Write-Host "Build Config: $BuildConfig"
Write-Host ""

# --- Step 1: Build ---
Write-Host "Step 1: Building..." -ForegroundColor Yellow
$msbuildArgs = @(
    $projectFile,
    "/p:Configuration=$BuildConfig",
    "/v:minimal",
    "/t:Build"
)
& msbuild @msbuildArgs
if ($LASTEXITCODE -ne 0) {
    Write-Host "BUILD FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Build succeeded." -ForegroundColor Green

if (-not (Test-Path $assemblyPath)) {
    Write-Host "ERROR: Assembly not found at $assemblyPath" -ForegroundColor Red
    exit 1
}
Write-Host "Assembly: $assemblyPath"
Write-Host ""

# --- Step 2: Authenticate ---
Write-Host "Step 2: Checking auth..." -ForegroundColor Yellow
try {
    $authList = pac auth list 2>&1
    if ($authList -notmatch $EnvironmentUrl) {
        Write-Host "Authenticating to $EnvironmentUrl..."
        pac auth create --url $EnvironmentUrl
    } else {
        Write-Host "Already authenticated."
    }
} catch {
    pac auth create --url $EnvironmentUrl
}
Write-Host ""

# --- Step 3: Push Assembly ---
Write-Host "Step 3: Pushing assembly to D365..." -ForegroundColor Yellow

if ($AssemblyId) {
    Write-Host "Updating existing assembly: $AssemblyId"
    pac plugin push --id $AssemblyId --path $assemblyPath
} else {
    Write-Host "Registering new assembly..."
    $pushResult = pac plugin push --path $assemblyPath 2>&1
    Write-Host $pushResult

    # Try to extract the assembly ID from output
    $idMatch = $pushResult | Select-String -Pattern "[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}"
    if ($idMatch) {
        $AssemblyId = $idMatch.Matches[0].Value
        Write-Host ""
        Write-Host "Assembly ID: $AssemblyId" -ForegroundColor Green
        Write-Host "Save this ID for future deployments!" -ForegroundColor Yellow
    }
}

if ($LASTEXITCODE -ne 0) {
    Write-Host "PUSH FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Assembly pushed successfully." -ForegroundColor Green
Write-Host ""

# --- Step 4: Register Plugin Steps ---
Write-Host "Step 4: Registering plugin steps..." -ForegroundColor Yellow
Write-Host ""
Write-Host "Plugin steps must be registered using the Plugin Registration Tool or Web API." -ForegroundColor DarkYellow
Write-Host "Register these steps manually if not already done:" -ForegroundColor DarkYellow
Write-Host ""
Write-Host "  Plugin Type: ADC.MppImport.Plugins.MppImportJobPlugin" -ForegroundColor White
Write-Host ""
Write-Host "  Step 1: Create of adc_mppimportjob" -ForegroundColor White
Write-Host "    Message:    Create" -ForegroundColor Gray
Write-Host "    Entity:     adc_mppimportjob" -ForegroundColor Gray
Write-Host "    Stage:      PostOperation (40)" -ForegroundColor Gray
Write-Host "    Mode:       Asynchronous (1)" -ForegroundColor Gray
Write-Host "    Rank:       1" -ForegroundColor Gray
Write-Host ""
Write-Host "  Step 2: Update of adc_mppimportjob" -ForegroundColor White
Write-Host "    Message:    Update" -ForegroundColor Gray
Write-Host "    Entity:     adc_mppimportjob" -ForegroundColor Gray
Write-Host "    Stage:      PostOperation (40)" -ForegroundColor Gray
Write-Host "    Mode:       Asynchronous (1)" -ForegroundColor Gray
Write-Host "    Rank:       1" -ForegroundColor Gray
Write-Host "    Filtering:  adc_status,adc_tick" -ForegroundColor Gray
Write-Host ""

# Attempt to register steps via Web API
Write-Host "Attempting automated step registration via Web API..." -ForegroundColor Yellow

function Get-AuthToken {
    $tokenOutput = (pac auth token --environment $EnvironmentUrl 2>&1) | Select-String -Pattern "^[A-Za-z0-9]" | Select-Object -First 1
    return $tokenOutput.ToString().Trim()
}

function Invoke-DataverseApi {
    param(
        [string]$Method,
        [string]$Path,
        [object]$Body
    )

    $token = Get-AuthToken
    $headers = @{
        "Authorization" = "Bearer $token"
        "Content-Type" = "application/json"
        "OData-MaxVersion" = "4.0"
        "OData-Version" = "4.0"
    }

    $uri = "$EnvironmentUrl/api/data/v9.2/$Path"
    $params = @{
        Uri = $uri
        Method = $Method
        Headers = $headers
        ContentType = "application/json"
    }

    if ($Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10)
    }

    return Invoke-RestMethod @params
}

# Find the plugin type ID
try {
    $pluginTypes = Invoke-DataverseApi -Method GET -Path "plugintypes?`$filter=typename eq 'ADC.MppImport.Plugins.MppImportJobPlugin'&`$select=plugintypeid,typename"

    if ($pluginTypes.value.Count -eq 0) {
        Write-Host "  Plugin type not found yet. It may take a moment after push." -ForegroundColor Yellow
        Write-Host "  Re-run this script or register steps manually." -ForegroundColor Yellow
    } else {
        $pluginTypeId = $pluginTypes.value[0].plugintypeid
        Write-Host "  Found plugin type: $pluginTypeId" -ForegroundColor Green

        # Find the Create message ID
        $createMsg = Invoke-DataverseApi -Method GET -Path "sdkmessages?`$filter=name eq 'Create'&`$select=sdkmessageid"
        $updateMsg = Invoke-DataverseApi -Method GET -Path "sdkmessages?`$filter=name eq 'Update'&`$select=sdkmessageid"

        # Find message filter for adc_mppimportjob
        $createFilter = Invoke-DataverseApi -Method GET -Path "sdkmessagefilters?`$filter=primaryobjecttypecode eq 'adc_mppimportjob' and sdkmessageid/sdkmessageid eq $($createMsg.value[0].sdkmessageid)&`$select=sdkmessagefilterid"
        $updateFilter = Invoke-DataverseApi -Method GET -Path "sdkmessagefilters?`$filter=primaryobjecttypecode eq 'adc_mppimportjob' and sdkmessageid/sdkmessageid eq $($updateMsg.value[0].sdkmessageid)&`$select=sdkmessagefilterid"

        if ($createFilter.value.Count -gt 0) {
            # Register Create step
            $createStep = @{
                name = "MppImportJobPlugin: Create of adc_mppimportjob"
                mode = 1  # Async
                rank = 1
                stage = 40  # PostOperation
                "plugintypeid@odata.bind" = "/plugintypes($pluginTypeId)"
                "sdkmessageid@odata.bind" = "/sdkmessages($($createMsg.value[0].sdkmessageid))"
                "sdkmessagefilterid@odata.bind" = "/sdkmessagefilters($($createFilter.value[0].sdkmessagefilterid))"
                asyncautodelete = $true
            }

            try {
                Invoke-DataverseApi -Method POST -Path "sdkmessageprocessingsteps" -Body $createStep
                Write-Host "  Registered Create step." -ForegroundColor Green
            } catch {
                Write-Host "  Create step: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        if ($updateFilter.value.Count -gt 0) {
            # Register Update step with filtering attributes
            $updateStep = @{
                name = "MppImportJobPlugin: Update of adc_mppimportjob"
                mode = 1  # Async
                rank = 1
                stage = 40  # PostOperation
                filteringattributes = "adc_status,adc_tick"
                "plugintypeid@odata.bind" = "/plugintypes($pluginTypeId)"
                "sdkmessageid@odata.bind" = "/sdkmessages($($updateMsg.value[0].sdkmessageid))"
                "sdkmessagefilterid@odata.bind" = "/sdkmessagefilters($($updateFilter.value[0].sdkmessagefilterid))"
                asyncautodelete = $true
            }

            try {
                Invoke-DataverseApi -Method POST -Path "sdkmessageprocessingsteps" -Body $updateStep
                Write-Host "  Registered Update step." -ForegroundColor Green
            } catch {
                Write-Host "  Update step: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        if ($createFilter.value.Count -eq 0 -or $updateFilter.value.Count -eq 0) {
            Write-Host "  Message filters not found for adc_mppimportjob." -ForegroundColor Yellow
            Write-Host "  Ensure the table is created first (run create-import-job-table.ps1)." -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "  Auto-registration failed: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "  Please register steps manually using Plugin Registration Tool." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Deployment Complete ===" -ForegroundColor Cyan
if ($AssemblyId) {
    Write-Host "Assembly ID: $AssemblyId"
    Write-Host ""
    Write-Host "For future deploys, run:" -ForegroundColor Gray
    Write-Host "  .\deploy-plugin.ps1 -EnvironmentUrl $EnvironmentUrl -AssemblyId $AssemblyId" -ForegroundColor White
}
