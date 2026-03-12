<#
.SYNOPSIS
    Adds duration variance tracking fields to the msdyn_projecttask entity in Dataverse.

.DESCRIPTION
    Creates 5 custom fields on msdyn_projecttask for tracking source MPP durations
    and explaining any PSS scheduling variance.

.PARAMETER EnvironmentUrl
    The Dataverse environment URL (e.g., https://orgfbe0a613.crm6.dynamics.com)

.EXAMPLE
    .\add-projecttask-variance-fields.ps1 -EnvironmentUrl https://orgfbe0a613.crm6.dynamics.com
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$EnvironmentUrl
)

$ErrorActionPreference = "Stop"

Write-Host "=== Add Duration Variance Fields to msdyn_projecttask ===" -ForegroundColor Cyan
Write-Host "Environment: $EnvironmentUrl"
Write-Host ""

$headers = @{
    "Content-Type" = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version" = "4.0"
}

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

function Add-ProjectTaskField {
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
            $attrDef["MinValue"] = 0
            $attrDef["MaxValue"] = 2147483647
        }
        "Decimal" {
            $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.DecimalAttributeMetadata"
            $attrDef["AttributeType"] = "Decimal"
            $attrDef["Precision"] = 2
            $attrDef["MinValue"] = -100000
            $attrDef["MaxValue"] = 100000
        }
        "String" {
            $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.StringAttributeMetadata"
            $attrDef["AttributeType"] = "String"
            $attrDef["FormatName"] = @{ Value = "Text" }
            $attrDef["MaxLength"] = if ($ExtraProps.MaxLength) { $ExtraProps.MaxLength } else { 200 }
        }
        "Boolean" {
            $attrDef["@odata.type"] = "Microsoft.Dynamics.CRM.BooleanAttributeMetadata"
            $attrDef["AttributeType"] = "Boolean"
            $attrDef["OptionSet"] = @{
                "@odata.type" = "Microsoft.Dynamics.CRM.BooleanOptionSetMetadata"
                TrueOption = @{
                    Value = 1
                    Label = @{
                        "@odata.type" = "Microsoft.Dynamics.CRM.Label"
                        LocalizedLabels = @(@{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            Label = "Yes"
                            LanguageCode = 1033
                        })
                    }
                }
                FalseOption = @{
                    Value = 0
                    Label = @{
                        "@odata.type" = "Microsoft.Dynamics.CRM.Label"
                        LocalizedLabels = @(@{
                            "@odata.type" = "Microsoft.Dynamics.CRM.LocalizedLabel"
                            Label = "No"
                            LanguageCode = 1033
                        })
                    }
                }
            }
        }
    }

    Write-Host "  Adding field: $SchemaName ($Type) - $Label..."
    try {
        Invoke-DataverseApi -Method POST -Path "EntityDefinitions(LogicalName='msdyn_projecttask')/Attributes" -Body $attrDef
        Write-Host "    OK" -ForegroundColor Green
    } catch {
        Write-Host "    $_" -ForegroundColor Yellow
    }
}

# Add the 5 variance tracking fields
Add-ProjectTaskField -SchemaName "adc_sourcedurationdays" -Label "Source Duration (Days)" -Type "Integer"
Add-ProjectTaskField -SchemaName "adc_sourcedurationhours" -Label "Source Duration (Hours)" -Type "Integer"
Add-ProjectTaskField -SchemaName "adc_durationvariancedays" -Label "Duration Variance (Days)" -Type "Decimal"
Add-ProjectTaskField -SchemaName "adc_durationvariancereason" -Label "Duration Variance Reason" -Type "String" -ExtraProps @{ MaxLength = 200 }
Add-ProjectTaskField -SchemaName "adc_issourcemilestone" -Label "Is Source Milestone" -Type "Boolean"

# Publish customizations
Write-Host "`nPublishing customizations..." -ForegroundColor Yellow
try {
    Invoke-DataverseApi -Method POST -Path "PublishAllXml" -Body @{}
    Write-Host "Published." -ForegroundColor Green
} catch {
    Write-Host "Publish: $_" -ForegroundColor Yellow
}

Write-Host "`n=== Done ===" -ForegroundColor Cyan
Write-Host "Fields added to msdyn_projecttask:"
Write-Host "  adc_sourcedurationdays    (Integer)  - Original MPP duration in days"
Write-Host "  adc_sourcedurationhours   (Integer)  - Original MPP duration in hours"
Write-Host "  adc_durationvariancedays  (Decimal)  - Difference: PSS duration - source duration"
Write-Host "  adc_durationvariancereason (String)  - Auto-populated explanation"
Write-Host "  adc_issourcemilestone     (Boolean)  - Whether source was 0-duration milestone"
