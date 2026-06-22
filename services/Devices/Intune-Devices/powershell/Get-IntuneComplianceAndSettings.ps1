#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Graph REST API for additional Intune compliance, configurations, detected apps, and tenant subscriptions.
.DESCRIPTION
    A cloud-native GRC Asset script designed to retrieve additional Intune details
    utilizing clean REST calls and exporting structured GRC results (JSON, CSV).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$CertificateBase64,

    [Parameter(Mandatory = $false)]
    [switch]$Interactive,

    [Parameter(Mandatory = $false)]
    [switch]$AiAgentMode
)

$ErrorActionPreference = 'Stop'

# Import our GRC Common Authentication and Export library
$commonModulePath = Join-Path -Path $PSScriptRoot -ChildPath "../../../../common/GRC-M365-Common.psm1"
if (Test-Path $commonModulePath) {
    Import-Module -Name $commonModulePath -Force
} else {
    Write-Error "Required GRC Common module not found at: $commonModulePath"
    return
}

# 1. Establish connection to M365 REST Graph context
try {
    if ($Interactive) {
        Connect-GRCEnvironment -Interactive
    } else {
        Connect-GRCEnvironment -TenantId $TenantId -ClientId $ClientId -CertificateBase64 $CertificateBase64
    }
} catch {
    Write-Error "Authentication failed: $_"
    return
}

# Helper to fetch all paginated values from an endpoint
function Get-GraphPaginatedData {
    param(
        [string]$Uri
    )
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    $currentUri = $Uri
    try {
        while ($currentUri) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $currentUri -ErrorAction Stop
            if ($response) {
                if ($response.value) {
                    foreach ($item in $response.value) {
                        $results.Add($item)
                    }
                } else {
                    # Endpoint might return a single object instead of collection (e.g., summaries)
                    $results.Add($response)
                    break
                }
                $currentUri = $response.'@odata.nextLink'
            } else {
                $currentUri = $null
            }
        }
    } catch {
        Write-Warning "Could not query endpoint $($Uri): $_"
    }
    return $results.ToArray()
}

# Helper to flatten nested objects and serialize complex types for CSV compatibility
function Flatten-Object {
    param(
        $InputObject
    )
    if ($null -eq $InputObject) { return $null }
    $flatObj = [ordered]@{}
    foreach ($prop in $InputObject.PSObject.Properties) {
        $name = $prop.Name
        # Skip OData metadata properties
        if ($name -like "@odata*") { continue }
        $val = $prop.Value
        if ($null -eq $val) {
            $flatObj[$name] = ""
        } elseif ($val -is [string] -or $val -is [valueType]) {
            $flatObj[$name] = $val
        } elseif ($val -is [array] -or $val -is [System.Collections.IEnumerable]) {
            $flatObj[$name] = ($val | ForEach-Object { $_.ToString() }) -join "; "
        } else {
            # Serialize nested objects to JSON string to prevent Export-Csv from crashing
            $flatObj[$name] = $val | ConvertTo-Json -Compress -Depth 2
        }
    }
    return [PSCustomObject]$flatObj
}

# Define endpoints mapping to their GRC AssetName and ServiceName
$endpoints = @(
    [PSCustomObject]@{ Service = "Devices"; Asset = "DeviceCompliancePolicies"; Uri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies" }
    [PSCustomObject]@{ Service = "Devices"; Asset = "DetectedApps"; Uri = "https://graph.microsoft.com/v1.0/deviceManagement/detectedApps" }
    [PSCustomObject]@{ Service = "Devices"; Asset = "DeviceCategories"; Uri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceCategories" }
    [PSCustomObject]@{ Service = "Devices"; Asset = "CompliancePolicyDeviceStateSummary"; Uri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicyDeviceStateSummary" }
    [PSCustomObject]@{ Service = "Devices"; Asset = "CompliancePolicySettingStateSummaries"; Uri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicySettingStateSummaries" }
    [PSCustomObject]@{ Service = "Devices"; Asset = "ConfigurationDeviceStateSummaries"; Uri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurationDeviceStateSummaries" }
    [PSCustomObject]@{ Service = "Devices"; Asset = "DeviceEnrollmentConfigurations"; Uri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations" }
    [PSCustomObject]@{ Service = "Devices"; Asset = "ManagedDeviceOverview"; Uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDeviceOverview" }
    [PSCustomObject]@{ Service = "Devices"; Asset = "RoleDefinitions"; Uri = "https://graph.microsoft.com/v1.0/deviceManagement/roleDefinitions" }
)

$combinedOutput = @{}

foreach ($ep in $endpoints) {
    Write-Verbose "Querying $($ep.Uri)..."
    $data = Get-GraphPaginatedData -Uri $ep.Uri
    
    # Flatten objects to prevent Export-Csv from crashing on complex nested types
    $flattenedData = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($item in $data) {
        if ($item) {
            $flattenedData.Add((Flatten-Object -InputObject $item))
        }
    }
    
    if ($AiAgentMode) {
        $combinedOutput[$ep.Asset] = $flattenedData.ToArray()
    } else {
        if ($flattenedData.Count -gt 0) {
            Export-GRCAssetData -ServiceName $ep.Service -AssetName $ep.Asset -Data $flattenedData.ToArray()
        }
    }
}

if ($AiAgentMode) {
    $combinedOutput | ConvertTo-Json -Depth 10
}
