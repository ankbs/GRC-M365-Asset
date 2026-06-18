#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Purview for sensitivity labels and DLP policy configurations.
.DESCRIPTION
    Retrieves key Purview Compliance and Information Protection metrics for GRC auditing.
    Exports structured JSON/CSV data.
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

# Import GRC Common library
$commonModulePath = Join-Path -Path $PSScriptRoot -ChildPath "../../../common/GRC-M365-Common.psm1"
if (Test-Path $commonModulePath) {
    Import-Module -Name $commonModulePath -Force
} else {
    Write-Error "Required GRC Common module not found at: $commonModulePath"
    return
}

# 1. Establish connection to Graph
try {
    if ($Interactive) {
        Connect-GRCEnvironment -Interactive
    } else {
        Connect-GRCEnvironment -TenantId $TenantId -ClientId $ClientId -CertificateBase64 $CertificateBase64
    }
} catch {
    Write-Error "Graph authentication failed: $_"
    return
}

# 2. Query Purview Settings using REST calls
$reportData = [Ordered]@{
    TotalSensitivityLabels   = 0
    SensitivityLabelNames    = ""
    SensitivityLabelsDetails = @()
    TotalDlpPolicies         = 0
    DlpPolicyNames           = ""
    DlpPoliciesDetails       = @()
}

try {
    # Query Sensitivity Labels (Try v1.0 security endpoint first, fallback to beta)
    $labelsResponse = $null
    $endpoints = @(
        "https://graph.microsoft.com/v1.0/security/dataSecurityAndGovernance/sensitivityLabels",
        "https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels"
    )
    foreach ($uri in $endpoints) {
        $labelsResponse = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
        if ($labelsResponse -and $labelsResponse.value) {
            break
        }
    }

    if ($labelsResponse -and $labelsResponse.value) {
        $reportData.TotalSensitivityLabels = @($labelsResponse.value).Count
        
        $labelNames = $labelsResponse.value | ForEach-Object {
            if ($_.name) { $_.name } else { $_.displayName }
        }
        $reportData.SensitivityLabelNames = ($labelNames | Where-Object { $_ }) -join '; '

        # Collect Sensitivity Label details
        $reportData.SensitivityLabelsDetails = $labelsResponse.value | ForEach-Object {
            $labelName = if ($_.name) { $_.name } else { $_.displayName }
            [Ordered]@{
                Id          = $_.id
                Name        = $labelName
                Description = $_.description
                IsActive    = $_.isActive
                Sensitivity = $_.sensitivity
            }
        }
    }

    # Query DLP Policies (using beta endpoints if v1.0 is limited)
    $dlpResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/informationProtection/dataLossPrevention/policies" -ErrorAction SilentlyContinue
    if ($dlpResponse -and $dlpResponse.value) {
        $reportData.TotalDlpPolicies = @($dlpResponse.value).Count
        $reportData.DlpPolicyNames = ($dlpResponse.value | Select-Object -ExpandProperty name) -join '; '
        # Collect DLP Policy details
        $reportData.DlpPoliciesDetails = $dlpResponse.value | ForEach-Object {
            [Ordered]@{
                Id          = $_.id
                Name        = $_.name
                Description = $_.description
                State       = $_.state
            }
        }
    }
} catch {
    Write-Warning "Could not query all Microsoft Purview endpoints: $_"
}

# 3. Handle Outputs based on execution scope
$exportObj = [PSCustomObject]$reportData
if ($AiAgentMode) {
    $exportObj | ConvertTo-Json -Depth 5
} else {
    Export-GRCAssetData -ServiceName "Purview" -AssetName "PurviewSummary" -Data @($exportObj)
}
