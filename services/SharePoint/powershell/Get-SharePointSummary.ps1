#Requires -Version 7.0
<#
.SYNOPSIS
    Queries SharePoint Online and OneDrive for site counts and sharing configuration policies.
.DESCRIPTION
    Retrieves key SharePoint Online governance settings via Graph REST API for GRC auditing.
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

# 2. Query SharePoint Settings and Site list using REST calls
$reportData = [Ordered]@{
    TotalSharepointSites    = 0
    ExternalSharingMode     = "Unknown"
    FileSharingCapability   = "Unknown"
    BlockMacAccess          = $false
}

try {
    # Query SharePoint Settings (v1.0 /admin/sharepoint/settings)
    $spSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/admin/sharepoint/settings" -ErrorAction SilentlyContinue
    if ($spSettings) {
        $reportData.ExternalSharingMode = if ($spSettings.sharingCapability) { $spSettings.sharingCapability } else { "Unknown" }
        $reportData.FileSharingCapability = if ($spSettings.sharingCapability) { $spSettings.sharingCapability } else { "Unknown" }
    }

    # Query total SharePoint Sites
    $sites = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites?`$select=id,name" -ErrorAction SilentlyContinue
    if ($sites -and $sites.value) {
        $reportData.TotalSharepointSites = @($sites.value).Count
    }

} catch {
    Write-Warning "Could not query all SharePoint Online settings: $_"
}

# 3. Handle Outputs based on execution scope
$exportObj = [PSCustomObject]$reportData
if ($AiAgentMode) {
    $exportObj | ConvertTo-Json -Depth 5
} else {
    Export-GRCAssetData -ServiceName "SharePoint" -AssetName "SharePointSummary" -Data @($exportObj)
}
