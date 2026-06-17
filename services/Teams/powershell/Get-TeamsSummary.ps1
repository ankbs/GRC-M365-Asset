#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Teams for team counts and GRC visibility settings.
.DESCRIPTION
    Retrieves key Teams governance metrics (total teams, public vs private counts) via Graph REST API.
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

# 2. Query Microsoft Teams metrics using REST calls
$reportData = [Ordered]@{
    TotalTeams       = 0
    PublicTeamsCount = 0
    PrivateTeamsCount= 0
}

try {
    # Query all Groups that are Teams
    $teamsUri = "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/any(x:x eq 'Team')&`$select=id,displayName,visibility"
    $teamsList = Invoke-MgGraphRequest -Method GET -Uri $teamsUri -ErrorAction Stop
    
    if ($teamsList -and $teamsList.value) {
        $reportData.TotalTeams = @($teamsList.value).Count
        
        $publicTeams = $teamsList.value | Where-Object { $_.visibility -eq 'Public' }
        $reportData.PublicTeamsCount = if ($publicTeams) { @($publicTeams).Count } else { 0 }
        
        $privateTeams = $teamsList.value | Where-Object { $_.visibility -eq 'Private' }
        $reportData.PrivateTeamsCount = if ($privateTeams) { @($privateTeams).Count } else { 0 }
    }
} catch {
    Write-Warning "Could not query Microsoft Teams endpoints: $_"
}

# 3. Handle Outputs based on execution scope
$exportObj = [PSCustomObject]$reportData
if ($AiAgentMode) {
    $exportObj | ConvertTo-Json -Depth 5
} else {
    Export-GRCAssetData -ServiceName "Teams" -AssetName "TeamsSummary" -Data @($exportObj)
}
