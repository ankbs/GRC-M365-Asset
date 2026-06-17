#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Graph REST API for Entra ID Tenant Details.
.DESCRIPTION
    A cloud-native GRC Asset script designed to retrieve Entra ID tenant information
    utilizing clean REST calls and exporting structured GRC results (JSON, CSV, HTML-ready).
    
    Includes execution optimizations for:
      - PowerShell (interactive admin)
      - Workflows (GitHub Actions runners)
      - AI Agents (silent JSON stdout/file integration)
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

# 2. Query Tenant Information using native Graph REST calls
Write-Verbose "Querying Microsoft Graph REST endpoint /v1.0/organization..."
$orgData = $null
$domainData = $null
$securityDefaults = $null

try {
    $orgData = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop
    $domainData = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains" -ErrorAction Stop
} catch {
    Write-Error "Failed to query core tenant REST endpoints: $_"
    return
}

try {
    $securityDefaults = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy" -ErrorAction SilentlyContinue
} catch {
    Write-Verbose "Could not retrieve Security Defaults Policy settings (possibly disabled or missing permission)."
}

# 3. Process data into structured PSCustomObject
$verifiedDomains = $domainData.value | Where-Object { $_.isVerified -eq $true } | ForEach-Object { $_.id }
$verifiedDomainsJoined = ($verifiedDomains | Sort-Object) -join '; '
$defaultDomain = ($domainData.value | Where-Object { $_.isDefault -eq $true } | Select-Object -First 1).id

$assetReport = [PSCustomObject]@{
    OrgDisplayName          = $orgData.value[0].displayName
    TenantId                = $orgData.value[0].id
    VerifiedDomains         = $verifiedDomainsJoined
    DefaultDomain           = $defaultDomain
    SecurityDefaultsEnabled = if ($null -ne $securityDefaults) { $securityDefaults.isEnabled } else { 'N/A' }
    CreatedDateTime         = $orgData.value[0].createdDateTime
    OnPremisesSyncEnabled   = $orgData.value[0].onPremisesSyncEnabled
}

# 4. Handle Outputs based on execution scope
if ($AiAgentMode) {
    # AI Agents need silent stdout directly as JSON string for easy processing
    $assetReport | ConvertTo-Json -Depth 5
} else {
    # Local admin and workflow environments get file exports
    Export-GRCAssetData -ServiceName "EntraID" -AssetName "TenantInfo" -Data @($assetReport)
}
