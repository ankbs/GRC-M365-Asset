#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Graph REST API for Entra ID (Azure AD) registered/joined devices.
.DESCRIPTION
    A cloud-native GRC Asset script designed to retrieve Entra ID device inventory details
    utilizing clean REST calls. Exports structured GRC results (JSON, CSV).

    Supports:
      - PowerShell (interactive admin)
      - Workflows (GitHub Actions runners)
      - AI Agents (silent JSON stdout/file integration)

.PARAMETER TenantId
    The target Entra ID Tenant ID.
.PARAMETER ClientId
    The App Registration client ID (unattended mode).
.PARAMETER CertificateBase64
    Base64 encoded certificate bytes for auth (unattended mode).
.PARAMETER Interactive
    Use interactive delegated login flow.
.PARAMETER AiAgentMode
    Silent output in JSON format directly to stdout.

.LINK
    https://learn.microsoft.com/en-us/graph/api/resources/device
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

# 2. Query Devices details using native Graph REST calls with Pagination
Write-Verbose "Querying Microsoft Graph REST endpoint /v1.0/devices..."
$devices = [System.Collections.Generic.List[PSCustomObject]]::new()
$deviceUri = "https://graph.microsoft.com/v1.0/devices?`$select=id,displayName,deviceId,operatingSystem,operatingSystemVersion,isManaged,trustType,approximateLastSignInDateTime,profileType"

try {
    while ($deviceUri) {
        $deviceResponse = Invoke-MgGraphRequest -Method GET -Uri $deviceUri -ErrorAction Stop
        if ($deviceResponse -and $deviceResponse.value) {
            foreach ($device in $deviceResponse.value) {
                $devObj = [PSCustomObject]@{
                    Id                           = $device.id
                    DisplayName                  = $device.displayName
                    DeviceId                     = $device.deviceId
                    OperatingSystem              = $device.operatingSystem
                    OperatingSystemVersion       = $device.operatingSystemVersion
                    IsManaged                    = if ($null -ne $device.isManaged) { $device.isManaged } else { $false }
                    TrustType                    = $device.trustType
                    ApproximateLastSignInDateTime = $device.approximateLastSignInDateTime
                    ProfileType                  = $device.profileType
                }
                $devices.Add($devObj)
            }
            $deviceUri = $deviceResponse.'@odata.nextLink'
        } else {
            $deviceUri = $null
        }
    }
} catch {
    Write-Error "Failed to query devices REST endpoint: $_"
    return
}

# 3. Handle Outputs based on execution scope
if ($AiAgentMode) {
    # AI Agents need silent stdout directly as JSON string for easy processing
    $devices.ToArray() | ConvertTo-Json -Depth 5
} else {
    # Local admin and workflow environments get file exports
    Export-GRCAssetData -ServiceName "Devices" -AssetName "EntraDevices" -Data ($devices.ToArray())
}
