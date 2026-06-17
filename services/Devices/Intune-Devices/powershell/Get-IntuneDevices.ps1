#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Graph REST API for Intune MDM-managed devices.
.DESCRIPTION
    A cloud-native GRC Asset script designed to retrieve Intune-managed device list details,
    compliance states, and hardware models utilizing clean REST calls. Exports structured GRC results (JSON, CSV).

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
    https://learn.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice
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

# 2. Query Managed Devices details using native Graph REST calls with Pagination
Write-Verbose "Querying Microsoft Graph REST endpoint /v1.0/deviceManagement/managedDevices..."
$managedDevices = [System.Collections.Generic.List[PSCustomObject]]::new()
$intuneUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=id,deviceName,operatingSystem,osVersion,complianceState,managementAgent,lastSyncDateTime,userPrincipalName,model,manufacturer,serialNumber"

try {
    while ($intuneUri) {
        $intuneResponse = Invoke-MgGraphRequest -Method GET -Uri $intuneUri -ErrorAction Stop
        if ($intuneResponse -and $intuneResponse.value) {
            foreach ($dev in $intuneResponse.value) {
                $devObj = [PSCustomObject]@{
                    Id                = $dev.id
                    DeviceName        = $dev.deviceName
                    OperatingSystem   = $dev.operatingSystem
                    OsVersion         = $dev.osVersion
                    ComplianceState   = $dev.complianceState
                    ManagementAgent   = $dev.managementAgent
                    LastSyncDateTime  = $dev.lastSyncDateTime
                    UserPrincipalName = $dev.userPrincipalName
                    Model             = $dev.model
                    Manufacturer      = $dev.manufacturer
                    SerialNumber      = $dev.serialNumber
                }
                $managedDevices.Add($devObj)
            }
            $intuneUri = $intuneResponse.'@odata.nextLink'
        } else {
            $intuneUri = $null
        }
    }
} catch {
    Write-Error "Failed to query Intune managedDevices REST endpoint: $_"
    return
}

# 3. Handle Outputs based on execution scope
if ($AiAgentMode) {
    # AI Agents need silent stdout directly as JSON string for easy processing
    $managedDevices.ToArray() | ConvertTo-Json -Depth 5
} else {
    # Local admin and workflow environments get file exports
    Export-GRCAssetData -ServiceName "Devices" -AssetName "IntuneDevices" -Data ($managedDevices.ToArray())
}
