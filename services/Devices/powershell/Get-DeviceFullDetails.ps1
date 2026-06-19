#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Graph REST API for deep details of all Devices.
.DESCRIPTION
    Retrieves and correlates devices from both Entra ID (hardware and registration state)
    and Intune MDM (compliance, serial numbers, last sync times). Exports results to JSON/CSV.
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

# Import our GRC Common module
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
    Write-Error "Authentication failed: $_"
    return
}

# 2. Query Entra ID Devices
Write-Verbose "Querying Entra ID Devices..."
$entraDevices = @{}
try {
    $devicesUri = "https://graph.microsoft.com/v1.0/devices?`$select=id,deviceId,displayName,operatingSystem,operatingSystemVersion,trustType,isCompliant,approximateLastSignInDateTime"
    while ($devicesUri) {
        $devResponse = Invoke-MgGraphRequest -Method GET -Uri $devicesUri -ErrorAction Stop
        if ($devResponse -and $devResponse.value) {
            foreach ($dev in $devResponse.value) {
                if ($dev.deviceId) {
                    $entraDevices[$dev.deviceId.ToString()] = $dev
                }
            }
            $devicesUri = $devResponse.'@odata.nextLink'
        } else {
            $devicesUri = $null
        }
    }
} catch {
    Write-Warning "Could not query Entra ID Devices: $_"
}

# 3. Query Intune Managed Devices
Write-Verbose "Querying Intune Managed Devices..."
$intuneDevices = @{}
try {
    $intuneUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=id,azureADDeviceId,deviceName,operatingSystem,osVersion,complianceState,enrollmentType,lastSyncDateTime,serialNumber,model,manufacturer,userPrincipalName,partnerThreatProtectionConnectionStatus,managementAgent"
    while ($intuneUri) {
        $intuneResponse = Invoke-MgGraphRequest -Method GET -Uri $intuneUri -ErrorAction Stop
        if ($intuneResponse -and $intuneResponse.value) {
            foreach ($dev in $intuneResponse.value) {
                if ($dev.azureADDeviceId) {
                    $intuneDevices[$dev.azureADDeviceId.ToString()] = $dev
                }
            }
            $intuneUri = $intuneResponse.'@odata.nextLink'
        } else {
            $intuneUri = $null
        }
    }
} catch {
    Write-Warning "Could not query Intune Devices: $_"
}

# 4. Correlate and Build Master Device List
$correlatedDevices = [System.Collections.Generic.List[PSCustomObject]]::new()
$allDeviceIds = @($entraDevices.Keys + $intuneDevices.Keys) | Select-Object -Unique

foreach ($dId in $allDeviceIds) {
    $eDev = if ($entraDevices.ContainsKey($dId)) { $entraDevices[$dId] } else { $null }
    $iDev = if ($intuneDevices.ContainsKey($dId)) { $intuneDevices[$dId] } else { $null }

    # Retrieve matching names/OS
    $displayName = if ($eDev) { $eDev.displayName } else { $iDev.deviceName }
    $os = if ($eDev) { $eDev.operatingSystem } else { $iDev.operatingSystem }
    $osVer = if ($eDev) { $eDev.operatingSystemVersion } else { $iDev.osVersion }
    
    # Trust type & active status
    $trustType = if ($eDev) { $eDev.trustType } else { "Unknown" }
    
    # Compliance check
    $isCompliant = $false
    if ($eDev -and $eDev.isCompliant) { $isCompliant = $true }
    if ($iDev -and $iDev.complianceState -eq "compliant") { $isCompliant = $true }

    # Defender Status from threat protection status / sense agent
    $defenderStatus = "Not Enrolled"
    if ($iDev) {
        $tpStatus = $iDev.partnerThreatProtectionConnectionStatus
        $agent = $iDev.managementAgent
        if ($tpStatus -eq 'activated' -or $agent -match 'microsoftSense') {
            $defenderStatus = "Active"
        } elseif ($tpStatus) {
            $defenderStatus = $tpStatus.ToString()
        }
    }

    $devObj = [PSCustomObject]@{
        DeviceId                     = $dId
        DisplayName                  = $displayName
        OperatingSystem              = $os
        OperatingSystemVersion       = $osVer
        TrustType                    = $trustType
        IsCompliant                  = $isCompliant
        IntuneDeviceId               = if ($iDev) { $iDev.id } else { "" }
        IntuneManaged                = if ($iDev) { $true } else { $false }
        EnrollmentType               = if ($iDev) { $iDev.enrollmentType } else { "Unknown" }
        LastSyncDateTime             = if ($iDev) { $iDev.lastSyncDateTime } else { if ($eDev) { $eDev.approximateLastSignInDateTime } else { "" } }
        SerialNumber                 = if ($iDev) { $iDev.serialNumber } else { "" }
        Model                        = if ($iDev) { $iDev.model } else { "" }
        Manufacturer                 = if ($iDev) { $iDev.manufacturer } else { "" }
        UserPrincipalName            = if ($iDev) { $iDev.userPrincipalName } else { "" }
        DefenderStatus               = $defenderStatus
    }
    $correlatedDevices.Add($devObj)
}

# 5. Handle Outputs
if ($AiAgentMode) {
    $correlatedDevices.ToArray() | ConvertTo-Json -Depth 5
} else {
    $exportDir = Join-Path -Path $PSScriptRoot -ChildPath "../../../exports/Devices/DeviceFullDetails"
    if (!(Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $jsonPath = Join-Path -Path $exportDir -ChildPath "DeviceFullDetails_${timestamp}.json"
    $csvPath  = Join-Path -Path $exportDir -ChildPath "DeviceFullDetails_${timestamp}.csv"

    $correlatedDevices.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding utf8
    $correlatedDevices.ToArray() | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

    Write-Host "Devices Full Details JSON written to: $jsonPath" -ForegroundColor Green
    Write-Host "Devices Full Details CSV written to: $csvPath" -ForegroundColor Green
}
