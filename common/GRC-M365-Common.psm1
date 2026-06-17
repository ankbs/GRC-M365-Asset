#Requires -Version 7.0
<#
.SYNOPSIS
    Core authentication, API query, and data export library for GRC-M365-Asset.
.DESCRIPTION
    Provides unified helper functions for connecting to Microsoft Graph and Exchange Online,
    querying endpoints safely, and exporting structured datasets (JSON, CSV, HTML reports).
#>

function Connect-GRCEnvironment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$TenantId,

        [Parameter(Mandatory = $false)]
        [string]$ClientId,

        [Parameter(Mandatory = $false)]
        [string]$CertificateBase64,

        [Parameter(Mandatory = $false)]
        [switch]$Interactive
    )

    # 1. Interactive OAuth login for local Admins
    if ($Interactive) {
        Write-Verbose "Initiating interactive user authentication..."
        Connect-MgGraph -Scopes "Directory.Read.All", "DeviceManagementManagedDevices.Read.All", "Reports.Read.All"
        return
    }

    # 2. Non-interactive certificate authentication for Cloud / Workflows
    if ($CertificateBase64 -and $ClientId -and $TenantId) {
        Write-Verbose "Initiating unattended certificate authentication..."
        $certBytes = [System.Convert]::FromBase64String($CertificateBase64)
        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certBytes)
        
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Certificate $cert
        return
    }

    Write-Error "Invalid authentication parameters. Specify either -Interactive or -CertificateBase64, -ClientId, and -TenantId."
}

function Export-GRCAssetData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ServiceName,

        [Parameter(Mandatory)]
        [string]$AssetName,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [PSCustomObject[]]$Data
    )

    $exportDir = Join-Path -Path $PSScriptRoot -ChildPath "../exports/$ServiceName/$AssetName"
    if (!(Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $jsonPath = Join-Path -Path $exportDir -ChildPath "${AssetName}_${timestamp}.json"
    $csvPath  = Join-Path -Path $exportDir -ChildPath "${AssetName}_${timestamp}.csv"

    # Export structured JSON & CSV
    $Data | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding utf8
    $Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

    Write-Output "GRC Asset exported successfully to: $jsonPath and $csvPath"
}

Export-ModuleMember -Function Connect-GRCEnvironment, Export-GRCAssetData
