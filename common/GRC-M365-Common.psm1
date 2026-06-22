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
        $keyStorageFlags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certBytes, "", $keyStorageFlags)
        
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Certificate $cert
        return
    }

    Write-Error "Invalid authentication parameters. Specify either -Interactive or -CertificateBase64, -ClientId, and -TenantId."
}

function Flatten-GRCValue {
    param(
        $InputObject
    )
    if ($null -eq $InputObject) { return $null }
    
    $objToParse = $InputObject
    if ($InputObject -is [System.Collections.IDictionary]) {
        $objToParse = [PSCustomObject]$InputObject
    }
    
    $flatObj = [ordered]@{}
    foreach ($prop in $objToParse.PSObject.Properties) {
        $name = $prop.Name
        # Skip OData metadata properties
        if ($name -like "@odata*") { continue }
        $val = $prop.Value
        if ($null -eq $val) {
            $flatObj[$name] = ""
        } elseif ($val -is [string] -or $val -is [valueType]) {
            $flatObj[$name] = $val
        } elseif ($val -is [System.Collections.IEnumerable] -and $val -isnot [System.Collections.IDictionary]) {
            $flatObj[$name] = ($val | ForEach-Object { if ($null -ne $_) { $_.ToString() } else { "" } }) -join "; "
        } else {
            # Serialize nested objects to JSON string to prevent Export-Csv from crashing
            $flatObj[$name] = $val | ConvertTo-Json -Compress -Depth 2
        }
    }
    return [PSCustomObject]$flatObj
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
        [object[]]$Data
    )

    $exportDir = Join-Path -Path $PSScriptRoot -ChildPath "../exports/$ServiceName/$AssetName"
    if (!(Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $jsonPath = Join-Path -Path $exportDir -ChildPath "${AssetName}_${timestamp}.json"
    $csvPath  = Join-Path -Path $exportDir -ChildPath "${AssetName}_${timestamp}.csv"

    # Export structured JSON (retains rich objects and dictionary structures)
    $Data | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding utf8

    # Flat array of objects for Export-Csv compatibility
    $flatData = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($item in $Data) {
        if ($null -ne $item) {
            $flatData.Add((Flatten-GRCValue -InputObject $item))
        }
    }

    $flatData.ToArray() | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

    Write-Output "GRC Asset exported successfully to: $jsonPath and $csvPath"
}

function Connect-GRCExchange {
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

    # Import ExchangeOnlineManagement module if not loaded
    if (-not (Get-Module -Name ExchangeOnlineManagement)) {
        Import-Module -Name ExchangeOnlineManagement -Force
    }

    # 1. Interactive login for local Admins
    if ($Interactive) {
        Write-Verbose "Initiating interactive Exchange Online authentication..."
        Connect-ExchangeOnline -ShowBanner:$false
        return
    }

    # 2. Non-interactive certificate authentication for Cloud / Workflows
    if ($CertificateBase64 -and $ClientId -and $TenantId) {
        Write-Verbose "Initiating unattended Exchange Online certificate authentication..."
        $certBytes = [System.Convert]::FromBase64String($CertificateBase64)
        $keyStorageFlags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certBytes, "", $keyStorageFlags)
        
        # Ensure any existing Exchange sessions are cleaned up first to avoid session conflicts
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        
        Connect-ExchangeOnline -Certificate $cert -AppId $ClientId -Organization $TenantId -ShowBanner:$false
        return
    }

    Write-Error "Invalid authentication parameters for Exchange. Specify either -Interactive or -CertificateBase64, -ClientId, and -TenantId."
}

function Connect-GRCCompliance {
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

    if ($Interactive) {
        Write-Verbose "Initiating interactive Security & Compliance Center authentication..."
        Connect-IPPSSession -ShowBanner:$false
        return
    }

    if ($CertificateBase64 -and $ClientId -and $TenantId) {
        Write-Verbose "Initiating unattended Security & Compliance Center certificate authentication..."
        $certBytes = [System.Convert]::FromBase64String($CertificateBase64)
        $keyStorageFlags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($certBytes, "", $keyStorageFlags)
        
        # Connect-IPPSSession requires the primary domain name (e.g. *.onmicrosoft.com) for -Organization.
        # If a Tenant ID GUID is passed, it throws a NullReferenceException ("Object reference not set to an instance of an object").
        $orgDomain = $TenantId
        if ($TenantId -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
            try {
                Write-Verbose "Resolving primary organization domain from Graph for IPPS connection..."
                $orgResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction SilentlyContinue
                if ($orgResponse -and $orgResponse.value) {
                    $defaultDomain = $orgResponse.value[0].verifiedDomains | Where-Object { $_.isDefault } | Select-Object -ExpandProperty name
                    if ($defaultDomain) {
                        $orgDomain = $defaultDomain
                        Write-Verbose "Resolved organization domain: $orgDomain"
                    }
                }
            } catch {
                Write-Warning "Failed to resolve organization domain name: $_"
            }
        }
        
        # Ensure any existing Exchange/Compliance sessions are cleaned up first to avoid session conflicts
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        
        Connect-IPPSSession -Certificate $cert -AppId $ClientId -Organization $orgDomain -CommandName Get-ComplianceTag, Get-DlpCompliancePolicy -ShowBanner:$false
        return
    }

    Write-Error "Invalid authentication parameters for IPPS. Specify either -Interactive or -CertificateBase64, -ClientId, and -TenantId."
}

Export-ModuleMember -Function Connect-GRCEnvironment, Connect-GRCExchange, Connect-GRCCompliance, Export-GRCAssetData
