#Requires -Version 7.2

<#
.SYNOPSIS
    Reads Microsoft Purview retention labels through Security & Compliance PowerShell.

.DESCRIPTION
    This script is designed for Linux-based GitHub Actions runners.
    It uses the existing repository secrets created by Initialize-GRCEnvironment.ps1:

    - GRC_CLIENT_ID
    - GRC_TENANT_ID
    - GRC_CERTIFICATE
    - Optional: GRC_ORGANIZATION
    - Optional: GRC_CERTIFICATE_PASSWORD

    The script does not call Microsoft Graph retentionLabels because the Graph endpoint
    does not support app-only permissions. Instead, it connects to Security & Compliance
    PowerShell with Connect-IPPSSession and reads labels with Get-ComplianceTag.

.PARAMETER OutputDirectory
    Directory where the JSON and CSV output files are written.

.PARAMETER IncludeRawObject
    Adds a compressed JSON representation of the raw Get-ComplianceTag object per label.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $OutputDirectory = './output',

    [Parameter()]
    [switch] $IncludeRawObject
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-GrcRequiredValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string] $Value,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Name
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "Required environment variable '$Name' is missing or empty."
    }
}

function Get-GrcOrganizationName {
    [CmdletBinding()]
    param()

    if (-not [string]::IsNullOrWhiteSpace($env:GRC_ORGANIZATION)) {
        return $env:GRC_ORGANIZATION
    }

    if (-not [string]::IsNullOrWhiteSpace($env:GRC_TENANT_ID) -and $env:GRC_TENANT_ID -match '\.onmicrosoft\.com$') {
        return $env:GRC_TENANT_ID
    }

    throw @'
No usable organization domain was found.

Connect-IPPSSession requires the primary *.onmicrosoft.com tenant domain in -Organization.
Set one of these GitHub repository secrets:
- GRC_ORGANIZATION = yourtenant.onmicrosoft.com
- or keep GRC_TENANT_ID as yourtenant.onmicrosoft.com if your setup script already stores the domain there.
'@
}

function ConvertFrom-GrcBase64Pfx {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Base64Pfx,

        [Parameter()]
        [AllowEmptyString()]
        [string] $Password = ''
    )

    $certificateBytes = [Convert]::FromBase64String($Base64Pfx)

    $keyStorageFlags =
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet -bor
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable

    $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $certificateBytes,
        $Password,
        $keyStorageFlags
    )

    if (-not $certificate.HasPrivateKey) {
        throw 'The certificate loaded from GRC_CERTIFICATE does not contain a private key. Use a base64 encoded PFX file, not a CER file.'
    }

    return $certificate
}

function Connect-GrcComplianceSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $ClientId,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Organization,

        [Parameter(Mandatory)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate
    )

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $module = Get-Module ExchangeOnlineManagement -ListAvailable |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $module) {
        throw 'ExchangeOnlineManagement module is not available.'
    }

    Write-Host "ExchangeOnlineManagement version: $($module.Version)"
    Write-Host "Connecting to Security & Compliance PowerShell for organization '$Organization'."

    if (-not $IsWindows) {
        $Global:IsWindows = $true
    }

    $logDirectory = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'grc-ipps-logs'

    if (-not (Test-Path -Path $logDirectory)) {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }

    Connect-IPPSSession `
        -AppId $ClientId `
        -Certificate $Certificate `
        -Organization $Organization `
        -CommandName Get-ComplianceTag `
        -ShowBanner:$false `
        -EnableErrorReporting `
        -LogDirectoryPath $logDirectory
}

function Get-GrcObjectPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object] $InputObject,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Name
    )

    $property = $InputObject.PSObject.Properties[$Name]

    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Get-GrcRetentionLabelSummary {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch] $IncludeRawObject
    )

    $labels = @(Get-ComplianceTag -IncludingLabelState)

    foreach ($label in $labels) {
        $summary = [ordered]@{
            Name               = Get-GrcObjectPropertyValue -InputObject $label -Name 'Name'
            Guid               = Get-GrcObjectPropertyValue -InputObject $label -Name 'Guid'
            ImmutableId        = Get-GrcObjectPropertyValue -InputObject $label -Name 'ImmutableId'
            Priority           = Get-GrcObjectPropertyValue -InputObject $label -Name 'Priority'
            Workload           = Get-GrcObjectPropertyValue -InputObject $label -Name 'Workload'
            RetentionAction    = Get-GrcObjectPropertyValue -InputObject $label -Name 'RetentionAction'
            RetentionDuration  = Get-GrcObjectPropertyValue -InputObject $label -Name 'RetentionDuration'
            RetentionType      = Get-GrcObjectPropertyValue -InputObject $label -Name 'RetentionType'
            IsRecordLabel      = Get-GrcObjectPropertyValue -InputObject $label -Name 'IsRecordLabel'
            Disabled           = Get-GrcObjectPropertyValue -InputObject $label -Name 'Disabled'
            Comment            = Get-GrcObjectPropertyValue -InputObject $label -Name 'Comment'
            CreatedBy          = Get-GrcObjectPropertyValue -InputObject $label -Name 'CreatedBy'
            WhenCreatedUtc     = Get-GrcObjectPropertyValue -InputObject $label -Name 'WhenCreatedUTC'
            LastModifiedBy     = Get-GrcObjectPropertyValue -InputObject $label -Name 'LastModifiedBy'
            WhenChangedUtc     = Get-GrcObjectPropertyValue -InputObject $label -Name 'WhenChangedUTC'
        }

        if ($IncludeRawObject.IsPresent) {
            $summary.RawObjectJson = $label | ConvertTo-Json -Depth 30 -Compress
        }

        [pscustomobject] $summary
    }
}

function Save-GrcOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]] $Labels,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Directory,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Organization
    )

    if (-not (Test-Path -Path $Directory)) {
        New-Item -Path $Directory -ItemType Directory -Force | Out-Null
    }

    $timestamp = (Get-Date).ToUniversalTime().ToString('yyyyMMdd-HHmmss')
    $jsonPath = Join-Path -Path $Directory -ChildPath "purview-retention-labels-$timestamp.json"
    $csvPath = Join-Path -Path $Directory -ChildPath "purview-retention-labels-$timestamp.csv"

    $payload = [pscustomobject]@{
        ExportedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        Organization  = $Organization
        LabelCount    = $Labels.Count
        Labels        = $Labels
    }

    $payload |
        ConvertTo-Json -Depth 30 |
        Out-File -FilePath $jsonPath -Encoding utf8

    $Labels |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

    [pscustomobject]@{
        JsonPath   = $jsonPath
        CsvPath    = $csvPath
        LabelCount = $Labels.Count
    }
}

try {
    Test-GrcRequiredValue -Value $env:GRC_CLIENT_ID -Name 'GRC_CLIENT_ID'
    Test-GrcRequiredValue -Value $env:GRC_CERTIFICATE -Name 'GRC_CERTIFICATE'

    $organization = Get-GrcOrganizationName

    $certificatePassword = ''
    if (-not [string]::IsNullOrEmpty($env:GRC_CERTIFICATE_PASSWORD)) {
        $certificatePassword = $env:GRC_CERTIFICATE_PASSWORD
    }

    Write-Host 'Loading PFX certificate from GRC_CERTIFICATE secret.'
    $certificate = ConvertFrom-GrcBase64Pfx `
        -Base64Pfx $env:GRC_CERTIFICATE `
        -Password $certificatePassword

    Connect-GrcComplianceSession `
        -ClientId $env:GRC_CLIENT_ID `
        -Organization $organization `
        -Certificate $certificate

    Write-Host 'Reading Purview retention labels with Get-ComplianceTag.'
    $labels = @(Get-GrcRetentionLabelSummary -IncludeRawObject:$IncludeRawObject)

    $result = Save-GrcOutput `
        -Labels $labels `
        -Directory $OutputDirectory `
        -Organization $organization

    Write-Host "Label count: $($result.LabelCount)"
    Write-Host "JSON output: $($result.JsonPath)"
    Write-Host "CSV output : $($result.CsvPath)"
}
catch {
    Write-Warning 'Purview label-only export failed.'
    Write-Warning "Exception type: $($_.Exception.GetType().FullName)"
    Write-Warning "Exception message: $($_.Exception.Message)"

    if ($_.ScriptStackTrace) {
        Write-Warning 'Script stack trace:'
        Write-Warning $_.ScriptStackTrace
    }

    throw
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
