#Requires -Version 5.1

<#
.SYNOPSIS
    Reads Microsoft Purview retention labels through Security & Compliance PowerShell on a Windows GitHub runner.

.DESCRIPTION
    This script is intended for GitHub Actions with runs-on: windows-latest and shell: powershell.
    It imports the base64 encoded PFX from the GRC_CERTIFICATE secret into the CurrentUser certificate
    store, connects to Security & Compliance PowerShell with Connect-IPPSSession and the certificate
    thumbprint, reads retention labels with Get-ComplianceTag, and writes JSON/CSV plus IPPS logs to
    the output directory.

    This avoids the Linux runner failure seen inside ExchangeOnlineManagement/Connect-IPPSSession.

.REQUIRED GITHUB SECRETS
    GRC_CLIENT_ID
    GRC_CERTIFICATE
    GRC_TENANT_ID or GRC_ORGANIZATION

.OPTIONAL GITHUB SECRET
    GRC_CERTIFICATE_PASSWORD

.PARAMETER OutputDirectory
    Directory where the JSON, CSV, and IPPS logs are written.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $OutputDirectory = '.\output'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ImportedCertificateThumbprint = $null
$script:TemporaryCertificatePath = $null

function Test-GrcRequiredValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string] $Value,

        [Parameter(Mandatory = $true)]
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

function ConvertTo-GrcSecureString {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowEmptyString()]
        [string] $PlainText = ''
    )

    if ([string]::IsNullOrEmpty($PlainText)) {
        return New-Object System.Security.SecureString
    }

    return ConvertTo-SecureString -String $PlainText -AsPlainText -Force
}

function New-GrcTemporaryPfxFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $Base64Pfx
    )

    $certificateBytes = [Convert]::FromBase64String($Base64Pfx)
    $certificateDirectory = Join-Path -Path $env:RUNNER_TEMP -ChildPath 'grc-certificates'

    if (-not (Test-Path -Path $certificateDirectory)) {
        New-Item -Path $certificateDirectory -ItemType Directory -Force | Out-Null
    }

    $certificatePath = Join-Path -Path $certificateDirectory -ChildPath "grc-purview-$([guid]::NewGuid()).pfx"
    [System.IO.File]::WriteAllBytes($certificatePath, $certificateBytes)

    return $certificatePath
}

function Import-GrcPfxCertificate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $CertificateFilePath,

        [Parameter(Mandatory = $true)]
        [securestring] $CertificatePassword
    )

    $certificate = Import-PfxCertificate `
        -FilePath $CertificateFilePath `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -Password $CertificatePassword `
        -Exportable

    if ($null -eq $certificate) {
        throw 'The PFX certificate could not be imported into Cert:\CurrentUser\My.'
    }

    if (-not $certificate.HasPrivateKey) {
        throw 'The imported certificate does not contain a private key. Use a PFX file, not a CER file.'
    }

    return $certificate
}

function Get-GrcCertificateDiagnostics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate
    )

    [pscustomobject]@{
        Subject            = $Certificate.Subject
        Issuer             = $Certificate.Issuer
        Thumbprint         = $Certificate.Thumbprint
        NotBeforeUtc       = $Certificate.NotBefore.ToUniversalTime().ToString('o')
        NotAfterUtc        = $Certificate.NotAfter.ToUniversalTime().ToString('o')
        HasPrivateKey      = $Certificate.HasPrivateKey
        PublicKeyAlgorithm = $Certificate.PublicKey.Oid.FriendlyName
        SignatureAlgorithm = $Certificate.SignatureAlgorithm.FriendlyName
    }
}

function Connect-GrcComplianceSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $ClientId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $Organization,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $CertificateThumbprint,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $LogDirectory
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
    Write-Host "Writing IPPS logs to '$LogDirectory'."

    if (-not (Test-Path -Path $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }

    Connect-IPPSSession `
        -AppId $ClientId `
        -CertificateThumbprint $CertificateThumbprint `
        -Organization $Organization `
        -CommandName Get-ComplianceTag `
        -ShowBanner:$false `
        -EnableErrorReporting `
        -LogDirectoryPath $LogDirectory
}

function Get-GrcObjectPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object] $InputObject,

        [Parameter(Mandatory = $true)]
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
    param()

    $labels = @(Get-ComplianceTag -IncludingLabelState)

    foreach ($label in $labels) {
        [pscustomobject] [ordered]@{
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
    }
}

function Save-GrcOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object[]] $Labels,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $Directory,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $Organization
    )

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

function Remove-GrcImportedCertificate {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowEmptyString()]
        [string] $Thumbprint
    )

    if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
        return
    }

    $certificatePath = "Cert:\CurrentUser\My\$Thumbprint"

    if (Test-Path -Path $certificatePath) {
        Remove-Item -Path $certificatePath -Force -ErrorAction SilentlyContinue
    }
}

try {
    Test-GrcRequiredValue -Value $env:GRC_CLIENT_ID -Name 'GRC_CLIENT_ID'
    Test-GrcRequiredValue -Value $env:GRC_CERTIFICATE -Name 'GRC_CERTIFICATE'

    if (-not (Test-Path -Path $OutputDirectory)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
    }

    $logDirectory = Join-Path -Path $OutputDirectory -ChildPath 'ipps-logs'

    if (-not (Test-Path -Path $logDirectory)) {
        New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null
    }

    $organization = Get-GrcOrganizationName

    $certificatePasswordValue = ''
    if (-not [string]::IsNullOrEmpty($env:GRC_CERTIFICATE_PASSWORD)) {
        $certificatePasswordValue = $env:GRC_CERTIFICATE_PASSWORD
    }

    Write-Host 'Writing temporary PFX certificate file from GRC_CERTIFICATE secret.'
    $script:TemporaryCertificatePath = New-GrcTemporaryPfxFile -Base64Pfx $env:GRC_CERTIFICATE

    Write-Host 'Importing PFX certificate into Cert:\CurrentUser\My.'
    $certificatePassword = ConvertTo-GrcSecureString -PlainText $certificatePasswordValue
    $certificate = Import-GrcPfxCertificate `
        -CertificateFilePath $script:TemporaryCertificatePath `
        -CertificatePassword $certificatePassword

    $script:ImportedCertificateThumbprint = $certificate.Thumbprint

    $certificateDiagnostics = Get-GrcCertificateDiagnostics -Certificate $certificate
    $certificateDiagnostics |
        ConvertTo-Json -Depth 10 |
        Out-File -FilePath (Join-Path -Path $OutputDirectory -ChildPath 'certificate-diagnostics.json') -Encoding utf8

    Write-Host "Certificate subject: $($certificateDiagnostics.Subject)"
    Write-Host "Certificate thumbprint: $($certificateDiagnostics.Thumbprint)"
    Write-Host "Certificate has private key: $($certificateDiagnostics.HasPrivateKey)"

    Connect-GrcComplianceSession `
        -ClientId $env:GRC_CLIENT_ID `
        -Organization $organization `
        -CertificateThumbprint $certificate.Thumbprint `
        -LogDirectory $logDirectory

    Write-Host 'Reading Purview retention labels with Get-ComplianceTag.'
    $labels = @(Get-GrcRetentionLabelSummary)

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

    Remove-GrcImportedCertificate -Thumbprint $script:ImportedCertificateThumbprint

    if (-not [string]::IsNullOrWhiteSpace($script:TemporaryCertificatePath) -and (Test-Path -Path $script:TemporaryCertificatePath)) {
        Remove-Item -Path $script:TemporaryCertificatePath -Force -ErrorAction SilentlyContinue
    }
}
