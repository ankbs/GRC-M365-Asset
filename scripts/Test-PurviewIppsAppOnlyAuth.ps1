#Requires -Version 7.2

<#
.SYNOPSIS
    Diagnoses Microsoft 365 app-only authentication for Exchange Online and Security & Compliance PowerShell.

.DESCRIPTION
    This script is intended for GitHub Actions on windows-latest with shell: pwsh.
    It imports a base64 encoded PFX from GRC_CERTIFICATE into Cert:\CurrentUser\My,
    verifies certificate diagnostics, tries Microsoft Graph app-only connection,
    tries Connect-ExchangeOnline, then tries Connect-IPPSSession.

    The goal is to separate:
    - certificate import problems
    - Graph/app registration/certificate trust problems
    - Exchange Online app-only authentication problems
    - Security & Compliance PowerShell/IPPS authentication problems

.PARAMETER OutputDirectory
    Directory where diagnostic JSON, logs, and test results are written.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $OutputDirectory = './output'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ImportedCertificateThumbprint = $null
$script:TemporaryCertificatePath = $null

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

    throw 'Set GRC_ORGANIZATION to the primary tenant domain, for example mycloudofficedev.onmicrosoft.com.'
}

function ConvertTo-GrcSecureString {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowEmptyString()]
        [string] $PlainText = ''
    )

    if ([string]::IsNullOrEmpty($PlainText)) {
        return [System.Security.SecureString]::new()
    }

    return ConvertTo-SecureString -String $PlainText -AsPlainText -Force
}

function New-GrcTemporaryPfxFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Base64Pfx
    )

    $certificateBytes = [Convert]::FromBase64String($Base64Pfx)
    $certificateDirectory = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'grc-certificates'

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
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $CertificateFilePath,

        [Parameter(Mandatory)]
        [securestring] $CertificatePassword
    )

    $certificate = Import-PfxCertificate `
        -FilePath $CertificateFilePath `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -Password $CertificatePassword `
        -Exportable

    if ($null -eq $certificate) {
        throw 'PFX import returned no certificate.'
    }

    if (-not $certificate.HasPrivateKey) {
        throw 'The imported certificate does not contain a private key.'
    }

    return $certificate
}

function Get-GrcCertificateDiagnostics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate
    )

    $rsaKeyType = $null
    try {
        $rsaKey = $Certificate.GetRSAPrivateKey()
        if ($null -ne $rsaKey) {
            $rsaKeyType = $rsaKey.GetType().FullName
        }
    }
    catch {
        $rsaKeyType = "Unavailable: $($_.Exception.Message)"
    }

    [pscustomobject]@{
        Subject               = $Certificate.Subject
        Issuer                = $Certificate.Issuer
        Thumbprint            = $Certificate.Thumbprint
        NotBeforeUtc          = $Certificate.NotBefore.ToUniversalTime().ToString('o')
        NotAfterUtc           = $Certificate.NotAfter.ToUniversalTime().ToString('o')
        HasPrivateKey         = $Certificate.HasPrivateKey
        PublicKeyAlgorithm    = $Certificate.PublicKey.Oid.FriendlyName
        SignatureAlgorithm    = $Certificate.SignatureAlgorithm.FriendlyName
        PrivateKeyRuntimeType = $rsaKeyType
    }
}

function Invoke-GrcStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Name,

        [Parameter(Mandatory)]
        [scriptblock] $ScriptBlock
    )

    $startedAt = Get-Date

    try {
        Write-Host "START: $Name"
        $result = & $ScriptBlock
        $finishedAt = Get-Date

        [pscustomobject]@{
            Name          = $Name
            Success       = $true
            StartedAtUtc  = $startedAt.ToUniversalTime().ToString('o')
            FinishedAtUtc = $finishedAt.ToUniversalTime().ToString('o')
            ErrorType     = $null
            ErrorMessage  = $null
            ScriptTrace   = $null
            Result        = $result
        }
    }
    catch {
        $finishedAt = Get-Date
        Write-Warning "FAILED: $Name"
        Write-Warning "Exception type: $($_.Exception.GetType().FullName)"
        Write-Warning "Exception message: $($_.Exception.Message)"

        [pscustomobject]@{
            Name          = $Name
            Success       = $false
            StartedAtUtc  = $startedAt.ToUniversalTime().ToString('o')
            FinishedAtUtc = $finishedAt.ToUniversalTime().ToString('o')
            ErrorType     = $_.Exception.GetType().FullName
            ErrorMessage  = $_.Exception.Message
            ScriptTrace   = $_.ScriptStackTrace
            Result        = $null
        }
    }
}

function Connect-GrcGraphAppOnly {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Tenant,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $ClientId,

        [Parameter(Mandatory)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate
    )

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    Connect-MgGraph `
        -TenantId $Tenant `
        -ClientId $ClientId `
        -Certificate $Certificate `
        -NoWelcome

    $context = Get-MgContext

    [pscustomobject]@{
        TenantId = $context.TenantId
        ClientId = $context.ClientId
        AuthType = $context.AuthType
        Scopes   = $context.Scopes -join ', '
    }
}

function Get-GrcAppRegistrationDiagnostics {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $ClientId
    )

    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

    $servicePrincipal = Get-MgServicePrincipal -Filter "appId eq '$ClientId'"

    if ($null -eq $servicePrincipal) {
        throw "Service principal for AppId '$ClientId' was not found."
    }

    $assignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipal.Id -All)

    [pscustomobject]@{
        ServicePrincipalId = $servicePrincipal.Id
        AppId              = $servicePrincipal.AppId
        DisplayName        = $servicePrincipal.DisplayName
        AppRoleAssignments = $assignments | Select-Object ResourceDisplayName, AppRoleId, PrincipalDisplayName
    }
}

function Connect-GrcExchangeOnline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $ClientId,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Organization,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $CertificateThumbprint,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $LogDirectory
    )

    Import-Module ExchangeOnlineManagement -RequiredVersion 3.9.0 -ErrorAction Stop

    Connect-ExchangeOnline `
        -AppId $ClientId `
        -CertificateThumbprint $CertificateThumbprint `
        -Organization $Organization `
        -ShowBanner:$false `
        -SkipLoadingFormatData `
        -DisableWAM `
        -EnableErrorReporting `
        -LogDirectoryPath $LogDirectory

    Get-ConnectionInformation |
        Select-Object State, ConnectionUri, TokenStatus, ModuleName, ModuleVersion
}

function Connect-GrcIppsSession {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $ClientId,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Organization,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $CertificateThumbprint,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $LogDirectory
    )

    Import-Module ExchangeOnlineManagement -RequiredVersion 3.9.0 -ErrorAction Stop

    Connect-IPPSSession `
        -AppId $ClientId `
        -CertificateThumbprint $CertificateThumbprint `
        -Organization $Organization `
        -CommandName Get-ComplianceTag `
        -ShowBanner:$false `
        -DisableWAM `
        -EnableErrorReporting `
        -LogDirectoryPath $LogDirectory

    $labels = @(Get-ComplianceTag -IncludingLabelState | Select-Object Name, RetentionAction, RetentionDuration -First 10)

    [pscustomobject]@{
        LabelPreviewCount = $labels.Count
        LabelPreview      = $labels
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

    $logDirectory = Join-Path -Path $OutputDirectory -ChildPath 'module-logs'
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

    $steps = @()

    $steps += Invoke-GrcStep -Name 'Graph app-only connection' -ScriptBlock {
        Connect-GrcGraphAppOnly `
            -Tenant $organization `
            -ClientId $env:GRC_CLIENT_ID `
            -Certificate $certificate
    }

    $steps += Invoke-GrcStep -Name 'Graph app registration diagnostics' -ScriptBlock {
        Get-GrcAppRegistrationDiagnostics -ClientId $env:GRC_CLIENT_ID
    }

    $steps += Invoke-GrcStep -Name 'Exchange Online app-only connection' -ScriptBlock {
        Connect-GrcExchangeOnline `
            -ClientId $env:GRC_CLIENT_ID `
            -Organization $organization `
            -CertificateThumbprint $certificate.Thumbprint `
            -LogDirectory $logDirectory
    }

    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

    $steps += Invoke-GrcStep -Name 'Security and Compliance IPPS connection' -ScriptBlock {
        Connect-GrcIppsSession `
            -ClientId $env:GRC_CLIENT_ID `
            -Organization $organization `
            -CertificateThumbprint $certificate.Thumbprint `
            -LogDirectory $logDirectory
    }

    $steps |
        ConvertTo-Json -Depth 30 |
        Out-File -FilePath (Join-Path -Path $OutputDirectory -ChildPath 'auth-diagnostics-result.json') -Encoding utf8

    $failedSteps = @($steps | Where-Object { -not $_.Success })

    if ($failedSteps.Count -gt 0) {
        Write-Warning "$($failedSteps.Count) diagnostic step(s) failed. See auth-diagnostics-result.json and module-logs artifact."
        exit 1
    }

    Write-Host 'All diagnostic steps completed successfully.'
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue

    Remove-GrcImportedCertificate -Thumbprint $script:ImportedCertificateThumbprint

    if (-not [string]::IsNullOrWhiteSpace($script:TemporaryCertificatePath) -and (Test-Path -Path $script:TemporaryCertificatePath)) {
        Remove-Item -Path $script:TemporaryCertificatePath -Force -ErrorAction SilentlyContinue
    }
}
