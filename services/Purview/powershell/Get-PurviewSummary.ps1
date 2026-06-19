#Requires -Version 7.2
<#
.SYNOPSIS
    Queries Microsoft Purview for sensitivity labels, DLP policies, and retention labels.

.DESCRIPTION
    Retrieves key Purview Compliance and Information Protection metrics for GRC auditing.
    Exports structured JSON/CSV data.

    Authentication for Security & Compliance PowerShell:
    - Uses Import-PfxCertificate into Cert:\CurrentUser\My (Windows certificate store).
    - Connects with -CertificateThumbprint and -DisableWAM.
    - Requires ExchangeOnlineManagement 3.9.0 (pinned).
    - Designed for windows-latest GitHub Actions runner.

    Reason: Connect-IPPSSession with CSP-provider certificates requires the Windows
    certificate store. EphemeralKeySet (in-memory) is not compatible with CSP-provider
    private keys on Linux runners.

    Graph is used for sensitivity labels only (supported with app-only permissions).
    Retention labels and DLP policies use Security & Compliance PowerShell (Get-ComplianceTag,
    Get-DlpCompliancePolicy) because the Graph retentionLabels endpoint does not support
    application permissions.

.PARAMETER TenantId
    Azure AD Tenant ID (GUID) or primary .onmicrosoft.com domain name.

.PARAMETER ClientId
    Application (client) ID of the app registration.

.PARAMETER CertificateBase64
    Base64-encoded PFX certificate (with private key) for app-only authentication.

.PARAMETER CertificatePassword
    Optional password for the PFX certificate. Leave empty if not password-protected.

.PARAMETER Interactive
    Use interactive user authentication instead of certificate-based app-only auth.

.PARAMETER AiAgentMode
    When set, outputs JSON to stdout instead of writing to the exports directory.
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
    [string]$CertificatePassword = '',

    [Parameter(Mandatory = $false)]
    [switch]$Interactive,

    [Parameter(Mandatory = $false)]
    [switch]$AiAgentMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Track certificate thumbprint for cleanup in finally block
$script:ImportedCertificateThumbprint = $null
$script:TemporaryCertificatePath = $null

# Import GRC Common library
$commonModulePath = Join-Path -Path $PSScriptRoot -ChildPath '../../../common/GRC-M365-Common.psm1'
if (Test-Path $commonModulePath) {
    Import-Module -Name $commonModulePath -Force
} else {
    Write-Error "Required GRC Common module not found at: $commonModulePath"
    return
}

#region --- Helper Functions ---

function Get-GrcOrganizationName {
    <#
    .SYNOPSIS
        Resolves the primary .onmicrosoft.com domain name for Connect-IPPSSession.
    .DESCRIPTION
        Connect-IPPSSession requires -Organization to be the primary onmicrosoft.com
        domain, not a GUID. Precedence: GRC_ORGANIZATION env var, then GRC_TENANT_ID
        if it ends with .onmicrosoft.com, then resolved via Microsoft Graph.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId
    )

    # 1. Explicit GRC_ORGANIZATION secret takes precedence
    if (-not [string]::IsNullOrWhiteSpace($env:GRC_ORGANIZATION)) {
        Write-Verbose "Using GRC_ORGANIZATION secret: $($env:GRC_ORGANIZATION)"
        return $env:GRC_ORGANIZATION
    }

    # 2. If TenantId is already a domain name (ends with .onmicrosoft.com)
    if ($TenantId -match '\.onmicrosoft\.com$') {
        Write-Verbose "TenantId is already an onmicrosoft.com domain: $TenantId"
        return $TenantId
    }

    # 3. Resolve via Microsoft Graph (requires prior Connect-MgGraph)
    try {
        Write-Verbose 'Resolving primary organization domain from Graph for IPPS connection...'
        $orgResponse = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
        if ($orgResponse -and $orgResponse.value) {
            $defaultDomain = $orgResponse.value[0].verifiedDomains |
                Where-Object { $_.isDefault } |
                Select-Object -ExpandProperty name
            if ($defaultDomain) {
                Write-Verbose "Resolved organization domain: $defaultDomain"
                return $defaultDomain
            }
        }
    } catch {
        Write-Warning "Failed to resolve organization domain via Graph: $_"
    }

    throw @'
No usable organization domain found for Connect-IPPSSession.

Connect-IPPSSession requires the primary *.onmicrosoft.com tenant domain in -Organization.
Set one of these GitHub repository secrets:
  GRC_ORGANIZATION = yourtenant.onmicrosoft.com
  or set GRC_TENANT_ID = yourtenant.onmicrosoft.com
'@
}

function New-GrcTemporaryPfxFile {
    <#
    .SYNOPSIS
        Writes the base64-encoded PFX to a temporary file.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Base64Pfx
    )

    $certBytes  = [Convert]::FromBase64String($Base64Pfx)
    $certDir    = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'grc-certificates'

    if (-not (Test-Path -Path $certDir)) {
        New-Item -Path $certDir -ItemType Directory -Force | Out-Null
    }

    $certPath = Join-Path -Path $certDir -ChildPath "grc-purview-$([guid]::NewGuid()).pfx"
    [System.IO.File]::WriteAllBytes($certPath, $certBytes)
    return $certPath
}

function Import-GrcPfxCertificate {
    <#
    .SYNOPSIS
        Imports the PFX into Cert:\CurrentUser\My and returns the certificate object.
    .NOTES
        Windows only. Required for CSP-provider certificates which cannot be loaded
        into memory via EphemeralKeySet on non-Windows platforms.
    #>
    [CmdletBinding()]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$CertificateFilePath,

        [Parameter(Mandatory)]
        [securestring]$CertificatePassword
    )

    $certificate = Import-PfxCertificate `
        -FilePath $CertificateFilePath `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -Password $CertificatePassword `
        -Exportable

    if ($null -eq $certificate) {
        throw 'Import-PfxCertificate returned no certificate.'
    }

    if (-not $certificate.HasPrivateKey) {
        throw 'The imported certificate does not contain a private key. Use a PFX file, not a CER file.'
    }

    return $certificate
}

function Remove-GrcImportedCertificate {
    <#
    .SYNOPSIS
        Removes the certificate from Cert:\CurrentUser\My after use.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowEmptyString()]
        [string]$Thumbprint
    )

    if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
        return
    }

    $certPath = "Cert:\CurrentUser\My\$Thumbprint"
    if (Test-Path -Path $certPath) {
        Remove-Item -Path $certPath -Force -ErrorAction SilentlyContinue
        Write-Verbose "Removed certificate $Thumbprint from CurrentUser\My store."
    }
}

function Connect-GrcComplianceSession {
    <#
    .SYNOPSIS
        Connects to Security & Compliance PowerShell using CertificateThumbprint.
    .NOTES
        Requires ExchangeOnlineManagement 3.9.0 (pinned).
        Uses -DisableWAM to prevent interactive login prompts on GitHub Actions runners.
        Only loads the cmdlets needed for GRC collection via -CommandName.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ClientId,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Organization,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$CertificateThumbprint,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$LogDirectory
    )

    # Ensure pinned module version is available
    $module = Get-Module ExchangeOnlineManagement -ListAvailable |
        Where-Object { $_.Version -eq [version]'3.9.0' } |
        Select-Object -First 1

    if ($null -eq $module) {
        throw 'ExchangeOnlineManagement 3.9.0 is not available. Run: Install-Module ExchangeOnlineManagement -RequiredVersion 3.9.0'
    }

    Import-Module ExchangeOnlineManagement -RequiredVersion 3.9.0 -ErrorAction Stop

    Write-Host "ExchangeOnlineManagement version: $($module.Version)"
    Write-Host "Connecting to Security & Compliance PowerShell for organization '$Organization'."
    Write-Host "Writing IPPS logs to '$LogDirectory'."

    if (-not (Test-Path -Path $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }

    Connect-IPPSSession `
        -AppId                $ClientId `
        -CertificateThumbprint $CertificateThumbprint `
        -Organization         $Organization `
        -CommandName          Get-ComplianceTag, Get-DlpCompliancePolicy `
        -ShowBanner:$false `
        -DisableWAM `
        -EnableErrorReporting `
        -LogDirectoryPath     $LogDirectory
}

function Get-GrcSafeProperty {
    <#
    .SYNOPSIS
        Safely retrieves a property value from a PSCustomObject or Hashtable key
        without throwing PropertyNotFoundException under StrictMode Latest.
    #>
    param(
        [object]$InputObject,
        [string]$Name
    )
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) {
            return $InputObject[$Name]
        }
        foreach ($k in $InputObject.Keys) {
            if ($k.ToString() -ieq $Name) {
                return $InputObject[$k]
            }
        }
        return $null
    }
    $prop = $InputObject.PSObject.Properties[$Name]
    if ($prop) {
        return $prop.Value
    }
    return $null
}

#endregion

try {
    # 1. Establish connection to Microsoft Graph
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

    # 2. Resolve organization domain for IPPS
    $orgDomain = Get-GrcOrganizationName -TenantId $TenantId

    # 3. Prepare exports directory for IPPS logs
    $exportDir = Join-Path -Path $PSScriptRoot -ChildPath '../../../exports/Purview/PurviewSummary'
    if (-not (Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    $ippsLogDir = Join-Path -Path $exportDir -ChildPath 'ipps-logs'

    # 4. Load certificate and connect to Security & Compliance PowerShell
    $ippsConnected = $false

    if (-not $Interactive -and -not [string]::IsNullOrEmpty($CertificateBase64)) {
        try {
            Write-Host 'Writing temporary PFX certificate file...'
            $script:TemporaryCertificatePath = New-GrcTemporaryPfxFile -Base64Pfx $CertificateBase64

            $certPassword = if ([string]::IsNullOrEmpty($CertificatePassword)) {
                [System.Security.SecureString]::new()
            } else {
                ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
            }

            Write-Host "Importing PFX certificate into Cert:\CurrentUser\My..."
            $certificate = Import-GrcPfxCertificate `
                -CertificateFilePath $script:TemporaryCertificatePath `
                -CertificatePassword $certPassword

            $script:ImportedCertificateThumbprint = $certificate.Thumbprint
            Write-Host "Certificate thumbprint: $($certificate.Thumbprint)"
            Write-Host "Certificate has private key: $($certificate.HasPrivateKey)"

            Connect-GrcComplianceSession `
                -ClientId              $ClientId `
                -Organization          $orgDomain `
                -CertificateThumbprint $certificate.Thumbprint `
                -LogDirectory          $ippsLogDir

            $ippsConnected = $true
        } catch {
            Write-Warning "Security & Compliance Center authentication failed (DLP policies and retention labels won't be collected): $_"
            if ($_.Exception) {
                Write-Warning "Exception Type: $($_.Exception.GetType().FullName)"
                Write-Warning "Exception Message: $($_.Exception.Message)"
                if ($_.Exception.StackTrace) {
                    Write-Warning "Exception StackTrace: $($_.Exception.StackTrace)"
                }
                if ($_.Exception.InnerException) {
                    Write-Warning "Inner Exception: $($_.Exception.InnerException.Message)"
                }
            }
        }
    } elseif ($Interactive) {
        try {
            Connect-GRCCompliance -Interactive
            $ippsConnected = $true
        } catch {
            Write-Warning "Interactive Security & Compliance Center authentication failed: $_"
        }
    }

    # 5. Build report data structure
    $reportData = [Ordered]@{
        TotalSensitivityLabels   = 0
        SensitivityLabelNames    = ''
        SensitivityLabelsDetails = @()
        TotalDlpPolicies         = 0
        DlpPolicyNames           = ''
        DlpPoliciesDetails       = @()
        TotalRetentionLabels     = 0
        RetentionLabelsDetails   = @()
        UserSensitivityLabels    = @()
        LabelPolicySettings      = [Ordered]@{
            IsMandatory                    = $false
            DefaultLabelId                 = ''
            DowngradeJustificationRequired = $false
        }
    }

    # 6. Query Sensitivity Labels via Graph (app-only supported)
    try {
        $labelsResponse = $null
        $endpoints = @(
            'https://graph.microsoft.com/v1.0/security/dataSecurityAndGovernance/sensitivityLabels',
            'https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels'
        )
        foreach ($uri in $endpoints) {
            $labelsResponse = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction SilentlyContinue
            if ($labelsResponse -and $labelsResponse.value) {
                break
            }
        }

        if ($labelsResponse -and $labelsResponse.value) {
            $reportData.TotalSensitivityLabels = @($labelsResponse.value).Count

            $labelNames = $labelsResponse.value | ForEach-Object {
                $displayName = Get-GrcSafeProperty -InputObject $_ -Name 'displayName'
                $name = Get-GrcSafeProperty -InputObject $_ -Name 'name'
                if ($displayName) { $displayName } else { $name }
            }
            $reportData.SensitivityLabelNames = ($labelNames | Where-Object { $_ }) -join '; '

            $reportData.SensitivityLabelsDetails = $labelsResponse.value | ForEach-Object {
                $item = $_
                $displayName = Get-GrcSafeProperty -InputObject $item -Name 'displayName'
                $name = Get-GrcSafeProperty -InputObject $item -Name 'name'
                $labelName = if ($displayName) { $displayName } else { $name }
                [Ordered]@{
                    Id          = Get-GrcSafeProperty -InputObject $item -Name 'id'
                    Name        = $labelName
                    Description = Get-GrcSafeProperty -InputObject $item -Name 'description'
                    IsActive    = Get-GrcSafeProperty -InputObject $item -Name 'isActive'
                    Sensitivity = Get-GrcSafeProperty -InputObject $item -Name 'sensitivity'
                    Color       = Get-GrcSafeProperty -InputObject $item -Name 'color'
                }
            }
        }

    } catch {
        Write-Warning "Could not query Sensitivity Labels via Graph: $_"
    }

    # 7. Query DLP Policies via Security & Compliance PowerShell
    if ($ippsConnected) {
        try {
            $dlpPolicies = Get-DlpCompliancePolicy -ErrorAction Stop
            $reportData.TotalDlpPolicies = @($dlpPolicies).Count
            $reportData.DlpPolicyNames = ($dlpPolicies | Select-Object -ExpandProperty Name) -join '; '
            $reportData.DlpPoliciesDetails = $dlpPolicies | ForEach-Object {
                [Ordered]@{
                    Id          = $_.Identity.ToString()
                    Name        = $_.Name
                    Description = $_.Comment
                    State       = $_.Mode.ToString()
                }
            }
        } catch {
            Write-Warning "Could not query DLP policies via Security & Compliance: $_"
        }
    }

    # 8. Query Retention Labels via Security & Compliance PowerShell
    # NOTE: Graph /security/labels/retentionLabels does NOT support application permissions.
    #       Always use Get-ComplianceTag via IPPS for app-only authentication.
    if ($ippsConnected) {
        try {
            $retentionLabels = Get-ComplianceTag -IncludingLabelState -ErrorAction Stop
            $reportData.TotalRetentionLabels = @($retentionLabels).Count
            $reportData.RetentionLabelsDetails = $retentionLabels | ForEach-Object {
                $label = $_
                [Ordered]@{
                    Id                            = $label.Identity.ToString()
                    Name                          = $label.Name
                    Guid                          = Get-Member -InputObject $label -Name 'Guid' -ErrorAction SilentlyContinue | ForEach-Object { $label.Guid }
                    RetentionAction               = $label.RetentionAction
                    RetentionDuration             = $label.RetentionDuration
                    RetentionType                 = $label.RetentionType
                    IsRecordLabel                 = $label.IsRecordLabel
                    Disabled                      = $label.Disabled
                    Workload                      = $label.Workload
                    Comment                       = $label.Comment
                    WhenCreatedUtc                = $label.WhenCreatedUTC
                    WhenChangedUtc                = $label.WhenChangedUTC
                }
            }
        } catch {
            Write-Warning "Could not query Retention Labels via Security & Compliance: $_"
        }
    }

    # 9. Query User-Specific Sensitivity Labels (first 5 users, via Graph)
    try {
        $usersResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users?`$top=5&`$select=id,userPrincipalName" -ErrorAction SilentlyContinue
        if ($usersResponse -and $usersResponse.value) {
            $reportData.UserSensitivityLabels = $usersResponse.value | ForEach-Object {
                $u = $_
                $userLabels = @()
                $userLabelsRes = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$($u.id)/security/informationProtection/sensitivityLabels" -ErrorAction SilentlyContinue
                if ($userLabelsRes -and $userLabelsRes.value) {
                    $userLabels = $userLabelsRes.value | ForEach-Object { if ($_.name) { $_.name } else { $_.displayName } }
                }
                [Ordered]@{
                    UserPrincipalName = $u.userPrincipalName
                    UserId            = $u.id
                    AvailableLabels   = ($userLabels -join '; ')
                }
            }
        }
    } catch {
        Write-Warning "Could not query User Sensitivity Labels: $_"
    }

    # 10. Query Label Policy Settings via Graph
    try {
        $policySettings = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/security/informationProtection/labelPolicySettings' -ErrorAction SilentlyContinue
        if ($policySettings) {
            $reportData.LabelPolicySettings = [Ordered]@{
                IsMandatory                    = Get-GrcSafeProperty -InputObject $policySettings -Name 'isMandatory'
                DefaultLabelId                 = Get-GrcSafeProperty -InputObject $policySettings -Name 'defaultLabelId'
                DowngradeJustificationRequired = Get-GrcSafeProperty -InputObject $policySettings -Name 'downgradeSensitivityLabelJustificationRequired'
                MandatoryLabelEnabled          = Get-GrcSafeProperty -InputObject $policySettings -Name 'mandatoryLabelEnabled'
                OutlookRecommendedLabelEnabled = Get-GrcSafeProperty -InputObject $policySettings -Name 'outlookRecommendedLabelEnabled'
            }
        }
    } catch {
        Write-Warning "Could not query Label Policy Settings: $_"
    }

    # 11. Export results
    $exportObj = [PSCustomObject]$reportData
    if ($AiAgentMode) {
        $exportObj | ConvertTo-Json -Depth 5
    } else {
        Export-GRCAssetData -ServiceName 'Purview' -AssetName 'PurviewSummary' -Data @($exportObj)
    }

} finally {
    # Disconnect sessions
    if ($ippsConnected) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    # Remove certificate from Windows store
    Remove-GrcImportedCertificate -Thumbprint $script:ImportedCertificateThumbprint

    # Remove temporary PFX file
    if (-not [string]::IsNullOrWhiteSpace($script:TemporaryCertificatePath) -and (Test-Path -Path $script:TemporaryCertificatePath)) {
        Remove-Item -Path $script:TemporaryCertificatePath -Force -ErrorAction SilentlyContinue
    }
}
