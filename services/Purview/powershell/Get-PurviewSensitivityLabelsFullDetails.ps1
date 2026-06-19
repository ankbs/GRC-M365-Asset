#Requires -Version 7.2
<#
.SYNOPSIS
    Queries Microsoft Purview for full detailed sensitivity label configurations.
.DESCRIPTION
    Retrieves deep details of Microsoft Purview Sensitivity Labels, including:
    - Base properties (Guid, Priority, Color, ParentGuid, IsActive)
    - Copilot Protection status (BlockContentAnalysisServices)
    - Encryption & Rights Management (RMS) Template details, protection type, and rights definitions
    - Locale settings (translations of display names and tooltips)
    - Workload scopes (Files, Emails, Meetings, Sites, SchematizedData)
    - Scoped publishing details (users, groups, AUs) by analyzing Get-Label and Get-LabelPolicy
    
    Exports structured JSON and flattened CSV to exports/Purview/SensitivityLabelsFullDetails/.
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
    [switch]$Interactive
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
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId
    )

    if (-not [string]::IsNullOrWhiteSpace($env:GRC_ORGANIZATION)) {
        return $env:GRC_ORGANIZATION
    }

    if ($TenantId -match '\.onmicrosoft\.com$') {
        return $TenantId
    }

    try {
        $orgResponse = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
        if ($orgResponse -and $orgResponse.value) {
            $defaultDomain = $orgResponse.value[0].verifiedDomains |
                Where-Object { $_.isDefault } |
                Select-Object -ExpandProperty name
            if ($defaultDomain) {
                return $defaultDomain
            }
        }
    } catch {
        Write-Warning "Failed to resolve organization domain via Graph: $_"
    }

    throw "No usable organization domain found for Connect-IPPSSession."
}

function New-GrcTemporaryPfxFile {
    param([string]$Base64Pfx)
    $certBytes  = [Convert]::FromBase64String($Base64Pfx)
    $certDir    = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'grc-certificates'
    if (-not (Test-Path -Path $certDir)) {
        New-Item -Path $certDir -ItemType Directory -Force | Out-Null
    }
    $certPath = Join-Path -Path $certDir -ChildPath "grc-purview-details-$([guid]::NewGuid()).pfx"
    [System.IO.File]::WriteAllBytes($certPath, $certBytes)
    return $certPath
}

function Import-GrcPfxCertificate {
    param(
        [string]$CertificateFilePath,
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
    return $certificate
}

function Remove-GrcImportedCertificate {
    param([string]$Thumbprint)
    if ([string]::IsNullOrWhiteSpace($Thumbprint)) { return }
    $certPath = "Cert:\CurrentUser\My\$Thumbprint"
    if (Test-Path -Path $certPath) {
        Remove-Item -Path $certPath -Force -ErrorAction SilentlyContinue
    }
}

function Connect-GrcComplianceSession {
    param(
        [string]$ClientId,
        [string]$Organization,
        [string]$CertificateThumbprint,
        [string]$LogDirectory
    )
    $module = Get-Module ExchangeOnlineManagement -ListAvailable |
        Where-Object { $_.Version -eq [version]'3.9.0' } |
        Select-Object -First 1

    if ($null -eq $module) {
        throw 'ExchangeOnlineManagement 3.9.0 is not available.'
    }
    Import-Module ExchangeOnlineManagement -RequiredVersion 3.9.0 -ErrorAction Stop

    if (-not (Test-Path -Path $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
    }

    Connect-IPPSSession `
        -AppId                $ClientId `
        -CertificateThumbprint $CertificateThumbprint `
        -Organization         $Organization `
        -CommandName          Get-Label, Get-LabelPolicy `
        -ShowBanner:$false `
        -DisableWAM `
        -EnableErrorReporting `
        -LogDirectoryPath     $LogDirectory
}

function Get-GrcSafeProperty {
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
    # 1. Establish Graph connection for tenant queries
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

    $orgDomain = Get-GrcOrganizationName -TenantId $TenantId

    $exportDir = Join-Path -Path $PSScriptRoot -ChildPath '../../../exports/Purview/SensitivityLabelsFullDetails'
    if (-not (Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    $ippsLogDir = Join-Path -Path $exportDir -ChildPath 'ipps-logs'

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

            Connect-GrcComplianceSession `
                -ClientId              $ClientId `
                -Organization          $orgDomain `
                -CertificateThumbprint $certificate.Thumbprint `
                -LogDirectory          $ippsLogDir

            $ippsConnected = $true
        } catch {
            Write-Error "Security & Compliance Center authentication failed: $_"
            return
        }
    } elseif ($Interactive) {
        try {
            # Standard compliance interactive connection
            Connect-IPPSSession -ShowBanner:$false
            $ippsConnected = $true
        } catch {
            Write-Error "Interactive Security & Compliance Center authentication failed: $_"
            return
        }
    }

    if (-not $ippsConnected) {
        Write-Error "Could not establish IPPSSession. Exiting."
        return
    }

    Write-Host "Retrieving Sensitivity Labels and Policies from Microsoft Purview..." -ForegroundColor Cyan
    $ippsLabels = @(Get-Label -IncludeDetailedLabelActions -SkipValidations)
    $ippsPolicies = @(Get-LabelPolicy)
    Write-Host "Found $($ippsLabels.Count) sensitivity labels and $($ippsPolicies.Count) label policies."

    # Parse and transform data for both JSON and CSV
    $jsonDetails = @()
    $csvDetails = @()

    foreach ($lbl in $ippsLabels) {
        # Parse settings array (Key=Value format)
        $settingsDict = @{}
        $lblSettings = Get-GrcSafeProperty -InputObject $lbl -Name 'Settings'
        if ($lblSettings) {
            foreach ($setting in $lblSettings) {
                if ($setting -match '^([^=]+)=(.*)$') {
                    $settingsDict[$Matches[1]] = $Matches[2]
                }
            }
        }

        # 1. Base Properties
        $guidObj = Get-GrcSafeProperty -InputObject $lbl -Name 'Guid'
        $guid = if ($guidObj) { $guidObj.ToString() } else { "" }
        $name = Get-GrcSafeProperty -InputObject $lbl -Name 'Name'
        $displayName = Get-GrcSafeProperty -InputObject $lbl -Name 'DisplayName'
        $comment = Get-GrcSafeProperty -InputObject $lbl -Name 'Comment'
        $isActive = Get-GrcSafeProperty -InputObject $lbl -Name 'IsActive'
        $priority = Get-GrcSafeProperty -InputObject $lbl -Name 'Priority'
        $color = Get-GrcSafeProperty -InputObject $lbl -Name 'Color'
        
        $parentGuidObj = Get-GrcSafeProperty -InputObject $lbl -Name 'ParentGuid'
        $parentGuid = if ($parentGuidObj) { $parentGuidObj.ToString() } else { "" }

        # 2. Copilot Protection
        $blockCopilot = $settingsDict.ContainsKey('BlockContentAnalysisServices') -and $settingsDict['BlockContentAnalysisServices'] -ieq 'True'

        # 3. Workload Scopes (with backwards compatibility for SchematizedData)
        $contentType = Get-GrcSafeProperty -InputObject $lbl -Name 'ContentType'
        $scopeFiles = $contentType -match 'File'
        $scopeEmails = $contentType -match 'Email'
        $scopeMeetings = $contentType -match 'Meeting'
        $scopeSites = $contentType -match 'Site' -or $contentType -match 'UnifiedGroup'
        
        # SchematizedData: Old models used 'SchematizedData' in ContentType.
        # New models configure ScopeSchematizedData in advanced settings/LabelActions.
        $scopeSchematizedData = $false
        if ($contentType -match 'SchematizedData') {
            $scopeSchematizedData = $true
        }
        if ($settingsDict.ContainsKey('ScopeSchematizedData') -and $settingsDict['ScopeSchematizedData'] -ieq 'True') {
            $scopeSchematizedData = $true
        }
        $lblLabelActions = Get-GrcSafeProperty -InputObject $lbl -Name 'LabelActions'
        if ($lblLabelActions) {
            foreach ($action in $lblLabelActions) {
                $applicableTo = Get-GrcSafeProperty -InputObject $action -Name 'ApplicableTo'
                if ($applicableTo -match 'SchematizedData' -or $applicableTo -match 'Purview') {
                    $scopeSchematizedData = $true
                }
            }
        }

        # 4. Encryption & RMS Settings
        $encryptionEnabled = $false
        $protectionType = 'None'
        $templateId = ''
        $rightsDefinitions = @()
        $flatRightsStr = ''

        if ($lblLabelActions) {
            foreach ($action in $lblLabelActions) {
                # Look for Protect or Encryption action type
                $actionType = Get-GrcSafeProperty -InputObject $action -Name 'ActionType'
                $templateIdVal = Get-GrcSafeProperty -InputObject $action -Name 'TemplateId'
                if ($actionType -eq 'Protect' -or $actionType -eq 'Encryption' -or $null -ne $templateIdVal) {
                    $encryptionEnabled = $true
                    $templateId = $templateIdVal
                    $protectionTypeObj = Get-GrcSafeProperty -InputObject $action -Name 'ProtectionType'
                    $protectionType = if ($protectionTypeObj) { $protectionTypeObj } else { 'Template' }

                    $rightsDefs = Get-GrcSafeProperty -InputObject $action -Name 'RightsDefinitions'
                    if ($rightsDefs) {
                        # RightsDefinitions is typically an array of objects/hashtables
                        # showing identities and their granted rights
                        $rightsList = @()
                        $flatRights = @()
                        foreach ($def in $rightsDefs) {
                            $identity = Get-GrcSafeProperty -InputObject $def -Name 'Identity'
                            $rightsArr = Get-GrcSafeProperty -InputObject $def -Name 'Rights'
                            $rights = @($rightsArr) -join ','
                            $rightsList += @{
                                Identity = $identity
                                Rights   = $rightsArr
                            }
                            $flatRights += "$($identity):($rights)"
                        }
                        $rightsDefinitions = $rightsList
                        $flatRightsStr = $flatRights -join '; '
                    }
                }
            }
        }

        # 5. Locale Translations (Languages)
        $localeDict = [Ordered]@{}
        $flatLocaleList = @()
        $lblLocaleSettings = Get-GrcSafeProperty -InputObject $lbl -Name 'LocaleSettings'
        if ($lblLocaleSettings) {
            foreach ($locSetting in $lblLocaleSettings) {
                if ($locSetting -match '^([^=]+)_(.+?)=(.*)$') {
                    $propName = $Matches[1]
                    $langCode = $Matches[2]
                    $propVal  = $Matches[3]

                    if (-not $localeDict.ContainsKey($langCode)) {
                        $localeDict[$langCode] = [Ordered]@{}
                    }
                    $localeDict[$langCode][$propName] = $propVal
                }
            }
            foreach ($lang in $localeDict.Keys) {
                $dn = if ($localeDict[$lang].Contains('DisplayName')) { $localeDict[$lang]['DisplayName'] } else { '' }
                $desc = if ($localeDict[$lang].Contains('Description')) { $localeDict[$lang]['Description'] } else { '' }
                $flatLocaleList += "$lang:DN=`"$dn`",Desc=`"$desc`""
            }
        }
        $flatLocalesStr = $flatLocaleList -join '; '

        # 6. Publishing Policies & Audience scoping (isScopedToUser)
        $isScopedToUser = $false
        if ($settingsDict.ContainsKey('isScopedToUser') -and $settingsDict['isScopedToUser'] -ieq 'True') {
            $isScopedToUser = $true
        }

        # Find which policies distribute this label
        $scopedPolicies = @()
        $publishedToUsers = @()
        $publishedToGroups = @()
        
        foreach ($policy in $ippsPolicies) {
            if ($policy.Labels -contains $guid -or $policy.Labels -contains $name) {
                $scopedPolicies += $policy.Name
                # Get location bindings / target distribution
                if ($policy.UserLocation) {
                    foreach ($loc in $policy.UserLocation) {
                        if ($loc -eq 'All') {
                            $publishedToUsers += 'All Users'
                        } else {
                            $publishedToUsers += $loc
                        }
                    }
                }
                if ($policy.ModernGroupLocation) {
                    foreach ($grp in $policy.ModernGroupLocation) {
                        $publishedToGroups += $grp
                    }
                }
            }
        }
        $publishedToUsers = $publishedToUsers | Select-Object -Unique
        $publishedToGroups = $publishedToGroups | Select-Object -Unique

        # --- Populate JSON details ---
        $jsonDetails += [Ordered]@{
            Id                     = $guid
            Name                   = $name
            DisplayName            = $displayName
            Comment                = $comment
            IsActive               = $isActive
            Priority               = $priority
            Color                  = $color
            ParentGuid             = $parentGuid
            BlockCopilot           = $blockCopilot
            ScopeFiles             = $scopeFiles
            ScopeEmails            = $scopeEmails
            ScopeMeetings          = $scopeMeetings
            ScopeSites             = $scopeSites
            ScopeSchematizedData   = $scopeSchematizedData
            EncryptionEnabled      = $encryptionEnabled
            ProtectionType         = $protectionType
            TemplateId             = $templateId
            RightsDefinitions      = $rightsDefinitions
            LocaleSettings         = $localeDict
            IsScopedToUser         = $isScopedToUser
            PublishedPolicies      = $scopedPolicies
            PublishedToUsers       = $publishedToUsers
            PublishedToGroups      = $publishedToGroups
        }

        # --- Populate CSV details (flat strings only) ---
        $csvDetails += [Ordered]@{
            Id                     = $guid
            Name                   = $name
            DisplayName            = $displayName
            Comment                = $comment
            IsActive               = $isActive
            Priority               = $priority
            Color                  = $color
            ParentGuid             = $parentGuid
            BlockCopilot           = $blockCopilot
            ScopeFiles             = $scopeFiles
            ScopeEmails            = $scopeEmails
            ScopeMeetings          = $scopeMeetings
            ScopeSites             = $scopeSites
            ScopeSchematizedData   = $scopeSchematizedData
            EncryptionEnabled      = $encryptionEnabled
            ProtectionType         = $protectionType
            TemplateId             = $templateId
            RightsDefinitions      = $flatRightsStr
            LocaleSettings         = $flatLocalesStr
            IsScopedToUser         = $isScopedToUser
            PublishedPolicies      = ($scopedPolicies -join '; ')
            PublishedToUsers       = ($publishedToUsers -join '; ')
            PublishedToGroups      = ($publishedToGroups -join '; ')
        }
    }

    # Save exports
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $jsonPath = Join-Path -Path $exportDir -ChildPath "SensitivityLabelsFullDetails_${timestamp}.json"
    $csvPath  = Join-Path -Path $exportDir -ChildPath "SensitivityLabelsFullDetails_${timestamp}.csv"

    $jsonDetails | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding utf8
    $csvDetails | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

    Write-Host "JSON details written to: $jsonPath" -ForegroundColor Green
    Write-Host "CSV details written to: $csvPath" -ForegroundColor Green

} finally {
    if ($ippsConnected) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    Remove-GrcImportedCertificate -Thumbprint $script:ImportedCertificateThumbprint

    if (-not [string]::IsNullOrWhiteSpace($script:TemporaryCertificatePath) -and (Test-Path -Path $script:TemporaryCertificatePath)) {
        Remove-Item -Path $script:TemporaryCertificatePath -Force -ErrorAction SilentlyContinue
    }
}
