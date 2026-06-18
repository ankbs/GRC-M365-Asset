#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Purview for sensitivity labels and DLP policy configurations.
.DESCRIPTION
    Retrieves key Purview Compliance and Information Protection metrics for GRC auditing.
    Exports structured JSON/CSV data.
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

# Import GRC Common library
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
    Write-Error "Graph authentication failed: $_"
    return
}

# 1b. Establish connection to Security & Compliance Center for DLP Policies
$ippsConnected = $false
try {
    if ($Interactive) {
        Connect-GRCCompliance -Interactive
    } else {
        Connect-GRCCompliance -TenantId $TenantId -ClientId $ClientId -CertificateBase64 $CertificateBase64
    }
    $ippsConnected = $true
} catch {
    Write-Warning "Security & Compliance Center authentication failed (DLP policies won't be collected): $_"
    if ($_.Exception) {
        Write-Warning "Exception Type: $($_.Exception.GetType().FullName)"
        Write-Warning "Exception Message: $($_.Exception.Message)"
        Write-Warning "Exception StackTrace: $($_.Exception.StackTrace)"
        if ($_.Exception.InnerException) {
            Write-Warning "Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }
}

# 2. Query Purview Settings using REST calls
$reportData = [Ordered]@{
    TotalSensitivityLabels   = 0
    SensitivityLabelNames    = ""
    SensitivityLabelsDetails = @()
    TotalDlpPolicies         = 0
    DlpPolicyNames           = ""
    DlpPoliciesDetails       = @()
    TotalRetentionLabels     = 0
    RetentionLabelsDetails   = @()
    UserSensitivityLabels    = @()
    LabelPolicySettings      = [Ordered]@{
        IsMandatory                    = $false
        DefaultLabelId                 = ""
        DowngradeJustificationRequired = $false
    }
}

try {
    # Query Sensitivity Labels (Try v1.0 security endpoint first, fallback to beta)
    $labelsResponse = $null
    $endpoints = @(
        "https://graph.microsoft.com/v1.0/security/dataSecurityAndGovernance/sensitivityLabels",
        "https://graph.microsoft.com/beta/security/informationProtection/sensitivityLabels"
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
            if ($_.name) { $_.name } else { $_.displayName }
        }
        $reportData.SensitivityLabelNames = ($labelNames | Where-Object { $_ }) -join '; '

        # Collect Sensitivity Label details
        $reportData.SensitivityLabelsDetails = $labelsResponse.value | ForEach-Object {
            $labelName = if ($_.name) { $_.name } else { $_.displayName }
            [Ordered]@{
                Id          = $_.id
                Name        = $labelName
                Description = $_.description
                IsActive    = $_.isActive
                Sensitivity = $_.sensitivity
            }
        }
    }

    # Query DLP Policies (using Exchange Online Security & Compliance PowerShell)
    if ($ippsConnected) {
        try {
            $dlpPolicies = Get-DlpCompliancePolicy -ErrorAction Stop
            $reportData.TotalDlpPolicies = @($dlpPolicies).Count
            $reportData.DlpPolicyNames = ($dlpPolicies | Select-Object -ExpandProperty Name) -join '; '
            # Collect DLP Policy details
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

    # Query Retention Labels (Records Management)
    if ($ippsConnected) {
        try {
            $retentionLabels = Get-ComplianceTag -ErrorAction Stop
            $reportData.TotalRetentionLabels = @($retentionLabels).Count
            $reportData.RetentionLabelsDetails = $retentionLabels | ForEach-Object {
                [Ordered]@{
                    Id                            = $_.Identity.ToString()
                    DisplayName                   = $_.Name
                    BehaviorDuringRetentionPeriod = $_.RetentionType
                    ActionAfterRetentionPeriod    = $_.RetentionAction
                }
            }
        } catch {
            Write-Warning "Could not query retention labels via Security & Compliance: $_"
        }
    } else {
        $retResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/security/labels/retentionLabels" -ErrorAction SilentlyContinue
        if ($retResponse -and $retResponse.value) {
            $reportData.TotalRetentionLabels = @($retResponse.value).Count
            $reportData.RetentionLabelsDetails = $retResponse.value | ForEach-Object {
                [Ordered]@{
                    Id                            = $_.id
                    DisplayName                   = $_.displayName
                    BehaviorDuringRetentionPeriod = $_.behaviorDuringRetentionPeriod
                    ActionAfterRetentionPeriod    = $_.actionAfterRetentionPeriod
                }
            }
        }
    }

    # Query User-Specific Sensitivity Labels (for the first 5 active users)
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

    # Query Label Policy Settings (Classification policies)
    $policySettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/security/informationProtection/labelPolicySettings" -ErrorAction SilentlyContinue
    if ($policySettings) {
        $reportData.LabelPolicySettings = [Ordered]@{
            IsMandatory                    = $policySettings.isMandatory
            DefaultLabelId                 = $policySettings.defaultLabelId
            DowngradeJustificationRequired = $policySettings.downgradeSensitivityLabelJustificationRequired
        }
    }

} catch {
    Write-Warning "Could not query all Microsoft Purview endpoints: $_"
}

# Disconnect compliance sessions
if ($ippsConnected) {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
}

# 3. Handle Outputs based on execution scope
$exportObj = [PSCustomObject]$reportData
if ($AiAgentMode) {
    $exportObj | ConvertTo-Json -Depth 5
} else {
    Export-GRCAssetData -ServiceName "Purview" -AssetName "PurviewSummary" -Data @($exportObj)
}
