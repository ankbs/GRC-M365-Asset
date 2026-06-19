#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Graph REST API for Microsoft Secure Score and detailed recommendations.
.DESCRIPTION
    A cloud-native GRC Asset script designed to retrieve Microsoft Secure Score history
    and detailed control profiles (security recommendations) from Graph API.
    Groups recommendations by category and maps licensing requirements.
    Exports structured GRC results (JSON, CSV).

    Required Graph Permissions: SecurityEvents.Read.All

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
$commonModulePath = Join-Path -Path $PSScriptRoot -ChildPath "../../../common/GRC-M365-Common.psm1"
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
    Write-Error "Graph authentication failed: $_"
    return
}

# 2. Helper function to map license requirements
function Get-GrcLicenseRequirement {
    param(
        [string]$Service,
        [string]$Title,
        [string]$Category
    )
    if ($Title -match "Conditional Access" -or $Title -match "CA policy" -or $Title -match "bedingten Zugriff") {
        return "Microsoft Entra ID Premium Plan 1 (oder M365 Business Premium / E3 / E5)"
    }
    if ($Service -ieq "Azure Active Directory" -or $Service -match "Entra" -or $Service -match "Active Directory") {
        if ($Title -match "Identity Protection" -or $Title -match "Risk-based" -or $Title -match "Risiko") {
            return "Microsoft Entra ID Premium Plan 2 (oder M365 E5 / Security)"
        }
        if ($Title -match "Privileged Identity" -or $Title -match "PIM") {
            return "Microsoft Entra ID Premium Plan 2 (oder M365 E5 / Security)"
        }
        return "Microsoft Entra ID Free / Premium P1"
    }
    if ($Service -ieq "Intune" -or $Service -match "Intune" -or $Service -match "Device" -or $Service -match "Gerät") {
        return "Microsoft Intune (oder M365 Business Premium / E3 / E5)"
    }
    if ($Service -match "Exchange" -or $Service -match "EXO" -or $Service -match "Outlook") {
        return "Exchange Online Plan 1/2 (oder M365 Business Standard / E3 / E5)"
    }
    if ($Service -match "SharePoint" -or $Service -match "OneDrive") {
        return "SharePoint Online Plan 1/2 (oder M365 Business Standard / E3 / E5)"
    }
    if ($Service -match "Defender" -or $Service -match "Endpoint" -or $Service -match "Threat") {
        return "Microsoft Defender Plan 1/2 / Defender for Business (oder M365 Business Premium / E5)"
    }
    if ($Service -match "Purview" -or $Service -match "Compliance" -or $Service -match "DLP" -or $Service -match "Retention" -or $Service -match "Information Protection" -or $Service -match "Label") {
        return "Microsoft Purview / Information Protection Plan 1/2 (oder M365 E5 / Compliance)"
    }
    return "M365 Standard-Lizenz"
}

# 3. Retrieve Secure Scores Summary
Write-Host "Querying Microsoft Secure Score summary..." -ForegroundColor Cyan
$secureScoreResponse = $null
try {
    $secureScoreResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/security/secureScores"
} catch {
    Write-Warning "Could not query Secure Score summary: $_"
}

$summaryObj = [ordered]@{
    CurrentScore = 0
    MaxScore = 0
    Percentage = 0
    CreatedDateTime = (Get-Date).ToString("o")
    M365Average = 0
}

if ($secureScoreResponse -and $secureScoreResponse.value -and $secureScoreResponse.value.Count -gt 0) {
    $latestScore = $secureScoreResponse.value[0]
    $summaryObj.CurrentScore = $latestScore.currentScore
    $summaryObj.MaxScore = $latestScore.maxScore
    if ($latestScore.maxScore -gt 0) {
        $summaryObj.Percentage = [Math]::Round(($latestScore.currentScore / $latestScore.maxScore) * 100, 2)
    }
    $summaryObj.CreatedDateTime = $latestScore.createdDateTime
    
    # Extract comparative average if available
    if ($latestScore.averageComparativeScores) {
        foreach ($comp in $latestScore.averageComparativeScores) {
            if ($comp.basis -eq "allTenants") {
                $summaryObj.M365Average = [Math]::Round($comp.averageScore, 2)
            }
        }
    }
}

# 4. Retrieve Secure Control Profiles (Detailed Recommendations)
Write-Host "Querying Microsoft Secure Control Profiles (Recommendations)..." -ForegroundColor Cyan
$profilesList = [System.Collections.Generic.List[PSCustomObject]]::new()
$nextLink = "https://graph.microsoft.com/v1.0/security/secureControlProfiles"

try {
    while ($nextLink) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink
        if ($response -and $response.value) {
            foreach ($profile in $response.value) {
                # Map control profiles to clean recommendations
                $id = $profile.id
                $title = $profile.title
                $category = $profile.category
                $maxScore = $profile.maxScore
                $service = $profile.service
                $userImpact = $profile.userImpact
                $remediation = $profile.remediation
                $actionText = $profile.actionText
                $implementationStatus = $profile.implementationStatus
                
                # Strip HTML from remediation if any
                $remediationClean = if ($remediation) { 
                    [regex]::Replace($remediation, "<[^>]*>", "").Replace("`n", " ").Replace("`r", "").Trim() 
                } else { 
                    $actionText 
                }
                
                $license = Get-GrcLicenseRequirement -Service $service -Title $title -Category $category
                
                $profileObj = [ordered]@{
                    CheckId = $id
                    Title = $title
                    Category = $category
                    Service = $service
                    MaxScore = $maxScore
                    UserImpact = $userImpact
                    ImplementationStatus = $implementationStatus
                    LicenseRequired = $license
                    Remediation = $remediationClean
                }
                
                $profilesList.Add([PSCustomObject]$profileObj)
            }
        }
        $nextLink = if ($response -and $response.'@odata.nextLink') { $response.'@odata.nextLink' } else { $null }
    }
} catch {
    Write-Warning "Could not query Secure Control Profiles: $_"
}

# 5. Handle Outputs
if ($AiAgentMode) {
    # Combine outputs for AI Mode
    $combined = [ordered]@{
        Summary = $summaryObj
        Recommendations = $profilesList.ToArray()
    }
    $combined | ConvertTo-Json -Depth 10
} else {
    # 5.1 Export Summary
    $summaryDir = Join-Path -Path $PSScriptRoot -ChildPath "../../../exports/SecurityScore/SecurityScoreSummary"
    if (!(Test-Path $summaryDir)) {
        New-Item -ItemType Directory -Path $summaryDir -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $summaryJsonPath = Join-Path -Path $summaryDir -ChildPath "SecurityScoreSummary_${timestamp}.json"
    $summaryCsvPath  = Join-Path -Path $summaryDir -ChildPath "SecurityScoreSummary_${timestamp}.csv"
    
    [PSCustomObject]$summaryObj | ConvertTo-Json | Set-Content -Path $summaryJsonPath -Encoding utf8
    [PSCustomObject]$summaryObj | Export-Csv -Path $summaryCsvPath -NoTypeInformation -Encoding utf8
    
    # 5.2 Export Recommendations
    $detailsDir = Join-Path -Path $PSScriptRoot -ChildPath "../../../exports/SecurityScore/SecurityScoreDetails"
    if (!(Test-Path $detailsDir)) {
        New-Item -ItemType Directory -Path $detailsDir -Force | Out-Null
    }
    $detailsJsonPath = Join-Path -Path $detailsDir -ChildPath "SecurityScoreDetails_${timestamp}.json"
    $detailsCsvPath  = Join-Path -Path $detailsDir -ChildPath "SecurityScoreDetails_${timestamp}.csv"
    
    $profilesList.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $detailsJsonPath -Encoding utf8
    $profilesList.ToArray() | Export-Csv -Path $detailsCsvPath -NoTypeInformation -Encoding utf8
    
    Write-Host "Secure Score Summary JSON written to: $summaryJsonPath" -ForegroundColor Green
    Write-Host "Secure Score Details JSON written to: $detailsJsonPath" -ForegroundColor Green
}
