#Requires -Version 7.0
<#
.SYNOPSIS
    Local orchestrator to run all M365 GRC Asset collectors, compile the HTML report, and open it.
.DESCRIPTION
    Runs all Entra ID and Device collectors locally (either interactively or using the certificate
    created during initialization), executes the report generator, and opens the resulting
    dashboard in the default browser.
.PARAMETER TenantId
    The target Entra ID Tenant ID.
.PARAMETER ClientId
    The App Registration client ID (for unattended certificate authentication).
.PARAMETER Interactive
    Force interactive browser logon instead of certificate authentication.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'

Write-Host "=== Starting Local GRC Asset Collection & Reporting ===" -ForegroundColor Cyan

# 1. Resolve Authentication Parameters
$certToUse = $null
$authArgs = @{}

if ($Interactive) {
    Write-Host "Forcing interactive login mode..." -ForegroundColor Yellow
    $authArgs["Interactive"] = $true
} else {
    # If ClientId and TenantId are provided, search for local certificate matching the App
    if ($ClientId -and $TenantId) {
        Write-Host "Searching for local certificate for ClientID $ClientId..." -ForegroundColor Cyan
        
        # Check local certificates
        $certs = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Subject -like "*GRC-M365-Asset*" }
        if ($certs.Count -gt 0) {
            $certToUse = $certs[0]
            Write-Host "Found certificate: $($certToUse.Subject) [Thumbprint: $($certToUse.Thumbprint)]" -ForegroundColor Green
            
            $certBytes = $certToUse.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, "")
            $certBase64 = [System.Convert]::ToBase64String($certBytes)
            
            $authArgs["TenantId"] = $TenantId
            $authArgs["ClientId"] = $ClientId
            $authArgs["CertificateBase64"] = $certBase64
        }
    }
    
    if (-not $certToUse) {
        Write-Host "No local certificate details provided or found. Defaulting to Interactive Login Mode..." -ForegroundColor Yellow
        $authArgs["Interactive"] = $true
    }
}

# 2. Run All Collectors Sequentially
$rootPath = $PSScriptRoot

try {
    # 1. Tenant Info
    Write-Host "`nRunning Tenant Info Collector..." -ForegroundColor Yellow
    & "$rootPath/services/EntraID/Users/powershell/Get-TenantInfo.ps1" @authArgs
    
    # 2. Users Summary & Full Details
    Write-Host "`nRunning User Summary Collector..." -ForegroundColor Yellow
    & "$rootPath/services/EntraID/Users/powershell/Get-UserSummary.ps1" @authArgs
    Write-Host "`nRunning User Full Details Collector..." -ForegroundColor Yellow
    & "$rootPath/services/EntraID/Users/powershell/Get-UserFullDetails.ps1" @authArgs
    
    # 3. Groups Summary & Full Details
    Write-Host "`nRunning Group Summary Collector..." -ForegroundColor Yellow
    & "$rootPath/services/EntraID/Groups/powershell/Get-GroupSummary.ps1" @authArgs
    Write-Host "`nRunning Group Full Details Collector..." -ForegroundColor Yellow
    & "$rootPath/services/EntraID/Groups/powershell/Get-GroupFullDetails.ps1" @authArgs
    
    # 4. Entra Devices
    Write-Host "`nRunning Entra ID Devices Collector..." -ForegroundColor Yellow
    & "$rootPath/services/Devices/EntraID-Devices/powershell/Get-EntraDevices.ps1" @authArgs
    
    # 5. Intune Devices
    Write-Host "`nRunning Intune Managed Devices Collector..." -ForegroundColor Yellow
    & "$rootPath/services/Devices/Intune-Devices/powershell/Get-IntuneDevices.ps1" @authArgs

    # 5.1. Intune Compliance & Settings Extra
    Write-Host "`nRunning Intune Compliance and Settings Extra Collector..." -ForegroundColor Yellow
    & "$rootPath/services/Devices/Intune-Devices/powershell/Get-IntuneComplianceAndSettings.ps1" @authArgs
    # 6. Defender Devices & Combined Device Full Details
    Write-Host "`nRunning Defender Endpoint Collector..." -ForegroundColor Yellow
    & "$rootPath/services/Devices/Defender-Devices/powershell/Get-DefenderDevices.ps1" @authArgs
    Write-Host "`nRunning Device Full Details Collector..." -ForegroundColor Yellow
    & "$rootPath/services/Devices/powershell/Get-DeviceFullDetails.ps1" @authArgs

    # 7. Exchange Online Summary & Full Details
    Write-Host "`nRunning Exchange Online Summary Collector..." -ForegroundColor Yellow
    & "$rootPath/services/ExchangeOnline/powershell/Get-ExchangeSummary.ps1" @authArgs
    Write-Host "`nRunning Exchange Online Full Details Collector..." -ForegroundColor Yellow
    & "$rootPath/services/ExchangeOnline/powershell/Get-ExchangeFullDetails.ps1" @authArgs

    # 8. SharePoint Online Summary & Full Details
    Write-Host "`nRunning SharePoint Online Summary Collector..." -ForegroundColor Yellow
    & "$rootPath/services/SharePoint/powershell/Get-SharePointSummary.ps1" @authArgs
    Write-Host "`nRunning SharePoint Online Full Details Collector..." -ForegroundColor Yellow
    & "$rootPath/services/SharePoint/powershell/Get-SharePointFullDetails.ps1" @authArgs

    # 9. Microsoft Teams Summary & Full Details
    Write-Host "`nRunning Microsoft Teams Summary Collector..." -ForegroundColor Yellow
    & "$rootPath/services/Teams/powershell/Get-TeamsSummary.ps1" @authArgs
    Write-Host "`nRunning Microsoft Teams Full Details Collector..." -ForegroundColor Yellow
    & "$rootPath/services/Teams/powershell/Get-TeamsFullDetails.ps1" @authArgs

    # 10. Entra ID Governance
    Write-Host "`nRunning Entra ID Governance Summary Collector..." -ForegroundColor Yellow
    & "$rootPath/services/EntraID/Governance/powershell/Get-EntraGovernanceSummary.ps1" @authArgs

    # 11. Purview Summary & Full Details
    Write-Host "`nRunning Purview Summary Collector..." -ForegroundColor Yellow
    & "$rootPath/services/Purview/powershell/Get-PurviewSummary.ps1" @authArgs
    Write-Host "`nRunning Purview Sensitivity Labels Full Details Collector..." -ForegroundColor Yellow
    & "$rootPath/services/Purview/powershell/Get-PurviewSensitivityLabelsFullDetails.ps1" @authArgs

    # 12. Security Score & Recommendations
    Write-Host "`nRunning Security Score & Recommendations Collector..." -ForegroundColor Yellow
    & "$rootPath/services/SecurityScore/powershell/Get-SecurityScoreDetails.ps1" @authArgs

} catch {
    Write-Error "Error during collection run: $_"
    return
}

# 3. Compile the GRC HTML Report
Write-Host "`n=== Compiling HTML Report ===" -ForegroundColor Cyan
try {
    & "$rootPath/common/Generate-GRCReport.ps1"
} catch {
    Write-Error "Failed to compile GRC HTML Report: $_"
    return
}

# 4. Open the generated report in the default browser
$reportFile = Join-Path -Path $rootPath -ChildPath "docs/index.html"
if (Test-Path $reportFile) {
    Write-Host "`nOpening GRC Audit Report in default browser..." -ForegroundColor Green
    Start-Process $reportFile
} else {
    Write-Error "Report file not found at: $reportFile"
}
