#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Exchange Online for mailbox counts, transport rules, spam/malware policy configuration.
.DESCRIPTION
    Retrieves key Exchange Online governance metrics for GRC auditing, using certificate or
    interactive authentication. Exports structured JSON/CSV data.
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

# 1. Establish connection to Exchange Online
try {
    if ($Interactive) {
        Connect-GRCExchange -Interactive
    } else {
        Connect-GRCExchange -TenantId $TenantId -ClientId $ClientId -CertificateBase64 $CertificateBase64
    }
} catch {
    Write-Error "Exchange Online authentication failed: $_"
    return
}

# 2. Query Exchange Online details
$reportData = [Ordered]@{
    TotalUserMailboxes       = 0
    TotalSharedMailboxes     = 0
    TotalTransportRules      = 0
    TransportRuleNames       = ""
    DkimEnabledDomainsCount  = 0
    AntimalwarePoliciesCount = 0
}

try {
    # Mailbox summary counts
    $mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue
    if ($mailboxes) {
        $reportData.TotalUserMailboxes = ($mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' }).Count
        $reportData.TotalSharedMailboxes = ($mailboxes | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' }).Count
    }

    # Transport rules (Mail Flow Rules)
    $rules = Get-TransportRule -ErrorAction SilentlyContinue
    if ($rules) {
        $reportData.TotalTransportRules = @($rules).Count
        $reportData.TransportRuleNames = ($rules | Select-Object -ExpandProperty Name) -join '; '
    }

    # DKIM configuration
    $dkim = Get-DkimSigningConfig -ErrorAction SilentlyContinue
    if ($dkim) {
        $reportData.DkimEnabledDomainsCount = ($dkim | Where-Object { $_.Enabled -eq $true }).Count
    }

    # Anti-malware policies
    $antiMalware = Get-MalwareFilterPolicy -ErrorAction SilentlyContinue
    if ($antiMalware) {
        $reportData.AntimalwarePoliciesCount = @($antiMalware).Count
    }

} catch {
    Write-Warning "Could not query all Exchange configuration endpoints: $_"
} finally {
    # Always disconnect from Exchange Online to clean up sessions
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}

# 3. Handle Outputs based on execution scope
$exportObj = [PSCustomObject]$reportData
if ($AiAgentMode) {
    $exportObj | ConvertTo-Json -Depth 5
} else {
    Export-GRCAssetData -ServiceName "ExchangeOnline" -AssetName "ExchangeSummary" -Data @($exportObj)
}
