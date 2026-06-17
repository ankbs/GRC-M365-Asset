#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Graph REST API for Entra ID Group details and statistics.
.DESCRIPTION
    A cloud-native GRC Asset script designed to retrieve Entra ID security and Microsoft 365 groups
    utilizing clean REST calls. Exports structured GRC results (JSON, CSV).

    Supports:
      - PowerShell (interactive admin)
      - Workflows (GitHub Actions runners)
      - AI Agents (silent JSON stdout/file integration)

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

.LINK
    https://learn.microsoft.com/en-us/graph/api/resources/group
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
$commonModulePath = Join-Path -Path $PSScriptRoot -ChildPath "../../../../common/GRC-M365-Common.psm1"
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
    Write-Error "Authentication failed: $_"
    return
}

# 2. Query Groups details using native Graph REST calls with Pagination
Write-Verbose "Querying Microsoft Graph REST endpoint /v1.0/groups..."
$groups = [System.Collections.Generic.List[PSCustomObject]]::new()
$groupUri = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,mailEnabled,securityEnabled,groupTypes,visibility,createdDateTime,onPremisesSyncEnabled"

try {
    while ($groupUri) {
        $groupResponse = Invoke-MgGraphRequest -Method GET -Uri $groupUri -ErrorAction Stop
        if ($groupResponse -and $groupResponse.value) {
            foreach ($group in $groupResponse.value) {
                # Determine Group classification
                $classification = "Security Group"
                if ($group.groupTypes -contains "Unified") {
                    $classification = "Microsoft 365 Group"
                } elseif ($group.mailEnabled -and $group.securityEnabled) {
                    $classification = "Mail-enabled Security Group"
                } elseif ($group.mailEnabled) {
                    $classification = "Distribution Group"
                }

                $groupObj = [PSCustomObject]@{
                    Id                     = $group.id
                    DisplayName            = $group.displayName
                    MailEnabled            = $group.mailEnabled
                    SecurityEnabled        = $group.securityEnabled
                    GroupClassification    = $classification
                    Visibility             = $group.visibility
                    CreatedDateTime        = $group.createdDateTime
                    OnPremisesSyncEnabled  = if ($null -ne $group.onPremisesSyncEnabled) { $group.onPremisesSyncEnabled } else { $false }
                }
                $groups.Add($groupObj)
            }
            $groupUri = $groupResponse.'@odata.nextLink'
        } else {
            $groupUri = $null
        }
    }
} catch {
    Write-Error "Failed to query groups REST endpoint: $_"
    return
}

# 3. Handle Outputs based on execution scope
if ($AiAgentMode) {
    # AI Agents need silent stdout directly as JSON string for easy processing
    $groups.ToArray() | ConvertTo-Json -Depth 5
} else {
    # Local admin and workflow environments get file exports
    Export-GRCAssetData -ServiceName "EntraID" -AssetName "Groups" -Data ($groups.ToArray())
}
