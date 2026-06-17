#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Graph REST API for Entra ID User details and summary statistics.
.DESCRIPTION
    A cloud-native GRC Asset script designed to retrieve Entra ID user list details
    utilizing clean REST calls, including licenses, account state, and correlates with
    MFA registration status when permissions allow. Exports structured GRC results (JSON, CSV).

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
    https://learn.microsoft.com/en-us/graph/api/resources/user
    https://learn.microsoft.com/en-us/graph/api/reportroot-list-credentialuserregistrationdetails
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

# 2. Retrieve MFA and authentication registration status (correlative audit)
Write-Verbose "Querying Microsoft Graph REST endpoint /v1.0/reports/authenticationMethods/userRegistrationDetails..."
$mfaDetails = @{}
try {
    $mfaUri = "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails"
    while ($mfaUri) {
        $mfaResponse = Invoke-MgGraphRequest -Method GET -Uri $mfaUri -ErrorAction Stop
        if ($mfaResponse -and $mfaResponse.value) {
            foreach ($item in $mfaResponse.value) {
                if ($item.userPrincipalName) {
                    $mfaDetails[$item.userPrincipalName] = $item
                }
            }
            $mfaUri = $mfaResponse.'@odata.nextLink'
        } else {
            $mfaUri = $null
        }
    }
} catch {
    Write-Verbose "Could not retrieve user registration details (requires Reports.Read.All or UserAuthenticationMethod.Read.All): $_"
}

# 3. Retrieve User profiles using native Graph REST calls with Pagination
Write-Verbose "Querying Microsoft Graph REST endpoint /v1.0/users..."
$users = [System.Collections.Generic.List[PSCustomObject]]::new()
$userUri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,accountEnabled,createdDateTime,userType,onPremisesSyncEnabled,jobTitle,department,assignedLicenses"

try {
    while ($userUri) {
        $userResponse = Invoke-MgGraphRequest -Method GET -Uri $userUri -ErrorAction Stop
        if ($userResponse -and $userResponse.value) {
            foreach ($user in $userResponse.value) {
                $upn = $user.userPrincipalName
                $mfaInfo = if ($mfaDetails.ContainsKey($upn)) { $mfaDetails[$upn] } else { $null }

                $userObj = [PSCustomObject]@{
                    Id                     = $user.id
                    DisplayName            = $user.displayName
                    UserPrincipalName      = $upn
                    AccountEnabled         = $user.accountEnabled
                    CreatedDateTime        = $user.createdDateTime
                    UserType               = $user.userType
                    OnPremisesSyncEnabled  = if ($null -ne $user.onPremisesSyncEnabled) { $user.onPremisesSyncEnabled } else { $false }
                    JobTitle               = $user.jobTitle
                    Department             = $user.department
                    LicenseCount           = if ($user.assignedLicenses) { $user.assignedLicenses.Count } else { 0 }
                    IsMfaRegistered        = if ($null -ne $mfaInfo) { $mfaInfo.isMfaRegistered } else { 'Unknown' }
                    IsSsprRegistered       = if ($null -ne $mfaInfo) { $mfaInfo.isSsprRegistered } else { 'Unknown' }
                    IsSsprEnabled          = if ($null -ne $mfaInfo) { $mfaInfo.isSsprEnabled } else { 'Unknown' }
                    IsSsprCapable          = if ($null -ne $mfaInfo) { $mfaInfo.isSsprCapable } else { 'Unknown' }
                    IsMfaCapable           = if ($null -ne $mfaInfo) { $mfaInfo.isMfaCapable } else { 'Unknown' }
                }
                $users.Add($userObj)
            }
            $userUri = $userResponse.'@odata.nextLink'
        } else {
            $userUri = $null
        }
    }
} catch {
    Write-Error "Failed to query users REST endpoint: $_"
    return
}

# 4. Handle Outputs based on execution scope
if ($AiAgentMode) {
    # AI Agents need silent stdout directly as JSON string for easy processing
    $users.ToArray() | ConvertTo-Json -Depth 5
} else {
    # Local admin and workflow environments get file exports
    Export-GRCAssetData -ServiceName "EntraID" -AssetName "Users" -Data ($users.ToArray())
}
