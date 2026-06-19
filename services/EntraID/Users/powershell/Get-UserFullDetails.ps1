#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Graph REST API for deep details of all Entra ID Users.
.DESCRIPTION
    Retrieves full user details including licenses (resolved SKU names), directory roles,
    managers, group memberships, and MFA/SSPR status. Exports results to JSON/CSV.
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

# Import our GRC Common module
$commonModulePath = Join-Path -Path $PSScriptRoot -ChildPath "../../../../common/GRC-M365-Common.psm1"
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
    Write-Error "Authentication failed: $_"
    return
}

# 2. Map SkuIds to SkuPartNumbers (friendly license names)
Write-Verbose "Querying subscribed SKUs for license resolution..."
$skuMap = @{}
try {
    $skuUri = "https://graph.microsoft.com/v1.0/subscribedSkus"
    $skuResponse = Invoke-MgGraphRequest -Method GET -Uri $skuUri -ErrorAction SilentlyContinue
    if ($skuResponse -and $skuResponse.value) {
        foreach ($sku in $skuResponse.value) {
            if ($sku.skuId -and $sku.skuPartNumber) {
                $skuMap[$sku.skuId.ToString()] = $sku.skuPartNumber
            }
        }
    }
} catch {
    Write-Verbose "Failed to map subscribed SKUs: $_"
}

# 3. Retrieve User MFA/SSPR Registration status
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
    Write-Verbose "Could not retrieve user registration details: $_"
}

# 4. Map Directory Roles to Users
$userRoles = @{}
try {
    Write-Verbose "Querying directory roles and members..."
    $rolesUri = "https://graph.microsoft.com/v1.0/directoryRoles?`$expand=members(`$select=id)"
    $rolesResponse = Invoke-MgGraphRequest -Method GET -Uri $rolesUri -ErrorAction SilentlyContinue
    if ($rolesResponse -and $rolesResponse.value) {
        foreach ($role in $rolesResponse.value) {
            $roleName = $role.displayName
            if ($role.members) {
                foreach ($member in $role.members) {
                    $mId = $member.id
                    if (-not $userRoles.ContainsKey($mId)) {
                        $userRoles[$mId] = [System.Collections.Generic.List[string]]::new()
                    }
                    $userRoles[$mId].Add($roleName)
                }
            }
        }
    }
} catch {
    Write-Verbose "Failed to query directory roles: $_"
}

# 5. Map Group Memberships to Users (Group-to-User lookup to avoid N+1 queries)
$userGroups = @{}
try {
    Write-Verbose "Querying group memberships..."
    $groupsUri = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName&`$expand=members(`$select=id)"
    while ($groupsUri) {
        $groupsResponse = Invoke-MgGraphRequest -Method GET -Uri $groupsUri -ErrorAction Stop
        if ($groupsResponse -and $groupsResponse.value) {
            foreach ($grp in $groupsResponse.value) {
                $grpName = $grp.displayName
                if ($grp.members) {
                    foreach ($m in $grp.members) {
                        $mId = $m.id
                        if (-not $userGroups.ContainsKey($mId)) {
                            $userGroups[$mId] = [System.Collections.Generic.List[string]]::new()
                        }
                        $userGroups[$mId].Add($grpName)
                    }
                }
            }
            $groupsUri = $groupsResponse.'@odata.nextLink'
        } else {
            $groupsUri = $null
        }
    }
} catch {
    Write-Verbose "Failed to map group memberships: $_"
}

# 6. Retrieve Users with expanded manager details and assigned licenses
Write-Verbose "Querying detailed user list from Graph..."
$usersList = [System.Collections.Generic.List[PSCustomObject]]::new()
$userUri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,mail,accountEnabled,createdDateTime,userType,onPremisesSyncEnabled,jobTitle,department,assignedLicenses&`$expand=manager(`$select=displayName,userPrincipalName)"

try {
    while ($userUri) {
        $userResponse = Invoke-MgGraphRequest -Method GET -Uri $userUri -ErrorAction Stop
        if ($userResponse -and $userResponse.value) {
            foreach ($user in $userResponse.value) {
                $upn = $user.userPrincipalName
                $uId = $user.id
                
                # Resolve licenses
                $licenses = @()
                if ($user.assignedLicenses) {
                    foreach ($lic in $user.assignedLicenses) {
                        $skuId = $lic.skuId.ToString()
                        if ($skuMap.ContainsKey($skuId)) {
                            $licenses += $skuMap[$skuId]
                        } else {
                            $licenses += $skuId
                        }
                    }
                }

                # MFA/SSPR Info
                $mfaInfo = if ($mfaDetails.ContainsKey($upn)) { $mfaDetails[$upn] } else { $null }

                # Directory Roles
                $roles = if ($userRoles.ContainsKey($uId)) { $userRoles[$uId].ToArray() } else { @() }

                # Group memberships
                $groups = if ($userGroups.ContainsKey($uId)) { $userGroups[$uId].ToArray() } else { @() }

                # Manager UPN/Name
                $managerName = ""
                $managerUpn = ""
                if ($user.manager) {
                    $managerName = $user.manager.displayName
                    $managerUpn = $user.manager.userPrincipalName
                }

                $userObj = [PSCustomObject]@{
                    Id                     = $uId
                    DisplayName            = $user.displayName
                    UserPrincipalName      = $upn
                    Mail                   = $user.mail
                    AccountEnabled         = $user.accountEnabled
                    CreatedDateTime        = $user.createdDateTime
                    UserType               = $user.userType
                    OnPremisesSyncEnabled  = if ($null -ne $user.onPremisesSyncEnabled) { $user.onPremisesSyncEnabled } else { $false }
                    JobTitle               = $user.jobTitle
                    Department             = $user.department
                    AssignedLicenses       = $licenses
                    LicenseCount           = $licenses.Count
                    DirectoryRoles         = $roles
                    GroupMemberships       = $groups
                    ManagerName            = $managerName
                    ManagerUserPrincipalName = $managerUpn
                    IsMfaRegistered        = if ($null -ne $mfaInfo) { $mfaInfo.isMfaRegistered } else { 'Unknown' }
                    IsSsprRegistered       = if ($null -ne $mfaInfo) { $mfaInfo.isSsprRegistered } else { 'Unknown' }
                    IsSsprEnabled          = if ($null -ne $mfaInfo) { $mfaInfo.isSsprEnabled } else { 'Unknown' }
                    IsSsprCapable          = if ($null -ne $mfaInfo) { $mfaInfo.isSsprCapable } else { 'Unknown' }
                    IsMfaCapable           = if ($null -ne $mfaInfo) { $mfaInfo.isMfaCapable } else { 'Unknown' }
                }
                $usersList.Add($userObj)
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

# 7. Convert Lists to flat strings for CSV export (if we need to structure outputs)
$csvFormattedList = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($u in $usersList) {
    $csvObj = $u.PSObject.Copy()
    $csvObj.AssignedLicenses = $u.AssignedLicenses -join '; '
    $csvObj.DirectoryRoles = $u.DirectoryRoles -join '; '
    $csvObj.GroupMemberships = $u.GroupMemberships -join '; '
    $csvFormattedList.Add($csvObj)
}

# 8. Handle Outputs
if ($AiAgentMode) {
    $usersList.ToArray() | ConvertTo-Json -Depth 5
} else {
    # Custom save to match SensitivityLabelsFullDetails pattern
    $exportDir = Join-Path -Path $PSScriptRoot -ChildPath "../../../../exports/EntraID/Users/UsersFullDetails"
    if (!(Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $jsonPath = Join-Path -Path $exportDir -ChildPath "UsersFullDetails_${timestamp}.json"
    $csvPath  = Join-Path -Path $exportDir -ChildPath "UsersFullDetails_${timestamp}.csv"

    $usersList.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding utf8
    $csvFormattedList.ToArray() | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

    Write-Host "Users Full Details JSON written to: $jsonPath" -ForegroundColor Green
    Write-Host "Users Full Details CSV written to: $csvPath" -ForegroundColor Green
}
