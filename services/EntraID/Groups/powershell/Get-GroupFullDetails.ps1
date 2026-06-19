#Requires -Version 7.0
<#
.SYNOPSIS
    Queries Microsoft Graph REST API for deep details of all Entra ID Groups.
.DESCRIPTION
    Retrieves full group details including classifications, owner details, member count,
    and member details. Exports results to JSON/CSV.
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

# 2. Query Groups details using native Graph REST calls with Pagination
Write-Verbose "Querying Microsoft Graph REST endpoint /v1.0/groups..."
$groupsList = [System.Collections.Generic.List[PSCustomObject]]::new()
$groupUri = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,description,mailEnabled,securityEnabled,groupTypes,visibility,createdDateTime,onPremisesSyncEnabled"

try {
    while ($groupUri) {
        $groupResponse = Invoke-MgGraphRequest -Method GET -Uri $groupUri -ErrorAction Stop
        if ($groupResponse -and $groupResponse.value) {
            foreach ($group in $groupResponse.value) {
                $gId = $group.id
                
                # Determine Group classification
                $classification = "Security Group"
                if ($group.groupTypes -contains "Unified") {
                    $classification = "Microsoft 365 Group"
                } elseif ($group.mailEnabled -and $group.securityEnabled) {
                    $classification = "Mail-enabled Security Group"
                } elseif ($group.mailEnabled) {
                    $classification = "Distribution Group"
                }

                # Query Owners for this group
                $owners = @()
                try {
                    $ownerUri = "https://graph.microsoft.com/v1.0/groups/$gId/owners?`$select=displayName,userPrincipalName"
                    $ownerRes = Invoke-MgGraphRequest -Method GET -Uri $ownerUri -ErrorAction SilentlyContinue
                    if ($ownerRes -and $ownerRes.value) {
                        foreach ($owner in $ownerRes.value) {
                            $oName = if ($owner.userPrincipalName) { $owner.userPrincipalName } else { $owner.displayName }
                            if ($oName) { $owners += $oName }
                        }
                    }
                } catch {}

                # Query Members for this group
                $members = @()
                try {
                    $memberUri = "https://graph.microsoft.com/v1.0/groups/$gId/members?`$select=displayName,userPrincipalName"
                    while ($memberUri) {
                        $memberRes = Invoke-MgGraphRequest -Method GET -Uri $memberUri -ErrorAction Stop
                        if ($memberRes -and $memberRes.value) {
                            foreach ($member in $memberRes.value) {
                                $mName = if ($member.userPrincipalName) { $member.userPrincipalName } else { $member.displayName }
                                if ($mName) { $members += $mName }
                            }
                            $memberUri = $memberRes.'@odata.nextLink'
                        } else {
                            $memberUri = $null
                        }
                    }
                } catch {}

                $groupObj = [PSCustomObject]@{
                    Id                     = $gId
                    DisplayName            = $group.displayName
                    Description            = $group.description
                    MailEnabled            = $group.mailEnabled
                    SecurityEnabled        = $group.securityEnabled
                    GroupClassification    = $classification
                    Visibility             = $group.visibility
                    CreatedDateTime        = $group.createdDateTime
                    OnPremisesSyncEnabled  = if ($null -ne $group.onPremisesSyncEnabled) { $group.onPremisesSyncEnabled } else { $false }
                    Owners                 = $owners
                    MemberCount            = $members.Count
                    Members                = $members
                }
                $groupsList.Add($groupObj)
            }
            $groupUri = $groupResponse.'@odata.nextLink'
        } else {
            $groupUri = $null
        }
    }
} catch {
    Write-Error "Failed to query groups: $_"
    return
}

# 3. Convert Lists to flat strings for CSV export
$csvFormattedList = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($g in $groupsList) {
    $csvObj = $g.PSObject.Copy()
    $csvObj.Owners = $g.Owners -join '; '
    $csvObj.Members = $g.Members -join '; '
    $csvFormattedList.Add($csvObj)
}

# 4. Handle Outputs
if ($AiAgentMode) {
    $groupsList.ToArray() | ConvertTo-Json -Depth 5
} else {
    $exportDir = Join-Path -Path $PSScriptRoot -ChildPath "../../../../exports/EntraID/Groups/GroupsFullDetails"
    if (!(Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $jsonPath = Join-Path -Path $exportDir -ChildPath "GroupsFullDetails_${timestamp}.json"
    $csvPath  = Join-Path -Path $exportDir -ChildPath "GroupsFullDetails_${timestamp}.csv"

    $groupsList.ToArray() | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding utf8
    $csvFormattedList.ToArray() | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8

    Write-Host "Groups Full Details JSON written to: $jsonPath" -ForegroundColor Green
    Write-Host "Groups Full Details CSV written to: $csvPath" -ForegroundColor Green
}
