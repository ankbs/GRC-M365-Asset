#Requires -Version 7.0
<#
.SYNOPSIS
    Cleans up and removes the GRC-M365-Asset App Registration, Service Principal, and local certificate.
.DESCRIPTION
    A cleanup script that connects to Microsoft Graph and Exchange Online to undo all configurations
    performed by Initialize-GRCEnvironment.ps1:
      1. Removes the Service Principal from Entra ID directory roles.
      2. Removes the Service Principal from Exchange Online role groups.
      3. Deletes the Service Principal and App Registration.
      4. Optionally removes the self-signed certificate from the local user certificate store.

.PARAMETER TenantId
    The target Entra ID Tenant ID.
.PARAMETER AdminUpn
    The administrative user UPN to run the delegated session for cleanup.
.PARAMETER AppSuffix
    The suffix used when creating the App (e.g. 'Cloud', 'Agent', 'Local'). Defaults to 'Reader'.
.PARAMETER RemoveLocalCertificate
    Switch to also delete the generated self-signed certificate from the local Cert:\CurrentUser\My store.

.LINK
    https://learn.microsoft.com/en-us/graph/api/application-delete
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$AdminUpn,

    [Parameter(Mandatory = $false)]
    [string]$AppSuffix = "Reader",

    [Parameter(Mandatory = $false)]
    [switch]$RemoveLocalCertificate
)

$ErrorActionPreference = "Stop"

$appDisplayName = "GRC-M365-Asset-$AppSuffix"

Write-Host "=== GRC-M365-Asset Environment Cleanup & Uninstallation ===" -ForegroundColor Yellow
Write-Host "Target App Registration: $appDisplayName" -ForegroundColor Yellow

# 1. Install/Load required modules
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Applications", "ExchangeOnlineManagement")
foreach ($m in $requiredModules) {
    if (!(Get-Module -ListAvailable -Name $m)) {
        Write-Host "Installing module $m..." -ForegroundColor Yellow
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber
    }
}

# 2. Connect to Microsoft Graph (Delegated)
Write-Host "Connecting to Microsoft Graph (delegated as $AdminUpn)..." -ForegroundColor Cyan
Connect-MgGraph -TenantId $TenantId -Scopes 'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'RoleManagement.ReadWrite.Directory', 'Directory.Read.All' -NoWelcome

# 3. Locate App and Service Principal
$app = Get-MgApplication -Filter "displayName eq '$appDisplayName'" -ErrorAction SilentlyContinue
if ($app) {
    $sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
    
    if ($sp) {
        # Remove Entra ID Directory Roles (assignments)
        Write-Host "Checking Entra ID Directory Role assignments..." -ForegroundColor Cyan
        # Listing active directory roles
        $dirRoles = Get-MgDirectoryRole
        foreach ($role in $dirRoles) {
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction SilentlyContinue
            if ($members -and $members.Id -contains $sp.Id) {
                Write-Host "Removing from Entra ID Role: $($role.DisplayName)" -ForegroundColor Yellow
                if ($PSCmdlet.ShouldProcess("Service Principal $($sp.DisplayName)", "Remove from Directory Role $($role.DisplayName)")) {
                    Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -DirectoryObjectId $sp.Id
                }
            }
        }
        
        # Delete Service Principal
        Write-Host "Deleting Service Principal (App ID: $($app.AppId))..." -ForegroundColor Yellow
        if ($PSCmdlet.ShouldProcess("Service Principal $($sp.DisplayName)", "Delete")) {
            Remove-MgServicePrincipal -ServicePrincipalId $sp.Id
        }
    }
    
    # Delete App Registration
    Write-Host "Deleting App Registration '$appDisplayName'..." -ForegroundColor Yellow
    if ($PSCmdlet.ShouldProcess("Application $appDisplayName", "Delete")) {
        Remove-MgApplication -ApplicationId $app.Id
    }
} else {
    Write-Host "No App Registration found with display name '$appDisplayName'." -ForegroundColor Gray
}

# Disconnect Graph
Disconnect-MgGraph

# 4. Connect to Exchange Online to clean up Role Groups
Write-Host "Connecting to Exchange Online (delegated as $AdminUpn)..." -ForegroundColor Cyan
Connect-ExchangeOnline -UserPrincipalName $AdminUpn -ShowBanner:$false

$requiredExoRoleGroups = @(
    "View-Only Organization Management",
    "Compliance Management"
)

foreach ($rg in $requiredExoRoleGroups) {
    try {
        $members = Get-RoleGroupMember -Identity $rg -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
        if ($members -contains $appDisplayName) {
            Write-Host "Removing $appDisplayName from Exchange Online Role Group: $rg" -ForegroundColor Yellow
            if ($PSCmdlet.ShouldProcess("Exchange Role Group $rg", "Remove member $appDisplayName")) {
                Remove-RoleGroupMember -Identity $rg -Member $appDisplayName -Confirm:$false
            }
        }
    } catch {
        Write-Verbose "Could not inspect EXO Role Group ${rg}: $_"
    }
}

Disconnect-ExchangeOnline -Confirm:$false

# 5. Clean up local certificate
if ($RemoveLocalCertificate) {
    Write-Host "Locating local self-signed certificate for CN=$appDisplayName..." -ForegroundColor Cyan
    $certs = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Subject -eq "CN=$appDisplayName" }
    foreach ($cert in $certs) {
        Write-Host "Removing certificate with Thumbprint: $($cert.Thumbprint) from Cert:\CurrentUser\My" -ForegroundColor Yellow
        if ($PSCmdlet.ShouldProcess("Certificate CN=$appDisplayName", "Delete from store")) {
            Remove-Item -Path $cert.PSPath -Force
        }
    }
}

Write-Host "=== Uninstallation and cleanup finished! ===" -ForegroundColor Green
