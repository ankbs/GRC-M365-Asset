#Requires -Version 7.0
<#
.SYNOPSIS
    Configures M365 GRC Asset App Registration permissions, certificates, and Azure roles.
.DESCRIPTION
    A modular GRC setup bootstrapper to automatically provision Azure App Registrations,
    Microsoft Graph API permissions (delegated-like app scopes), and Exchange Online RBAC Roles.
    
    This script replaces the manual setup instructions by fully automating the steps described in
    "D:\_GRC_Agent\GRC-M365-Asset\Setup\Add-M365AssessmentPermissions.txt".

.PARAMETER TenantId
    The target Entra ID Tenant ID (e.g. 'contoso.onmicrosoft.com').
.PARAMETER CreateNew
    When specified, automatically bootstraps a new App Registration and self-signed certificate.
.PARAMETER AdminUpn
    The administrative user UPN to run the delegated session (App creation & EXO role group binding).
.PARAMETER GitHubPat
    The user's GitHub Personal Access Token (PAT) to upload credentials directly to the Fork secrets.
.PARAMETER GitHubOwner
    The GitHub repository owner of the fork.
.PARAMETER GitHubRepo
    The GitHub repository name of the fork (defaults to 'GRC-M365-Asset').
    
.LINK
    https://learn.microsoft.com/en-us/graph/api/resources/application
    https://learn.microsoft.com/en-us/exchange/permissions-in-exchange-online/role-groups
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [switch]$CreateNew,

    [Parameter(Mandatory = $true)]
    [string]$AdminUpn,

    [Parameter(Mandatory = $true)]
    [string]$GitHubPat,

    [Parameter(Mandatory = $true)]
    [string]$GitHubOwner,

    [Parameter(Mandatory = $false)]
    [string]$GitHubRepo = "GRC-M365-Asset",

    [Parameter(Mandatory = $false)]
    [string]$AppSuffix = "Reader"
)

$ErrorActionPreference = "Stop"

# Define target required Graph Permissions (aligned with Add-M365AssessmentPermissions.txt)
$requiredGraphPermissions = @(
    "Organization.Read.All",
    "Domain.Read.All",
    "Group.Read.All",
    "User.Read.All",
    "AuditLog.Read.All",
    "UserAuthenticationMethod.Read.All",
    "RoleManagement.Read.Directory",
    "Policy.Read.All",
    "Application.Read.All",
    "Directory.Read.All",
    "DeviceManagementManagedDevices.Read.All",
    "DeviceManagementConfiguration.Read.All",
    "DeviceManagementRBAC.Read.All",
    "DeviceManagementApps.Read.All",
    "SecurityEvents.Read.All",
    "SharePointTenantSettings.Read.All",
    "TeamSettings.Read.All",
    "TeamworkAppSettings.Read.All",
    "Team.ReadBasic.All",
    "TeamMember.Read.All",
    "Channel.ReadBasic.All",
    "Reports.Read.All",
    "Sites.Read.All",
    "EntitlementManagement.Read.All",
    "InformationProtectionPolicy.Read.All",
    "RecordsManagement.Read.All",
    "SensitivityLabels.Read.All"
)

# Define Exchange Online Role Groups
$requiredExoRoleGroups = @(
    "View-Only Organization Management",
    "Compliance Management"
)

# Define Entra ID Compliance Directory Roles
$requiredComplianceRoles = @(
    @{ Name = 'Compliance Administrator'; Id = '17315797-102d-40b4-93e0-432062caca18' },
    @{ Name = 'Security Reader';          Id = '5d6b6bb7-de71-4623-b4af-96380a352509' },
    @{ Name = 'Global Reader';            Id = 'f2ef992c-3afb-46b9-b7cf-a126ee74c451' }
)

Write-Host "=== M365 GRC Asset Installer & Permision Manager ===" -ForegroundColor Cyan

# 1. Install prerequisites (Common Graph & EXO modules)
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Applications", "ExchangeOnlineManagement")
foreach ($m in $requiredModules) {
    if (!(Get-Module -ListAvailable -Name $m)) {
        Write-Host "Installing module $m..." -ForegroundColor Yellow
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber
    }
}

# 2. Authenticate Delegated to Graph for creation and roles binding
Write-Host "Connecting to Microsoft Graph (delegated as $AdminUpn)..." -ForegroundColor Cyan
Connect-MgGraph -TenantId $TenantId -Scopes 'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'RoleManagement.ReadWrite.Directory', 'Directory.Read.All' -NoWelcome -UseDeviceAuthentication

# 3. Create self-signed certificate and register App Registration
$appDisplayName = "GRC-M365-Asset-$AppSuffix"
$app = Get-MgApplication -Filter "displayName eq '$appDisplayName'" -ErrorAction SilentlyContinue

if (!$app -and $CreateNew) {
    Write-Host "Generating self-signed certificate..." -ForegroundColor Cyan
    $cert = New-SelfSignedCertificate -Subject "CN=$appDisplayName" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature
    $CertificateThumbprint = $cert.Thumbprint
    
    # Export certificate data (with empty password to ensure private key is included)
    $certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, "")
    $certBase64 = [System.Convert]::ToBase64String($certBytes)
    $cerRawData = $cert.RawData
    
    $keyCredential = @{
        Type          = 'AsymmetricX509Cert'
        Usage         = 'Verify'
        Key           = $cerRawData
        DisplayName   = "CN=$appDisplayName"
        StartDateTime = $cert.NotBefore.ToUniversalTime().ToString('o')
        EndDateTime   = $cert.NotAfter.ToUniversalTime().ToString('o')
    }
    
    Write-Host "Creating App Registration '$appDisplayName' in Entra ID..." -ForegroundColor Cyan
    $app = New-MgApplication -DisplayName $appDisplayName -SignInAudience 'AzureADMyOrg' -KeyCredentials @($keyCredential)
    $sp = New-MgServicePrincipal -AppId $app.AppId
    Write-Host "App Registration created successfully. ClientID: $($app.AppId)" -ForegroundColor Green
} else {
    if ($app) {
        $sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'"
        Write-Host "Found existing App Registration '$appDisplayName' (ClientID: $($app.AppId))" -ForegroundColor Yellow
        
        # Export existing local certificate to ensure GitHub secrets are synced
        $certs = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Subject -eq "CN=$appDisplayName" }
        if ($certs.Count -gt 0) {
            $cert = $certs[0]
            $certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, "")
            $certBase64 = [System.Convert]::ToBase64String($certBytes)
            Write-Host "Found existing local certificate CN=$appDisplayName, preparing to sync secrets..." -ForegroundColor Cyan
        }
    } else {
        Write-Error "No existing app found and -CreateNew switch was not specified."
        return
    }
}

# 4. Assign Microsoft Graph Permissions (Application Permission Roles)
Write-Host "Assigning Graph API permissions..." -ForegroundColor Cyan
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$roleLookup = @{}
foreach ($r in $graphSp.AppRoles) { $roleLookup[$r.Value] = $r.Id }

$existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id | Where-Object { $_.ResourceId -eq $graphSp.Id } | Select-Object -ExpandProperty AppRoleId

foreach ($permName in $requiredGraphPermissions) {
    if ($roleLookup.ContainsKey($permName)) {
        $roleId = $roleLookup[$permName]
        if ($existingAssignments -notcontains $roleId) {
            Write-Host "Assigning Graph Permission: $permName" -ForegroundColor Cyan
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter @{
                PrincipalId = $sp.Id
                ResourceId  = $graphSp.Id
                AppRoleId   = $roleId
            } | Out-Null
        } else {
            Write-Host "Graph Permission already present: $permName" -ForegroundColor DarkGray
        }
    }
}

# 4b. Assign Exchange Online API Permission (Exchange.ManageAsApp)
Write-Host "Assigning Exchange Online API permissions..." -ForegroundColor Cyan
$exoSp = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'" -ErrorAction SilentlyContinue
if ($exoSp) {
    $exoRole = $exoSp.AppRoles | Where-Object { $_.Value -eq "Exchange.ManageAsApp" }
    if ($exoRole) {
        $existingExoAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id | Where-Object { $_.ResourceId -eq $exoSp.Id } | Select-Object -ExpandProperty AppRoleId
        if ($existingExoAssignments -notcontains $exoRole.Id) {
            Write-Host "Assigning Exchange Online Permission: Exchange.ManageAsApp" -ForegroundColor Cyan
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter @{
                PrincipalId = $sp.Id
                ResourceId  = $exoSp.Id
                AppRoleId   = $exoRole.Id
            } | Out-Null
        } else {
            Write-Host "Exchange Online Permission already present: Exchange.ManageAsApp" -ForegroundColor DarkGray
        }
    }
}

# 5. Assign Entra ID Compliance / Security Roles
Write-Host "Assigning Entra ID Directory Roles..." -ForegroundColor Cyan
foreach ($roleDef in $requiredComplianceRoles) {
    $roleName = $roleDef.Name
    $roleTemplateId = $roleDef.Id
    
    $dirRole = Get-MgDirectoryRole -Filter "roleTemplateId eq '$roleTemplateId'" -ErrorAction SilentlyContinue
    if (!$dirRole) {
        $dirRole = New-MgDirectoryRole -BodyParameter @{ roleTemplateId = $roleTemplateId }
    }
    
    $members = Get-MgDirectoryRoleMemberAsServicePrincipal -DirectoryRoleId $dirRole.Id | Select-Object -ExpandProperty Id
    if ($members -notcontains $sp.Id) {
        Write-Host "Assigning Entra ID Role: $roleName" -ForegroundColor Cyan
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId $dirRole.Id -BodyParameter @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($sp.Id)" } | Out-Null
    } else {
        Write-Host "Entra ID Role already present: $roleName" -ForegroundColor DarkGray
    }
}

# Disconnect Graph
Disconnect-MgGraph

# 6. Assign Exchange Online RBAC Role Groups
Write-Host "Connecting to Exchange Online (delegated as $AdminUpn)..." -ForegroundColor Cyan
Connect-ExchangeOnline -UserPrincipalName $AdminUpn -ShowBanner:$false

# Ensure the Service Principal is registered in Exchange Online (for CBA EXO role assignment)
$exoSp = Get-ServicePrincipal -Identity $app.AppId -ErrorAction SilentlyContinue
if (!$exoSp) {
    Write-Host "Registering Service Principal in Exchange Online..." -ForegroundColor Cyan
    New-ServicePrincipal -AppId $app.AppId -ObjectId $sp.Id -DisplayName $appDisplayName | Out-Null
    # Give directory synchronization a moment
    Start-Sleep -Seconds 5
}

foreach ($rg in $requiredExoRoleGroups) {
    $members = Get-RoleGroupMember -Identity $rg | Select-Object -ExpandProperty Name
    if ($members -notcontains $appDisplayName) {
        Write-Host "Adding $appDisplayName to EXO Role Group: $rg" -ForegroundColor Cyan
        Add-RoleGroupMember -Identity $rg -Member $appDisplayName
    } else {
        Write-Host "EXO Role Group membership already present: $rg" -ForegroundColor DarkGray
    }
}

Disconnect-ExchangeOnline -Confirm:$false

# 7. Push Secrets to GitHub Repository using customer PAT
if ($GitHubPat -and $certBase64) {
    Write-Host "Pushing PFX certificate as a GitHub secret..." -ForegroundColor Cyan
    $env:GH_TOKEN = $GitHubPat
    try {
        $null = gh secret set GRC_CERTIFICATE --body $certBase64 --repo "$GitHubOwner/$GitHubRepo"
        if ($LASTEXITCODE -ne 0) { throw "gh secret set GRC_CERTIFICATE failed." }
        $null = gh secret set GRC_CLIENT_ID --body $app.AppId --repo "$GitHubOwner/$GitHubRepo"
        if ($LASTEXITCODE -ne 0) { throw "gh secret set GRC_CLIENT_ID failed." }
        $null = gh secret set GRC_TENANT_ID --body $TenantId --repo "$GitHubOwner/$GitHubRepo"
        if ($LASTEXITCODE -ne 0) { throw "gh secret set GRC_TENANT_ID failed." }
        Write-Host "GitHub Secrets uploaded successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to upload GitHub Secrets: $_"
    } finally {
        $env:GH_TOKEN = $null
    }
}

Write-Host "=== Setup completed successfully! ===" -ForegroundColor Green
