#Requires -Version 7.0
<#
.SYNOPSIS
    Compiles all collected M365 GRC Asset JSON files into an interactive, premium HTML report.
.DESCRIPTION
    Looks up the latest JSON files exported under exports/, compiles metrics, and generates
    a static, highly-styled responsive HTML dashboard saved to docs/index.html for deployment
    via GitHub Pages. Handles missing files gracefully.
#>

$ErrorActionPreference = 'Stop'

Write-Host "=== Starting GRC HTML Report Compilation ===" -ForegroundColor Cyan

# 1. Helper function to find the latest JSON file for a given asset path
function Get-LatestJsonData {
    param(
        [string]$Path
    )
    if (Test-Path $Path) {
        $files = Get-ChildItem -Path $Path -Filter "*.json" | Sort-Object LastWriteTime -Descending
        if ($files.Count -gt 0) {
            $latest = $files[0].FullName
            Write-Host "Found latest export: $($files[0].Name)" -ForegroundColor DarkGray
            return Get-Content -Raw -Path $latest | ConvertFrom-Json -AsHashTable
        }
    }
    Write-Host "No export found under $Path" -ForegroundColor DarkYellow
    return $null
}

# 2. Locate and load the latest data files
$exportsRoot = Join-Path -Path $PSScriptRoot -ChildPath "../exports"
$tenantInfo = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "EntraID/TenantInfo")
$users      = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "EntraID/Users")
$groups     = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "EntraID/Groups")
$entraDev   = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "Devices/EntraDevices")
$intuneDev  = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "Devices/IntuneDevices")
$defDev     = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "Devices/DefenderDevices")
$exchange   = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "ExchangeOnline/ExchangeSummary")
$sharepoint = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "SharePoint/SharePointSummary")
$teams      = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "Teams/TeamsSummary")
$governance = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "EntraID/EntraGovernanceSummary")
$purview    = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "Purview/PurviewSummary")

# Load Deep Details (Full Details) data files
$usersFull      = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "EntraID/Users/UsersFullDetails")
$groupsFull     = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "EntraID/Groups/GroupsFullDetails")
$devicesFull    = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "Devices/DeviceFullDetails")
$exchangeFull   = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "ExchangeOnline/ExchangeFullDetails")
$sharepointFull = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "SharePoint/SharePointFullDetails")
$teamsFull      = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "Teams/TeamsFullDetails")
$purviewFull          = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "Purview/SensitivityLabelsFullDetails")
$securityScoreSummary = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "SecurityScore/SecurityScoreSummary")
$securityScoreDetails = Get-LatestJsonData -Path (Join-Path -Path $exportsRoot -ChildPath "SecurityScore/SecurityScoreDetails")

# 3. Calculate Summary Metrics
$tenantName = if ($tenantInfo) { $tenantInfo.OrgDisplayName } else { "Microsoft 365 Tenant" }
$tenantIdVal = if ($tenantInfo) { $tenantInfo.TenantId } else { "N/A" }
$securityDefaults = if ($tenantInfo) { $tenantInfo.SecurityDefaultsEnabled } else { "N/A" }
$verifiedDomains = if ($tenantInfo) { $tenantInfo.VerifiedDomains } else { "N/A" }

# User statistics
$totalUsers = 0
$activeUsers = 0
$disabledUsers = 0
$mfaUsers = 0
$noMfaUsers = 0
$mfaUnknown = 0

if ($users) {
    $totalUsers = @($users).Count
    foreach ($u in $users) {
        if ($u.AccountEnabled -eq $true) { $activeUsers++ } else { $disabledUsers++ }
        if ($u.IsMfaRegistered -eq $true -or $u.IsMfaRegistered -eq "True") {
            $mfaUsers++
        } elseif ($u.IsMfaRegistered -eq $false -or $u.IsMfaRegistered -eq "False") {
            $noMfaUsers++
        } else {
            $mfaUnknown++
        }
    }
}

# Group statistics
$totalGroups = 0
$securityGroups = 0
$m365Groups = 0
if ($groups) {
    $totalGroups = @($groups).Count
    foreach ($g in $groups) {
        if ($g.GroupTypes -contains 'Unified') { $m365Groups++ } else { $securityGroups++ }
    }
}

# Device statistics
$totalEntraDevices = if ($entraDev) { @($entraDev).Count } else { 0 }
$totalIntuneDevices = if ($intuneDev) { @($intuneDev).Count } else { 0 }
$totalDefenderDevices = if ($defDev) { @($defDev).Count } else { 0 }

# Exchange statistics
$exchUserMailboxes = if ($exchange) { $exchange.TotalUserMailboxes } else { 0 }
$exchSharedMailboxes = if ($exchange) { $exchange.TotalSharedMailboxes } else { 0 }
$exchTransportRules = if ($exchange) { $exchange.TotalTransportRules } else { 0 }
$exchDkimDomains = if ($exchange) { $exchange.DkimEnabledDomainsCount } else { 0 }
$exchAntiMalware = if ($exchange) { $exchange.AntimalwarePoliciesCount } else { 0 }

# SharePoint statistics
$spSites = if ($sharepoint) { $sharepoint.TotalSharepointSites } else { 0 }
$spSharingMode = if ($sharepoint) { $sharepoint.ExternalSharingMode } else { "N/A" }
$spSharingCap = if ($sharepoint) { $sharepoint.FileSharingCapability } else { "N/A" }

# Teams statistics
$teamsCount = if ($teams) { $teams.TotalTeams } else { 0 }
$teamsPublic = if ($teams) { $teams.PublicTeamsCount } else { 0 }
$teamsPrivate = if ($teams) { $teams.PrivateTeamsCount } else { 0 }

# Governance statistics
$caCount = if ($governance) { $governance.TotalConditionalAccessPolicies } else { 0 }
$caEnabled = if ($governance) { $governance.EnabledCAPoliciesCount } else { 0 }
$caReportOnly = if ($governance) { $governance.ReportOnlyCAPoliciesCount } else { 0 }
$apCount = if ($governance) { $governance.TotalAccessPackages } else { 0 }
$arCount = if ($governance) { $governance.TotalAccessReviews } else { 0 }

# Purview statistics
$purviewLabels = if ($purview) { $purview.TotalSensitivityLabels } else { 0 }
$purviewLabelsNames = if ($purview -and $purview.SensitivityLabelNames) { $purview.SensitivityLabelNames } else { "Keine" }
$purviewCopilotBlocked = if ($purview -and $purview.TotalCopilotBlockedLabels) { $purview.TotalCopilotBlockedLabels } else { 0 }
$purviewCopilotBlockedNames = if ($purview -and $purview.CopilotBlockedLabelNames) { $purview.CopilotBlockedLabelNames } else { "Keine" }
$purviewDlp = if ($purview) { $purview.TotalDlpPolicies } else { 0 }
$purviewDlpNames = if ($purview -and $purview.DlpPolicyNames) { $purview.DlpPolicyNames } else { "Keine" }
$purviewRetention = if ($purview -and $purview.TotalRetentionLabels) { $purview.TotalRetentionLabels } else { 0 }

# Policy settings
$mandatoryLabeling = "Nein"
$defaultLabel = "Keins"
$justificationReq = "Nein"
if ($purview -and $purview.LabelPolicySettings) {
    $mandatoryLabeling = if ($purview.LabelPolicySettings.IsMandatory -eq $true -or $purview.LabelPolicySettings.IsMandatory -eq "True") { "Ja" } else { "Nein" }
    $defaultLabel = if ($purview.LabelPolicySettings.DefaultLabelId) { $purview.LabelPolicySettings.DefaultLabelId } else { "Keins" }
    $justificationReq = if ($purview.LabelPolicySettings.DowngradeJustificationRequired -eq $true -or $purview.LabelPolicySettings.DowngradeJustificationRequired -eq "True") { "Ja" } else { "Nein" }
}

# Helper function to generate premium HTML tables from data objects
function Get-GrcHtmlTable {
    param(
        $Data,
        [string[]]$Properties,
        [string[]]$Headers
    )
    if ($null -eq $Data -or @($Data).Count -eq 0) {
        return "<p style='color: var(--text-muted); padding: 1rem;'>Keine Daten verfügbar.</p>"
    }
    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.Append("<div class='table-wrapper'><table class='data-table'><thead><tr>")
    foreach ($h in $Headers) {
        $null = $sb.Append("<th data-i18n-hdr='$h'>$h</th>")
    }
    $null = $sb.Append("</tr></thead><tbody>")
    foreach ($row in $Data) {
        $null = $sb.Append("<tr>")
        foreach ($p in $Properties) {
            $val = $row.$p
            $valStr = ""
            if ($null -eq $val) {
                $valStr = "-"
            } elseif ($val -is [bool]) {
                $valStr = if ($val) { "<span class='badge success' data-i18n-badge='Ja'>Ja</span>" } else { "<span class='badge danger' data-i18n-badge='Nein'>Nein</span>" }
            } elseif ($val -is [array] -or $val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                # Format license SKU codes, directory roles, group memberships, channel lists, Site Collection Admins
                $cleanItems = @()
                foreach ($item in $val) {
                    if ($item) {
                        $cleanItems += $item.ToString().Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
                    }
                }
                $valStr = $cleanItems -join '; '
            } else {
                $valStr = $val.ToString().Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
            }
            $null = $sb.Append("<td>$valStr</td>")
        }
        $null = $sb.Append("</tr>")
    }
    $null = $sb.Append("</tbody></table></div>")
    return $sb.ToString()
}

# Dedicated function to render Purview labels in a hierarchical parent/child tree
function Get-GrcPurviewTable {
    param(
        $Data
    )
    if ($null -eq $Data -or @($Data).Count -eq 0) {
        return "<p style='color: var(--text-muted); padding: 1rem;'>Keine Daten verfügbar.</p>"
    }

    # Step 1: Index labels by Id and Name
    $labelsById = @{}
    $labelsByName = @{}
    foreach ($item in $Data) {
        if ($item.Id) { $labelsById[$item.Id.ToLower()] = $item }
        if ($item.DisplayName) { $labelsByName[$item.DisplayName.ToLower()] = $item }
    }

    # Step 2: Establish Parent-Child relationships
    $parentToChildren = @{}
    $rootLabels = [System.Collections.Generic.List[object]]::new()
    $childToParent = @{}

    foreach ($item in $Data) {
        $id = $item.Id.ToLower()
        $parentGuid = if ($item.ParentGuid) { $item.ParentGuid.ToLower() } else { "" }
        $parentId = ""

        # ParentGuid match
        if ($parentGuid -and $labelsById.ContainsKey($parentGuid)) {
            $parentId = $parentGuid
        } else {
            # Regex match for parenthetical sublabels, e.g. "Vertraulich (verschlüsselt) - 0010" -> parent "Vertraulich - 0010"
            if ($item.DisplayName -match "^(.+?)\s*\((.+?)\)\s*(-\s*\d+)?$") {
                $prefix = $Matches[1].Trim()
                $suffix = if ($Matches[3]) { $Matches[3].Trim() } else { "" }
                $parentName = if ($suffix) { "$prefix $suffix" } else { $prefix }
                $parentName2 = if ($suffix) { "$prefix - $($suffix.Replace('-','').Trim())" } else { $prefix }

                if ($labelsByName.ContainsKey($parentName.ToLower())) {
                    $parentId = $labelsByName[$parentName.ToLower()].Id.ToLower()
                } elseif ($labelsByName.ContainsKey($parentName2.ToLower())) {
                    $parentId = $labelsByName[$parentName2.ToLower()].Id.ToLower()
                }
            }
        }

        if ($parentId) {
            $childToParent[$id] = $parentId
            if (-not $parentToChildren.ContainsKey($parentId)) {
                $parentToChildren[$parentId] = [System.Collections.Generic.List[object]]::new()
            }
            $parentToChildren[$parentId].Add($item)
        }
    }

    # Root labels
    foreach ($item in $Data) {
        $id = $item.Id.ToLower()
        if (-not $childToParent.ContainsKey($id)) {
            $rootLabels.Add($item)
        }
    }

    # Sort roots by priority
    $sortedRoots = $rootLabels | Sort-Object Priority

    # Flatten the tree
    $flatList = [System.Collections.Generic.List[object]]::new()
    $traverse = {
        param($node, $level)
        $nodeWrapper = [ordered]@{
            Node = $node
            IndentLevel = $level
        }
        $flatList.Add($nodeWrapper)
        $id = $node.Id.ToLower()
        if ($parentToChildren.ContainsKey($id)) {
            $sortedChildren = $parentToChildren[$id] | Sort-Object Priority
            foreach ($child in $sortedChildren) {
                & $traverse $child ($level + 1)
            }
        }
    }

    foreach ($root in $sortedRoots) {
        & $traverse $root 0
    }

    # Generate HTML
    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.Append("<div class='table-wrapper'><table class='data-table'><thead><tr>")
    $headers = @("Labelname", "Copilot Sperre", "Scope Files", "Scope Emails", "Scope Sites", "Verschlüsselt", "Richtlinien")
    foreach ($h in $headers) {
        $null = $sb.Append("<th data-i18n-hdr='$h'>$h</th>")
    }
    $null = $sb.Append("</tr></thead><tbody>")

    $properties = @("BlockCopilot", "ScopeFiles", "ScopeEmails", "ScopeSites", "EncryptionEnabled", "PublishedPolicies")

    foreach ($rowWrapper in $flatList) {
        $row = $rowWrapper.Node
        $level = $rowWrapper.IndentLevel

        $rowStyle = ""
        if ($level -eq 0) {
            $rowStyle = "style='background-color: rgba(255,255,255,0.02); font-weight: 600; color: #ffffff;'"
        } else {
            $rowStyle = "style='color: var(--text-muted);'"
        }

        $null = $sb.Append("<tr $rowStyle>")

        # Name column with indentation and arrow
        $nameStr = $row.DisplayName.ToString().Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
        if ($level -gt 0) {
            $indentPx = $level * 20
            $nameStr = "<span style='padding-left: $($indentPx)px; color: var(--accent-secondary); font-size: 0.95em;'>↳ $nameStr</span>"
        }

        $null = $sb.Append("<td>$nameStr</td>")

        # Other columns
        foreach ($p in $properties) {
            $val = $row.$p
            $valStr = ""
            if ($null -eq $val) {
                $valStr = "-"
            } elseif ($val -is [bool] -or $val -eq "True" -or $val -eq "False") {
                $boolVal = ($val -eq $true -or $val -eq "True")
                $valStr = if ($boolVal) { "<span class='badge success' data-i18n-badge='Ja'>Ja</span>" } else { "<span class='badge danger' data-i18n-badge='Nein'>Nein</span>" }
            } elseif ($val -is [array] -or $val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                $cleanItems = @()
                foreach ($item in $val) {
                    if ($item) {
                        $cleanItems += $item.ToString().Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
                    }
                }
                $valStr = $cleanItems -join '; '
            } else {
                $valStr = $val.ToString().Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
            }
            $null = $sb.Append("<td>$valStr</td>")
        }
        $null = $sb.Append("</tr>")
    }

    $null = $sb.Append("</tbody></table></div>")
    return $sb.ToString()
}

# Dedicated function to render Microsoft Secure Score and grouped recommendations
function Get-GrcSecureScoreSectionHtml {
    param(
        $Summary,
        $Details
    )
    if ($null -eq $Summary -and ($null -eq $Details -or @($Details).Count -eq 0)) {
        return "<p style='color: var(--text-muted); padding: 1rem;'>Keine Secure Score Daten verfügbar.</p>"
    }

    $sb = [System.Text.StringBuilder]::new()

    # 1. Add Summary Overview Grid if Summary is available
    if ($Summary) {
        $pct = $Summary.Percentage
        $curr = $Summary.CurrentScore
        $max = $Summary.MaxScore
        $avg = $Summary.M365Average
        $diff = $curr - $avg
        $diffText = if ($diff -ge 0) { "$($diff) Pkt. über dem Durchschnitt" } else { "$([Math]::Abs($diff)) Pkt. unter dem Durchschnitt" }
        $diffClass = if ($diff -ge 0) { "success-text" } else { "danger-text" }
        
        # Donut Chart SVG math: circumference of r=53 is 333.01
        $dashArray = 333.01
        $dashOffset = [Math]::Round($dashArray * (1 - ($pct / 100)), 2)
        $donutColorClass = if ($pct -ge 80) { "donut-success" } elseif ($pct -ge 50) { "donut-warning" } else { "donut-danger" }

        $null = $sb.Append(@"
        <!-- Secure Score Overview Dash -->
        <div class="grid-2" style="margin-top: 1rem; margin-bottom: 2rem;">
            <!-- Left: Donut Chart -->
            <div class="card" style="display: flex; align-items: center; justify-content: center; gap: 2rem; padding: 2rem;">
                <div class="id-donut-chart">
                    <svg class="donut-chart" width="140" height="140" viewBox="0 0 130 130">
                        <circle class="donut-track" cx="65" cy="65" r="53" fill="none" stroke-width="10"/>
                        <circle class="donut-fill $donutColorClass" cx="65" cy="65" r="53" fill="none" stroke-width="10"
                                stroke-dasharray="$dashArray" stroke-dashoffset="$dashOffset" stroke-linecap="round" transform="rotate(-90 65 65)"/>
                        <text class="donut-text" x="65" y="65" text-anchor="middle" dominant-baseline="central">$pct%</text>
                    </svg>
                </div>
                <div style="flex: 1; min-width: 200px;">
                    <h3 style="font-size: 1.1rem; font-weight: 700; margin-bottom: 0.5rem; color: #ffffff;" data-i18n="sec_score_status">Secure Score Status</h3>
                    <p style="font-size: 0.9rem; color: var(--text-muted); line-height: 1.4;" data-i18n-score-text="$pct|$curr|$max">
                        Ihr aktueller Sicherheitsindex liegt bei <strong style="color: #ffffff;">$pct%</strong>. 
                        Das entspricht <strong style="color: #ffffff;">$curr von $max</strong> möglichen Punkten.
                    </p>
                </div>
            </div>

            <!-- Right: Comparisons & Metrics -->
            <div class="card" style="display: flex; flex-direction: column; justify-content: space-between; padding: 1.5rem;">
                <div class="id-donut-stack">
                    <div class="id-donut-item">
                        <span style="font-size: 1.25rem;">📊</span>
                        <div class="id-donut-info">
                            <div class="id-donut-title" data-i18n="global_average">M365 globaler Durchschnitt</div>
                            <div class="id-donut-detail">$avg Pkt. ($([Math]::Round(($avg / $max) * 100, 2))% von max. $max Pkt.)</div>
                        </div>
                    </div>
                    <div class="id-donut-item">
                        <span style="font-size: 1.25rem;">⚖️</span>
                        <div class="id-donut-info">
                            <div class="id-donut-title" data-i18n="deviation_average">Abweichung zum Durchschnitt</div>
                            <div class="id-donut-detail $diffClass" style="font-weight: 600;" data-i18n-diff="$diff">$diffText</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
"@
)
    }

    # 2. Group Details by Microsoft Product / Category
    if ($Details -and @($Details).Count -gt 0) {
        # Grouping dictionary
        $groups = [ordered]@{
            "🔑 Microsoft Entra ID (Identität & Zugriff)" = [System.Collections.Generic.List[object]]::new()
            "💻 Microsoft Intune & Geräteverwaltung"     = [System.Collections.Generic.List[object]]::new()
            "🛡️ Microsoft Defender (Bedrohungsschutz)"   = [System.Collections.Generic.List[object]]::new()
            "🔒 Microsoft Purview (Compliance & DLP)"    = [System.Collections.Generic.List[object]]::new()
            "📬 Exchange, SharePoint & Teams (Kollaboration)" = [System.Collections.Generic.List[object]]::new()
            "⚙️ Sonstige M365 Sicherheitsdienste"         = [System.Collections.Generic.List[object]]::new()
        }

        foreach ($rec in $Details) {
            $svc = $rec.Service
            $title = $rec.Title
            $targetList = $groups["⚙️ Sonstige M365 Sicherheitsdienste"]

            if ($svc -match "Azure Active Directory" -or $svc -match "Entra" -or $title -match "MFA" -or $title -match "Conditional Access") {
                $targetList = $groups["🔑 Microsoft Entra ID (Identität & Zugriff)"]
            } elseif ($svc -match "Intune" -or $svc -match "Device" -or $rec.Category -match "Device") {
                $targetList = $groups["💻 Microsoft Intune & Geräteverwaltung"]
            } elseif ($svc -match "Defender" -or $svc -match "Endpoint" -or $svc -match "Threat" -or $title -match "malware" -or $title -match "phishing") {
                $targetList = $groups["🛡️ Microsoft Defender (Bedrohungsschutz)"]
            } elseif ($svc -match "Purview" -or $svc -match "DLP" -or $svc -match "Compliance" -or $svc -match "Retention" -or $title -match "DLP") {
                $targetList = $groups["🔒 Microsoft Purview (Compliance & DLP)"]
            } elseif ($svc -match "Exchange" -or $svc -match "SharePoint" -or $svc -match "Teams" -or $svc -match "Yammer" -or $svc -match "Skype") {
                $targetList = $groups["📬 Exchange, SharePoint & Teams (Kollaboration)"]
            }

            $targetList.Add($rec)
        }

        # Render each group as a collapsible details section with a sub-table
        foreach ($groupName in $groups.Keys) {
            $list = $groups[$groupName]
            if ($list.Count -eq 0) { continue }

            $null = $sb.Append("<details class='collector-detail'><summary>$groupName ($($list.Count) Empfehlungen)</summary>")
            $null = $sb.Append("<div class='table-wrapper'><table class='data-table'>")
            $null = $sb.Append("<thead><tr><th data-i18n-hdr='Empfehlung'>Empfehlung</th><th data-i18n-hdr='Max. Punkte'>Max. Punkte</th><th data-i18n-hdr='Auswirkung'>Auswirkung</th><th data-i18n-hdr='Status'>Status</th><th data-i18n-hdr='Lizenzbedarf'>Lizenzbedarf</th><th data-i18n-hdr='Handlungsanweisung'>Handlungsanweisung</th></tr></thead><tbody>")

            foreach ($item in $list) {
                # Format status badge
                $status = $item.ImplementationStatus
                $badgeClass = "badge danger"
                $badgeText = "Nicht umgesetzt"
                if ($status -ieq "implemented" -or $status -ieq "completed") {
                    $badgeClass = "badge success"
                    $badgeText = "Umgesetzt"
                } elseif ($status -match "alternative" -or $status -match "planned" -or $status -match "thirdParty") {
                    $badgeClass = "badge warning"
                    $badgeText = "Teilweise / Alternativ"
                }

                $titleStr = $item.Title.ToString().Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
                $maxScoreStr = $item.MaxScore
                $impactStr = $item.UserImpact
                $licenseStr = $item.LicenseRequired.ToString().Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
                $remediationStr = if ($item.Remediation) { 
                    $item.Remediation.ToString().Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;') 
                } else { 
                    "-" 
                }

                $null = $sb.Append("<tr>")
                $null = $sb.Append("<td style='font-weight: 500;'>$titleStr</td>")
                $null = $sb.Append("<td>$maxScoreStr</td>")
                $null = $sb.Append("<td>$impactStr</td>")
                $null = $sb.Append("<td><span class='$badgeClass' data-i18n-badge='$badgeText'>$badgeText</span></td>")
                $null = $sb.Append("<td style='font-size: 0.85rem; color: #a5b4fc; font-weight: 500;'>$licenseStr</td>")
                $null = $sb.Append("<td style='font-size: 0.85rem; max-width: 320px; white-space: normal; line-height: 1.45; color: var(--text-muted);'>$remediationStr</td>")
                $null = $sb.Append("</tr>")
            }

            $null = $sb.Append("</tbody></table></div></details>")
        }
    }

    return $sb.ToString()
}

# Pre-generate detailed tables for HTML interpolation
$usersTableHtml      = Get-GrcHtmlTable -Data $usersFull -Properties @("DisplayName", "UserPrincipalName", "AccountEnabled", "IsMfaRegistered", "DirectoryRoles", "ManagerName") -Headers @("Name", "UPN", "Aktiv", "MFA", "Admin-Rollen", "Vorgesetzter")
$groupsTableHtml     = Get-GrcHtmlTable -Data $groupsFull -Properties @("DisplayName", "GroupClassification", "Visibility", "Owners", "MemberCount") -Headers @("Gruppenname", "Klassifizierung", "Sichtbarkeit", "Besitzer", "Mitglieder")
$devicesTableHtml    = Get-GrcHtmlTable -Data $devicesFull -Properties @("DisplayName", "OperatingSystem", "OperatingSystemVersion", "TrustType", "IsCompliant", "IntuneManaged", "DefenderStatus") -Headers @("Gerätename", "Betriebssystem", "OS-Version", "Trust-Typ", "Konform", "Intune MDM", "Defender Status")
$exchangeTableHtml   = Get-GrcHtmlTable -Data $exchangeFull -Properties @("MailboxAddress", "MailboxType", "ProhibitSendQuota", "FullAccessPermissions", "SendAsPermissions", "ForwardingAddress") -Headers @("Postfach-Adresse", "Typ", "Quota Limit", "Vollzugriff (Delegiert)", "Send As (Delegiert)", "Weiterleitung")
$sharepointTableHtml = Get-GrcHtmlTable -Data $sharepointFull -Properties @("DisplayName", "WebUrl", "SiteOwners", "StorageUsedBytes") -Headers @("Website-Name", "URL", "Administratoren", "Speicher (Bytes)")
$teamsTableHtml      = Get-GrcHtmlTable -Data $teamsFull -Properties @("DisplayName", "Visibility", "Owners", "MemberCount", "GuestCount", "Channels") -Headers @("Teamname", "Typ", "Besitzer", "Mitglieder", "Gäste", "Kanäle")
$purviewTableHtml    = Get-GrcPurviewTable -Data $purviewFull
$secureScoreHtml     = Get-GrcSecureScoreSectionHtml -Summary $securityScoreSummary -Details $securityScoreDetails

# 4. Generate HTML Content (Highly styled with Outfit typography, glassmorphism, responsive grid)
$htmlContent = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>M365 GRC Asset Audit Report - $tenantName</title>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        :root {
            --bg-color: #0b0f19;
            --card-bg: rgba(17, 24, 39, 0.7);
            --card-border: rgba(255, 255, 255, 0.08);
            --text-color: #f3f4f6;
            --text-muted: #9ca3af;
            --accent-primary: #6366f1;
            --accent-secondary: #06b6d4;
            --success: #10b981;
            --danger: #ef4444;
            --warning: #f59e0b;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Outfit', sans-serif;
        }

        body {
            background-color: var(--bg-color);
            background-image: 
                radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.12) 0px, transparent 50%),
                radial-gradient(at 100% 100%, rgba(6, 182, 212, 0.08) 0px, transparent 50%);
            color: var(--text-color);
            min-height: 100vh;
            padding: 2rem;
            line-height: 1.5;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
        }

        /* Header section */
        header {
            margin-bottom: 2.5rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 1rem;
            border-bottom: 1px solid var(--card-border);
            padding-bottom: 1.5rem;
        }

        .header-title h1 {
            font-size: 2rem;
            font-weight: 800;
            background: linear-gradient(135deg, #a5b4fc 0%, #6366f1 50%, #22d3ee 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .header-title p {
            color: var(--text-muted);
            margin-top: 0.25rem;
            font-size: 0.95rem;
        }

        .timestamp-badge {
            background: rgba(99, 102, 241, 0.15);
            border: 1px solid rgba(99, 102, 241, 0.3);
            color: #a5b4fc;
            padding: 0.5rem 1rem;
            border-radius: 9999px;
            font-size: 0.85rem;
            font-weight: 600;
        }

        /* Responsive Grid */
        .grid-3 {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 1.5rem;
            margin-bottom: 2rem;
        }

        .grid-2 {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(480px, 1fr));
            gap: 1.5rem;
            margin-bottom: 2rem;
        }

        /* Premium Cards */
        .card {
            background: var(--card-bg);
            border: 1px solid var(--card-border);
            border-radius: 16px;
            padding: 1.5rem;
            backdrop-filter: blur(12px);
            transition: transform 0.2s, box-shadow 0.2s;
        }

        .card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 30px rgba(0, 0, 0, 0.4);
            border-color: rgba(99, 102, 241, 0.25);
        }

        .card h2 {
            font-size: 1.25rem;
            font-weight: 700;
            margin-bottom: 1.25rem;
            color: #ffffff;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            border-bottom: 1px solid rgba(255,255,255,0.05);
            padding-bottom: 0.5rem;
        }

        /* Metric Lists */
        .metric-row {
            display: flex;
            justify-content: space-between;
            padding: 0.75rem 0;
            border-bottom: 1px solid rgba(255,255,255,0.03);
            font-size: 0.95rem;
        }

        .metric-row:last-child {
            border-bottom: none;
        }

        .metric-label {
            color: var(--text-muted);
            font-weight: 500;
        }

        .metric-value {
            font-weight: 600;
            color: #ffffff;
        }

        .metric-value.success { color: var(--success); }
        .metric-value.danger { color: var(--danger); }
        .metric-value.warning { color: var(--warning); }

        /* Summary statistics bar */
        .stats-summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
            gap: 1rem;
            margin-bottom: 2rem;
        }

        .stat-box {
            background: rgba(255,255,255,0.02);
            border: 1px solid var(--card-border);
            border-radius: 12px;
            padding: 1rem;
            text-align: center;
        }

        .stat-box .num {
            font-size: 1.75rem;
            font-weight: 800;
            color: var(--accent-secondary);
        }

        .stat-box .lbl {
            font-size: 0.75rem;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.05em;
            margin-top: 0.25rem;
        }

        /* Chart Canvas wrapper */
        .chart-wrapper {
            position: relative;
            height: 220px;
            width: 100%;
            display: flex;
            justify-content: center;
            align-items: center;
        }

        /* Footer */
        footer {
            text-align: center;
            color: var(--text-muted);
            font-size: 0.85rem;
            margin-top: 4rem;
            padding-top: 1.5rem;
            border-top: 1px solid var(--card-border);
        }

        /* Collapsible details styling */
        .collector-detail {
            background: rgba(255, 255, 255, 0.02);
            border: 1px solid var(--card-border);
            border-radius: 12px;
            margin-top: 1.5rem;
            margin-bottom: 1.5rem;
            backdrop-filter: blur(12px);
            overflow: hidden;
        }

        .collector-detail > summary {
            cursor: pointer;
            padding: 1rem;
            font-weight: 600;
            color: #ffffff;
            list-style: none;
            display: flex;
            justify-content: space-between;
            align-items: center;
            user-select: none;
            transition: background-color 0.2s;
        }

        .collector-detail > summary:hover {
            background-color: rgba(255, 255, 255, 0.04);
        }

        .collector-detail > summary::-webkit-details-marker {
            display: none;
        }

        .collector-detail > summary::after {
            content: '▼';
            font-size: 0.8rem;
            color: var(--text-muted);
            transition: transform 0.2s;
        }

        .collector-detail[open] > summary::after {
            transform: rotate(180deg);
        }

        .collector-detail[open] > summary {
            border-bottom: 1px solid var(--card-border);
            background-color: rgba(255, 255, 255, 0.03);
        }

        /* Detail Tables styling */
        .table-wrapper {
            overflow-x: auto;
            max-height: 450px;
            overflow-y: auto;
        }

        .data-table {
            width: 100%;
            border-collapse: collapse;
            text-align: left;
            font-size: 0.9rem;
        }

        .data-table th, .data-table td {
            padding: 0.75rem 1rem;
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
        }

        .data-table th {
            background-color: rgba(17, 24, 39, 0.9);
            color: #ffffff;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 2;
        }

        .data-table tbody tr:hover {
            background-color: rgba(255, 255, 255, 0.02);
        }

        /* Badges */
        .badge {
            display: inline-block;
            padding: 0.15rem 0.5rem;
            border-radius: 4px;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: uppercase;
        }

        .badge.success {
            background-color: rgba(16, 185, 129, 0.15);
            color: #34d399;
            border: 1px solid rgba(16, 185, 129, 0.3);
        }

        .badge.danger {
            background-color: rgba(239, 68, 68, 0.15);
            color: #f87171;
            border: 1px solid rgba(239, 68, 68, 0.3);
        }

        .badge.warning {
            background-color: rgba(245, 158, 11, 0.15);
            color: #fbbf24;
            border: 1px solid rgba(245, 158, 11, 0.3);
        }

        /* Secure Score Donut & Layout Styles */
        .id-donut-stack {
            display: flex;
            flex-direction: column;
            gap: 1rem;
        }
        .id-donut-item {
            display: flex;
            align-items: center;
            gap: 1rem;
            padding: 0.75rem 1rem;
            background: rgba(255, 255, 255, 0.02);
            border-radius: 8px;
            border: 1px solid var(--card-border);
        }
        .id-donut-chart { flex-shrink: 0; }
        .id-donut-info { min-width: 0; }
        .id-donut-title {
            font-size: 0.95rem;
            font-weight: 600;
            color: #ffffff;
        }
        .id-donut-detail {
            font-size: 0.85rem;
            color: var(--text-muted);
            margin-top: 0.25rem;
        }
        .donut-chart { display: block; margin: 0 auto; }
        .donut-track { stroke: rgba(255, 255, 255, 0.08); }
        .donut-fill { transition: stroke-dashoffset 0.6s ease; }
        .donut-success { stroke: var(--success); }
        .donut-warning { stroke: var(--warning); }
        .donut-danger { stroke: var(--danger); }
        .donut-text { font-size: 20px; font-weight: 700; fill: #ffffff; font-family: inherit; }
        .success-text { color: var(--success) !important; }
        .danger-text { color: var(--danger) !important; }
        .warning-text { color: var(--warning) !important; }

        /* Language Selector styling */
        .lang-selector {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--card-border);
            padding: 0.35rem 0.75rem;
            border-radius: 8px;
            font-size: 0.85rem;
            color: var(--text-color);
        }
        .lang-selector select {
            background: transparent;
            border: none;
            color: #ffffff;
            font-weight: 600;
            cursor: pointer;
            outline: none;
            font-family: inherit;
        }
        .lang-selector select option {
            background: #0b0f19;
            color: #ffffff;
        }

        @media(max-width: 600px) {
            body { padding: 1rem; }
            .grid-2 { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="header-title">
                <h1 data-i18n="report_title">M365 GRC Asset Audit Report</h1>
                <p><span data-i18n="lbl_tenant">Mandant</span>: <strong>$tenantName</strong> ($tenantIdVal)</p>
            </div>
            <div style="display: flex; align-items: center; gap: 1rem;">
                <div class="lang-selector">
                    🌐 <select id="languageSelect" onchange="switchLanguage(this.value)">
                        <option value="de" selected>Deutsch</option>
                        <option value="en">English</option>
                    </select>
                </div>
                <div class="timestamp-badge">
                    📅 <span data-i18n="lbl_generated_at">Generiert am</span>: $(Get-Date -Format "dd.MM.yyyy HH:mm")
                </div>
            </div>
        </header>

        <!-- Global Count Summaries -->
        <div class="stats-summary">
            <div class="stat-box">
                <div class="num">$totalUsers</div>
                <div class="lbl" data-i18n="stat_users_total">Benutzer gesamt</div>
            </div>
            <div class="stat-box">
                <div class="num">$activeUsers</div>
                <div class="lbl" data-i18n="stat_active_accounts">Aktive Konten</div>
            </div>
            <div class="stat-box">
                <div class="num">$totalGroups</div>
                <div class="lbl" data-i18n="stat_groups_total">Gruppen gesamt</div>
            </div>
            <div class="stat-box">
                <div class="num">$totalEntraDevices</div>
                <div class="lbl" data-i18n="stat_entra_devices">Entra ID Devices</div>
            </div>
            <div class="stat-box">
                <div class="num">$totalIntuneDevices</div>
                <div class="lbl" data-i18n="stat_intune_managed">Intune Managed</div>
            </div>
            <div class="stat-box">
                <div class="num">$totalDefenderDevices</div>
                <div class="lbl" data-i18n="stat_defender_endpoints">Defender Endpoints</div>
            </div>
        </div>

        <div class="grid-3">
            <!-- Tenant Metadata Card -->
            <div class="card">
                <h2 data-i18n="card_tenant_details">🏢 Mandanten-Details</h2>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_name">Name</span>
                    <span class="metric-value">$tenantName</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_tenant_id">Tenant ID</span>
                    <span class="metric-value" style="font-size: 0.8rem; font-family: monospace;">$tenantIdVal</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_security_defaults">Identity Security Defaults</span>
                    <span class="metric-value $(if ($securityDefaults -eq $true) { 'success' } else { 'danger' })">
                        $(if ($securityDefaults -eq $true) { '<span data-i18n="lbl_security_defaults_active">Aktiviert</span>' } elseif ($securityDefaults -eq $false) { '<span data-i18n="lbl_security_defaults_inactive">Deaktiviert</span>' } else { 'N/A' })
                    </span>
                </div>
                <div class="metric-row" style="flex-direction: column; gap: 0.25rem;">
                    <span class="metric-label" data-i18n="lbl_verified_domains">Verifizierte Domains:</span>
                    <span class="metric-value" style="font-size: 0.8rem; color: var(--text-muted); word-break: break-all; margin-top: 0.25rem;">$verifiedDomains</span>
                </div>
            </div>

            <!-- Identity GRC Card -->
            <div class="card">
                <h2 data-i18n="card_mfa_sec">MFA Absicherung</h2>
                <div class="chart-wrapper">
                    <canvas id="mfaChart"></canvas>
                </div>
            </div>

            <!-- Groups Audit Card -->
            <div class="card">
                <h2 data-i18n="card_groups_struct">👥 Gruppen-Struktur</h2>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_m365_groups">M365 / Unified Groups</span>
                    <span class="metric-value">$m365Groups</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_sec_groups">Sicherheitsgruppen</span>
                    <span class="metric-value">$securityGroups</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_total_groups">Gruppen gesamt</span>
                    <span class="metric-value">$totalGroups</span>
                </div>
            </div>
        </div>

        <details class="collector-detail">
            <summary data-i18n="sum_users_details">👤 Benutzer-Details (Entra ID Users)</summary>
            $usersTableHtml
        </details>
        <details class="collector-detail">
            <summary data-i18n="sum_groups_details">👥 Gruppen-Details (Entra ID Groups)</summary>
            $groupsTableHtml
        </details>

        <div class="grid-2">
            <!-- User Status Breakdown Card -->
            <div class="card">
                <h2 data-i18n="lbl_user_status">👤 Benutzerkonten-Status</h2>
                <div class="chart-wrapper">
                    <canvas id="userStatusChart"></canvas>
                </div>
            </div>

            <!-- Device GRC Overlap Card -->
            <div class="card">
                <h2 data-i18n="lbl_hardware_audit">💻 Endgeräte GRC-Audit</h2>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_registered_joined">In Entra ID registriert/joined (Hardware Asset)</span>
                    <span class="metric-value">$totalEntraDevices</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_managed_intune">In Intune verwaltet (Compliance erzwingbar)</span>
                    <span class="metric-value" style="color: var(--accent-secondary);">$totalIntuneDevices</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_enrolled_defender">In Defender for Endpoint erfasst (EDR Abdeckung)</span>
                    <span class="metric-value" style="color: var(--success);">$totalDefenderDevices</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_intune_coverage">Intune-Abdeckung (relativ to Entra ID)</span>
                    <span class="metric-value">
                        $(if ($totalEntraDevices -gt 0) { "{0:P1}" -f ($totalIntuneDevices / $totalEntraDevices) } else { "0%" })
                    </span>
                </div>
            </div>
        </div>

        <details class="collector-detail">
            <summary data-i18n="sum_devices_details">💻 Geräte-Details (Hardware & Compliance)</summary>
            $devicesTableHtml
        </details>

        <!-- Microsoft Secure Score & Recommendations -->
        <h2 style="margin-top: 2.5rem; margin-bottom: 1rem; font-size: 1.5rem; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 0.5rem; color: #ffffff;"><span data-i18n="lbl_sec_score_recommendations">🎯 Microsoft Secure Score & Handlungsempfehlungen</span></h2>
        $secureScoreHtml

        <!-- M365 Collaboration & Mail GRC Row -->
        <h2 style="margin-top: 2.5rem; margin-bottom: 1rem; font-size: 1.5rem; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 0.5rem; color: #ffffff;"><span data-i18n="lbl_collab_mail_audit">📬 Kollaboration & E-Mail GRC-Audit</span></h2>
        <div class="grid-3">
            <!-- Exchange Online Card -->
            <div class="card">
                <h2 data-i18n="sum_exchange_details">✉️ Exchange Online</h2>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_user_mailboxes">Benutzer-Postfächer</span>
                    <span class="metric-value">$exchUserMailboxes</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_shared_mailboxes">Gemeinsame Postfächer (Shared)</span>
                    <span class="metric-value">$exchSharedMailboxes</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_transport_rules">Mailflow Transport-Regeln</span>
                    <span class="metric-value warning">$exchTransportRules</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_dkim_domains">DKIM-geschützte Domains</span>
                    <span class="metric-value success">$exchDkimDomains</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_antimalware_policies">Anti-Malware Richtlinien</span>
                    <span class="metric-value">$exchAntiMalware</span>
                </div>
            </div>

            <!-- SharePoint & OneDrive Card -->
            <div class="card">
                <h2 data-i18n="sum_sharepoint_details">🌐 SharePoint & OneDrive</h2>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_sp_sites">Aktive SharePoint-Sites</span>
                    <span class="metric-value">$spSites</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_sharing_policy">Externer Freigabe-Modus</span>
                    <span class="metric-value" style="font-size: 0.85rem; font-family: monospace;">$spSharingMode</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_sharing_capabilities">Freigabe-Berechtigung</span>
                    <span class="metric-value" style="font-size: 0.85rem; font-family: monospace;">$spSharingCap</span>
                </div>
            </div>

            <!-- Microsoft Teams Card -->
            <div class="card">
                <h2 data-i18n="sum_teams_details">💬 Microsoft Teams</h2>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_teams_total">Teams gesamt</span>
                    <span class="metric-value">$teamsCount</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_public_teams">Öffentliche Teams (Risiko)</span>
                    <span class="metric-value $(if ($teamsPublic -gt 0) { 'warning' } else { '' })">$teamsPublic</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_private_teams">Private Teams</span>
                    <span class="metric-value success">$teamsPrivate</span>
                </div>
            </div>
        </div>

        <details class="collector-detail">
            <summary data-i18n="sum_exchange_table">✉️ Exchange Online Postfach-Details</summary>
            $exchangeTableHtml
        </details>
        <details class="collector-detail">
            <summary data-i18n="sum_sharepoint_table">🌐 SharePoint Online Website-Details</summary>
            $sharepointTableHtml
        </details>
        <details class="collector-detail">
            <summary data-i18n="sum_teams_table">👥 Teams-Details (Kanallisten & Gäste)</summary>
            $teamsTableHtml
        </details>

        <!-- M365 Governance & Purview GRC Row -->
        <h2 style="margin-top: 2.5rem; margin-bottom: 1rem; font-size: 1.5rem; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 0.5rem; color: #ffffff;"><span data-i18n="lbl_governance_compliance_audit">🛡️ Identity Governance & Compliance GRC-Audit</span></h2>
        <div class="grid-2">
            <!-- Entra ID Governance Card -->
            <div class="card">
                <h2 data-i18n="lbl_entra_governance">🔑 Identity Governance (Entra ID)</h2>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_ca_policies">Bedingter Zugriff (CA-Richtlinien)</span>
                    <span class="metric-value">$caCount <span data-i18n="lbl_policies">Richtlinien</span></span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_ca_enabled">CA Richtlinien Aktiviert</span>
                    <span class="metric-value success">$caEnabled</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_ca_report_only">CA Richtlinien im Report-only Modus</span>
                    <span class="metric-value warning">$caReportOnly</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_access_packages">Zugriffspakete (Access Packages)</span>
                    <span class="metric-value">$apCount</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_access_reviews">Zugriffsüberprüfungen (Access Reviews)</span>
                    <span class="metric-value">$arCount</span>
                </div>
            </div>

            <!-- Purview Information Protection Card -->
            <div class="card">
                <h2 data-i18n="lbl_purview_compliance">🔒 Microsoft Purview Compliance & Information Protection</h2>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_sensitivity_labels">Vertraulichkeitslabels (Sensitivity Labels)</span>
                    <span class="metric-value">$purviewLabels</span>
                </div>
                <div class="metric-row" style="flex-direction: column; gap: 0.25rem;">
                    <span class="metric-label" data-i18n="lbl_label_names">Labelnamen:</span>
                    <span class="metric-value" style="font-size: 0.85rem; color: var(--text-muted);">$purviewLabelsNames</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_copilot_exclusion">Copilot-Ausschluss aktiv (BlockContentAnalysisServices)</span>
                    <span class="metric-value $(if ($purviewCopilotBlocked -gt 0) { 'warning' } else { '' })">$purviewCopilotBlocked <span data-i18n="lbl_labels">Labels</span></span>
                </div>
                <div class="metric-row" style="flex-direction: column; gap: 0.25rem;">
                    <span class="metric-label" data-i18n="lbl_copilot_blocked_labels">Copilot-ausgeschlossene Labels:</span>
                    <span class="metric-value" style="font-size: 0.85rem; color: var(--text-muted);">$purviewCopilotBlockedNames</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_dlp_policies_count">DLP Richtlinien (Data Loss Prevention)</span>
                    <span class="metric-value">$purviewDlp</span>
                </div>
                <div class="metric-row" style="flex-direction: column; gap: 0.25rem;">
                    <span class="metric-label" data-i18n="lbl_dlp_policy_names">DLP Richtliniennamen:</span>
                    <span class="metric-value" style="font-size: 0.85rem; color: var(--text-muted);">$purviewDlpNames</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_retention_labels_count">Aufbewahrungsbezeichnungen (Retention Labels)</span>
                    <span class="metric-value">$purviewRetention</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_mandatory_labeling_status">Labeling Pflicht (Mandatory)</span>
                    <span class="metric-value $(if ($mandatoryLabeling -eq 'Ja') { 'success' } else { '' })">
                        $(if ($mandatoryLabeling -eq 'Ja') { '<span data-i18n="lbl_yes">Ja</span>' } else { '<span data-i18n="lbl_no">Nein</span>' })
                    </span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_default_label_status">Standard-Label</span>
                    <span class="metric-value" style="font-size: 0.85rem; font-family: monospace;">$defaultLabel</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label" data-i18n="lbl_downgrade_justification_status">Begründungspflicht bei Herabstufung</span>
                    <span class="metric-value $(if ($justificationReq -eq 'Ja') { 'success' } else { '' })">
                        $(if ($justificationReq -eq 'Ja') { '<span data-i18n="lbl_yes">Ja</span>' } else { '<span data-i18n="lbl_no">Nein</span>' })
                    </span>
                </div>
            </div>
        </div>

        <details class="collector-detail">
            <summary data-i18n="sum_purview_table">🔒 Purview Vertraulichkeitslabels-Details</summary>
            $purviewTableHtml
        </details>

        <footer data-i18n="footer_text">
            M365 GRC Assistant Onboarding Portal · Erstellt von Michael Kirst-Neshva
        </footer>
    </div>

    <!-- Internationalization and Charts Script -->
    <script>
        const GrcTranslations = {
          de: {
            "report_title": "M365 GRC Asset Audit Report",
            "lbl_tenant": "Mandant",
            "lbl_generated_at": "Generiert am",
            "stat_users_total": "Benutzer gesamt",
            "stat_active_accounts": "Aktive Konten",
            "stat_groups_total": "Gruppen gesamt",
            "stat_entra_devices": "Entra ID Devices",
            "stat_intune_managed": "Intune Managed",
            "stat_defender_endpoints": "Defender Endpoints",
            "card_tenant_details": "Mandanten-Details",
            "lbl_name": "Name",
            "lbl_tenant_id": "Tenant ID",
            "lbl_security_defaults": "Identity Security Defaults",
            "lbl_security_defaults_active": "Aktiviert",
            "lbl_security_defaults_inactive": "Deaktiviert",
            "lbl_verified_domains": "Verifizierte Domains:",
            "card_mfa_sec": "MFA Absicherung",
            "card_groups_struct": "Gruppen-Struktur",
            "lbl_m365_groups": "M365 / Unified Groups",
            "lbl_sec_groups": "Sicherheitsgruppen",
            "lbl_total_groups": "Gruppen gesamt",
            "sum_users_details": "👤 Benutzer-Details (Entra ID Users)",
            "sum_groups_details": "👥 Gruppen-Details (Entra ID Groups)",
            "lbl_user_status": "👤 Benutzerkonten-Status",
            "lbl_hardware_audit": "💻 Endgeräte GRC-Audit",
            "lbl_registered_joined": "In Entra ID registriert/joined (Hardware Asset)",
            "lbl_managed_intune": "In Intune verwaltet (Compliance erzwingbar)",
            "lbl_enrolled_defender": "In Defender for Endpoint erfasst (EDR Abdeckung)",
            "lbl_intune_coverage": "Intune-Abdeckung (relativ to Entra ID)",
            "sum_devices_details": "💻 Geräte-Details (Hardware & Compliance)",
            "lbl_sec_score_recommendations": "🎯 Microsoft Secure Score & Handlungsempfehlungen",
            "lbl_collab_mail_audit": "📬 Kollaboration & E-Mail GRC-Audit",
            "sum_exchange_details": "✉️ Exchange Online",
            "lbl_user_mailboxes": "Benutzer-Postfächer",
            "lbl_shared_mailboxes": "Gemeinsame Postfächer (Shared)",
            "lbl_transport_rules": "Mailflow Transport-Regeln",
            "lbl_dkim_domains": "DKIM-geschützte Domains",
            "lbl_antimalware_policies": "Anti-Malware Richtlinien",
            "sum_sharepoint_details": "🌐 SharePoint & OneDrive",
            "lbl_sp_sites": "Aktive SharePoint-Sites",
            "lbl_sharing_policy": "Externer Freigabe-Modus",
            "lbl_sharing_capabilities": "Freigabe-Berechtigung",
            "sum_teams_details": "💬 Microsoft Teams",
            "lbl_teams_total": "Teams gesamt",
            "lbl_public_teams": "Öffentliche Teams (Risiko)",
            "lbl_private_teams": "Private Teams",
            "sum_exchange_table": "✉️ Exchange Online Postfach-Details",
            "sum_sharepoint_table": "🌐 SharePoint Online Website-Details",
            "sum_teams_table": "👥 Teams-Details (Kanallisten & Gäste)",
            "lbl_governance_compliance_audit": "🛡️ Identity Governance & Compliance GRC-Audit",
            "lbl_entra_governance": "🔑 Identity Governance (Entra ID)",
            "lbl_ca_policies": "Bedingter Zugriff (CA-Richtlinien)",
            "lbl_policies": "Richtlinien",
            "lbl_ca_enabled": "CA Richtlinien Aktiviert",
            "lbl_ca_report_only": "CA Richtlinien im Report-only Modus",
            "lbl_access_packages": "Zugriffspakete (Access Packages)",
            "lbl_access_reviews": "Zugriffsüberprüfungen (Access Reviews)",
            "lbl_purview_compliance": "🔒 Microsoft Purview Compliance & Information Protection",
            "lbl_sensitivity_labels": "Vertraulichkeitslabels (Sensitivity Labels)",
            "lbl_label_names": "Labelnamen:",
            "lbl_copilot_exclusion": "Copilot-Ausschluss aktiv (BlockContentAnalysisServices)",
            "lbl_labels": "Labels",
            "lbl_copilot_blocked_labels": "Copilot-ausgeschlossene Labels:",
            "lbl_dlp_policies_count": "DLP Richtlinien (Data Loss Prevention)",
            "lbl_dlp_policy_names": "DLP Richtliniennamen:",
            "lbl_retention_labels_count": "Aufbewahrungsbezeichnungen (Retention Labels)",
            "lbl_mandatory_labeling_status": "Labeling Pflicht (Mandatory)",
            "lbl_yes": "Ja",
            "lbl_no": "Nein",
            "lbl_default_label_status": "Standard-Label",
            "lbl_downgrade_justification_status": "Begründungspflicht bei Herabstufung",
            "sum_purview_table": "🔒 Purview Sensitivity Labels Details",
            "sec_score_status": "Secure Score Status",
            "global_average": "M365 globaler Durchschnitt",
            "deviation_average": "Abweichung zum Durchschnitt",
            "footer_text": "M365 GRC Assistant Onboarding Portal · Erstellt von Michael Kirst-Neshva",
            // Table headers auto translation
            "Name": "Name", "UPN": "UPN", "Aktiv": "Aktiv", "MFA": "MFA", "Admin-Rollen": "Admin-Rollen", "Vorgesetzter": "Vorgesetzter",
            "Gruppenname": "Gruppenname", "Klassifizierung": "Klassifizierung", "Sichtbarkeit": "Sichtbarkeit", "Besitzer": "Besitzer", "Mitglieder": "Mitglieder",
            "Gerätename": "Gerätename", "Betriebssystem": "Betriebssystem", "OS-Version": "OS-Version", "Trust-Typ": "Trust-Typ", "Konform": "Konform", "Intune MDM": "Intune MDM", "Defender Status": "Defender Status",
            "Postfach-Adresse": "Postfach-Adresse", "Typ": "Typ", "Quota Limit": "Quota Limit", "Vollzugriff (Delegiert)": "Vollzugriff (Delegiert)", "Send As (Delegiert)": "Send As (Delegiert)", "Weiterleitung": "Weiterleitung",
            "Website-Name": "Website-Name", "URL": "URL", "Administratoren": "Administratoren", "Speicher (Bytes)": "Speicher (Bytes)",
            "Teamname": "Teamname", "Gäste": "Gäste", "Kanäle": "Kanäle",
            "Empfehlung": "Empfehlung", "Max. Punkte": "Max. Punkte", "Auswirkung": "Auswirkung", "Status": "Status", "Lizenzbedarf": "Lizenzbedarf", "Handlungsanweisung": "Handlungsanweisung",
            // Badges / States auto translation
            "Umgesetzt": "Umgesetzt", "Teilweise / Alternativ": "Teilweise / Alternativ", "Nicht umgesetzt": "Nicht umgesetzt",
            "Ja": "Ja", "Nein": "Nein",
            // Purview Table headers
            "Labelname": "Labelname", "Copilot Sperre": "Copilot Sperre", "Scope Files": "Scope Files", "Scope Emails": "Scope Emails", "Scope Sites": "Scope Sites", "Verschlüsselt": "Verschlüsselt", "Richtlinien": "Richtlinien"
          },
          en: {
            "report_title": "M365 GRC Asset Audit Report",
            "lbl_tenant": "Tenant",
            "lbl_generated_at": "Generated at",
            "stat_users_total": "Total Users",
            "stat_active_accounts": "Active Accounts",
            "stat_groups_total": "Total Groups",
            "stat_entra_devices": "Entra ID Devices",
            "stat_intune_managed": "Intune Managed",
            "stat_defender_endpoints": "Defender Endpoints",
            "card_tenant_details": "Tenant Details",
            "lbl_name": "Name",
            "lbl_tenant_id": "Tenant ID",
            "lbl_security_defaults": "Identity Security Defaults",
            "lbl_security_defaults_active": "Enabled",
            "lbl_security_defaults_inactive": "Disabled",
            "lbl_verified_domains": "Verified Domains:",
            "card_mfa_sec": "MFA Security Status",
            "card_groups_struct": "Groups Structure",
            "lbl_m365_groups": "M365 / Unified Groups",
            "lbl_sec_groups": "Security Groups",
            "lbl_total_groups": "Total Groups",
            "sum_users_details": "👤 User Details (Entra ID Users)",
            "sum_groups_details": "👥 Groups Details (Entra ID Groups)",
            "lbl_user_status": "👤 User Account Status",
            "lbl_hardware_audit": "💻 Devices GRC Audit",
            "lbl_registered_joined": "Registered/Joined in Entra ID (Hardware Asset)",
            "lbl_managed_intune": "Managed in Intune (Compliance enforceable)",
            "lbl_enrolled_defender": "Enrolled in Defender for Endpoint (EDR Coverage)",
            "lbl_intune_coverage": "Intune Coverage (relative to Entra ID)",
            "sum_devices_details": "💻 Device Details (Hardware & Compliance)",
            "lbl_sec_score_recommendations": "🎯 Microsoft Secure Score & Recommendations",
            "lbl_collab_mail_audit": "📬 Collaboration & Mail GRC Audit",
            "sum_exchange_details": "✉️ Exchange Online",
            "lbl_user_mailboxes": "User Mailboxes",
            "lbl_shared_mailboxes": "Shared Mailboxes",
            "lbl_transport_rules": "Mailflow Transport Rules",
            "lbl_dkim_domains": "DKIM-Protected Domains",
            "lbl_antimalware_policies": "Anti-Malware Policies",
            "sum_sharepoint_details": "🌐 SharePoint & OneDrive",
            "lbl_sp_sites": "Active SharePoint Sites",
            "lbl_sharing_policy": "External Sharing Policy",
            "lbl_sharing_capabilities": "File Sharing Capabilities",
            "sum_teams_details": "💬 Microsoft Teams",
            "lbl_teams_total": "Total Teams",
            "lbl_public_teams": "Public Teams (Risk)",
            "lbl_private_teams": "Private Teams",
            "sum_exchange_table": "✉️ Exchange Online Mailbox Details",
            "sum_sharepoint_table": "🌐 SharePoint Online Site Details",
            "sum_teams_table": "👥 Teams Details (Channel Lists & Guests)",
            "lbl_governance_compliance_audit": "🛡️ Identity Governance & Compliance GRC Audit",
            "lbl_entra_governance": "🔑 Identity Governance (Entra ID)",
            "lbl_ca_policies": "Conditional Access (CA Policies)",
            "lbl_policies": "Policies",
            "lbl_ca_enabled": "CA Policies Enabled",
            "lbl_ca_report_only": "CA Policies in Report-only Mode",
            "lbl_access_packages": "Access Packages",
            "lbl_access_reviews": "Access Reviews",
            "lbl_purview_compliance": "🔒 Microsoft Purview Compliance & Information Protection",
            "lbl_sensitivity_labels": "Sensitivity Labels (Classification)",
            "lbl_label_names": "Label Names:",
            "lbl_copilot_exclusion": "Copilot Exclusion Active (BlockContentAnalysisServices)",
            "lbl_labels": "Labels",
            "lbl_copilot_blocked_labels": "Copilot Blocked Labels:",
            "lbl_dlp_policies_count": "DLP Policies (Data Loss Prevention)",
            "lbl_dlp_policy_names": "DLP Policy Names:",
            "lbl_retention_labels_count": "Retention Labels",
            "lbl_mandatory_labeling_status": "Mandatory Labeling",
            "lbl_yes": "Yes",
            "lbl_no": "No",
            "lbl_default_label_status": "Default Label",
            "lbl_downgrade_justification_status": "Downgrade Justification Required",
            "sum_purview_table": "🔒 Purview Sensitivity Labels Details",
            "sec_score_status": "Secure Score Status",
            "global_average": "M365 Global Average",
            "deviation_average": "Deviation to Average",
            "footer_text": "M365 GRC Assistant Onboarding Portal · Created by Michael Kirst-Neshva",
            // Table headers auto translation
            "Name": "Name", "UPN": "UPN", "Aktiv": "Active", "MFA": "MFA", "Admin-Rollen": "Admin Roles", "Vorgesetzter": "Manager",
            "Gruppenname": "Group Name", "Klassifizierung": "Classification", "Sichtbarkeit": "Visibility", "Besitzer": "Owners", "Mitglieder": "Members",
            "Gerätename": "Device Name", "Betriebssystem": "Operating System", "OS-Version": "OS Version", "Trust-Typ": "Trust Type", "Konform": "Compliant", "Intune MDM": "Intune MDM", "Defender Status": "Defender Status",
            "Postfach-Adresse": "Mailbox Address", "Typ": "Type", "Quota Limit": "Quota Limit", "Vollzugriff (Delegiert)": "Full Access (Delegated)", "Send As (Delegiert)": "Send As (Delegated)", "Weiterleitung": "Forwarding",
            "Website-Name": "Site Name", "URL": "URL", "Administratoren": "Administrators", "Speicher (Bytes)": "Storage (Bytes)",
            "Teamname": "Team Name", "Gäste": "Guests", "Kanäle": "Channels",
            "Empfehlung": "Recommendation", "Max. Punkte": "Max. Score", "Auswirkung": "User Impact", "Status": "Status", "Lizenzbedarf": "License Required", "Handlungsanweisung": "Remediation",
            // Badges / States auto translation
            "Umgesetzt": "Implemented", "Teilweise / Alternativ": "Partial / Alternative", "Nicht umgesetzt": "Not Implemented",
            "Ja": "Yes", "Nein": "No",
            // Purview Table headers
            "Labelname": "Label Name", "Copilot Sperre": "Copilot Block", "Scope Files": "Scope Files", "Scope Emails": "Scope Emails", "Scope Sites": "Scope Sites", "Verschlüsselt": "Encrypted", "Richtlinien": "Policies"
          }
        };

        let mfaChartInstance, userStatusChartInstance;

        function updateCharts(lang) {
          const mfaLabels = {
            de: ['MFA Registriert', 'MFA Nicht registriert', 'Unbekannt'],
            en: ['MFA Registered', 'MFA Not Registered', 'Unknown']
          };
          const statusLabels = {
            de: ['Aktive Konten', 'Deaktivierte Konten'],
            en: ['Active Accounts', 'Disabled Accounts']
          };
          
          if (mfaChartInstance) {
            mfaChartInstance.data.labels = mfaLabels[lang];
            mfaChartInstance.update();
          }
          if (userStatusChartInstance) {
            userStatusChartInstance.data.labels = statusLabels[lang];
            userStatusChartInstance.update();
          }
        }

        function applyLanguage(lang) {
          document.documentElement.lang = lang;
          
          // Translate direct data-i18n tags
          document.querySelectorAll("[data-i18n]").forEach(elem => {
            const key = elem.getAttribute("data-i18n");
            if (GrcTranslations[lang] && GrcTranslations[lang][key] !== undefined) {
              elem.innerHTML = GrcTranslations[lang][key];
            }
          });

          // Translate table headers automatically
          document.querySelectorAll("th").forEach(th => {
            const key = th.getAttribute("data-i18n-hdr") || th.textContent.trim();
            if (!th.getAttribute("data-i18n-hdr")) {
              th.setAttribute("data-i18n-hdr", key);
            }
            if (GrcTranslations[lang] && GrcTranslations[lang][key] !== undefined) {
              th.innerHTML = GrcTranslations[lang][key];
            }
          });

          // Translate badges
          document.querySelectorAll(".badge").forEach(badge => {
            const key = badge.getAttribute("data-i18n-badge") || badge.textContent.trim();
            if (!badge.getAttribute("data-i18n-badge")) {
              badge.setAttribute("data-i18n-badge", key);
            }
            if (GrcTranslations[lang] && GrcTranslations[lang][key] !== undefined) {
              badge.innerHTML = GrcTranslations[lang][key];
            }
          });

          // Handle dynamic score values translation
          document.querySelectorAll("[data-i18n-score-text]").forEach(elem => {
            const parts = elem.getAttribute("data-i18n-score-text").split("|");
            const pct = parts[0];
            const curr = parts[1];
            const max = parts[2];
            if (lang === "de") {
              elem.innerHTML = 'Ihr aktueller Sicherheitsindex liegt bei <strong style="color: #ffffff;">' + pct + '%</strong>. Das entspricht <strong style="color: #ffffff;">' + curr + ' von ' + max + '</strong> möglichen Punkten.';
            } else {
              elem.innerHTML = 'Your current security index is <strong style="color: #ffffff;">' + pct + '%</strong>. This represents <strong style="color: #ffffff;">' + curr + ' of ' + max + '</strong> possible points.';
            }
          });

          document.querySelectorAll("[data-i18n-diff]").forEach(elem => {
            const diffVal = parseFloat(elem.getAttribute("data-i18n-diff"));
            const absDiff = Math.abs(diffVal);
            if (lang === "de") {
              elem.textContent = diffVal >= 0 ? absDiff + ' Pkt. über dem Durchschnitt' : absDiff + ' Pkt. unter dem Durchschnitt';
            } else {
              elem.textContent = diffVal >= 0 ? absDiff + ' pts above global average' : absDiff + ' pts below global average';
            }
          });

          updateCharts(lang);
        }

        function switchLanguage(lang) {
          localStorage.setItem("grc_lang", lang);
          window.location.reload();
        }

        // Initialize Charts
        mfaChartInstance = new Chart(document.getElementById('mfaChart'), {
            type: 'doughnut',
            data: {
                labels: ['MFA Registriert', 'MFA Nicht registriert', 'Unbekannt'],
                datasets: [{
                    data: [$mfaUsers, $noMfaUsers, $mfaUnknown],
                    backgroundColor: ['#10b981', '#ef4444', '#6b7280'],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { color: '#f3f4f6', boxWidth: 12, font: { family: 'Outfit' } }
                    }
                }
            }
        });

        userStatusChartInstance = new Chart(document.getElementById('userStatusChart'), {
            type: 'doughnut',
            data: {
                labels: ['Aktive Konten', 'Deaktivierte Konten'],
                datasets: [{
                    data: [$activeUsers, $disabledUsers],
                    backgroundColor: ['#6366f1', '#f59e0b'],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: { color: '#f3f4f6', boxWidth: 12, font: { family: 'Outfit' } }
                    }
                }
            }
        });

        // Initialize Language on load
        document.addEventListener("DOMContentLoaded", () => {
          const savedLang = localStorage.getItem("grc_lang") || "de";
          document.getElementById("languageSelect").value = savedLang;
          applyLanguage(savedLang);
        });
    </script>
</body>
</html>
"@

# 5. Write index.html to docs/ directory
$docsDir = Join-Path -Path $PSScriptRoot -ChildPath "../docs"
if (!(Test-Path $docsDir)) {
    New-Item -ItemType Directory -Path $docsDir -Force | Out-Null
}

$reportPath = Join-Path -Path $docsDir -ChildPath "index.html"
$htmlContent | Set-Content -Path $reportPath -Encoding utf8

Write-Host "=== GRC HTML Report generated successfully at: $reportPath ===" -ForegroundColor Green
