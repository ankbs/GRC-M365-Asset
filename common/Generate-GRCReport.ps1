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
            return Get-Content -Raw -Path $latest | ConvertFrom-Json
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
                <h1>M365 GRC Asset Audit Report</h1>
                <p>Mandant: <strong>$tenantName</strong> ($tenantIdVal)</p>
            </div>
            <div class="timestamp-badge">
                📅 Generiert am: $(Get-Date -Format "dd.MM.yyyy HH:mm")
            </div>
        </header>

        <!-- Global Count Summaries -->
        <div class="stats-summary">
            <div class="stat-box">
                <div class="num">$totalUsers</div>
                <div class="lbl">Benutzer gesamt</div>
            </div>
            <div class="stat-box">
                <div class="num">$activeUsers</div>
                <div class="lbl">Aktive Konten</div>
            </div>
            <div class="stat-box">
                <div class="num">$totalGroups</div>
                <div class="lbl">Gruppen gesamt</div>
            </div>
            <div class="stat-box">
                <div class="num">$totalEntraDevices</div>
                <div class="lbl">Entra ID Devices</div>
            </div>
            <div class="stat-box">
                <div class="num">$totalIntuneDevices</div>
                <div class="lbl">Intune Managed</div>
            </div>
            <div class="stat-box">
                <div class="num">$totalDefenderDevices</div>
                <div class="lbl">Defender Endpoints</div>
            </div>
        </div>

        <div class="grid-3">
            <!-- Tenant Metadata Card -->
            <div class="card">
                <h2>🏢 Mandanten-Details</h2>
                <div class="metric-row">
                    <span class="metric-label">Name</span>
                    <span class="metric-value">$tenantName</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label">Tenant ID</span>
                    <span class="metric-value" style="font-size: 0.8rem; font-family: monospace;">$tenantIdVal</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label">Identity Security Defaults</span>
                    <span class="metric-value $(if ($securityDefaults -eq $true) { 'success' } else { 'danger' })">
                        $(if ($securityDefaults -eq $true) { 'Aktiviert' } elseif ($securityDefaults -eq $false) { 'Deaktiviert' } else { 'N/A' })
                    </span>
                </div>
                <div class="metric-row" style="flex-direction: column; gap: 0.25rem;">
                    <span class="metric-label">Verifizierte Domains:</span>
                    <span class="metric-value" style="font-size: 0.8rem; color: var(--text-muted); word-break: break-all; margin-top: 0.25rem;">$verifiedDomains</span>
                </div>
            </div>

            <!-- Identity GRC Card -->
            <div class="card">
                <h2>MFA Absicherung</h2>
                <div class="chart-wrapper">
                    <canvas id="mfaChart"></canvas>
                </div>
            </div>

            <!-- Groups Audit Card -->
            <div class="card">
                <h2>👥 Gruppen-Struktur</h2>
                <div class="metric-row">
                    <span class="metric-label">M365 / Unified Groups</span>
                    <span class="metric-value">$m365Groups</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label">Sicherheitsgruppen</span>
                    <span class="metric-value">$securityGroups</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label">Gruppen gesamt</span>
                    <span class="metric-value">$totalGroups</span>
                </div>
            </div>
        </div>

        <div class="grid-2">
            <!-- User Status Breakdown Card -->
            <div class="card">
                <h2>👤 Benutzerkonten-Status</h2>
                <div class="chart-wrapper">
                    <canvas id="userStatusChart"></canvas>
                </div>
            </div>

            <!-- Device GRC Overlap Card -->
            <div class="card">
                <h2>💻 Endgeräte GRC-Audit</h2>
                <div class="metric-row">
                    <span class="metric-label">In Entra ID registriert/joined (Hardware Asset)</span>
                    <span class="metric-value">$totalEntraDevices</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label">In Intune verwaltet (Compliance erzwingbar)</span>
                    <span class="metric-value" style="color: var(--accent-secondary);">$totalIntuneDevices</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label">In Defender for Endpoint erfasst (EDR Abdeckung)</span>
                    <span class="metric-value" style="color: var(--success);">$totalDefenderDevices</span>
                </div>
                <div class="metric-row">
                    <span class="metric-label">Intune-Abdeckung (relativ to Entra ID)</span>
                    <span class="metric-value">
                        $(if ($totalEntraDevices -gt 0) { "{0:P1}" -f ($totalIntuneDevices / $totalEntraDevices) } else { "0%" })
                    </span>
                </div>
            </div>
        </div>

        <footer>
            M365 GRC Assistant Onboarding Portal · Erstellt von Antigravity AI
        </footer>
    </div>

    <!-- Chart Configuration Script -->
    <script>
        // MFA Chart
        new Chart(document.getElementById('mfaChart'), {
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

        // User Status Chart
        new Chart(document.getElementById('userStatusChart'), {
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
