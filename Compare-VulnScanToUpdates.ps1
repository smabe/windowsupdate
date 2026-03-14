<#
.SYNOPSIS
    Correlates a Qualys vulnerability scan report (XLSX) with Windows Update installation results (all_updates.csv).

.DESCRIPTION
    Reads a Qualys scan report and the all_updates.csv output from the Windows Update script,
    then determines which vulnerabilities have been remediated by installed updates.
    Supports cumulative update supersession — if a host has a newer cumulative update installed
    (e.g., 2026-03) than the missing KB's month (e.g., 2025-12), the vulnerability is marked
    "Remediated (Cumulative)" rather than "Not Remediated".
    Produces an XLSX report with Host Summary, Vulnerability Detail, Unmatched Hosts,
    and Cumulative Coverage sheets.

.PARAMETER VulnReportPath
    Path to the Qualys vulnerability scan XLSX file.

.PARAMETER UpdatesCsvPath
    Path to the all_updates.csv file from the Windows Update script.

.PARAMETER OutputPath
    Path for the output report. Defaults to the same directory as the vuln report.

.PARAMETER ExportCsv
    Also export a detailed CSV alongside the XLSX report.

.PARAMETER HostnameColumn
    Which vuln report column to use for hostname matching: "DNS" (col 3) or "NetBIOS" (col 4).
    Default: DNS, with fallback to NetBIOS if DNS is empty.

.EXAMPLE
    .\Compare-VulnScanToUpdates.ps1 -VulnReportPath ".\scan_report.xlsx" -UpdatesCsvPath ".\all_updates.csv"

.EXAMPLE
    .\Compare-VulnScanToUpdates.ps1 -VulnReportPath ".\scan_report.xlsx" -UpdatesCsvPath ".\all_updates.csv" -ExportCsv
#>
param(
    [Parameter(Mandatory)][string]$VulnReportPath,
    [Parameter(Mandatory)][string]$UpdatesCsvPath,
    [string]$OutputPath,
    [switch]$ExportCsv,
    [ValidateSet("DNS","NetBIOS")][string]$HostnameColumn = "DNS",
    [switch]$IncludePreviousLogs
)

# ============================================================================
# VALIDATION
# ============================================================================

if (-not (Test-Path $VulnReportPath)) {
    Write-Host "ERROR: Vulnerability report not found: $VulnReportPath" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $UpdatesCsvPath)) {
    Write-Host "ERROR: Updates CSV not found: $UpdatesCsvPath" -ForegroundColor Red
    exit 1
}

$VulnReportPath = (Resolve-Path $VulnReportPath).Path
$UpdatesCsvPath = (Resolve-Path $UpdatesCsvPath).Path

if (-not $OutputPath) {
    $dir = Split-Path $VulnReportPath -Parent
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($VulnReportPath)
    $OutputPath = Join-Path $dir "${baseName}_Remediation_Report.xlsx"
}
# Excel COM requires absolute paths for SaveAs
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)

# ============================================================================
# HELPER: Normalize hostname (strip domain, lowercase, trim)
# ============================================================================
function Get-NormalizedHostname {
    param([string]$Name)
    $n = $Name.Trim().ToLower()
    if ($n -match '\.') { $n = ($n -split '\.')[0] }
    return $n
}

# ============================================================================
# HELPER: Extract product family from cumulative update title
# ============================================================================
function Get-CumulativeProductFamily {
    param([string]$Title)
    if ($Title -match 'Cumulative Update for \.NET|\.NET Framework') { return "DotNet" }
    if ($Title -match 'Cumulative.*Update.*Windows|Security Update.*Windows') { return "Windows" }
    if ($Title -match 'Cumulative Update for') { return "Other" }
    if ($Title -match 'Security Update \(KB') { return "Windows" }  # e.g., "2026-03 Security Update (KB5079473)"
    if ($Title -match 'Windows\s+(10|11|Server)\b') { return "Windows" }  # e.g., "Windows 11, version 25H2"
    return $null
}

# ============================================================================
# READ ALL_UPDATES.CSV
# ============================================================================

Write-Host "`n=== Reading Windows Update results ===" -ForegroundColor Cyan
$updates = Import-Csv $UpdatesCsvPath

# Validate required columns
$requiredCols = @("ComputerName", "KB", "Status")
$csvCols = $updates[0].PSObject.Properties.Name
$missingCols = $requiredCols | Where-Object { $csvCols -notcontains $_ }
if ($missingCols) {
    Write-Host "ERROR: all_updates.csv is missing required columns: $($missingCols -join ', ')" -ForegroundColor Red
    Write-Host "  Expected: $($requiredCols -join ', ')" -ForegroundColor Yellow
    exit 1
}
$hasTitle = $csvCols -contains "Title"
if (-not $hasTitle) {
    Write-Host "  Warning: 'Title' column not found — cumulative update detection disabled" -ForegroundColor Yellow
}
$hasIP = $csvCols -contains "ComputerIP"
if (-not $hasIP) {
    Write-Host "  Warning: 'ComputerIP' column not found — falling back to hostname matching" -ForegroundColor Yellow
}

# Derive session year-month from folder name (e.g. "Session_20260313_221433" → "2026-03")
# Used as fallback when update Title lacks a date pattern
$sessionYearMonth = $null
$sessionDirName = Split-Path (Split-Path $UpdatesCsvPath -Parent) -Leaf
if ($sessionDirName -match 'Session_(\d{4})(\d{2})\d{2}') {
    $sessionYearMonth = "{0}-{1}" -f $Matches[1], $Matches[2]
}

# Trim all string fields
$updates | ForEach-Object {
    foreach ($prop in $_.PSObject.Properties) {
        if ($prop.Value -is [string]) { $prop.Value = $prop.Value.Trim() }
    }
}

# Build lookup: key = "ip|KB" → status (primary), with hostname fallback
# Also build per-host cumulative update tracker keyed by IP
$updateLookup = @{}
$updateIPs = @{}          # IP → hostname (for display in unmatched)
$updateHosts = @{}        # hostname_lower → $true (fallback matching)
$hostCumulatives = @{}    # matchKey → @( @{Product; YearMonth; KB; Title} )

foreach ($u in $updates) {
    $hostShort = Get-NormalizedHostname $u.ComputerName
    $ip = if ($hasIP -and $u.ComputerIP) { $u.ComputerIP.Trim() } else { $null }
    $kb = ($u.KB -replace '\s','').ToUpper()
    if (-not $kb) { continue }

    # Primary key: IP-based (reliable across naming conventions)
    # Fallback key: hostname-based (when IP not available)
    $matchKey = if ($ip) { $ip } else { $hostShort }
    $key = "$matchKey|$kb"
    $status = $u.Status

    # Keep "Installed" as the winning status
    if ($updateLookup[$key] -ne "Installed") {
        $updateLookup[$key] = $status
    }
    if ($ip) { $updateIPs[$ip] = $u.ComputerName }
    $updateHosts[$hostShort] = $true

    # Track cumulative updates by parsing title for YYYY-MM pattern
    if ($hasTitle -and $u.Title -and $status -eq "Installed") {
        if ($u.Title -match '(\d{4})-(\d{1,2})\s+(?:Cumulative|Security)\s+Update') {
            $yearMonth = "{0}-{1:D2}" -f $Matches[1], [int]$Matches[2]
            $productFamily = Get-CumulativeProductFamily $u.Title
            if ($productFamily) {
                if (-not $hostCumulatives[$matchKey]) { $hostCumulatives[$matchKey] = @() }
                $hostCumulatives[$matchKey] += @{
                    Product   = $productFamily
                    YearMonth = $yearMonth
                    KB        = $kb
                    Title     = $u.Title
                }
            }
        }
        # Fallback: Title has product info but no date (e.g. "Windows 11, version 25H2")
        # Use session date as the cumulative month
        elseif ($sessionYearMonth) {
            $productFamily = Get-CumulativeProductFamily $u.Title
            if ($productFamily) {
                if (-not $hostCumulatives[$matchKey]) { $hostCumulatives[$matchKey] = @() }
                $hostCumulatives[$matchKey] += @{
                    Product   = $productFamily
                    YearMonth = $sessionYearMonth
                    KB        = $kb
                    Title     = $u.Title
                }
            }
        }
    }
}

$matchMode = if ($hasIP) { "IP address" } else { "hostname" }
Write-Host "  Loaded $($updates.Count) update records for $($updateIPs.Count) IPs ($($updateHosts.Count) hostnames)"
Write-Host "  Matching mode: $matchMode"
Write-Host "  Unique matchkey+KB combinations: $($updateLookup.Count)"
$hostsWithCumulatives = ($hostCumulatives.Keys | Measure-Object).Count
if ($hostsWithCumulatives -gt 0) {
    Write-Host "  Hosts with cumulative updates tracked: $hostsWithCumulatives" -ForegroundColor DarkCyan
}

# Build KB → cumulative info map so hotfix-only KBs can inherit cumulative metadata
$knownCumulativeKBs = @{}
foreach ($entries in $hostCumulatives.Values) {
    foreach ($cum in $entries) {
        if (-not $knownCumulativeKBs.ContainsKey($cum.KB)) {
            $knownCumulativeKBs[$cum.KB] = @{
                Product   = $cum.Product
                YearMonth = $cum.YearMonth
                Title     = $cum.Title
            }
        }
    }
}

# Build KB → Title map from all_updates.csv (for product family detection in hotfix data)
$kbTitleMap = @{}
if ($hasTitle) {
    foreach ($u in $updates) {
        $kb = ($u.KB -replace '\s','').ToUpper()
        if ($kb -and $u.Title -and -not $kbTitleMap.ContainsKey($kb)) {
            $kbTitleMap[$kb] = $u.Title
        }
    }
}

# ============================================================================
# INCORPORATE PREVIOUS_UPDATELOG FILES (from earlier deployment runs)
# ============================================================================

if ($IncludePreviousLogs) {
    $sessionDir = Split-Path $UpdatesCsvPath -Parent
    $prevLogFiles = Get-ChildItem -Path $sessionDir -Filter "previous_updatelog_*.csv" -ErrorAction SilentlyContinue
    $computerSumPath = Join-Path $sessionDir "computer_summary.csv"

    if ($prevLogFiles.Count -gt 0) {
        Write-Host "`n=== Incorporating previous run logs ===" -ForegroundColor Cyan
        Write-Host "  Found $($prevLogFiles.Count) previous_updatelog files"

        # Build Name → IP mapping from computer_summary.csv
        $nameToIP = @{}
        if (Test-Path $computerSumPath) {
            Import-Csv $computerSumPath | ForEach-Object {
                $compName = $_.ComputerName.Trim()
                $compIP = $_.IP.Trim()
                if ($compName -and $compIP) { $nameToIP[$compName] = $compIP }
            }
            Write-Host "  Loaded $($nameToIP.Count) Name→IP mappings from computer_summary.csv"
        } else {
            Write-Host "  Warning: computer_summary.csv not found — previous logs will use hostname matching only" -ForegroundColor Yellow
        }

        $prevAdded = 0
        $prevSkipped = 0
        foreach ($prevFile in $prevLogFiles) {
            # Extract the computer name from filename: "previous_updatelog_<ComputerName>.csv"
            $compName = $prevFile.BaseName -replace '^previous_updatelog_', ''
            $compIP = $nameToIP[$compName]

            # Determine match key (same logic as primary CSV)
            $prevMatchKey = if ($compIP) { $compIP } else { (Get-NormalizedHostname $compName) }

            $prevUpdates = Import-Csv $prevFile.FullName -ErrorAction SilentlyContinue
            if (-not $prevUpdates) { continue }

            foreach ($pu in $prevUpdates) {
                if (-not $pu.KB) { continue }
                $kb = ($pu.KB -replace '\s','').ToUpper()
                if (-not $kb) { continue }

                # Determine status — previous logs use "Result" not "Status"
                $status = "Unknown"
                if ($pu.Result) { $status = $pu.Result.Trim() }
                elseif ($pu.InstallResult) { $status = $pu.InstallResult.Trim() }
                elseif ($pu.Status) { $status = $pu.Status.Trim() }

                $key = "$prevMatchKey|$kb"

                # Only add if not already in the lookup (primary CSV takes precedence)
                if (-not $updateLookup.ContainsKey($key)) {
                    $updateLookup[$key] = $status
                    $prevAdded++

                    # Also track for IP/host mapping
                    if ($compIP) { $updateIPs[$compIP] = $compName }
                    $updateHosts[(Get-NormalizedHostname $compName)] = $true
                } else {
                    $prevSkipped++
                }

                # Track cumulative updates from previous logs too
                $title = if ($pu.Title) { $pu.Title.Trim() } else { "" }
                if ($title -and $status -eq "Installed") {
                    if ($title -match '(\d{4})-(\d{1,2})\s+(?:Cumulative|Security)\s+Update') {
                        $yearMonth = "{0}-{1:D2}" -f $Matches[1], [int]$Matches[2]
                        $productFamily = Get-CumulativeProductFamily $title
                        if ($productFamily) {
                            # Check if we already have this exact cumulative tracked
                            $existingCums = $hostCumulatives[$prevMatchKey]
                            $alreadyTracked = $false
                            if ($existingCums) {
                                foreach ($ec in $existingCums) {
                                    if ($ec.KB -eq $kb -and $ec.YearMonth -eq $yearMonth) { $alreadyTracked = $true; break }
                                }
                            }
                            if (-not $alreadyTracked) {
                                if (-not $hostCumulatives[$prevMatchKey]) { $hostCumulatives[$prevMatchKey] = @() }
                                $hostCumulatives[$prevMatchKey] += @{
                                    Product   = $productFamily
                                    YearMonth = $yearMonth
                                    KB        = $kb
                                    Title     = $title
                                }
                            }
                        }
                    }
                }
            }
        }

        Write-Host "  Added $prevAdded new KB entries from previous logs (skipped $prevSkipped duplicates)"
        $hostsWithCumulatives = ($hostCumulatives.Keys | Measure-Object).Count
        Write-Host "  Total hosts with cumulative updates tracked: $hostsWithCumulatives" -ForegroundColor DarkCyan
        Write-Host "  Total unique matchkey+KB combinations: $($updateLookup.Count)"
    } else {
        Write-Host "`n  No previous_updatelog files found in session directory" -ForegroundColor Yellow
    }
}

# ============================================================================
# READ INSTALLED_HOTFIXES.CSV (full hotfix history from Get-HotFix)
# ============================================================================

# Refresh known cumulatives map (previous logs may have added more)
if ($IncludePreviousLogs) {
    foreach ($entries in $hostCumulatives.Values) {
        foreach ($cum in $entries) {
            if (-not $knownCumulativeKBs.ContainsKey($cum.KB)) {
                $knownCumulativeKBs[$cum.KB] = @{
                    Product   = $cum.Product
                    YearMonth = $cum.YearMonth
                    Title     = $cum.Title
                }
            }
        }
    }
}

$sessionDir = if ($IncludePreviousLogs) { $sessionDir } else { Split-Path $UpdatesCsvPath -Parent }
$hotfixCsvPath = Join-Path $sessionDir "installed_hotfixes.csv"
if (Test-Path $hotfixCsvPath) {
    Write-Host "`n=== Reading installed hotfix history ===" -ForegroundColor Cyan
    $hotfixData = Import-Csv $hotfixCsvPath
    $hfAdded = 0
    $hfSkipped = 0
    $hfCumAdded = 0
    foreach ($hf in $hotfixData) {
        $kb = ($hf.KB -replace '\s','').ToUpper()
        if (-not $kb) { continue }
        $ip = if ($hf.ComputerIP) { $hf.ComputerIP.Trim() } else { $null }
        $hfMatchKey = if ($ip) { $ip } else { (Get-NormalizedHostname $hf.ComputerName) }
        $key = "$hfMatchKey|$kb"

        if (-not $updateLookup.ContainsKey($key)) {
            $updateLookup[$key] = "Installed"
            $hfAdded++
            if ($ip) { $updateIPs[$ip] = $hf.ComputerName }
            $updateHosts[(Get-NormalizedHostname $hf.ComputerName)] = $true
        } else {
            $hfSkipped++
        }

        # If this KB is a known cumulative, track it for supersession on this host
        if ($knownCumulativeKBs.ContainsKey($kb)) {
            $cumInfo = $knownCumulativeKBs[$kb]
            if (-not $hostCumulatives[$hfMatchKey]) { $hostCumulatives[$hfMatchKey] = @() }
            $alreadyTracked = $hostCumulatives[$hfMatchKey] | Where-Object { $_.KB -eq $kb }
            if (-not $alreadyTracked) {
                $hostCumulatives[$hfMatchKey] += @{
                    Product   = $cumInfo.Product
                    YearMonth = $cumInfo.YearMonth
                    KB        = $kb
                    Title     = $cumInfo.Title
                }
                $hfCumAdded++
            }
        }
        # Detect cumulatives from hotfix metadata (Description + InstalledOn)
        # When the all_updates.csv Title lacks a date pattern (e.g. "Windows 11, version 25H2"),
        # use the Get-HotFix InstalledOn date + Description="Security Update" to identify cumulatives
        elseif ($hf.Description -eq 'Security Update' -and $hf.InstalledOn) {
            $hfTitle = $kbTitleMap[$kb]
            # Default to "Windows" when no title available — "Security Update" from Get-HotFix
            # on Windows machines is the monthly cumulative rollup in the vast majority of cases.
            # The InstalledOn date drives the actual supersession comparison.
            $productFamily = if ($hfTitle) { Get-CumulativeProductFamily $hfTitle } else { "Windows" }
            if ($productFamily -and $hf.InstalledOn -match '^(\d{4})-(\d{1,2})') {
                $hfYearMonth = "{0}-{1:D2}" -f $Matches[1], [int]$Matches[2]
                if (-not $hostCumulatives[$hfMatchKey]) { $hostCumulatives[$hfMatchKey] = @() }
                $alreadyTracked = $hostCumulatives[$hfMatchKey] | Where-Object { $_.KB -eq $kb }
                if (-not $alreadyTracked) {
                    $hostCumulatives[$hfMatchKey] += @{
                        Product   = $productFamily
                        YearMonth = $hfYearMonth
                        KB        = $kb
                        Title     = if ($hfTitle) { $hfTitle } else { "Security Update $kb" }
                    }
                    $hfCumAdded++
                    # Also register in knownCumulativeKBs so other hosts can inherit
                    if (-not $knownCumulativeKBs.ContainsKey($kb)) {
                        $knownCumulativeKBs[$kb] = @{
                            Product   = $productFamily
                            YearMonth = $hfYearMonth
                            Title     = if ($hfTitle) { $hfTitle } else { "Security Update $kb" }
                        }
                    }
                }
            }
        }
    }
    Write-Host "  Added $hfAdded KB entries from hotfix history (skipped $hfSkipped already known)"
    if ($hfCumAdded -gt 0) {
        Write-Host "  Cumulative supersession extended to $hfCumAdded additional host+KB combinations" -ForegroundColor DarkCyan
    }
    Write-Host "  Total unique matchkey+KB combinations: $($updateLookup.Count)"
} else {
    Write-Host "`n  No installed_hotfixes.csv found (run update script with latest version to generate)" -ForegroundColor DarkGray
}

# ============================================================================
# BUILD IP ALIAS MAP (multi-NIC matching from computer_summary.csv)
# ============================================================================

$ipAliases = @{}   # any IP on machine → deployment IP (the primary match key)
$machineHostnames = @{}  # deployment IP → actual machine hostname

$computerSumPath2 = Join-Path $sessionDir "computer_summary.csv"
if (Test-Path $computerSumPath2) {
    $compSummary = Import-Csv $computerSumPath2
    $hasAllIPs = ($compSummary[0].PSObject.Properties.Name -contains "AllIPAddresses")
    $hasMachineHostname = ($compSummary[0].PSObject.Properties.Name -contains "MachineHostname")

    if ($hasAllIPs -or $hasMachineHostname) {
        Write-Host "`n=== Building IP alias and hostname maps ===" -ForegroundColor Cyan
    }

    foreach ($cs in $compSummary) {
        $deployIP = $cs.IP.Trim()
        if (-not $deployIP) { continue }

        # Ensure friendly name (ComputerName from deployment CSV) is in the IP lookup
        if ($cs.ComputerName -and -not $updateIPs[$deployIP]) {
            $updateIPs[$deployIP] = $cs.ComputerName.Trim()
        }

        # Build IP alias map: every IP on this machine maps back to the deployment IP
        if ($hasAllIPs -and $cs.AllIPAddresses) {
            $allIPs = $cs.AllIPAddresses -split ','
            foreach ($altIP in $allIPs) {
                $altIP = $altIP.Trim()
                if ($altIP -and $altIP -ne $deployIP) {
                    $ipAliases[$altIP] = $deployIP
                }
            }
        }

        # Build machine hostname map
        if ($hasMachineHostname -and $cs.MachineHostname) {
            $machineHostnames[$deployIP] = $cs.MachineHostname.Trim()
            # Also allow matching by machine hostname
            $mhLower = (Get-NormalizedHostname $cs.MachineHostname)
            if ($mhLower) { $ipAliases[$mhLower] = $deployIP }
        }
    }

    if ($hasAllIPs) {
        Write-Host "  IP aliases registered: $($ipAliases.Count) (alternate IPs and hostnames → deployment IP)"
    }
    if ($hasMachineHostname) {
        Write-Host "  Machine hostnames loaded: $($machineHostnames.Count)"
    }
}

# ============================================================================
# SESSION-VERIFIED CUMULATIVE COVERAGE
# When a host has Status=NoUpdatesNeeded or Completed but no cumulative tracked
# in $hostCumulatives, use the session date as proof of coverage.
# Get-HotFix (Win32_QuickFixEngineering) doesn't always list the latest
# cumulative, but PSWindowsUpdate's status confirms the machine is current.
# ============================================================================

if ($sessionYearMonth -and (Test-Path $computerSumPath2)) {
    $sessionVerified = 0
    foreach ($cs in $compSummary) {
        $csStatus = $cs.Status.Trim()
        if ($csStatus -notin @('NoUpdatesNeeded', 'Completed')) { continue }

        $deployIP = $cs.IP.Trim()
        if (-not $deployIP) { continue }

        # Only add if this host doesn't already have a Windows cumulative for the session month or newer
        $existing = $hostCumulatives[$deployIP]
        $alreadyCovered = $false
        if ($existing) {
            foreach ($cum in $existing) {
                if ($cum.Product -eq 'Windows' -and $cum.YearMonth -ge $sessionYearMonth) {
                    $alreadyCovered = $true
                    break
                }
            }
        }

        if (-not $alreadyCovered) {
            if (-not $hostCumulatives[$deployIP]) { $hostCumulatives[$deployIP] = @() }
            $hostCumulatives[$deployIP] += @{
                Product   = 'Windows'
                YearMonth = $sessionYearMonth
                KB        = '(session-verified)'
                Title     = "Session status: $csStatus"
            }
            $sessionVerified++
        }
    }

    if ($sessionVerified -gt 0) {
        Write-Host "`n=== Session-verified cumulative coverage ===" -ForegroundColor Cyan
        Write-Host "  $sessionVerified hosts verified current via session status (NoUpdatesNeeded/Completed)" -ForegroundColor DarkCyan
        Write-Host "  Coverage month: $sessionYearMonth (from session folder name)" -ForegroundColor DarkCyan
    }
}

# ============================================================================
# READ QUALYS VULNERABILITY REPORT (XLSX via COM)
# ============================================================================

Write-Host "`n=== Reading Qualys vulnerability report ===" -ForegroundColor Cyan

# Initialize collections outside try block so they survive COM failures
$vulnData = @()
$vulnHosts = @{}
$vulnIPs = @{}
$skippedEmpty = 0
$cumulativeRemediations = 0
$noResultsHosts = @()       # Hosts where Qualys had "No results available"

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "  Opening: $VulnReportPath"
    $wb = $excel.Workbooks.Open($VulnReportPath)
    if (-not $wb) {
        Write-Host "ERROR: Excel failed to open the workbook. Ensure the file is not locked." -ForegroundColor Red
        exit 1
    }
    $ws = $wb.Sheets.Item(1)
    $lastRow = $ws.UsedRange.Rows.Count

    Write-Host "  Sheet: $($ws.Name), Rows: $($lastRow - 1) (excluding header)"

    # Column indices (1-based) — defaults, may be overridden by header detection
    $colIP       = 1
    $colDNS      = 3
    $colNetBIOS  = 4
    $colTitle    = 9
    $colSeverity = 12
    $colCVE      = 21
    $colSolution = 26
    $colResults  = 29
    $colOS       = 6   # Column F: OS info, or "No results available" summary rows

    # ── Auto-detect column positions from header row ──
    $lastCol = $ws.UsedRange.Columns.Count
    $headerMap = @{
        'IP'       = 'colIP'
        'DNS'      = 'colDNS'
        'NetBIOS'  = 'colNetBIOS'
        'Title'    = 'colTitle'
        'Severity' = 'colSeverity'
        'CVE ID'   = 'colCVE'
        'Solution' = 'colSolution'
        'Results'  = 'colResults'
        'OS'       = 'colOS'
    }
    $detected = @{}
    for ($c = 1; $c -le [Math]::Min($lastCol, 40); $c++) {
        $headerText = "$($ws.Cells.Item(1, $c).Text)".Trim()
        foreach ($hName in $headerMap.Keys) {
            if ($headerText -eq $hName -or $headerText -ieq $hName) {
                $detected[$hName] = $c
            }
        }
    }

    # Apply detected positions if we found the critical ones (IP + Results at minimum)
    if ($detected.ContainsKey('IP') -and $detected.ContainsKey('Results')) {
        Write-Host "  Auto-detected column positions from header row" -ForegroundColor DarkCyan
        if ($detected['IP'])       { $colIP       = $detected['IP'] }
        if ($detected['DNS'])      { $colDNS      = $detected['DNS'] }
        if ($detected['NetBIOS'])  { $colNetBIOS  = $detected['NetBIOS'] }
        if ($detected['Title'])    { $colTitle     = $detected['Title'] }
        if ($detected['Severity']) { $colSeverity  = $detected['Severity'] }
        if ($detected['CVE ID'])   { $colCVE       = $detected['CVE ID'] }
        if ($detected['Solution']) { $colSolution  = $detected['Solution'] }
        if ($detected['Results'])  { $colResults   = $detected['Results'] }
        if ($detected['OS'])       { $colOS        = $detected['OS'] }

        $missingHeaders = $headerMap.Keys | Where-Object { -not $detected.ContainsKey($_) }
        if ($missingHeaders) {
            Write-Host "  Warning: Headers not found (using defaults): $($missingHeaders -join ', ')" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  Using hardcoded column positions (header auto-detection did not find IP+Results)" -ForegroundColor Yellow
    }

    Write-Host "  Columns: IP=$colIP DNS=$colDNS NetBIOS=$colNetBIOS Title=$colTitle Severity=$colSeverity CVE=$colCVE Solution=$colSolution Results=$colResults"

    # ── Pass 1: Read all rows and build global KB→date map ──
    $monthNames = @{
        'january'=1;'february'=2;'march'=3;'april'=4;'may'=5;'june'=6
        'july'=7;'august'=8;'september'=9;'october'=10;'november'=11;'december'=12
    }
    $kbDateMap = @{}       # KB → earliest known YYYY-MM date context
    $rawRows = @()         # Store parsed row data for second pass

    Write-Host "  Pass 1: Reading rows and building KB date map..."
    for ($r = 2; $r -le $lastRow; $r++) {
        $ip      = "$($ws.Cells.Item($r, $colIP).Text)".Trim()
        $colFText = "$($ws.Cells.Item($r, $colOS).Text)".Trim()

        # Detect Qualys summary rows at the bottom: "No results available for these hosts"
        # or "No vulnerabilities match your filters for these hosts"
        # These rows have comma-separated IPs in column A and the status message in column F
        if ($colFText -match 'No results available|No vulnerabilities match') {
            # Parse comma-separated IPs (may include ranges like "48.40.19.205-48.40.19.206")
            $ipList = $ip -split ',\s*' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            foreach ($noResIP in $ipList) {
                # Skip IP ranges (e.g., "48.40.19.205-48.40.19.206") — just take first IP
                $singleIP = ($noResIP -split '-')[0].Trim()
                if ($singleIP -match '^\d+\.\d+\.\d+\.\d+$') {
                    $noResultsHosts += [PSCustomObject]@{
                        IP     = $singleIP
                        Reason = $colFText
                    }
                    # Mark as seen in scan so deployment-side unmatched check doesn't double-list
                    if (-not $vulnIPs.ContainsKey($singleIP)) {
                        $vulnIPs[$singleIP] = "(no scan results)"
                    }
                }
            }
            continue
        }

        $dns     = "$($ws.Cells.Item($r, $colDNS).Text)".Trim()
        $netbios = "$($ws.Cells.Item($r, $colNetBIOS).Text)".Trim()
        $title   = "$($ws.Cells.Item($r, $colTitle).Text)".Trim()
        $sev     = "$($ws.Cells.Item($r, $colSeverity).Text)".Trim()
        $cve     = "$($ws.Cells.Item($r, $colCVE).Text)".Trim()
        $results = "$($ws.Cells.Item($r, $colResults).Text)".Trim()
        $solution = "$($ws.Cells.Item($r, $colSolution).Text)".Trim()

        # Determine hostname
        if ($HostnameColumn -eq "DNS") {
            $hostname = if ($dns) { $dns } else { $netbios }
        } else {
            $hostname = if ($netbios) { $netbios } else { $dns }
        }
        if (-not $hostname -and -not $title) { $skippedEmpty++; continue }

        $hostLower = Get-NormalizedHostname $hostname
        # Resolve IP aliases: if Qualys scanned a different interface, map to deployment IP
        $resolvedIP = if ($ip -and $ipAliases[$ip]) { $ipAliases[$ip] } else { $ip }
        $matchKey = if ($resolvedIP) { $resolvedIP } elseif ($ipAliases[$hostLower]) { $ipAliases[$hostLower] } else { $hostLower }
        $vulnIPs[$ip] = $hostname
        $vulnHosts[$hostLower] = $true

        # Extract required KBs from Results column (primary source)
        $requiredKBs = @()
        if ($results -match 'Missing') {
            $kbMatches = [regex]::Matches($results, 'Missing\s+(?:Hot)?Patch/KB:\s*(KB\s*\d+)', 'IgnoreCase')
            foreach ($m in $kbMatches) {
                $kb = ($m.Groups[1].Value -replace '\s','').ToUpper()
                if ($requiredKBs -notcontains $kb) { $requiredKBs += $kb }
            }
        }

        # Fallback: extract KBs from Solution column if Results didn't yield any
        if ($requiredKBs.Count -eq 0 -and $solution) {
            $solKBs = [regex]::Matches($solution, '(KB\s*\d{6,7})', 'IgnoreCase')
            foreach ($m in $solKBs) {
                $kb = ($m.Groups[1].Value -replace '\s','').ToUpper()
                if ($requiredKBs -notcontains $kb) { $requiredKBs += $kb }
            }
        }

        # Extract date context from this row and map to KBs
        $rowDateMonth = $null
        # Try title first: "Month YYYY" pattern (most reliable — never contains CVE numbers)
        if ($title -match '(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{4})') {
            $mName = $Matches[1].ToLower()
            $yr = $Matches[2]
            $rowDateMonth = "{0}-{1:D2}" -f $yr, $monthNames[$mName]
        }
        # Try Results column: YYYY-MM pattern (validate month 1-12 to exclude CVE numbers like CVE-2026-20805)
        if (-not $rowDateMonth -and $results -match '(\d{4})-(\d{1,2})' -and [int]$Matches[2] -ge 1 -and [int]$Matches[2] -le 12) {
            $rowDateMonth = "{0}-{1:D2}" -f $Matches[1], [int]$Matches[2]
        }
        # Try Solution column: YYYY-MM pattern (validate month 1-12)
        if (-not $rowDateMonth -and $solution -match '(\d{4})-(\d{1,2})' -and [int]$Matches[2] -ge 1 -and [int]$Matches[2] -le 12) {
            $rowDateMonth = "{0}-{1:D2}" -f $Matches[1], [int]$Matches[2]
        }

        # Map each KB to its date context (keep the earliest/oldest date per KB)
        if ($rowDateMonth) {
            foreach ($kb in $requiredKBs) {
                if (-not $kbDateMap[$kb] -or $kbDateMap[$kb] -gt $rowDateMonth) {
                    $kbDateMap[$kb] = $rowDateMonth
                }
            }
        }

        # Store parsed data for second pass
        $rawRows += [PSCustomObject]@{
            Hostname    = $hostname
            IP          = $ip
            Title       = $title
            Severity    = $sev
            CVE         = $cve
            Results     = $results
            Solution    = $solution
            MatchKey    = $matchKey
            RequiredKBs = $requiredKBs
        }
    }

    Write-Host "  Built date map for $($kbDateMap.Count) unique KBs"

    # ── Pass 2: Determine remediation status using global KB date map ──
    Write-Host "  Pass 2: Correlating vulnerabilities with updates..."
    foreach ($row in $rawRows) {
        $matchKey = $row.MatchKey
        $requiredKBs = $row.RequiredKBs

        $remediationStatus = "Manual Review"
        $kbDetails = @()
        $cumulativeCoverage = ""

        if ($requiredKBs.Count -gt 0) {
            $allInstalled = $true
            $anyNotCovered = $false

            foreach ($kb in $requiredKBs) {
                $key = "$matchKey|$kb"
                $updateStatus = $updateLookup[$key]
                if ($updateStatus -eq "Installed" -or $updateStatus -eq "Pending") {
                    $detailLabel = if ($updateStatus -eq "Pending") { "Installed (Pending Reboot)" } else { "Installed" }
                    $kbDetails += "$kb=$detailLabel"
                } else {
                    $allInstalled = $false

                    # ── Cumulative supersession check ──
                    $coveredByCumulative = $false
                    if ($hostCumulatives[$matchKey]) {
                        # Use global KB date map (populated across ALL Qualys entries)
                        $kbDateMonth = $kbDateMap[$kb]

                        # Fallback: try row-specific date extraction (Title first, most reliable)
                        if (-not $kbDateMonth -and $row.Title -match '(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{4})') {
                            $mName = $Matches[1].ToLower()
                            $yr = $Matches[2]
                            $kbDateMonth = "{0}-{1:D2}" -f $yr, $monthNames[$mName]
                        }
                        if (-not $kbDateMonth) {
                            if ($row.Results -match "(\d{4})-(\d{1,2}).*$([regex]::Escape($kb))|$([regex]::Escape($kb)).*(\d{4})-(\d{1,2})") {
                                $year = if ($Matches[1]) { $Matches[1] } else { $Matches[3] }
                                $month = if ($Matches[2]) { $Matches[2] } else { $Matches[4] }
                                # Validate month 1-12 to exclude CVE numbers (e.g. CVE-2026-20805)
                                if ([int]$month -ge 1 -and [int]$month -le 12) {
                                    $kbDateMonth = "{0}-{1:D2}" -f $year, [int]$month
                                }
                            }
                        }
                        if (-not $kbDateMonth -and $row.Solution -match '(\d{4})-(\d{1,2})' -and [int]$Matches[2] -ge 1 -and [int]$Matches[2] -le 12) {
                            $kbDateMonth = "{0}-{1:D2}" -f $Matches[1], [int]$Matches[2]
                        }

                        if ($kbDateMonth) {
                            # Determine the product family of the missing KB from Qualys context
                            $missingProduct = Get-CumulativeProductFamily $row.Title
                            if (-not $missingProduct -and $row.Solution) {
                                $missingProduct = Get-CumulativeProductFamily $row.Solution
                            }
                            # Default: if we can't determine family, only allow Windows cumulatives
                            # (most vulns are OS-level; this prevents .NET cumulatives from falsely covering OS KBs)
                            if (-not $missingProduct) { $missingProduct = "Windows" }

                            # Check if any installed cumulative for this host is >= the KB's month AND same product family
                            foreach ($cum in $hostCumulatives[$matchKey]) {
                                if ($cum.Product -eq $missingProduct -and $cum.YearMonth -ge $kbDateMonth) {
                                    $coveredByCumulative = $true
                                    $cumulativeCoverage = "$($cum.KB) ($($cum.YearMonth))"
                                    break
                                }
                            }
                        }
                    }

                    if ($coveredByCumulative) {
                        $kbDetails += "$kb=Superseded by $cumulativeCoverage"
                    } elseif ($updateStatus) {
                        $kbDetails += "$kb=$updateStatus"
                        $anyNotCovered = $true
                    } else {
                        $kbDetails += "$kb=Not Found"
                        $anyNotCovered = $true
                    }
                }
            }

            if ($allInstalled) {
                $remediationStatus = "Remediated"
            } elseif (-not $anyNotCovered) {
                $remediationStatus = "Remediated (Cumulative)"
                $cumulativeRemediations++
            } else {
                $remediationStatus = "Not Remediated"
            }
        }

        $vulnData += [PSCustomObject]@{
            Hostname          = $row.Hostname
            IP                = $row.IP
            Title             = $row.Title
            Severity          = $row.Severity
            CVE               = $row.CVE
            RequiredKBs       = ($requiredKBs -join ", ")
            RemediationStatus = $remediationStatus
            KBDetails         = ($kbDetails -join "; ")
        }
    }

    Write-Host "  Parsed $($vulnData.Count) vulnerability entries across $($vulnHosts.Count) hosts"
    if ($skippedEmpty -gt 0) { Write-Host "  Skipped $skippedEmpty empty rows" }
    if ($cumulativeRemediations -gt 0) {
        Write-Host "  Cumulative supersession matches: $cumulativeRemediations" -ForegroundColor DarkCyan
    }
    if ($noResultsHosts.Count -gt 0) {
        Write-Host "  Hosts with no scan results: $($noResultsHosts.Count)" -ForegroundColor Yellow
    }

} finally {
    try { if ($wb) { $wb.Close($false) } } catch {}
    try {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ============================================================================
# BUILD SUMMARIES
# ============================================================================

# --- Host Summary (grouped by IP for accurate matching) ---
$hostSummary = @()
$groupedByIP = $vulnData | Group-Object IP
foreach ($g in $groupedByIP) {
    $total      = $g.Count
    $remediated = ($g.Group | Where-Object { $_.RemediationStatus -eq "Remediated" }).Count
    $cumRemediated = ($g.Group | Where-Object { $_.RemediationStatus -eq "Remediated (Cumulative)" }).Count
    $notRemediated = ($g.Group | Where-Object { $_.RemediationStatus -eq "Not Remediated" }).Count
    $manual     = ($g.Group | Where-Object { $_.RemediationStatus -eq "Manual Review" }).Count
    $totalRemediated = $remediated + $cumRemediated
    $pct        = if ($total -gt 0) { [math]::Round(($totalRemediated / $total) * 100, 1) } else { 0 }

    # Find latest cumulative for this host (keyed by IP)
    $groupIP = $g.Name
    $latestCum = ""
    if ($hostCumulatives[$groupIP]) {
        $latest = $hostCumulatives[$groupIP] | Sort-Object { $_.YearMonth } -Descending | Select-Object -First 1
        $latestCum = "$($latest.KB) ($($latest.YearMonth))"
    }
    # Use the hostname from the first entry for display
    $displayName = ($g.Group | Select-Object -First 1).Hostname

    # Collect unique missing KBs from "Not Remediated" entries for this host
    $missingKBs = @()
    $notRemEntries = $g.Group | Where-Object { $_.RemediationStatus -eq "Not Remediated" }
    foreach ($nre in $notRemEntries) {
        if ($nre.RequiredKBs) {
            foreach ($kb in ($nre.RequiredKBs -split ',\s*')) {
                $kb = $kb.Trim().ToUpper()
                if ($kb -and $missingKBs -notcontains $kb) { $missingKBs += $kb }
            }
        }
    }

    # Also collect unique vuln titles from "Manual Review" entries (no KB, need other action)
    $manualItems = @()
    $manualEntries = $g.Group | Where-Object { $_.RemediationStatus -eq "Manual Review" }
    foreach ($mre in $manualEntries) {
        if ($mre.Title -and $manualItems -notcontains $mre.Title) { $manualItems += $mre.Title }
    }

    # Build combined outstanding items string
    $outstandingParts = @()
    if ($missingKBs.Count -gt 0) { $outstandingParts += $missingKBs -join ", " }
    if ($manualItems.Count -gt 0) { $outstandingParts += $manualItems -join "; " }
    $missingKBsStr = $outstandingParts -join " | "

    # Look up the friendly site name from deployment CSV
    $siteName = if ($updateIPs[$groupIP]) { $updateIPs[$groupIP] } else { "" }

    $hostSummary += [PSCustomObject]@{
        Hostname         = $displayName
        SiteName         = $siteName
        IP               = $groupIP
        TotalVulns       = $total
        Remediated       = $remediated
        CumRemediated    = $cumRemediated
        NotRemediated    = $notRemediated
        ManualReview     = $manual
        PctRemediated    = $pct
        LatestCumulative = $latestCum
        MissingKBs       = $missingKBsStr
    }
}
$hostSummary = $hostSummary | Sort-Object PctRemediated

# --- Unmatched Hosts (by IP) ---
$unmatchedVuln = @()
$unmatchedUpdate = @()
foreach ($vip in $vulnIPs.Keys) {
    if ($vip -and -not $updateIPs[$vip]) {
        $unmatchedVuln += [PSCustomObject]@{ Hostname = $vulnIPs[$vip]; IP = $vip; Source = "Vuln Scan"; Note = "No matching IP in all_updates.csv" }
    }
}
foreach ($uip in $updateIPs.Keys) {
    if ($uip -and -not $vulnIPs[$uip]) {
        $unmatchedUpdate += [PSCustomObject]@{ Hostname = $updateIPs[$uip]; IP = $uip; Source = "Updates CSV"; Note = "No matching IP in vuln scan" }
    }
}
# Add "No results" hosts from Qualys summary rows
$unmatchedNoResults = @()
foreach ($nrh in $noResultsHosts) {
    $unmatchedNoResults += [PSCustomObject]@{ Hostname = ""; IP = $nrh.IP; Source = "Vuln Scan"; Note = $nrh.Reason }
}
$unmatchedAll = $unmatchedVuln + $unmatchedUpdate + $unmatchedNoResults

# ============================================================================
# CONSOLE SUMMARY
# ============================================================================

$totalVulns = $vulnData.Count
$totalRemediated = ($vulnData | Where-Object { $_.RemediationStatus -eq "Remediated" }).Count
$totalCumRemediated = ($vulnData | Where-Object { $_.RemediationStatus -eq "Remediated (Cumulative)" }).Count
$totalNotRemediated = ($vulnData | Where-Object { $_.RemediationStatus -eq "Not Remediated" }).Count
$totalManual = ($vulnData | Where-Object { $_.RemediationStatus -eq "Manual Review" }).Count

Write-Host "`n=== Correlation Summary ===" -ForegroundColor Green
Write-Host "  Total vulnerability entries : $totalVulns"
Write-Host "  Remediated (exact KB)       : $totalRemediated" -ForegroundColor Green
Write-Host "  Remediated (cumulative)     : $totalCumRemediated" -ForegroundColor DarkCyan
Write-Host "  Not Remediated              : $totalNotRemediated" -ForegroundColor Red
Write-Host "  Manual Review (no KB info)  : $totalManual" -ForegroundColor Yellow
Write-Host "  Unmatched hosts             : $($unmatchedAll.Count)" -ForegroundColor $(if ($unmatchedAll.Count -gt 0) { "Yellow" } else { "Green" })
if ($noResultsHosts.Count -gt 0) {
    Write-Host "  No scan results (Qualys)    : $($noResultsHosts.Count) IPs" -ForegroundColor Yellow
}

if ($totalVulns -gt 0) {
    $overallPct = [math]::Round((($totalRemediated + $totalCumRemediated) / $totalVulns) * 100, 1)
    Write-Host "`n  Overall remediation rate    : $overallPct%" -ForegroundColor $(if ($overallPct -ge 90) { "Green" } elseif ($overallPct -ge 50) { "Yellow" } else { "Red" })
}

# ============================================================================
# WRITE OUTPUT REPORT (XLSX via COM)
# ============================================================================

Write-Host "`n=== Writing report ===" -ForegroundColor Cyan

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $outWb = $excel.Workbooks.Add()

    # --- Sheet 1: Host Summary ---
    $ws1 = $outWb.Sheets.Item(1)
    $ws1.Name = "Host Summary"
    $headers1 = @("Hostname", "Site Name", "IP", "Total Vulns", "Remediated", "Remediated (Cumulative)", "Not Remediated", "Manual Review", "% Remediated", "Latest Cumulative", "Outstanding Items")
    for ($c = 0; $c -lt $headers1.Count; $c++) {
        $ws1.Cells.Item(1, $c + 1) = $headers1[$c]
        $ws1.Cells.Item(1, $c + 1).Font.Bold = $true
        $ws1.Cells.Item(1, $c + 1).Interior.Color = 0x783C28  # Dark blue
        $ws1.Cells.Item(1, $c + 1).Font.Color = 0xFFFFFF       # White text
    }
    $row = 2
    foreach ($h in $hostSummary) {
        $ws1.Cells.Item($row, 1) = $h.Hostname
        $ws1.Cells.Item($row, 2) = $h.SiteName
        $ws1.Cells.Item($row, 3) = $h.IP
        $ws1.Cells.Item($row, 4) = $h.TotalVulns
        $ws1.Cells.Item($row, 5) = $h.Remediated
        $ws1.Cells.Item($row, 6) = $h.CumRemediated
        $ws1.Cells.Item($row, 7) = $h.NotRemediated
        $ws1.Cells.Item($row, 8) = $h.ManualReview
        $ws1.Cells.Item($row, 9) = "$($h.PctRemediated)%"
        $ws1.Cells.Item($row, 10) = $h.LatestCumulative
        $ws1.Cells.Item($row, 11) = $h.MissingKBs
        # Color code the % column
        if ($h.PctRemediated -ge 90) {
            $ws1.Cells.Item($row, 9).Interior.Color = 0x90EE90  # Light green
        } elseif ($h.PctRemediated -ge 50) {
            $ws1.Cells.Item($row, 9).Interior.Color = 0xADDCFF  # Light yellow (BGR)
        } else {
            $ws1.Cells.Item($row, 9).Interior.Color = 0x8080FF  # Light red (BGR)
        }
        $row++
    }
    $ws1.Columns.Item("A:K").AutoFit() | Out-Null
    # Cap Outstanding Items column width so it doesn't get absurdly wide
    if ($ws1.Columns.Item("K").ColumnWidth -gt 80) { $ws1.Columns.Item("K").ColumnWidth = 80 }

    # --- Sheet 2: Vulnerability Detail ---
    $ws2 = $outWb.Sheets.Add([System.Reflection.Missing]::Value, $ws1)
    $ws2.Name = "Vulnerability Detail"
    $headers2 = @("Hostname", "IP", "CVE", "Title", "Severity", "Required KBs", "Remediation Status", "KB Details")
    for ($c = 0; $c -lt $headers2.Count; $c++) {
        $ws2.Cells.Item(1, $c + 1) = $headers2[$c]
        $ws2.Cells.Item(1, $c + 1).Font.Bold = $true
        $ws2.Cells.Item(1, $c + 1).Interior.Color = 0x783C28
        $ws2.Cells.Item(1, $c + 1).Font.Color = 0xFFFFFF
    }
    $row = 2
    foreach ($v in ($vulnData | Sort-Object Hostname, Severity, CVE)) {
        $ws2.Cells.Item($row, 1) = $v.Hostname
        $ws2.Cells.Item($row, 2) = $v.IP
        $ws2.Cells.Item($row, 3) = $v.CVE
        $ws2.Cells.Item($row, 4) = $v.Title
        $ws2.Cells.Item($row, 5) = $v.Severity
        $ws2.Cells.Item($row, 6) = $v.RequiredKBs
        $ws2.Cells.Item($row, 7) = $v.RemediationStatus
        $ws2.Cells.Item($row, 8) = $v.KBDetails
        # Color code status
        switch ($v.RemediationStatus) {
            "Remediated"              { $ws2.Cells.Item($row, 7).Interior.Color = 0x90EE90 }  # Green
            "Remediated (Cumulative)" { $ws2.Cells.Item($row, 7).Interior.Color = 0xFFE0B0 }  # Light blue (BGR)
            "Not Remediated"          { $ws2.Cells.Item($row, 7).Interior.Color = 0x8080FF }  # Red
            "Manual Review"           { $ws2.Cells.Item($row, 7).Interior.Color = 0xADDCFF }  # Yellow
        }
        $row++
    }
    $ws2.Columns.Item("A:H").AutoFit() | Out-Null
    # Cap column D (Title) width
    if ($ws2.Columns.Item("D").ColumnWidth -gt 60) { $ws2.Columns.Item("D").ColumnWidth = 60 }

    # --- Sheet 3: Unmatched Hosts ---
    $ws3 = $outWb.Sheets.Add([System.Reflection.Missing]::Value, $ws2)
    $ws3.Name = "Unmatched Hosts"
    $headers3 = @("Hostname", "IP", "Source", "Note")
    for ($c = 0; $c -lt $headers3.Count; $c++) {
        $ws3.Cells.Item(1, $c + 1) = $headers3[$c]
        $ws3.Cells.Item(1, $c + 1).Font.Bold = $true
        $ws3.Cells.Item(1, $c + 1).Interior.Color = 0x783C28
        $ws3.Cells.Item(1, $c + 1).Font.Color = 0xFFFFFF
    }
    if ($unmatchedAll.Count -gt 0) {
        $row = 2
        foreach ($u in $unmatchedAll) {
            $ws3.Cells.Item($row, 1) = $u.Hostname
            $ws3.Cells.Item($row, 2) = $u.IP
            $ws3.Cells.Item($row, 3) = $u.Source
            $ws3.Cells.Item($row, 4) = $u.Note
            $row++
        }
    } else {
        $ws3.Cells.Item(2, 1) = "All hosts matched successfully"
        $ws3.Cells.Item(2, 1).Font.Color = 0x008000
    }
    $ws3.Columns.Item("A:D").AutoFit() | Out-Null

    # --- Sheet 4: Cumulative Coverage ---
    $ws4 = $outWb.Sheets.Add([System.Reflection.Missing]::Value, $ws3)
    $ws4.Name = "Cumulative Coverage"
    $headers4 = @("Hostname", "Product Family", "Year-Month", "KB", "Title")
    for ($c = 0; $c -lt $headers4.Count; $c++) {
        $ws4.Cells.Item(1, $c + 1) = $headers4[$c]
        $ws4.Cells.Item(1, $c + 1).Font.Bold = $true
        $ws4.Cells.Item(1, $c + 1).Interior.Color = 0x783C28
        $ws4.Cells.Item(1, $c + 1).Font.Color = 0xFFFFFF
    }
    $row = 2
    $cumDataWritten = $false
    foreach ($hostKey in ($hostCumulatives.Keys | Sort-Object)) {
        foreach ($cum in ($hostCumulatives[$hostKey] | Sort-Object { $_.YearMonth } -Descending)) {
            $ws4.Cells.Item($row, 1) = $hostKey
            $ws4.Cells.Item($row, 2) = $cum.Product
            $ws4.Cells.Item($row, 3) = $cum.YearMonth
            $ws4.Cells.Item($row, 4) = $cum.KB
            $ws4.Cells.Item($row, 5) = $cum.Title
            $row++
            $cumDataWritten = $true
        }
    }
    if (-not $cumDataWritten) {
        $ws4.Cells.Item(2, 1) = "No cumulative updates found in all_updates.csv"
        $ws4.Cells.Item(2, 1).Font.Color = 0x0000FF
    }
    $ws4.Columns.Item("A:E").AutoFit() | Out-Null
    if ($ws4.Columns.Item("E").ColumnWidth -gt 80) { $ws4.Columns.Item("E").ColumnWidth = 80 }

    # Remove any extra default sheets
    while ($outWb.Sheets.Count -gt 4) {
        $outWb.Sheets.Item($outWb.Sheets.Count).Delete()
    }

    # Select the first sheet (Host Summary) so the report opens on it
    $outWb.Sheets.Item(1).Activate()

    # Save
    if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
    $outWb.SaveAs($OutputPath, 51)  # 51 = xlOpenXMLWorkbook (.xlsx)
    Write-Host "  Report saved: $OutputPath" -ForegroundColor Green

} finally {
    try { if ($outWb) { $outWb.Close($false) } } catch {}
    try {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ============================================================================
# OPTIONAL CSV EXPORT
# ============================================================================

if ($ExportCsv) {
    $csvPath = [System.IO.Path]::ChangeExtension($OutputPath, ".csv")
    $vulnData | Export-Csv $csvPath -NoTypeInformation
    Write-Host "  CSV exported: $csvPath" -ForegroundColor Green
}

Write-Host "`nDone." -ForegroundColor Green
