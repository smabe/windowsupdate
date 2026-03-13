# Complete Windows Update Deployment Script
# Features: Smart job monitoring, detailed HTML reporting, error handling
# Version: FINAL - All features integrated
 
#Requires -Version 7.0  # For ForEach-Object -Parallel
 
param(
    [Parameter(Mandatory=$false)]
    [string]$CSVPath = "$env:USERPROFILE\Desktop",

    [Parameter(Mandatory=$false)]
    [string]$ReportPath = "C:\WindowsUpdateReports",

    [Parameter(Mandatory=$false)]
    [int]$MaxWaitMinutes = 180,

    [Parameter(Mandatory=$false)]
    [int]$CheckIntervalSeconds = 30,

    # Path on remote machines used for update log and marker files
    [Parameter(Mandatory=$false)]
    [string]$RemoteTempPath = "C:\temp",

    # Seconds the remote update log must be unmodified before treating it as stable (heuristic fallback)
    [Parameter(Mandatory=$false)]
    [int]$LogStabilitySeconds = 60,

    # Max additional attempts (beyond the first) for transient network failures
    [Parameter(Mandatory=$false)]
    [int]$MaxRetries = 2,

    # Maximum number of concurrent Phase 1 job-start operations
    [Parameter(Mandatory=$false)]
    [int]$ThrottleLimit = 50,

    # Regenerate the HTML report from an existing session folder (skips Phases 1-3)
    [switch]$ReportOnly,

    # Path to an existing session folder containing CSVs (required with -ReportOnly)
    [Parameter(Mandatory=$false)]
    [string]$SessionPath,

    # Save credentials to Windows Credential Manager after successful validation
    [switch]$SaveCredential,

    # Remove saved credentials from Windows Credential Manager and exit
    [switch]$ClearCredential
)
 
# ── Windows Credential Manager (native P/Invoke) ──────────────────────────────
$script:WUCredentialTarget = "WindowsUpdateScript"

if (-not ([System.Management.Automation.PSTypeName]'WU.CredentialManager').Type) {
    Add-Type -Namespace 'WU' -Name 'CredentialManager' -MemberDefinition @'
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool CredWrite(ref CREDENTIAL credential, uint flags);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool CredRead(string targetName, uint type, uint flags, out IntPtr credential);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool CredDelete(string targetName, uint type, uint flags);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern void CredFree(IntPtr credential);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct CREDENTIAL {
            public uint Flags;
            public uint Type;
            public string TargetName;
            public string Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public uint CredentialBlobSize;
            public IntPtr CredentialBlob;
            public uint Persist;
            public uint AttributeCount;
            public IntPtr Attributes;
            public string TargetAlias;
            public string UserName;
        }

        public const uint CRED_TYPE_GENERIC = 1;
        public const uint CRED_PERSIST_LOCAL_MACHINE = 2;
'@
}

function Save-WUCredential {
    param([PSCredential]$Credential)
    $passwordBytes = [System.Text.Encoding]::Unicode.GetBytes($Credential.GetNetworkCredential().Password)
    $passwordPtr = [System.Runtime.InteropServices.Marshal]::AllocHGlobal($passwordBytes.Length)
    try {
        [System.Runtime.InteropServices.Marshal]::Copy($passwordBytes, 0, $passwordPtr, $passwordBytes.Length)
        $cred = New-Object WU.CredentialManager+CREDENTIAL
        $cred.Type = [WU.CredentialManager]::CRED_TYPE_GENERIC
        $cred.TargetName = $script:WUCredentialTarget
        $cred.UserName = $Credential.UserName
        $cred.CredentialBlobSize = [uint32]$passwordBytes.Length
        $cred.CredentialBlob = $passwordPtr
        $cred.Persist = [WU.CredentialManager]::CRED_PERSIST_LOCAL_MACHINE
        $cred.Comment = "Windows Update Deployment Script"
        $result = [WU.CredentialManager]::CredWrite([ref]$cred, 0)
        if (-not $result) { throw "CredWrite failed: error $([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())" }
        return $true
    } finally {
        [System.Runtime.InteropServices.Marshal]::FreeHGlobal($passwordPtr)
    }
}

function Get-WUSavedCredential {
    $credPtr = [IntPtr]::Zero
    try {
        $result = [WU.CredentialManager]::CredRead($script:WUCredentialTarget, [WU.CredentialManager]::CRED_TYPE_GENERIC, 0, [ref]$credPtr)
        if (-not $result) { return $null }
        $credStruct = [System.Runtime.InteropServices.Marshal]::PtrToStructure($credPtr, [Type][WU.CredentialManager+CREDENTIAL])
        $password = ""
        if ($credStruct.CredentialBlobSize -gt 0 -and $credStruct.CredentialBlob -ne [IntPtr]::Zero) {
            $passwordBytes = New-Object byte[] $credStruct.CredentialBlobSize
            [System.Runtime.InteropServices.Marshal]::Copy($credStruct.CredentialBlob, $passwordBytes, 0, $credStruct.CredentialBlobSize)
            $password = [System.Text.Encoding]::Unicode.GetString($passwordBytes)
        }
        $secPass = ConvertTo-SecureString $password -AsPlainText -Force
        return New-Object PSCredential($credStruct.UserName, $secPass)
    } finally {
        if ($credPtr -ne [IntPtr]::Zero) { [WU.CredentialManager]::CredFree($credPtr) }
    }
}

function Remove-WUSavedCredential {
    [WU.CredentialManager]::CredDelete($script:WUCredentialTarget, [WU.CredentialManager]::CRED_TYPE_GENERIC, 0) | Out-Null
}

# ── Handle -ClearCredential ──
if ($ClearCredential) {
    $existing = Get-WUSavedCredential
    if ($existing) {
        Remove-WUSavedCredential
        Write-Host "Saved credentials removed from Windows Credential Manager." -ForegroundColor Green
    } else {
        Write-Host "No saved credentials found." -ForegroundColor Yellow
    }
    exit
}

# ── Helper: Test WinRM reachability using Test-WSMan (more reliable than raw TCP) ──
function Test-S2WinRM {
    param([string]$ComputerName)
    try {
        $null = Test-WSMan -ComputerName $ComputerName -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

if ($ReportOnly) {
    # ============================================
    # REPORT-ONLY MODE: Load existing session data
    # ============================================
    if (-not $SessionPath -or -not (Test-Path $SessionPath)) {
        Write-Host "ERROR: -SessionPath must point to an existing session folder when using -ReportOnly" -ForegroundColor Red
        exit
    }
    $summaryFile = Join-Path $SessionPath "computer_summary.csv"
    if (-not (Test-Path $summaryFile)) {
        Write-Host "ERROR: computer_summary.csv not found in $SessionPath" -ForegroundColor Red
        exit
    }

    $sessionReportPath = $SessionPath
    # Extract timestamp from folder name (Session_YYYYMMDD_HHMMSS) or use current
    if ($SessionPath -match '(\d{8}_\d{6})') {
        $timestamp = $Matches[1]
    } else {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    }

    Write-Host "`nReport-Only Mode: Regenerating HTML from $SessionPath" -ForegroundColor Cyan

    # Load CSVs
    $computerSummary = Import-Csv $summaryFile
    $allUpdatesFile = Join-Path $SessionPath "all_updates.csv"
    if (Test-Path $allUpdatesFile) {
        $allUpdateData = @(Import-Csv $allUpdatesFile)
    } else {
        $allUpdateData = @()
    }
    $rerunFile = Join-Path $SessionPath "rerun_computers.csv"
    if (Test-Path $rerunFile) {
        $rerunList = @(Import-Csv $rerunFile)
    } else {
        $rerunList = $null
    }

    # Deduplicate all_updates.csv (existing data has duplicates from PSWindowsUpdate)
    $seen = @{}
    $allUpdateData = @($allUpdateData | Where-Object {
        $key = "$($_.ComputerName)|$(if ($_.KB -and $_.KB -ne '') { $_.KB } else { $_.Title })"
        if ($seen.ContainsKey($key)) { return $false }
        $seen[$key] = $true
        return $true
    })

    # Ensure Verified field exists on all update records (old CSVs may lack it)
    $allUpdateData | ForEach-Object {
        if (-not ($_.PSObject.Properties.Name -contains 'Verified')) {
            $_ | Add-Member -NotePropertyName Verified -NotePropertyValue 'N/A' -Force
        }
    }

    # Ensure PreviousRunArchive field exists on summary objects (old CSVs may lack it)
    $computerSummary | ForEach-Object {
        if (-not ($_.PSObject.Properties.Name -contains 'PreviousRunArchive')) {
            $_ | Add-Member -NotePropertyName PreviousRunArchive -NotePropertyValue '' -Force
        }
    }

    # Recalculate per-computer summary counts from deduplicated data
    $updatesByComputer = $allUpdateData | Group-Object ComputerName
    foreach ($summary in $computerSummary) {
        $group = $updatesByComputer | Where-Object { $_.Name -eq $summary.ComputerName }
        if ($group) {
            $updates = $group.Group
            $summary.TotalUpdates = $updates.Count
            $summary.Installed = ($updates | Where-Object { $_.Status -in @('Installed') }).Count
            $summary.Failed = ($updates | Where-Object { $_.Status -in @('Failed', 'Aborted') }).Count
            $summary.Skipped = ($updates | Where-Object { $_.Status -eq 'Skipped' }).Count
            $summary.InstalledWithErrors = ($updates | Where-Object { $_.Status -eq 'InstalledWithErrors' }).Count
        }
    }

    Write-Host "Loaded $($computerSummary.Count) computers, $($allUpdateData.Count) unique updates" -ForegroundColor Green

} else {
# ============================================
# FULL DEPLOYMENT MODE: Phases 1-3
# ============================================

Add-Type -AssemblyName System.Windows.Forms

# Clean up any leftover local temp update-log files when the script exits (including Ctrl+C)
Register-EngineEvent PowerShell.Exiting -Action {
    Get-ChildItem "$env:TEMP\updatelog_*.csv" -ErrorAction SilentlyContinue |
        Remove-Item -Force -ErrorAction SilentlyContinue
} | Out-Null

# ============================================
# INITIALIZATION
# ============================================

# Credentials — try Windows Credential Manager first, fall back to prompt
$cred = Get-WUSavedCredential
if ($cred) {
    Write-Host "Using saved credentials for '$($cred.UserName)' from Windows Credential Manager" -ForegroundColor Green
    Write-Host "  (Run with -ClearCredential to remove, or -SaveCredential with new creds to update)" -ForegroundColor DarkGray
} else {
    $cred = Get-Credential -Message "Enter administrator credentials for remote computers"
    if (-not $cred) { Write-Host "No credentials provided. Exiting." -ForegroundColor Red; exit }
    if ($SaveCredential) {
        try {
            Save-WUCredential -Credential $cred
            Write-Host "Credentials saved to Windows Credential Manager" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not save credentials — $($_.Exception.Message)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  Tip: Run with -SaveCredential to remember these credentials" -ForegroundColor DarkGray
    }
}
Write-Host "Choose CSV file for computer list..." -ForegroundColor Cyan
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
$openFileDialog.Title = "Select CSV File for Computer List"
if ($openFileDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true } )) -eq [System.Windows.Forms.DialogResult]::OK) {
    $CSVPath = $openFileDialog.FileName
} else {
    Write-Host "No CSV file selected. Exiting..." -ForegroundColor Red
    exit
}
 
# Create report directory
if (!(Test-Path $ReportPath)) {
    New-Item -Path $ReportPath -ItemType Directory -Force | Out-Null
}
 
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$sessionReportPath = "$ReportPath\Session_$timestamp"
New-Item -Path $sessionReportPath -ItemType Directory -Force | Out-Null
 
Write-Host "`n" -NoNewline
Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║         Windows Update Deployment with Smart Monitoring      ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "Session ID: $timestamp" -ForegroundColor Yellow
Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow
Write-Host "Report Path: $sessionReportPath" -ForegroundColor Yellow
Write-Host ""
 
# ============================================
# SMART MONITORING FUNCTION
# ============================================
 
function Wait-ForUpdateJobs {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Computers,

        [Parameter(Mandatory=$true)]
        [PSCredential]$Credential,

        [Parameter(Mandatory=$false)]
        [int]$MaxWaitMinutes = 60,

        [Parameter(Mandatory=$false)]
        [int]$CheckIntervalSeconds = 30,

        # Path on remote machines where update log and marker files live
        [Parameter(Mandatory=$false)]
        [string]$RemoteTempPath = "C:\temp",

        # Seconds the log must be unmodified before treating it as stable (heuristic fallback)
        [Parameter(Mandatory=$false)]
        [int]$LogStabilitySeconds = 60,

        # Max additional retry attempts when a monitoring Invoke-Command fails transiently
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 2,

        # Computers that failed Phase 1 — checked every 3rd iteration for late arrival
        [Parameter(Mandatory=$false)]
        [object[]]$LateArrivals = @(),

        # Shared job-start scriptblock used to start late-arriving computers
        [Parameter(Mandatory=$false)]
        [scriptblock]$JobStartScript = $null,

        # Minutes a machine may sit "Waiting to start" before a job-start retry is attempted
        [Parameter(Mandatory=$false)]
        [int]$RetryAfterMinutes = 30,

        # Max retry attempts for a stuck "Waiting to start" machine before releasing as JobStartFailed
        [Parameter(Mandatory=$false)]
        [int]$MaxJobStartRetries = 3
    )
    Write-Host "`n═══ Monitoring Update Jobs ═══" -ForegroundColor Cyan
    Write-Host "Maximum wait time: $MaxWaitMinutes minutes" -ForegroundColor Yellow
    Write-Host "Check interval: $CheckIntervalSeconds seconds" -ForegroundColor Yellow
    Write-Host ""
    
    $startTime = Get-Date
    $endTime = $startTime.AddMinutes($MaxWaitMinutes)
    $completedComputers = @{}
    $computerStatus = @{}
    $ignoredComputers = @{}
    $wuJobDetectionWarned = @{}
    $ignoredCount = 0
    
    # Initialize tracking
    foreach ($computer in $Computers) {
        $completedComputers[$computer.Name] = $false
        $computerStatus[$computer.Name] = "Pending"
    }
 
    # ── WSMan pre-check ──────────────────────────────────────────────────────────
    Write-Host "Running WSMan connectivity check against $($Computers.Count) computers..." -ForegroundColor Cyan
    foreach ($computer in $Computers) {
        $target = if ($computer.IP) { $computer.IP } else { $computer.Name }
        $reachable = Test-S2WinRM -ComputerName $target

        if (-not $reachable) {
            Write-Host "  ⚠️  $($computer.Name): Not reachable on WinRM port — will be skipped" -ForegroundColor Yellow
            $ignoredComputers[$computer.Name] = $true
            $completedComputers[$computer.Name] = $true
            $computerStatus[$computer.Name] = "Ignored-WSMAN"
            $ignoredCount++
        }
    }


    $totalToMonitor = $Computers.Count - $ignoredCount
    if ($totalToMonitor -eq 0) {
        Write-Host "All target computers failed WSMan; nothing to monitor. Exiting monitoring." -ForegroundColor Yellow
        return @{ Completed = $completedComputers; Status = $computerStatus }
    }
    
    # Mutable list so late-arriving computers can be added mid-session
    $allMonitored = [System.Collections.ArrayList]@($Computers)

    # Late-arrival tracking
    $lateRetryCount = @{}
    $loopCount = 0
    # Convert to ArrayList so we can remove entries as machines get started
    $pendingLateArrivals = [System.Collections.ArrayList]@($LateArrivals)

    # Job-start retry tracking (for machines that passed Phase 1 but never ran their task)
    $waitingStartTime   = @{}  # first time each machine entered "Waiting to start"
    $jobStartRetryCount = @{}  # cumulative retry attempts per machine

    # Visual separator — everything above this line is preserved across refreshes
    Write-Host "`n═══ Live Monitoring (refreshes every ${CheckIntervalSeconds}s) ═══" -ForegroundColor DarkGray
    $monitorTop = [Console]::CursorTop

    # Monitor loop
    while ((Get-Date) -lt $endTime) {
        $stillRunning = 0
        $completed = 0
        $failed = 0

        # Reset cursor to monitoring area — preserves Phase 1 output above
        [Console]::SetCursorPosition(0, $monitorTop)
        $clearLine = " " * [Console]::WindowWidth
        $clearEnd = [math]::Min($monitorTop + [Console]::WindowHeight, [Console]::BufferHeight - 1)
        for ($ci = $monitorTop; $ci -lt $clearEnd; $ci++) {
            [Console]::SetCursorPosition(0, $ci)
            [Console]::Write($clearLine)
        }
        [Console]::SetCursorPosition(0, $monitorTop)

        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Checking status..." -ForegroundColor Gray
        $loopCount++

        # ── Late-arrival retry (every 3rd loop, ~90s at default interval) ────────
        # Check if any Phase 1 failures have come online and start their jobs.
        if ($pendingLateArrivals.Count -gt 0 -and $null -ne $JobStartScript -and ($loopCount % 3 -eq 0)) {
            $toRemove = [System.Collections.ArrayList]@()
            foreach ($la in $pendingLateArrivals) {
                if (-not $lateRetryCount.ContainsKey($la.Name)) { $lateRetryCount[$la.Name] = 0 }
                $lateRetryCount[$la.Name]++
                if ($lateRetryCount[$la.Name] -gt 20) {
                    $toRemove.Add($la) | Out-Null  # give up — will stay as JobStartFailed
                    continue
                }
                # Quick WinRM check
                $laReachable = Test-S2WinRM -ComputerName $la.IP

                if ($laReachable) {
                    Write-Host "  🔁 $($la.Name): Now reachable — starting job (late arrival)" -ForegroundColor Magenta
                    try {
                        Invoke-Command -ComputerName $la.IP -Credential $Credential -ErrorAction Stop `
                            -ArgumentList $RemoteTempPath -ScriptBlock $JobStartScript
                        $allMonitored.Add($la) | Out-Null
                        $completedComputers[$la.Name] = $false
                        $computerStatus[$la.Name] = "Waiting"
                        $toRemove.Add($la) | Out-Null
                        Write-Host "  ✓ $($la.Name): Late job started successfully" -ForegroundColor Green
                    } catch {
                        Write-Host "  ✗ $($la.Name): Late start failed — $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
            foreach ($r in $toRemove) { $pendingLateArrivals.Remove($r) | Out-Null }
        }

        foreach ($computer in $allMonitored) {
            # Skip hosts that failed the WSMan pre-check
            if ($ignoredComputers[$computer.Name]) {
                continue
            }
 
            # Skip if already completed
            if ($completedComputers[$computer.Name]) {
                $completed++
                continue
            }
            
            # ── Quick WinRM pre-check to avoid expensive Invoke-Command on offline machines ──
            $target = if ($computer.IP) { $computer.IP } else { $computer.Name }
            if (-not (Test-S2WinRM -ComputerName $target)) {
                $computerStatus[$computer.Name] = "Unreachable"
                $stillRunning++
                continue
            }

            # ── Remote status query with retry on transient failures ────────────
            $jobStatus = $null
            for ($retryAttempt = 0; $retryAttempt -le $MaxRetries; $retryAttempt++) {
                try {
                    $jobStatus = Invoke-Command -ComputerName $computer.IP -Credential $Credential `
                        -ArgumentList $RemoteTempPath, $LogStabilitySeconds `
                        -ScriptBlock {
                            param($remoteTempPath, $logStabilitySeconds)

                            # Check for PSWindowsUpdate scheduled tasks
                            $wuTasks = @()
                            try {
                                $wuTasks = Get-ScheduledTask -TaskName "*PSWindowsUpdate*" -ErrorAction SilentlyContinue |
                                           Where-Object { $_.State -eq 'Running' }
                            } catch { }

                            # Check via Get-WUJob; surface failure so caller can warn operator
                            $wuJobsRunning = $false
                            $wuJobDetectionFailed = $false
                            try {
                                Import-Module PSWindowsUpdate -ErrorAction Stop
                                $jobs = Get-WUJob -ErrorAction SilentlyContinue
                                $wuJobsRunning = ($null -ne $jobs -and $jobs.Count -gt 0)
                            } catch {
                                $wuJobDetectionFailed = $true
                            }

                            # Primary completion signal: explicit marker written when update job finishes
                            $completedMarkerExists = Test-Path "$remoteTempPath\update_completed.txt"

                            # Session anchor: $jobStartBlock writes update_started.txt fresh each run,
                            # AFTER deleting old artifacts. Any file written before this anchor is stale.
                            $startedTime = $null
                            if (Test-Path "$remoteTempPath\update_started.txt") {
                                $startedTime = (Get-Item "$remoteTempPath\update_started.txt").LastWriteTime
                            }
                            # Completion is only valid if the marker was written after this session's job start
                            $completedIsCurrentSession = $null -ne $startedTime -and
                                $completedMarkerExists -and
                                (Get-Item "$remoteTempPath\update_completed.txt").LastWriteTime -ge $startedTime

                            # Heuristic: log has not been written for $logStabilitySeconds
                            $logPath = "$remoteTempPath\updatelog.csv"
                            $logExists = Test-Path $logPath
                            $logComplete = $false
                            $logSize = 0
                            $logIsCurrentSession = $false
                            if ($logExists) {
                                $logFile = Get-Item $logPath
                                $logAge = (Get-Date) - $logFile.LastWriteTime
                                $logComplete = $logAge.TotalSeconds -gt $logStabilitySeconds
                                $logSize = $logFile.Length
                                $logIsCurrentSession = $null -ne $startedTime -and $logFile.LastWriteTime -ge $startedTime
                            }

                            # Check for active Windows Update processes
                            $updateProcesses = @()
                            try {
                                $updateProcesses = Get-Process -Name "TrustedInstaller", "TiWorker", "wuauclt", "msiexec", "wusa" -ErrorAction SilentlyContinue
                            } catch { }

                            # Check for WUJob start failure marker
                            $wuJobError = $null
                            $wuJobErrorPath = "$remoteTempPath\wujob_error.txt"
                            if (Test-Path $wuJobErrorPath) {
                                $wuJobError = (Get-Content $wuJobErrorPath -Raw).Trim()
                            }

                            return @{
                                TasksRunning                = ($wuTasks.Count -gt 0)
                                JobsRunning                 = $wuJobsRunning
                                WUJobDetectionFailed        = $wuJobDetectionFailed
                                CompletedMarkerExists       = $completedMarkerExists
                                CompletedIsCurrentSession   = $completedIsCurrentSession
                                LogExists                   = $logExists
                                LogIsCurrentSession         = $logIsCurrentSession
                                LogComplete                 = $logComplete
                                LogSize                     = $logSize
                                UpdateProcessesRunning      = ($updateProcesses.Count -gt 0)
                                WUJobError                  = $wuJobError
                                ComputerName                = $env:COMPUTERNAME
                            }
                        } -ErrorAction Stop
                    break  # success — exit retry loop
                } catch {
                    if ($retryAttempt -lt $MaxRetries) {
                        Start-Sleep -Seconds (5 * ($retryAttempt + 1))
                    }
                }
            }

            if (-not $jobStatus) {
                Write-Host "  ⚠️  $($computer.Name): Unreachable after $($MaxRetries + 1) attempt(s) (may be rebooting)" -ForegroundColor DarkYellow
                $computerStatus[$computer.Name] = "Unreachable"
                $stillRunning++
                continue
            }

            # Warn once if PSWindowsUpdate module is missing on remote host
            if ($jobStatus.WUJobDetectionFailed -and -not $wuJobDetectionWarned[$computer.Name]) {
                Write-Host "  ℹ️  $($computer.Name): PSWindowsUpdate unavailable on remote — task/process detection only" -ForegroundColor DarkYellow
                $wuJobDetectionWarned[$computer.Name] = $true
            }

            # ── Check for WUJob start failure ────────────────────────────────
            if ($jobStatus.WUJobError) {
                Write-Host "  ⛔ $($computer.Name): WUJob failed to start — $($jobStatus.WUJobError)" -ForegroundColor Red
                $computerStatus[$computer.Name] = "JobStartFailed"
                $completedComputers[$computer.Name] = $true
                $completed++
                continue
            }

            # ── Completion logic ─────────────────────────────────────────────
            # Primary: explicit completion marker AND log — both must be from the current session
            # (newer than update_started.txt) to prevent false positives from previous-run artifacts
            $definitelyDone = $jobStatus.CompletedIsCurrentSession -and $jobStatus.LogIsCurrentSession
            # Fallback heuristic: log is stable, current-session, no tasks or WU jobs running
            $heuristicDone  = $jobStatus.LogIsCurrentSession -and $jobStatus.LogComplete -and
                              -not $jobStatus.TasksRunning -and -not $jobStatus.JobsRunning -and
                              -not $jobStatus.UpdateProcessesRunning

            if ($definitelyDone -or $heuristicDone) {
                $how = if ($definitelyDone) { "(marker)" } else { "(heuristic)" }
                Write-Host "  ✅ $($computer.Name): Completed $how (Log: $([math]::Round($jobStatus.LogSize/1KB, 1))KB)" -ForegroundColor Green
                $completedComputers[$computer.Name] = $true
                $computerStatus[$computer.Name] = "Completed"
                $completed++
            } elseif ($jobStatus.TasksRunning -or $jobStatus.JobsRunning -or $jobStatus.UpdateProcessesRunning) {
                $runningWhat = @()
                if ($jobStatus.TasksRunning)           { $runningWhat += "Tasks" }
                if ($jobStatus.JobsRunning)            { $runningWhat += "Jobs" }
                if ($jobStatus.UpdateProcessesRunning) { $runningWhat += "Processes" }
                Write-Host "  ⏳ $($computer.Name): Running ($($runningWhat -join ', '))" -ForegroundColor Yellow
                $computerStatus[$computer.Name] = "Running"
                $stillRunning++
            } elseif ($jobStatus.LogExists -and -not $jobStatus.LogComplete) {
                Write-Host "  📝 $($computer.Name): Writing results (log age < ${LogStabilitySeconds}s)..." -ForegroundColor Cyan
                $computerStatus[$computer.Name] = "Writing"
                $stillRunning++
            } else {
                # Track elapsed wait time for this machine
                if (-not $waitingStartTime.ContainsKey($computer.Name)) {
                    $waitingStartTime[$computer.Name] = Get-Date
                }
                $waitingMinutes = [int]((Get-Date) - $waitingStartTime[$computer.Name]).TotalMinutes

                if ($null -ne $JobStartScript -and $waitingMinutes -ge $RetryAfterMinutes) {
                    if (-not $jobStartRetryCount.ContainsKey($computer.Name)) { $jobStartRetryCount[$computer.Name] = 0 }

                    if ($jobStartRetryCount[$computer.Name] -ge $MaxJobStartRetries) {
                        # All retries exhausted — release as JobStartFailed (lands on rerun list)
                        Write-Host "  ⛔ $($computer.Name): Waiting ${waitingMinutes}m, $MaxJobStartRetries retries exhausted — releasing as JobStartFailed" -ForegroundColor Red
                        $computerStatus[$computer.Name] = "JobStartFailed"
                        $completedComputers[$computer.Name] = $true
                        $completed++
                    } else {
                        $jobStartRetryCount[$computer.Name]++
                        $attempt = $jobStartRetryCount[$computer.Name]
                        Write-Host "  🔄 $($computer.Name): No activity for ${waitingMinutes}m — retrying job start ($attempt/$MaxJobStartRetries)" -ForegroundColor Yellow
                        # Reset timer regardless of outcome — prevents immediate re-trigger next loop
                        $waitingStartTime[$computer.Name] = Get-Date
                        try {
                            Invoke-Command -ComputerName $computer.IP -Credential $Credential -ErrorAction Stop `
                                -ArgumentList $RemoteTempPath -ScriptBlock $JobStartScript
                            Write-Host "    ✓ Job re-started on $($computer.Name)" -ForegroundColor Green
                        } catch {
                            Write-Host "    ✗ Re-start failed: $($_.Exception.Message)" -ForegroundColor Red
                        }
                        # Keep monitoring regardless — success or failure we check next loop
                        $computerStatus[$computer.Name] = "Waiting"
                        $stillRunning++
                    }
                } else {
                    $waitMsg = if ($waitingMinutes -gt 0) { " (${waitingMinutes}m)" } else { "" }
                    Write-Host "  ⏸️  $($computer.Name): Waiting to start${waitMsg}" -ForegroundColor Gray
                    $computerStatus[$computer.Name] = "Waiting"
                    $stillRunning++
                }
            }
        }
        
        # Progress summary
        $elapsed = [int]((Get-Date) - $startTime).TotalMinutes
        $percentComplete = ($completed / $allMonitored.Count) * 100

        Write-Host ""
        Write-Host "Progress: $completed/$($allMonitored.Count) completed ($([math]::Round($percentComplete, 1))%)" -ForegroundColor Cyan
        Write-Host "Elapsed: $elapsed minutes | Remaining: $stillRunning computers" -ForegroundColor Gray
        
        # Visual progress bar
        $barLength = 50
        $filledLength = [math]::Round(($percentComplete / 100) * $barLength)
        $bar = "█" * $filledLength + "░" * ($barLength - $filledLength)
        Write-Host "[$bar] $([math]::Round($percentComplete, 1))%" -ForegroundColor Green
        
        # Check if all done
        if ($stillRunning -eq 0) {
            Write-Host "`n✅ All update jobs completed!" -ForegroundColor Green
            break
        }
        
        # Wait before next check
        Write-Host "`nNext check in $CheckIntervalSeconds seconds..." -ForegroundColor DarkGray
        Start-Sleep -Seconds $CheckIntervalSeconds
    }
    
    # Timeout handling
    if ((Get-Date) -ge $endTime) {
        Write-Host "`n⚠️  Maximum wait time reached ($MaxWaitMinutes minutes)" -ForegroundColor Yellow
        $notCompleted = @()
        foreach ($computer in $Computers) {
            if (-not $completedComputers[$computer.Name]) {
                $notCompleted += "$($computer.Name) ($($computerStatus[$computer.Name]))"
            }
        }
        if ($notCompleted.Count -gt 0) {
            Write-Host "The following computers did not complete:" -ForegroundColor Yellow
            $notCompleted | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
        }
    }
    
    return @{
        Completed = $completedComputers
        Status = $computerStatus
    }
}
 
# ============================================
# PHASE 1: START UPDATE JOBS
# ============================================
 
Write-Host "═══ Phase 1: Starting Update Jobs ═══" -ForegroundColor Cyan
Write-Host ""
 
# Load and validate computers
$computers = Import-Csv $CSVPath
if (-not $computers -or $computers.Count -eq 0) {
    Write-Host "CSV file is empty or contains no rows. Exiting..." -ForegroundColor Red
    exit
}
$requiredColumns = @("Name", "IP")
$csvColumns = $computers[0].PSObject.Properties.Name
$missingCols = $requiredColumns | Where-Object { $csvColumns -notcontains $_ }
if ($missingCols.Count -gt 0) {
    Write-Host "CSV is missing required column(s): $($missingCols -join ', '). Expected columns: Name, IP" -ForegroundColor Red
    exit
}
$computers = @($computers | Where-Object { $_.Name -and $_.IP })
if ($computers.Count -eq 0) {
    Write-Host "No rows with both a Name and IP value found in CSV. Exiting..." -ForegroundColor Red
    exit
}
Write-Host "Loaded $($computers.Count) computers from CSV" -ForegroundColor Green

# ── Credential pre-validation ─────────────────────────────────────────────
# Test credentials against a reachable machine before launching all parallel jobs.
# Catches bad passwords early instead of failing 91+ machines identically.
Write-Host "`nValidating credentials..." -ForegroundColor Cyan
$credValid = $false
$credTestCount = 0
$credTestMax = [math]::Min(3, $computers.Count)
foreach ($testTarget in $computers) {
    $credTestCount++
    $target = if ($testTarget.IP) { $testTarget.IP } else { $testTarget.Name }
    try {
        # Quick WinRM check first
        if (-not (Test-S2WinRM -ComputerName $target)) { continue }

        Invoke-Command -ComputerName $target -Credential $cred -ScriptBlock { $env:COMPUTERNAME } -ErrorAction Stop | Out-Null
        Write-Host "  Credentials validated against $($testTarget.Name)" -ForegroundColor Green
        $credValid = $true
        break
    } catch {
        if ($_.Exception.Message -match 'Access is denied|user name or password|logon failure|credentials') {
            Write-Host "  ERROR: Credential validation failed — $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "  Please verify your username and password. Exiting..." -ForegroundColor Red
            exit
        }
        # Non-credential error (machine-specific) — try next machine
    }
}
if (-not $credValid) {
    Write-Host "  Warning: Could not validate credentials (first $credTestCount machine(s) unreachable). Proceeding..." -ForegroundColor Yellow
}

# Job-start scriptblock — shared by Phase 1 and the late-arrival retry logic.
# Serialized as a string for Phase 1 (ForEach-Object -Parallel cannot pass scriptblocks
# via $using:), and passed directly as -JobStartScript to Wait-ForUpdateJobs.
$jobStartBlock = {
    param($remoteTempPath)
    # Ensure temp directory exists
    if (-not (Test-Path $remoteTempPath)) {
        New-Item $remoteTempPath -ItemType Directory -Force | Out-Null
    }

    # Archive previous run's artifacts if they exist and are from today (preserve for diff)
    $logPath = "$remoteTempPath\updatelog.csv"
    if (Test-Path $logPath) {
        $logTime = (Get-Item $logPath).LastWriteTime
        if ($logTime.Date -eq (Get-Date).Date) {
            $archiveDir = "$remoteTempPath\archive_$($logTime.ToString('HHmmss'))"
            New-Item $archiveDir -ItemType Directory -Force | Out-Null
            @("updatelog.csv", "update_completed.txt", "update_started.txt") | ForEach-Object {
                $p = "$remoteTempPath\$_"
                if (Test-Path $p) { Move-Item $p "$archiveDir\$_" -Force }
            }
        }
    }

    # Clean up archives from previous days to prevent accumulation
    Get-ChildItem "$remoteTempPath\archive_*" -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.CreationTime.Date -lt (Get-Date).Date } |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

    # Clear stale artifacts from any previous run to prevent false signals
    @("updatelog.csv", "update_completed.txt", "wujob_error.txt") | ForEach-Object {
        $p = "$remoteTempPath\$_"
        if (Test-Path $p) { Remove-Item $p -Force }
    }

    # Remove any existing PSWindowsUpdate scheduled tasks from previous runs
    # to prevent Invoke-WUJob from silently failing on conflict
    try {
        Get-ScheduledTask -TaskName "*PSWindowsUpdate*" -ErrorAction SilentlyContinue |
            ForEach-Object {
                Stop-ScheduledTask -InputObject $_ -ErrorAction SilentlyContinue
                Unregister-ScheduledTask -InputObject $_ -Confirm:$false -ErrorAction SilentlyContinue
            }
    } catch { }

    # Record job start time
    Get-Date | Out-File "$remoteTempPath\update_started.txt" -Force
    # Build the inner update scriptblock with the path baked in so the
    # scheduled task (which runs as SYSTEM in its own session) can use it.
    $updateScript = [scriptblock]::Create(@"
        Import-Module PSWindowsUpdate
        `$updates = Install-WindowsUpdate -AcceptAll -MicrosoftUpdate -ForceInstall -Install -IgnoreReboot
        if (`$updates) {
            `$updates | Export-Csv '$remoteTempPath\updatelog.csv' -NoTypeInformation -Force
        } else {
            @() | Export-Csv '$remoteTempPath\updatelog.csv' -NoTypeInformation -Force
        }
        Get-Date | Out-File '$remoteTempPath\update_completed.txt' -Force
"@)
    try {
        Invoke-WUJob -RunNow -Confirm:$false -Verbose -Script $updateScript
    } catch {
        # Surface the failure so Phase 1 and Phase 2 can detect it
        "WUJOB_FAILED: $($_.Exception.Message)" | Out-File "$remoteTempPath\wujob_error.txt" -Force
        throw
    }
}

# Serialize $jobStartBlock as a string so it can cross the ForEach-Object -Parallel boundary.
# PowerShell does not support passing scriptblocks via $using: in parallel loops.
$jobStartBlockStr = $jobStartBlock.ToString()

# Start updates on all computers in parallel — no throttle (each site is a separate network)
$jobResults = $computers | ForEach-Object -Parallel {
    $site          = $_.'Name'
    $IP            = $_.'IP'
    $cred          = $using:cred
    $rtPath        = $using:RemoteTempPath   # capture for use inside Invoke-Command
    $jobBlock      = [scriptblock]::Create($using:jobStartBlockStr)

    # Pre-flight: WinRM check before attempting full Invoke-Command.
    $target = if ($IP) { $IP } else { $site }
    $preFlightOk = $false
    try {
        $null = Test-WSMan -ComputerName $target -ErrorAction Stop
        $preFlightOk = $true
    } catch {}

    if (-not $preFlightOk) {
        Write-Host "  [$site] WinRM not reachable — skipping" -ForegroundColor DarkYellow
        return [PSCustomObject]@{ Site = $site; IP = $IP; Status = "Unreachable-PreFlight"; Error = "WinRM not reachable" }
    }

    try {
        Write-Host "[$site] Starting update job..." -ForegroundColor Yellow

        Invoke-Command -ComputerName $IP -Credential $cred -ErrorAction Stop `
            -ArgumentList $rtPath -ScriptBlock $jobBlock

        Write-Host "[$site] ✓ Update job started successfully" -ForegroundColor Green
        return [PSCustomObject]@{
            Site   = $site
            IP     = $IP
            Status = "Started"
            Error  = ""
        }
    }
    catch {
        Write-Host "[$site] ✗ Failed to start: $($_.Exception.Message)" -ForegroundColor Red
        return [PSCustomObject]@{
            Site   = $site
            IP     = $IP
            Status = "Failed"
            Error  = $_.Exception.Message
        }
    }
} -ThrottleLimit ([math]::Min($ThrottleLimit, $computers.Count))
 
# Save initial job status
$jobResults | Export-CSV -Path "$sessionReportPath\job_start_status.csv" -NoTypeInformation
$phase1Started = ($jobResults | Where-Object { $_.Status -eq "Started" }).Count
$phase1Skipped = ($jobResults | Where-Object { $_.Status -ne "Started" }).Count
Write-Host "`nAll job start attempts completed — started: $phase1Started  skipped/failed: $phase1Skipped" -ForegroundColor Green

# Exclude Phase 1 failures from monitoring — they will never show activity and would
# sit at "Waiting to start" for the full monitoring window otherwise.
$phase1FailedSites = $jobResults | Where-Object { $_.Status -ne "Started" } |
    Select-Object -ExpandProperty Site
$computersToMonitor  = $computers | Where-Object { $_.Name -notin $phase1FailedSites }
$lateArrivalComputers = $computers | Where-Object { $_.Name -in $phase1FailedSites }

# ============================================
# PHASE 2: SMART MONITORING
# ============================================

Write-Host "`n═══ Phase 2: Smart Job Monitoring ═══" -ForegroundColor Cyan

$monitoringResults = Wait-ForUpdateJobs -Computers $computersToMonitor -Credential $cred `
                                        -MaxWaitMinutes $MaxWaitMinutes `
                                        -CheckIntervalSeconds $CheckIntervalSeconds `
                                        -RemoteTempPath $RemoteTempPath `
                                        -LogStabilitySeconds $LogStabilitySeconds `
                                        -MaxRetries $MaxRetries `
                                        -LateArrivals $lateArrivalComputers `
                                        -JobStartScript $jobStartBlock

# Inject Phase 1 failures into monitoring results so Phase 3 handles them correctly.
# Skip computers that were successfully started as late arrivals during monitoring —
# those already have a real status (Completed/Incomplete/etc.) in $monitoringResults.
foreach ($f in ($jobResults | Where-Object { $_.Status -ne "Started" })) {
    $existingStatus = $monitoringResults.Status[$f.Site]
    $lateStarted = $existingStatus -and $existingStatus -notin @('JobStartFailed', 'Pending', $null)
    if (-not $lateStarted) {
        $monitoringResults.Status[$f.Site]    = "JobStartFailed"
        $monitoringResults.Completed[$f.Site] = $true
    }
}

# Replace PHASE 3 and PHASE 4 entirely with this code
# This version properly handles text-based result values like "Failed" and "Installed"
 
# ============================================
# PHASE 3: COLLECT RESULTS (FIXED VERSION)
# ============================================
 
Write-Host "`n╔══ Phase 3: Collecting Update Results ══╗" -ForegroundColor Cyan
Write-Host ""

$allUpdateData = @()
$computerSummary = @()

# ── Separate computers into collectable vs skip-early groups ──
$skipComputers = @()
$incompleteComputers = @()
$collectableComputers = @()

foreach ($computer in $computers) {
    $computerName = $computer.Name
    $status = $monitoringResults.Status[$computerName]
    if ($status -eq "JobStartFailed") {
        $skipComputers += $computer
    } elseif (-not $monitoringResults.Completed[$computerName]) {
        $incompleteComputers += $computer
    } else {
        $collectableComputers += $computer
    }
}

# Handle JobStartFailed computers immediately (no remote work needed)
foreach ($computer in $skipComputers) {
    $phase1Error = ($jobResults | Where-Object { $_.Site -eq $computer.Name } | Select-Object -First 1).Error
    $errMsg = if ($phase1Error) { "Job did not start: $phase1Error" } else { "Job did not start" }
    Write-Host "[$($computer.Name)] " -NoNewline
    Write-Host "JobStartFailed — $errMsg" -ForegroundColor DarkYellow
    $computerSummary += [PSCustomObject]@{
        ComputerName = $computer.Name; IP = $computer.IP; Status = "JobStartFailed"
        TotalUpdates = 0; Installed = 0; Failed = 0; Skipped = 0; InstalledWithErrors = 0
        RebootRequired = $false; CollectionError = $errMsg; PreviousRunArchive = ""
    }
}

# Handle incomplete computers immediately
foreach ($computer in $incompleteComputers) {
    Write-Host "[$($computer.Name)] " -NoNewline
    Write-Host "Skipped (job did not complete)" -ForegroundColor Yellow
    $computerSummary += [PSCustomObject]@{
        ComputerName = $computer.Name; IP = $computer.IP; Status = "Incomplete"
        TotalUpdates = 0; Installed = 0; Failed = 0; Skipped = 0; InstalledWithErrors = 0
        RebootRequired = $false; CollectionError = ""; PreviousRunArchive = ""
    }
}

if ($collectableComputers.Count -gt 0) {
    # ── Step 1: Batch PSSession creation ──
    $collectTimer = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "Creating sessions to $($collectableComputers.Count) computers... " -NoNewline

    # Build IP→Name lookup since Invoke-Command returns PSComputerName (the IP)
    $ipToName = @{}
    $nameToIP = @{}
    foreach ($c in $collectableComputers) { $ipToName[$c.IP] = $c.Name; $nameToIP[$c.Name] = $c.IP }

    $allIPs = @($collectableComputers | ForEach-Object { $_.IP })
    $sessions = @()
    $sessionErrors = @{}
    try {
        $sessions = @(New-PSSession -ComputerName $allIPs -Credential $cred -ErrorAction SilentlyContinue -ErrorVariable sessionCreateErrors)
    } catch { }

    # Track which IPs failed to connect
    $connectedIPs = @{}
    foreach ($s in $sessions) { $connectedIPs[$s.ComputerName] = $true }
    $failedIPs = $allIPs | Where-Object { -not $connectedIPs[$_] }

    $failedCount = $failedIPs.Count
    $connectedCount = $sessions.Count
    Write-Host "✅ $connectedCount connected" -ForegroundColor Green -NoNewline
    if ($failedCount -gt 0) {
        Write-Host ", " -NoNewline
        Write-Host "$failedCount failed" -ForegroundColor Red
    } else {
        Write-Host ""
    }

    # Add CollectionFailed summaries for machines we couldn't connect to
    foreach ($failedIP in $failedIPs) {
        $failedName = $ipToName[$failedIP]
        $errDetail = ""
        if ($sessionCreateErrors) {
            $matchingErr = $sessionCreateErrors | Where-Object { $_.TargetObject -eq $failedIP -or "$_" -match [regex]::Escape($failedIP) } | Select-Object -First 1
            if ($matchingErr) { $errDetail = $matchingErr.Exception.Message }
        }
        Write-Host "  [$failedName] " -NoNewline
        Write-Host "Could not connect: $errDetail" -ForegroundColor Red
        $computerSummary += [PSCustomObject]@{
            ComputerName = $failedName; IP = $failedIP; Status = "CollectionFailed"
            TotalUpdates = 0; Installed = 0; Failed = 0; Skipped = 0; InstalledWithErrors = 0
            RebootRequired = $false; CollectionError = "Session creation failed: $errDetail"; PreviousRunArchive = ""
        }
    }

    # ── Step 2: Single batched Invoke-Command for ALL remote data ──
    $remoteData = @{}
    if ($sessions.Count -gt 0) {
        Write-Host "Querying remote data (logs, hotfixes, reboot status)... " -NoNewline

        $remoteResults = @(Invoke-Command -Session $sessions -ArgumentList $RemoteTempPath -ScriptBlock {
            param($rtPath)
            $data = @{
                CsvContent       = $null
                CsvExists        = $false
                ArchiveCsvContent = $null
                ArchiveName      = $null
                CompletedMarker  = $false
                Hotfixes         = @()
                PendingKBs       = @()
                RebootRequired   = $false
                VerifyError      = $null
            }

            # Read CSV content directly (eliminates file copy overhead)
            $logPath = "$rtPath\updatelog.csv"
            if (Test-Path $logPath) {
                $data.CsvExists = $true
                $data.CsvContent = Get-Content $logPath -Raw
            }

            # Check completion marker
            $data.CompletedMarker = Test-Path "$rtPath\update_completed.txt"

            # Check for same-day archive (previous run logs)
            try {
                $archive = Get-ChildItem "$rtPath\archive_*" -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.CreationTime.Date -eq (Get-Date).Date } |
                    Sort-Object CreationTime -Descending | Select-Object -First 1
                if ($archive -and (Test-Path "$($archive.FullName)\updatelog.csv")) {
                    $data.ArchiveCsvContent = Get-Content "$($archive.FullName)\updatelog.csv" -Raw
                    $data.ArchiveName = $archive.Name
                }
            } catch { }

            # Verification: Get-HotFix (installed patches)
            try {
                $data.Hotfixes = @(Get-HotFix -ErrorAction Stop | ForEach-Object { $_.HotFixID })
            } catch {
                $data.VerifyError = "Get-HotFix: $($_.Exception.Message)"
            }

            # Verification: Pending updates (not yet installed)
            try {
                $searcher = New-Object -ComObject Microsoft.Update.Searcher
                $searchResult = $searcher.Search("IsInstalled=0")
                $data.PendingKBs = @($searchResult.Updates | ForEach-Object {
                    @($_.KBArticleIDs) | ForEach-Object { "KB$_" }
                })
            } catch { }

            # Reboot status from registry
            if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
                $data.RebootRequired = $true
            } elseif (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
                $data.RebootRequired = $true
            } else {
                $pfro = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
                if ($pfro.PendingFileRenameOperations) { $data.RebootRequired = $true }
            }

            return $data
        } -ErrorAction SilentlyContinue -ErrorVariable remoteQueryErrors)

        # Index results by IP (PSComputerName)
        foreach ($r in $remoteResults) {
            $remoteData[$r.PSComputerName] = $r
        }

        $collectTimer.Stop()
        Write-Host "✅ Done ($([math]::Round($collectTimer.Elapsed.TotalSeconds, 1))s)" -ForegroundColor Green

        # Clean up all sessions in batch
        $sessions | Remove-PSSession -ErrorAction SilentlyContinue
    }

    # ── Step 3: Process results locally (fast — no network calls) ──
    Write-Host "Processing results..."

    foreach ($computer in $collectableComputers) {
        $computerName = $computer.Name
        $computerIP = $computer.IP

        # Skip if already handled as connection failure
        if ($failedIPs -contains $computerIP) { continue }

        # Initialize summary
        $summary = [PSCustomObject]@{
            ComputerName = $computerName; IP = $computerIP
            Status = $monitoringResults.Status[$computerName]
            TotalUpdates = 0; Installed = 0; Failed = 0; Skipped = 0; InstalledWithErrors = 0
            RebootRequired = $false; CollectionError = ""; PreviousRunArchive = ""
        }

        $result = $remoteData[$computerIP]
        if (-not $result) {
            # Remote query returned no data for this machine
            Write-Host "  [$computerName] " -NoNewline
            Write-Host "No data returned from remote query" -ForegroundColor Yellow
            $summary.Status = "CollectionFailed"
            $summary.CollectionError = "Invoke-Command returned no data"
            $computerSummary += $summary
            continue
        }

        Write-Host "  [$computerName] " -NoNewline

        try {
            if ($result.CsvExists -and $result.CsvContent) {
                # Parse CSV from string content (no file copy needed)
                $updates = @($result.CsvContent | ConvertFrom-Csv -ErrorAction Stop)

                # Save archive CSV locally if present (for diff reports)
                if ($result.ArchiveCsvContent) {
                    $prevLocalFile = "$sessionReportPath\previous_updatelog_${computerName}.csv"
                    $result.ArchiveCsvContent | Set-Content -Path $prevLocalFile -Encoding UTF8 -ErrorAction SilentlyContinue
                    $summary.PreviousRunArchive = $result.ArchiveName
                }

                # Validate CSV schema
                if ($updates.Count -gt 0) {
                    $csvCols = $updates[0].PSObject.Properties.Name
                    $expectedCols = @('KB', 'Title', 'Result', 'Status')
                    $missingCols = $expectedCols | Where-Object { $csvCols -notcontains $_ }
                    if ($missingCols.Count -gt 0) {
                        Write-Host "Warning: CSV missing columns: $($missingCols -join ', ') " -ForegroundColor Yellow -NoNewline
                        $summary.CollectionError = "CSV schema: missing $($missingCols -join ', ')"
                    }
                }

                # Build verification lookup tables
                $installedKBs = @{}
                $pendingKBs = @{}
                $hasVerification = $false
                if ($result.Hotfixes -or $result.PendingKBs) {
                    $hasVerification = $true
                    foreach ($hf in $result.Hotfixes) {
                        if ($hf) { $installedKBs[$hf] = $true }
                    }
                    foreach ($kb in $result.PendingKBs) {
                        $pendingKBs[$kb] = $true
                    }
                }

                $rebootStatus = if ($null -ne $result.RebootRequired) { $result.RebootRequired } else { $false }

                # Process updates with comprehensive parsing
                if ($updates.Count -gt 0) {
                    $seenUpdates = @{}

                    foreach ($update in $updates) {
                        $resultValue = if ($update.Result) { $update.Result.ToString().Trim() } else { "" }
                        $statusValue = if ($update.Status) { $update.Status.ToString().Trim() } else { "" }

                        # Skip duplicates
                        $dedupKey = if ($update.KB -and $update.KB -ne '') { $update.KB } else { $update.Title }
                        if ($seenUpdates.ContainsKey($dedupKey)) { continue }
                        $seenUpdates[$dedupKey] = $true

                        $actualStatus = "Unknown"
                        $isFailure = $false
                        $isSuccess = $false
                        $isDefenderDefinition = $update.Title -match 'Security Intelligence Update for Microsoft Defender'

                        # Defender definition updates self-install via MpSigStub before WU can;
                        # WU reports failure/abort but the definitions are current. Mark as Skipped.
                        if ($isDefenderDefinition -and ($resultValue -match '^Fail|^Abort' -or
                            $resultValue -eq '4' -or $resultValue -eq '5')) {
                            $summary.Skipped++
                            $actualStatus = "Skipped"
                        }
                        elseif ($resultValue -eq "Failed" -or $resultValue -match "^Fail") {
                            $summary.Failed++; $actualStatus = "Failed"; $isFailure = $true
                        }
                        elseif ($resultValue -eq "Installed" -or $resultValue -match "^Install|^Success") {
                            $summary.Installed++; $actualStatus = "Installed"; $isSuccess = $true
                        }
                        elseif ($resultValue -match "Abort") {
                            $summary.Failed++; $actualStatus = "Aborted"; $isFailure = $true
                        }
                        elseif ($resultValue -match "Error") {
                            $summary.InstalledWithErrors++; $actualStatus = "InstalledWithErrors"
                        }
                        elseif ($resultValue -eq "2" -or $resultValue -eq 2) {
                            $summary.Installed++; $actualStatus = "Installed"; $isSuccess = $true
                        }
                        elseif ($resultValue -eq "3" -or $resultValue -eq 3) {
                            $summary.InstalledWithErrors++; $actualStatus = "InstalledWithErrors"
                        }
                        elseif ($resultValue -eq "4" -or $resultValue -eq 4) {
                            $summary.Failed++; $actualStatus = "Failed"; $isFailure = $true
                        }
                        elseif ($resultValue -eq "5" -or $resultValue -eq 5) {
                            $summary.Failed++; $actualStatus = "Aborted"; $isFailure = $true
                        }
                        elseif ($statusValue -match "Fail") {
                            $summary.Failed++; $actualStatus = "Failed"; $isFailure = $true
                        }
                        elseif ($statusValue -match "Install|Success") {
                            $summary.Installed++; $actualStatus = "Installed"; $isSuccess = $true
                        }

                        # Special case: Status is "Unknown" but Result is "Failed"
                        if ($statusValue -eq "Unknown" -and $resultValue -eq "Failed" -and !$isFailure) {
                            $summary.Failed++; $actualStatus = "Failed"; $isFailure = $true
                        }

                        # ── Verification: cross-reference KB against actual OS state ──
                        $verified = "N/A"
                        if ($update.KB -and $update.KB -ne '' -and $update.KB -ne 'N/A') {
                            if ($installedKBs.ContainsKey($update.KB)) {
                                $verified = "Yes"
                            } elseif ($pendingKBs.ContainsKey($update.KB)) {
                                $verified = "No"
                            } elseif ($hasVerification) {
                                $verified = "Pending"
                            }
                        }

                        # Auto-correct: if PSWindowsUpdate said "Failed" but OS confirms installed
                        if ($isFailure -and $verified -eq "Yes" -and -not $isDefenderDefinition) {
                            $summary.Failed--
                            $summary.Installed++
                            $actualStatus = "Installed"
                            $isFailure = $false
                            $isSuccess = $true
                        }

                        $allUpdateData += [PSCustomObject]@{
                            ComputerName = $computerName
                            ComputerIP = $computerIP
                            KB = $update.KB
                            Title = $update.Title
                            Size = $update.Size
                            Result = $resultValue
                            Status = $actualStatus
                            OriginalStatus = $statusValue
                            ComputerNameClean = $computerName -replace '[^a-zA-Z0-9]', '_'
                            IsFailure = $isFailure
                            Verified = $verified
                        }
                    }

                    $summary.TotalUpdates = $seenUpdates.Count
                    $summary.RebootRequired = $rebootStatus

                    Write-Host "$($summary.TotalUpdates) updates " -NoNewline
                    if ($summary.Installed -gt 0) {
                        Write-Host "(✅ $($summary.Installed) installed" -NoNewline -ForegroundColor Green
                    }
                    if ($summary.Failed -gt 0) {
                        if ($summary.Installed -gt 0) { Write-Host ", " -NoNewline } else { Write-Host "(" -NoNewline }
                        Write-Host "❌ $($summary.Failed) failed" -NoNewline -ForegroundColor Red
                    }
                    Write-Host ")" -ForegroundColor Green

                } else {
                    # CSV exists but is empty
                    if ($result.CompletedMarker) {
                        Write-Host "No updates needed" -ForegroundColor Green
                        $summary.Status = "NoUpdatesNeeded"
                    } else {
                        Write-Host "CSV empty, completion marker missing - inconclusive" -ForegroundColor Yellow
                        $summary.Status = "Inconclusive"
                    }
                }

            } elseif (-not $result.CsvExists) {
                Write-Host "Update log not found on remote" -ForegroundColor Yellow
                $summary.Status = "NoLogFile"
            } else {
                Write-Host "CSV exists but content was empty" -ForegroundColor Yellow
                if ($result.CompletedMarker) {
                    $summary.Status = "NoUpdatesNeeded"
                } else {
                    $summary.Status = "Inconclusive"
                }
            }
        } catch {
            Write-Host "Processing error: $($_.Exception.Message)" -ForegroundColor Red
            $summary.Status = "CollectionFailed"
            $summary.CollectionError = $_.Exception.Message
        }

        # Finalize status
        if ($summary.TotalUpdates -gt 0 -and $summary.Status -ne "CollectionFailed") {
            $summary.Status = "Completed"
        }

        $computerSummary += $summary
    }
}
 
# Save collected data
$computerSummary | Export-Csv "$sessionReportPath\computer_summary.csv" -NoTypeInformation
if ($allUpdateData.Count -gt 0) {
    $allUpdateData | Export-Csv "$sessionReportPath\all_updates.csv" -NoTypeInformation
}

# Re-run CSV: computers that need another pass (real failures or did not complete)
$rerunList = $computerSummary | Where-Object {
    $_.Failed -gt 0 -or
    $_.Status -in @('Incomplete', 'CollectionFailed', 'Unreachable', 'Ignored-WSMAN', 'JobStartFailed', 'Inconclusive')
}
if ($rerunList) {
    $rerunCsvPath = Join-Path $sessionReportPath "rerun_computers.csv"
    $rerunList |
        Select-Object @{N='Name'; E={$_.ComputerName}}, IP |
        Export-Csv -Path $rerunCsvPath -NoTypeInformation
}

Write-Host "`nCollection phase completed" -ForegroundColor Green
Write-Host "Total updates collected: $($allUpdateData.Count)" -ForegroundColor Cyan
$failedComputers = ($computerSummary | Where-Object { $_.Failed -gt 0 }).Count
if ($failedComputers -gt 0) {
    Write-Host "⚠️ $failedComputers computer(s) have failed updates" -ForegroundColor Red
}

} # End of Full Deployment Mode (else block)

# ============================================
# PHASE 4: GENERATE HTML REPORT
# ============================================
 
Write-Host "`n╔══ Phase 4: Generating HTML Report ══╗" -ForegroundColor Cyan
 
# Calculate overall statistics
$totalComputers = $computerSummary.Count
$totalCompleted = ($computerSummary | Where-Object { $_.Status -eq "Completed" }).Count
$totalUpdates = ($computerSummary | Measure-Object -Property TotalUpdates -Sum).Sum
$totalInstalled = ($computerSummary | Measure-Object -Property Installed -Sum).Sum
$totalFailed = ($computerSummary | Measure-Object -Property Failed -Sum).Sum
$computersNeedReboot = ($computerSummary | Where-Object { $_.RebootRequired -eq $true -or $_.RebootRequired -eq "True" }).Count

# Verification statistics
$verifiedCount = ($allUpdateData | Where-Object { $_.Verified -eq 'Yes' }).Count
$verifiableCount = ($allUpdateData | Where-Object { $_.Verified -ne 'N/A' -and $_.Verified -ne '' }).Count
$discrepancies = @($allUpdateData | Where-Object {
    ($_.Status -eq 'Installed' -and $_.Verified -eq 'No') -or
    ($_.Status -in @('Failed', 'Aborted') -and $_.Verified -eq 'Yes')
})
$discrepancyComputers = @($discrepancies | Select-Object -ExpandProperty ComputerName -Unique)

# Load previous run data for rerun diff (keyed by computer name)
$previousRunData = @{}
foreach ($cs in $computerSummary) {
    if ($cs.PreviousRunArchive -and $cs.PreviousRunArchive -ne '') {
        $prevFile = "$sessionReportPath\previous_updatelog_$($cs.ComputerName).csv"
        if (Test-Path $prevFile) {
            try {
                $prevUpdates = Import-Csv $prevFile -ErrorAction Stop
                $previousRunData[$cs.ComputerName] = $prevUpdates
            } catch { }
        }
    }
}
$rerunComputers = @($computerSummary | Where-Object { $_.PreviousRunArchive -ne '' })

# Get computers with failures for the alert section
$computersWithFailures = $computerSummary | Where-Object { $_.Failed -gt 0 } | Sort-Object Failed -Descending
 
Write-Host "Found $($computersWithFailures.Count) computers with failures" -ForegroundColor Yellow
 
# Group updates by computer for HTML
$computerGroups = $allUpdateData | Group-Object ComputerName
 
# Build list of unique updates (by Title or KB if available)
$uniqueUpdates = @{}
foreach ($u in $allUpdateData) {
    $key = if ($u.KB -and $u.KB -ne '') { $u.KB } else { $u.Title }
    if (-not $uniqueUpdates.ContainsKey($key)) {
        $uniqueUpdates[$key] = @{ Title = $u.Title; KB = $u.KB }
    }
}
 
# HTML encoding helper — prevents malformed HTML if computer names or update titles
# contain characters like <, >, &, or " (rare but possible with certain naming conventions).
function ConvertTo-HtmlEncoded([string]$text) {
    $text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
}

# Build JS array for the HTML re-run banner (injected as a const into the report)
$rerunJsArray = if ($rerunList -and $rerunList.Count -gt 0) {
    $entries = $rerunList | ForEach-Object {
        # ReportOnly loads CSV with Name column; live mode uses ComputerName
        $compName = if ($_.ComputerName) { $_.ComputerName } else { $_.Name }
        $reason = if ($_.Failed -gt 0) {
            "$($_.Failed) failed update$(if($_.Failed -ne 1){'s'})"
        } elseif ($_.Status) { $_.Status } else { "Needs rerun" }
        $safeName = $compName -replace '"', '\"' -replace '\\', '\\\\'
        $safeIP   = $_.IP -replace '"', '\"'
        "{name:`"$safeName`",ip:`"$safeIP`",reason:`"$reason`"}"
    }
    "[" + ($entries -join ",") + "]"
} else { "[]" }

# Generate HTML
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Windows Update Report - $(Get-Date -Format "yyyy-MM-dd HH:mm")</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: #f0f2f5; height: 100vh; display: flex; flex-direction: column; overflow: hidden;
        }
        /* Top bar */
        .topbar {
            background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
            color: white; padding: 16px 24px; flex-shrink: 0;
            display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px;
        }
        .topbar h1 { font-size: 1.4em; font-weight: 600; display: flex; align-items: center; gap: 10px; }
        .topbar-meta { font-size: 0.8em; opacity: 0.7; }
        /* Stats strip */
        .stats-strip {
            display: flex; flex-wrap: wrap; gap: 0; background: white; flex-shrink: 0;
            border-bottom: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.06);
        }
        .alerts-container { flex-shrink: 0; }
        .stat-item {
            flex: 1; min-width: 120px; padding: 14px 20px; text-align: center;
            border-right: 1px solid #e2e8f0; transition: background 0.2s;
        }
        .stat-item:last-child { border-right: none; }
        .stat-item:hover { background: #f8fafc; }
        .stat-value { font-size: 1.8em; font-weight: 700; line-height: 1.2; }
        .stat-label { font-size: 0.75em; text-transform: uppercase; letter-spacing: 0.5px; color: #64748b; margin-top: 2px; }
        .stat-item.blue .stat-value { color: #2563eb; }
        .stat-item.green .stat-value { color: #16a34a; }
        .stat-item.red .stat-value { color: #dc2626; }
        .stat-item.amber .stat-value { color: #d97706; }
        .stat-item.purple .stat-value { color: #7c3aed; }
        .stat-item.gray .stat-value { color: #94a3b8; }
        /* Alert banners */
        .alert-banner {
            margin: 12px 16px 0; border-radius: 10px; overflow: hidden;
            box-shadow: 0 1px 4px rgba(0,0,0,0.08);
        }
        .alert-header {
            padding: 10px 16px; display: flex; align-items: center; gap: 10px;
            cursor: pointer; user-select: none; font-weight: 600; font-size: 0.9em;
        }
        .alert-header .toggle-icon { margin-left: auto; transition: transform 0.2s; font-size: 0.8em; }
        .alert-header.collapsed .toggle-icon { transform: rotate(-90deg); }
        .alert-body { padding: 0 16px 12px; }
        .alert-body.collapsed { display: none; }
        .alert-banner.failure { background: #fef2f2; border: 1px solid #fecaca; }
        .alert-banner.failure .alert-header { color: #991b1b; }
        .alert-banner.connection { background: #fdf4ff; border: 1px solid #f5d0fe; }
        .alert-banner.connection .alert-header { color: #86198f; }
        .alert-banner.rerun { background: #fffbeb; border: 1px solid #fde68a; }
        .alert-banner.rerun .alert-header { color: #92400e; }
        .alert-chip {
            display: inline-flex; align-items: center; gap: 6px;
            padding: 4px 10px; border-radius: 6px; font-size: 0.82em;
            cursor: pointer; transition: all 0.15s; margin: 3px;
        }
        .alert-chip:hover { filter: brightness(0.95); transform: translateX(2px); }
        .alert-chip.fail { background: white; border: 1px solid #fca5a5; color: #991b1b; }
        .alert-chip.conn { background: white; border: 1px solid #e9d5ff; color: #86198f; }
        .alert-chip .chip-count { background: #ef4444; color: white; padding: 1px 7px; border-radius: 10px; font-weight: 700; font-size: 0.85em; }
        /* Dashboard layout */
        .dashboard { display: flex; flex: 1; min-height: 0; }
        .sidebar {
            width: 240px; flex-shrink: 0; background: white; border-right: 1px solid #e2e8f0;
            display: flex; flex-direction: column;
        }
        .sidebar-section { padding: 14px 16px; border-bottom: 1px solid #f1f5f9; }
        .sidebar-section h3 { font-size: 0.75em; text-transform: uppercase; letter-spacing: 0.5px; color: #94a3b8; margin-bottom: 8px; }
        .sidebar input[type="text"] {
            width: 100%; padding: 8px 10px; border: 1px solid #e2e8f0; border-radius: 6px;
            font-size: 0.85em; outline: none; transition: border-color 0.2s;
        }
        .sidebar input[type="text"]:focus { border-color: #3b82f6; box-shadow: 0 0 0 2px rgba(59,130,246,0.1); }
        /* Status filter pills */
        .status-pills { display: flex; flex-wrap: wrap; gap: 4px; }
        .status-pill {
            padding: 4px 10px; border-radius: 20px; font-size: 0.78em; font-weight: 500;
            cursor: pointer; border: 1px solid #e2e8f0; background: white; transition: all 0.15s;
        }
        .status-pill:hover { background: #f1f5f9; }
        .status-pill.active { background: #1e293b; color: white; border-color: #1e293b; }
        /* Sidebar buttons */
        .sidebar-btn {
            display: block; width: 100%; padding: 7px 10px; margin-bottom: 4px;
            border: 1px solid #e2e8f0; border-radius: 6px; background: white;
            font-size: 0.82em; cursor: pointer; text-align: left; transition: all 0.15s;
        }
        .sidebar-btn:hover { background: #f8fafc; border-color: #cbd5e1; }
        .sidebar-btn.danger { color: #dc2626; border-color: #fecaca; }
        .sidebar-btn.danger:hover { background: #fef2f2; }
        /* Update filter checkboxes */
        .update-filter-list { max-height: 250px; overflow-y: auto; }
        .update-filter-item {
            display: flex; align-items: center; gap: 6px; padding: 3px 0; font-size: 0.8em; color: #374151;
        }
        .update-filter-item input { flex-shrink: 0; }
        .update-filter-item span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        /* Main content area */
        .main-content { flex: 1; overflow: auto; padding: 0; }
        /* Computer table */
        .computer-table { width: 100%; border-collapse: collapse; }
        .computer-table thead { position: sticky; top: 0; z-index: 2; }
        .computer-table th {
            background: #f8fafc; color: #475569; padding: 8px 8px; text-align: left;
            font-size: 0.72em; text-transform: uppercase; letter-spacing: 0.3px; font-weight: 600;
            border-bottom: 2px solid #e2e8f0; white-space: nowrap; cursor: pointer; user-select: none;
        }
        .computer-table th:hover { background: #f1f5f9; }
        .computer-row {
            cursor: pointer; transition: background 0.1s;
        }
        .computer-row:hover { background: #f8fafc; }
        .computer-row td {
            padding: 7px 8px; border-bottom: 1px solid #f1f5f9;
            font-size: 0.82em; white-space: nowrap;
        }
        .computer-row.has-failures { background: #fff5f5; }
        .computer-row.has-failures:hover { background: #fef2f2; }
        .computer-row.unreachable { background: #fdf4ff; }
        .computer-row.expanded { background: #eff6ff; }
        .computer-row td:first-child { width: 30px; text-align: center; }
        .computer-name-cell { max-width: 220px; overflow: hidden; text-overflow: ellipsis; font-weight: 500; }
        .ip-cell { color: #64748b; font-family: 'SF Mono', Consolas, monospace; font-size: 0.78em; }
        .count-cell { text-align: center; font-weight: 600; min-width: 36px; }
        .count-cell.installed { color: #16a34a; }
        .count-cell.failed { color: #dc2626; }
        .count-cell.skipped { color: #94a3b8; }
        .badge-sm {
            display: inline-block; padding: 2px 8px; border-radius: 10px;
            font-size: 0.75em; font-weight: 600;
        }
        .badge-sm.reboot { background: #fef3c7; color: #92400e; }
        .badge-sm.status-ok { background: #dcfce7; color: #166534; }
        .badge-sm.status-fail { background: #fee2e2; color: #991b1b; }
        .badge-sm.status-warn { background: #fef3c7; color: #92400e; }
        .badge-sm.status-skip { background: #f1f5f9; color: #475569; }
        .expand-chevron { transition: transform 0.2s; font-size: 0.7em; color: #94a3b8; }
        .computer-row.expanded .expand-chevron { transform: rotate(90deg); }
        /* Detail panel (hidden row beneath each computer) */
        .detail-row { display: none; }
        .detail-row.visible { display: table-row; }
        .detail-panel {
            padding: 12px 20px 16px; background: #f8fafc; border-bottom: 2px solid #e2e8f0;
        }
        .detail-panel .reboot-note {
            background: #fef3c7; border-left: 3px solid #f59e0b; padding: 6px 12px;
            font-size: 0.85em; color: #92400e; margin-bottom: 10px; border-radius: 0 6px 6px 0;
        }
        .detail-panel .no-data { padding: 20px; text-align: center; color: #94a3b8; font-size: 0.9em; }
        /* Update sub-table inside detail panel */
        .update-table { width: 100%; border-collapse: collapse; font-size: 0.85em; }
        .update-table th {
            background: #e2e8f0; color: #475569; padding: 7px 10px; text-align: left;
            font-size: 0.8em; text-transform: uppercase; letter-spacing: 0.3px;
            position: static;
        }
        .update-table td { padding: 6px 10px; border-bottom: 1px solid #e2e8f0; }
        .update-table tr.failed-row { background: #fef2f2; }
        .update-table tr.hidden { display: none; }
        .kb-badge {
            background: #e0e7ff; color: #3730a3; padding: 2px 8px; border-radius: 4px;
            font-family: 'SF Mono', Consolas, monospace; font-size: 0.9em;
        }
        .status-icon { font-weight: 500; }
        .status-installed { color: #16a34a; }
        .status-failed { color: #dc2626; }
        .status-warning { color: #d97706; }
        /* Footer */
        .footer {
            text-align: center; padding: 16px; color: #94a3b8; font-size: 0.8em;
            border-top: 1px solid #e2e8f0; background: white;
        }
        /* Utility */
        .hidden { display: none !important; }
        @keyframes highlightRow { 0% { box-shadow: inset 0 0 0 2px #3b82f6; } 100% { box-shadow: none; } }
        /* Responsive */
        @media (max-width: 900px) {
            body { height: auto; overflow: auto; }
            .dashboard { flex-direction: column; flex: none; height: auto; }
            .sidebar { width: 100%; max-height: 300px; border-right: none; border-bottom: 1px solid #e2e8f0; overflow-y: auto; }
            .main-content { overflow: visible; }
            .stats-strip { flex-wrap: wrap; }
            .stat-item { min-width: 100px; }
        }
    </style>
</head>
<body>
    <div class="topbar">
        <h1><span>&#128260;</span> Windows Update Deployment Report</h1>
        <div class="topbar-meta">
            <div>$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</div>
            <div style="opacity:0.6;font-family:monospace;font-size:0.9em;">Session $timestamp</div>
        </div>
    </div>

    <div class="stats-strip">
        <div class="stat-item blue"><div class="stat-value">$totalComputers</div><div class="stat-label">Computers</div></div>
        <div class="stat-item green"><div class="stat-value">$totalCompleted</div><div class="stat-label">Completed</div></div>
        <div class="stat-item green"><div class="stat-value" id="stat-installed-value">$totalInstalled</div><div class="stat-label">Installed</div></div>
        <div class="stat-item red"><div class="stat-value" id="stat-failed-value">$totalFailed</div><div class="stat-label">Failed</div></div>
        <div class="stat-item amber"><div class="stat-value">$computersNeedReboot</div><div class="stat-label">Need Reboot</div></div>
        <div class="stat-item purple"><div class="stat-value" id="stat-total-value">$totalUpdates</div><div class="stat-label">Total Updates</div></div>
        <div class="stat-item $(if ($verifiableCount -gt 0 -and $verifiedCount -eq $verifiableCount) {'green'} elseif ($verifiedCount -gt 0) {'blue'} else {'gray'})"><div class="stat-value">$verifiedCount / $verifiableCount</div><div class="stat-label">Verified</div></div>
    </div>
"@

$html += "`n    <div class=`"alerts-container`">"

# Alert banners
if ($computersWithFailures.Count -gt 0) {
    $html += @"
    <div class="alert-banner failure">
        <div class="alert-header" onclick="toggleAlert(this)">
            <span>&#9888;&#65039;</span>
            <span>$($computersWithFailures.Count) computer$(if($computersWithFailures.Count -ne 1){'s'}) with failed updates</span>
            <span class="toggle-icon">&#9660;</span>
        </div>
        <div class="alert-body" id="failure-alert-body">
"@
    foreach ($fc in $computersWithFailures) {
        $cn = $fc.ComputerName -replace '[^a-zA-Z0-9]', '_'
        $safeName = ConvertTo-HtmlEncoded $fc.ComputerName
        $html += "            <span class=`"alert-chip fail`" data-computer=`"$cn`" onclick=`"scrollToComputer('$cn')`">$safeName <span class=`"chip-count`"><span class=`"failure-count-number`">$($fc.Failed)</span></span></span>`n"
    }
    $html += @"
        </div>
    </div>
"@
}

$computersWithConnectionFailures = $computerSummary | Where-Object { $_.Status -eq "Ignored-WSMAN" -or $_.Status -eq "Incomplete" -or $_.Status -eq "CollectionFailed" -or $_.Status -eq "JobStartFailed" } | Sort-Object ComputerName
if ($computersWithConnectionFailures.Count -gt 0) {
    $html += @"
    <div class="alert-banner connection">
        <div class="alert-header" onclick="toggleAlert(this)">
            <span>&#128225;</span>
            <span><span id="connection-failure-count">$($computersWithConnectionFailures.Count)</span> computer$(if($computersWithConnectionFailures.Count -ne 1){'s'}) unable to connect</span>
            <span class="toggle-icon">&#9660;</span>
        </div>
        <div class="alert-body">
"@
    foreach ($cc in $computersWithConnectionFailures) {
        $cn = $cc.ComputerName -replace '[^a-zA-Z0-9]', '_'
        $safeName = ConvertTo-HtmlEncoded $cc.ComputerName
        $statusReason = switch ($cc.Status) {
            "Ignored-WSMAN"    { "WSMan failed" }
            "Incomplete"       { "Incomplete" }
            "CollectionFailed" { "Collection failed" }
            "JobStartFailed"   { "Job start failed" }
            default            { $cc.Status }
        }
        $html += "            <span class=`"alert-chip conn`" onclick=`"scrollToComputer('$cn')`">$safeName &mdash; $statusReason</span>`n"
    }
    $html += @"
        </div>
    </div>
"@
}

# Verification discrepancy banner
if ($discrepancies.Count -gt 0) {
    $html += @"
    <div class="alert-banner connection" style="background:#fefce8;border-color:#fde047;">
        <div class="alert-header" style="color:#854d0e;" onclick="toggleAlert(this)">
            <span>&#128269;</span>
            <span>$($discrepancies.Count) update$(if($discrepancies.Count -ne 1){'s'}) with verification discrepancies</span>
            <span class="toggle-icon">&#9660;</span>
        </div>
        <div class="alert-body" style="font-size:0.85em;color:#854d0e;">
"@
    foreach ($dc in $discrepancyComputers) {
        $cn = $dc -replace '[^a-zA-Z0-9]', '_'
        $safeName = ConvertTo-HtmlEncoded $dc
        $dcCount = ($discrepancies | Where-Object { $_.ComputerName -eq $dc }).Count
        $html += "            <span class=`"alert-chip`" style=`"background:#fef9c3;border-color:#facc15;color:#854d0e;`" onclick=`"scrollToComputer('$cn')`">$safeName <span class=`"chip-count`" style=`"background:#eab308;`">$dcCount</span></span>`n"
    }
    $html += @"
        </div>
    </div>
"@
}

# Rerun banner (hidden by default, shown by JS)
$html += @"
    <div class="alert-banner rerun" id="rerun-banner" style="display:none">
        <div class="alert-header" onclick="toggleAlert(this)">
            <span>&#128260;</span>
            <span><span id="rerun-count"></span> computer(s) queued for re-run</span>
            <button onclick="event.stopPropagation();downloadRerunCSV()" style="margin-left:auto;padding:4px 12px;border-radius:6px;border:1px solid #d97706;background:white;color:#92400e;font-size:0.82em;cursor:pointer;">Download Re-run CSV</button>
            <span class="toggle-icon">&#9660;</span>
        </div>
        <div class="alert-body" id="rerun-detail" style="font-size:0.85em;color:#92400e;"></div>
    </div>
    </div>

    <div class="dashboard">
        <div class="sidebar">
            <div class="sidebar-section">
                <h3>Search</h3>
                <input type="text" id="search-input" placeholder="Computer name or IP..." onkeyup="filterComputers(this.value)">
            </div>
            <div class="sidebar-section">
                <h3>Status</h3>
                <div class="status-pills">
                    <span class="status-pill active" data-filter="all" onclick="filterByStatus('all',this)">All</span>
                    <span class="status-pill" data-filter="Completed" onclick="filterByStatus('Completed',this)">Completed</span>
                    <span class="status-pill" data-filter="NoUpdatesNeeded" onclick="filterByStatus('NoUpdatesNeeded',this)">No Updates</span>
                    <span class="status-pill" data-filter="failures" onclick="filterByStatus('failures',this)">Failures</span>
                    <span class="status-pill" data-filter="unreachable" onclick="filterByStatus('unreachable',this)">Unreachable</span>
                </div>
            </div>
            <div class="sidebar-section">
                <h3>Actions</h3>
                <button class="sidebar-btn" onclick="expandAll()">&#128194; Expand All</button>
                <button class="sidebar-btn" onclick="collapseAll()">&#128193; Collapse All</button>
                <button class="sidebar-btn" onclick="sortComputersByStatus()">&#128260; Sort by Status</button>
"@

if ($computersWithFailures.Count -gt 0) {
    $html += "                <button class=`"sidebar-btn danger`" onclick=`"expandFailures()`">&#9888;&#65039; Show Failures Only</button>`n"
}

$html += @"
            </div>
            <div class="sidebar-section" style="flex:1;overflow:hidden;display:flex;flex-direction:column;">
                <h3>Filter Updates
                    <label style="float:right;font-size:1.1em;text-transform:none;letter-spacing:0;font-weight:400;cursor:pointer;">
                        <input type="checkbox" id="select-all-updates" onclick="(this.checked?selectAllUpdates():selectNoneUpdates())" checked> All
                    </label>
                </h3>
                <div id="update-filters-body" class="update-filter-list" style="flex:1;overflow-y:auto;">
                    <div id="update-filters-failed-list"></div>
                    <div id="update-filters-normal-list" style="margin-top:6px;"></div>
                </div>
            </div>
            <div class="sidebar-section" style="font-size:0.78em;color:#94a3b8;">
                <span id="filtered-computer-count">$($computerSummary.Count)</span> / $($computerSummary.Count) computers shown
            </div>
        </div>

        <div class="main-content">
            <table class="computer-table">
                <thead>
                    <tr>
                        <th></th>
                        <th>Status</th>
                        <th>Computer</th>
                        <th>IP</th>
                        <th style="text-align:center">Installed</th>
                        <th style="text-align:center">Failed</th>
                        <th style="text-align:center">Skipped</th>
                        <th style="text-align:center">Total</th>
                        <th>Reboot</th>
                    </tr>
                </thead>
                <tbody id="computers-tbody">
"@

# Add each computer as a table row + hidden detail row
foreach ($summary in $computerSummary | Sort-Object @{Expression={if([int]$_.Failed -gt 0){0}elseif($_.Status -in @('Incomplete','CollectionFailed','Ignored-WSMAN','JobStartFailed')){1}else{2}}}, ComputerName) {
    $computerName = $summary.ComputerName
    $group = $computerGroups | Where-Object { $_.Name -eq $computerName }
    $computerData = if ($group) { $group.Group } else { @() }

    $cleanName = $computerName -replace '[^a-zA-Z0-9]', '_'
    $safeComputerName = ConvertTo-HtmlEncoded $computerName
    $hasFailures = [int]$summary.Failed -gt 0
    $isUnreachable = $summary.Status -in @("Ignored-WSMAN", "Incomplete", "CollectionFailed", "JobStartFailed", "Inconclusive")

    $statusBadge = switch ($summary.Status) {
        "Completed"        { '<span class="badge-sm status-ok">Completed</span>' }
        "NoUpdatesNeeded"  { '<span class="badge-sm status-skip">Up to Date</span>' }
        "Incomplete"       { '<span class="badge-sm status-warn">Incomplete</span>' }
        "Inconclusive"     { '<span class="badge-sm status-warn">Inconclusive</span>' }
        "Ignored-WSMAN"    { '<span class="badge-sm status-fail">Unreachable</span>' }
        "CollectionFailed" { '<span class="badge-sm status-fail">Collection Failed</span>' }
        "JobStartFailed"   { '<span class="badge-sm status-fail">Job Failed</span>' }
        default            { "<span class=`"badge-sm status-warn`">$($summary.Status)</span>" }
    }

    $rowClass = "computer-row"
    if ($hasFailures) { $rowClass += " has-failures" }
    if ($isUnreachable) { $rowClass += " unreachable" }

    $rebootBadge = if ($summary.RebootRequired -eq $true -or $summary.RebootRequired -eq "True") {
        '<span class="badge-sm reboot">Yes</span>'
    } else { "" }

    $rerunBadge = if ($summary.PreviousRunArchive -and $summary.PreviousRunArchive -ne '') {
        ' <span class="badge-sm" style="background:#e0e7ff;color:#4338ca;font-size:0.7em;">Rerun</span>'
    } else { "" }

    $installedCount = if ([int]$summary.Installed -gt 0) { $summary.Installed } else { "-" }
    $failedCount = if ([int]$summary.Failed -gt 0) { $summary.Failed } else { "-" }
    $skippedCount = if ([int]$summary.Skipped -gt 0) { $summary.Skipped } else { "-" }
    $totalCount = if ([int]$summary.TotalUpdates -gt 0) { $summary.TotalUpdates } else { "-" }

    $html += @"
                    <tr class="$rowClass" data-computer="$computerName" data-status="$($summary.Status)" id="computer-$cleanName" onclick="toggleComputer('$cleanName')">
                        <td><span class="expand-chevron">&#9654;</span></td>
                        <td>$statusBadge</td>
                        <td class="computer-name-cell" title="$safeComputerName">$safeComputerName$rerunBadge</td>
                        <td class="ip-cell">$(ConvertTo-HtmlEncoded $summary.IP)</td>
                        <td class="count-cell installed">$installedCount</td>
                        <td class="count-cell failed">$failedCount</td>
                        <td class="count-cell skipped">$skippedCount</td>
                        <td class="count-cell" style="font-weight:400;color:#64748b;">$totalCount</td>
                        <td>$rebootBadge</td>
                    </tr>
                    <tr class="detail-row" id="detail-$cleanName">
                        <td colspan="9" style="padding:0;">
                            <div class="detail-panel">
"@

    if ($summary.RebootRequired -eq $true -or $summary.RebootRequired -eq "True") {
        $html += "                                <div class=`"reboot-note`">&#9888;&#65039; This computer requires a reboot to complete update installation.</div>`n"
    }

    if ($computerData.Count -gt 0) {
        $html += @"
                                <table class="update-table">
                                    <thead><tr>
                                        <th style="width:10%">KB</th>
                                        <th style="width:44%">Update Title</th>
                                        <th style="width:8%">Size</th>
                                        <th style="width:14%">Status</th>
                                        <th style="width:12%">Result</th>
                                        <th style="width:12%">Verified</th>
                                    </tr></thead>
                                    <tbody>
"@
        foreach ($update in $computerData | Sort-Object KB) {
            $resultText = if ($update.Result) { $update.Result.ToString() } else { "" }
            $statusText = if ($update.Status) { $update.Status.ToString() } else { "" }

            $isFailedRow = $false
            $statusDisplay = switch -Regex ($statusText) {
                "^Skipped$"  { '<span class="status-icon" style="color:#94a3b8;">&#9197; Skipped</span>'; break }
                "^Failed$|^Aborted$" { $isFailedRow = $true; '<span class="status-icon status-failed">&#10060; ' + $statusText + '</span>'; break }
                "^Installed$" { '<span class="status-icon status-installed">&#9989; Installed</span>'; break }
                "^InstalledWithErrors$" { '<span class="status-icon status-warning">&#9888;&#65039; With Errors</span>'; break }
                default { '<span class="status-icon">&#10067; ' + $statusText + '</span>' }
            }

            $kbDisplay = if ($update.KB -and $update.KB -ne "N/A" -and $update.KB -ne "") {
                "<span class='kb-badge'>$($update.KB)</span>"
            } else {
                "<span style='color:#94a3b8;'>-</span>"
            }

            $rowClass = if ($isFailedRow) { ' class="failed-row"' } else { '' }
            $uKey = if ($update.KB -and $update.KB -ne '') { $update.KB } else { $update.Title -replace '\s+', '_' -replace '[^A-Za-z0-9_\-]', '' }
            $safeTitle = ConvertTo-HtmlEncoded $update.Title

            $verifiedVal = if ($update.Verified) { $update.Verified.ToString() } else { "N/A" }
            $verifiedDisplay = switch ($verifiedVal) {
                "Yes"     { '<span style="color:#16a34a;" title="Confirmed installed via Get-HotFix">&#10003; Yes</span>' }
                "No"      { '<span style="color:#dc2626;" title="Still listed as needed by Windows Update">&#10007; No</span>' }
                "Pending" { '<span style="color:#d97706;" title="Not in hotfix list yet — may need reboot">&#8987; Pending</span>' }
                default   { '<span style="color:#94a3b8;" title="No KB to verify (driver/other)">&mdash;</span>' }
            }

            $html += @"
                                        <tr$rowClass data-ukey='$uKey'>
                                            <td>$kbDisplay</td>
                                            <td>$safeTitle</td>
                                            <td>$($update.Size)</td>
                                            <td>$statusDisplay</td>
                                            <td style="text-align:center;color:$(if($isFailedRow){'#ef4444'}else{'#64748b'});">$resultText</td>
                                            <td style="text-align:center;">$verifiedDisplay</td>
                                        </tr>
"@
        }

        $html += @"
                                    </tbody>
                                </table>
"@
    } elseif ($summary.Status -eq "NoUpdatesNeeded") {
        $html += "                                <div class=`"no-data`">&#9989; System is up to date &mdash; no updates needed</div>`n"
    } elseif ($summary.CollectionError) {
        $safeError = ConvertTo-HtmlEncoded $summary.CollectionError
        $html += "                                <div class=`"no-data`" style=`"color:#dc2626;`">Error: $safeError</div>`n"
    } else {
        $html += "                                <div class=`"no-data`">No update data available</div>`n"
    }

    # Previous run diff section (for same-day reruns)
    if ($previousRunData.ContainsKey($computerName)) {
        $prevUpdates = $previousRunData[$computerName]
        $prevKBs = @{}
        foreach ($pu in $prevUpdates) {
            $pk = if ($pu.KB -and $pu.KB -ne '') { $pu.KB } else { $pu.Title }
            $prevKBs[$pk] = $pu
        }
        # Find new updates in this run that weren't in the previous run
        $newInThisRun = @($computerData | Where-Object {
            $k = if ($_.KB -and $_.KB -ne '') { $_.KB } else { $_.Title }
            -not $prevKBs.ContainsKey($k)
        })
        $newBadge = if ($newInThisRun.Count -gt 0) { " &mdash; <strong>$($newInThisRun.Count) new</strong>" } else { "" }

        $html += @"
                                <details class="prev-run-details" style="margin-top:12px;">
                                    <summary style="cursor:pointer;font-size:0.85em;color:#6366f1;font-weight:600;">&#128337; Previous Run ($($summary.PreviousRunArchive -replace 'archive_',''))$newBadge</summary>
                                    <table class="update-table" style="margin-top:6px;opacity:0.75;">
                                        <thead><tr>
                                            <th style="width:12%">KB</th>
                                            <th style="width:52%">Update Title</th>
                                            <th style="width:18%">Status</th>
                                            <th style="width:18%">Result</th>
                                        </tr></thead>
                                        <tbody>
"@
        foreach ($pu in $prevUpdates) {
            $puKb = if ($pu.KB -and $pu.KB -ne 'N/A' -and $pu.KB -ne '') {
                "<span class='kb-badge'>$(ConvertTo-HtmlEncoded $pu.KB)</span>"
            } else { "<span style='color:#94a3b8;'>-</span>" }
            $puTitle = ConvertTo-HtmlEncoded $pu.Title
            $puResult = if ($pu.Result) { ConvertTo-HtmlEncoded $pu.Result.ToString() } else { "" }
            $puStatus = if ($pu.Status) { ConvertTo-HtmlEncoded $pu.Status.ToString() } else { "" }
            $html += "                                            <tr><td>$puKb</td><td>$puTitle</td><td>$puStatus</td><td>$puResult</td></tr>`n"
        }
        $html += @"
                                        </tbody>
                                    </table>
                                </details>
"@
    }

    $html += @"
                            </div>
                        </td>
                    </tr>
"@
}

$html += @"
                </tbody>
            </table>
        </div>
    </div>

    <div class="footer">
        Windows Update Deployment Report &middot; Session $timestamp &middot; $sessionReportPath
    </div>

    <script>
        const uniqueUpdates = {
"@
foreach ($k in $uniqueUpdates.Keys) {
    $v = $uniqueUpdates[$k]
    $safeKey = $k -replace "'", "\'"
    $safeTitle = ($v.Title -replace "'", "\'")
    $null = $html += "            '$safeKey': '$safeTitle',`n"
}
$null = $html += "        };`n"
$null = $html += "        const rerunComputers = $rerunJsArray;`n"

$null = $html += @"

        function toggleAlert(header) {
            header.classList.toggle('collapsed');
            const body = header.nextElementSibling;
            if (body) body.classList.toggle('collapsed');
        }

        let currentStatusFilter = 'all';
        let sortDescending = false;

        function filterByStatus(status, pill) {
            currentStatusFilter = status;
            document.querySelectorAll('.status-pill').forEach(p => p.classList.remove('active'));
            if (pill) pill.classList.add('active');
            applyFilters();
        }

        function filterComputers(searchText) {
            applyFilters();
        }

        function applyFilters() {
            const searchText = (document.getElementById('search-input').value || '').toLowerCase();
            let visibleCount = 0;

            document.querySelectorAll('.computer-row').forEach(row => {
                const name = (row.getAttribute('data-computer') || '').toLowerCase();
                const status = row.getAttribute('data-status') || '';
                const ip = row.querySelector('.ip-cell') ? row.querySelector('.ip-cell').textContent.toLowerCase() : '';

                let show = true;

                // Status filter
                if (currentStatusFilter !== 'all') {
                    if (currentStatusFilter === 'failures') {
                        show = row.classList.contains('has-failures');
                    } else if (currentStatusFilter === 'unreachable') {
                        show = row.classList.contains('unreachable');
                    } else {
                        show = status === currentStatusFilter;
                    }
                }

                // Search filter
                if (show && searchText) {
                    show = name.includes(searchText) || ip.includes(searchText);
                }

                row.classList.toggle('hidden', !show);
                // Also hide the detail row
                const detailId = 'detail-' + row.id.replace('computer-', '');
                const detail = document.getElementById(detailId);
                if (detail && !show) {
                    detail.classList.remove('visible');
                    row.classList.remove('expanded');
                }
                if (show) visibleCount++;
            });

            const countEl = document.getElementById('filtered-computer-count');
            if (countEl) countEl.textContent = visibleCount;
        }

        function toggleComputer(cleanName) {
            const row = document.getElementById('computer-' + cleanName);
            const detail = document.getElementById('detail-' + cleanName);
            if (!row || !detail) return;
            const isExpanded = row.classList.contains('expanded');
            row.classList.toggle('expanded');
            detail.classList.toggle('visible');
            if (!isExpanded) {
                detail.querySelector('.detail-panel').scrollTop = 0;
            }
        }

        function expandAll() {
            document.querySelectorAll('.computer-row:not(.hidden)').forEach(row => {
                row.classList.add('expanded');
                const cleanName = row.id.replace('computer-', '');
                const detail = document.getElementById('detail-' + cleanName);
                if (detail) detail.classList.add('visible');
            });
        }

        function collapseAll() {
            document.querySelectorAll('.computer-row').forEach(row => row.classList.remove('expanded'));
            document.querySelectorAll('.detail-row').forEach(d => d.classList.remove('visible'));
        }

        function expandFailures() {
            filterByStatus('failures', document.querySelector('.status-pill[data-filter="failures"]'));
            document.querySelectorAll('.computer-row.has-failures:not(.hidden)').forEach(row => {
                row.classList.add('expanded');
                const cleanName = row.id.replace('computer-', '');
                const detail = document.getElementById('detail-' + cleanName);
                if (detail) detail.classList.add('visible');
            });
        }

        function sortComputersByStatus() {
            const tbody = document.getElementById('computers-tbody');
            const pairs = [];
            const rows = Array.from(tbody.children);
            for (let i = 0; i < rows.length; i += 2) {
                pairs.push({ row: rows[i], detail: rows[i+1] });
            }
            const priority = { 'Failed':1, 'JobStartFailed':2, 'CollectionFailed':3, 'Incomplete':4, 'Ignored-WSMAN':5, 'Completed':6, 'NoUpdatesNeeded':7 };
            pairs.sort((a, b) => {
                const as = a.row.getAttribute('data-status') || '';
                const bs = b.row.getAttribute('data-status') || '';
                // Sort failures-with-count first
                const af = a.row.classList.contains('has-failures') ? 0 : 1;
                const bf = b.row.classList.contains('has-failures') ? 0 : 1;
                if (af !== bf) return sortDescending ? bf - af : af - bf;
                const diff = (priority[as]||99) - (priority[bs]||99);
                return sortDescending ? -diff : diff;
            });
            pairs.forEach(p => { tbody.appendChild(p.row); tbody.appendChild(p.detail); });
            sortDescending = !sortDescending;
        }

        function scrollToComputer(cleanName) {
            // Reset status filter to All so the target is visible
            filterByStatus('all', document.querySelector('.status-pill[data-filter="all"]'));
            document.getElementById('search-input').value = '';
            applyFilters();

            const row = document.getElementById('computer-' + cleanName);
            if (row) {
                row.scrollIntoView({ behavior: 'smooth', block: 'center' });
                row.classList.add('expanded');
                const detail = document.getElementById('detail-' + cleanName);
                if (detail) detail.classList.add('visible');
                row.style.animation = 'none';
                setTimeout(() => { row.style.animation = 'highlightRow 1.5s ease-out'; }, 50);
            }
        }

        // Update filter system
        function renderUpdateFilters() {
            const failedList = document.getElementById('update-filters-failed-list');
            const normalList = document.getElementById('update-filters-normal-list');
            if (!failedList || !normalList) return;

            Object.keys(uniqueUpdates).forEach(key => {
                const item = document.createElement('label');
                item.className = 'update-filter-item';

                const cb = document.createElement('input');
                cb.type = 'checkbox';
                const title = uniqueUpdates[key] || '';
                const isDefender = /Security Intelligence Update for Microsoft Defender/i.test(title);
                cb.checked = !isDefender;
                cb.dataset.ukey = key;
                cb.addEventListener('change', updateFiltersFromUI);

                const span = document.createElement('span');
                span.textContent = title.length > 50 ? title.substring(0,47)+'...' : title;
                span.title = title;

                item.appendChild(cb);
                item.appendChild(span);

                const hasFailed = document.querySelectorAll('tr.failed-row[data-ukey="' + key + '"]').length > 0;
                if (hasFailed) {
                    span.style.color = '#991b1b';
                    failedList.appendChild(item);
                } else {
                    normalList.appendChild(item);
                }
            });

            const allCbs = Array.from(document.querySelectorAll('#update-filters-body input[type=checkbox]'));
            const selectAll = document.getElementById('select-all-updates');
            if (selectAll) selectAll.checked = allCbs.every(cb => cb.checked);
            updateFiltersFromUI();
        }

        function updateFiltersFromUI() {
            const enabled = new Set();
            document.querySelectorAll('#update-filters-body input[type=checkbox]').forEach(cb => {
                if (cb.checked) enabled.add(cb.dataset.ukey);
            });

            document.querySelectorAll('tr[data-ukey]').forEach(row => {
                row.classList.toggle('hidden', !enabled.has(row.getAttribute('data-ukey')));
            });

            // Update per-computer failure indicators
            document.querySelectorAll('.computer-row').forEach(row => {
                const cleanName = row.id.replace('computer-', '');
                const detail = document.getElementById('detail-' + cleanName);
                if (!detail) return;
                const visibleFailed = detail.querySelectorAll('tr.failed-row:not(.hidden)').length;
                const failedCell = row.querySelector('.count-cell.failed');
                if (failedCell) {
                    failedCell.textContent = visibleFailed > 0 ? visibleFailed : '-';
                }
                if (visibleFailed > 0) {
                    row.classList.add('has-failures');
                } else {
                    row.classList.remove('has-failures');
                }
            });

            // Update top stats
            const totalVisibleFailed = document.querySelectorAll('tr.failed-row:not(.hidden)').length;
            const totalVisibleUpdates = document.querySelectorAll('tr[data-ukey]:not(.hidden)').length;
            const sf = document.getElementById('stat-failed-value');
            const st = document.getElementById('stat-total-value');
            if (sf) sf.textContent = totalVisibleFailed;
            if (st) st.textContent = totalVisibleUpdates;

            // Update failure alert chips
            document.querySelectorAll('.alert-chip.fail').forEach(chip => {
                const compId = chip.getAttribute('data-computer');
                const detail = document.getElementById('detail-' + compId);
                if (!detail) return;
                const cnt = detail.querySelectorAll('tr.failed-row:not(.hidden)').length;
                const numSpan = chip.querySelector('.failure-count-number');
                if (cnt > 0) {
                    chip.classList.remove('hidden');
                    if (numSpan) numSpan.textContent = cnt;
                } else {
                    chip.classList.add('hidden');
                }
            });
        }

        function selectAllUpdates() {
            document.querySelectorAll('#update-filters-body input[type=checkbox]').forEach(cb => cb.checked = true);
            document.getElementById('select-all-updates').checked = true;
            updateFiltersFromUI();
        }

        function selectNoneUpdates() {
            document.querySelectorAll('#update-filters-body input[type=checkbox]').forEach(cb => cb.checked = false);
            document.getElementById('select-all-updates').checked = false;
            updateFiltersFromUI();
        }

        function downloadRerunCSV() {
            const lines = ['Name,IP'];
            rerunComputers.forEach(c => lines.push(c.name + ',' + c.ip));
            const blob = new Blob([lines.join('\r\n')], {type: 'text/csv'});
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = 'rerun_computers.csv';
            a.click();
        }

        document.addEventListener('DOMContentLoaded', function() {
            renderUpdateFilters();

            if (rerunComputers.length > 0) {
                const banner = document.getElementById('rerun-banner');
                if (banner) banner.style.display = '';
                const cnt = document.getElementById('rerun-count');
                if (cnt) cnt.textContent = rerunComputers.length;
                const detail = document.getElementById('rerun-detail');
                if (detail) detail.textContent = rerunComputers.map(c => c.name + ' (' + c.reason + ')').join('  \u2022  ');
            }
        });
    </script>
</body>
</html>
"@

# Save HTML report
$htmlPath = "$sessionReportPath\WindowsUpdateReport.html"
$html | Out-File $htmlPath -Encoding UTF8
# ============================================
# FINAL SUMMARY
# ============================================

Write-Host "`n" -NoNewline
Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║                    DEPLOYMENT COMPLETE                       ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host ""
Write-Host "📊 FINAL STATISTICS:" -ForegroundColor Cyan
Write-Host "───────────────────" -ForegroundColor Gray
Write-Host "  Total Computers:        $totalComputers" -ForegroundColor White
Write-Host "  Completed Successfully: $totalCompleted" -ForegroundColor Green
Write-Host "  Updates Installed:      $totalInstalled" -ForegroundColor Green
Write-Host "  Updates Failed:         $totalFailed" -ForegroundColor $(if ($totalFailed -gt 0) { "Red" } else { "Green" })
Write-Host "  Computers Need Reboot:  $computersNeedReboot" -ForegroundColor $(if ($computersNeedReboot -gt 0) { "Yellow" } else { "Green" })
Write-Host ""
Write-Host "📁 REPORTS GENERATED:" -ForegroundColor Cyan
Write-Host "────────────────────" -ForegroundColor Gray
Write-Host "  Session Folder:  $sessionReportPath" -ForegroundColor White
Write-Host "  HTML Report:     WindowsUpdateReport.html" -ForegroundColor White
Write-Host "  Summary CSV:     computer_summary.csv" -ForegroundColor White
Write-Host "  Details CSV:     all_updates.csv" -ForegroundColor White
if ($rerunList) {
    Write-Host "  Re-run CSV:      rerun_computers.csv ($($rerunList.Count) computer$(if($rerunList.Count -ne 1){'s'}) need another pass)" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "✨ Opening HTML report..." -ForegroundColor Cyan
 
# Open the HTML report
Start-Process $htmlPath
 
# Also open the session folder
Start-Process $sessionReportPath
 
Write-Host ""
Write-Host "Session completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow
Write-Host ""
