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

    # Seconds to wait for WSMan TCP reachability before marking a host unreachable
    [Parameter(Mandatory=$false)]
    [int]$WSManTimeoutSeconds = 10
)
 
Add-Type -AssemblyName System.Windows.Forms

# Clean up any leftover local temp update-log files when the script exits (including Ctrl+C)
Register-EngineEvent PowerShell.Exiting -Action {
    Get-ChildItem "$env:TEMP\updatelog_*.csv" -ErrorAction SilentlyContinue |
        Remove-Item -Force -ErrorAction SilentlyContinue
} | Out-Null

# ============================================
# INITIALIZATION
# ============================================
 
# Credentials
 
$cred = Get-Credential -Message "Enter administrator credentials for remote computers"
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

        # Seconds to wait for TCP WinRM port before marking a host unreachable
        [Parameter(Mandatory=$false)]
        [int]$WSManTimeoutSeconds = 10,

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
    # Fast TCP connect to port 5985 with configurable timeout avoids long DNS/network
    # hangs that Test-WSMan alone would cause on truly unreachable hosts.
    Write-Host "Running WSMan connectivity check against $($Computers.Count) computers (timeout: ${WSManTimeoutSeconds}s)..." -ForegroundColor Cyan
    foreach ($computer in $Computers) {
        $target = if ($computer.IP) { $computer.IP } else { $computer.Name }
        $reachable = $false
        try {
            $tcp = New-Object System.Net.Sockets.TcpClient
            $asyncResult = $tcp.BeginConnect($target, 5985, $null, $null)
            $reachable = $asyncResult.AsyncWaitHandle.WaitOne($WSManTimeoutSeconds * 1000, $false) -and $tcp.Connected
            $tcp.Close()
        } catch { $reachable = $false }

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

    # Monitor loop
    while ((Get-Date) -lt $endTime) {
        $stillRunning = 0
        $completed = 0
        $failed = 0
        
        Clear-Host
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
                # Quick TCP check
                $laReachable = $false
                try {
                    $tcp = New-Object System.Net.Sockets.TcpClient
                    $ar = $tcp.BeginConnect($la.IP, 5985, $null, $null)
                    $laReachable = $ar.AsyncWaitHandle.WaitOne($WSManTimeoutSeconds * 1000, $false) -and $tcp.Connected
                    $tcp.Close()
                } catch {}

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
                                $updateProcesses = Get-Process -Name "TrustedInstaller", "TiWorker", "wuauclt" -ErrorAction SilentlyContinue
                            } catch { }

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

            # ── Completion logic ─────────────────────────────────────────────
            # Primary: explicit completion marker AND log — both must be from the current session
            # (newer than update_started.txt) to prevent false positives from previous-run artifacts
            $definitelyDone = $jobStatus.CompletedIsCurrentSession -and $jobStatus.LogIsCurrentSession
            # Fallback heuristic: log is stable, current-session, no tasks or WU jobs running
            $heuristicDone  = $jobStatus.LogIsCurrentSession -and $jobStatus.LogComplete -and
                              -not $jobStatus.TasksRunning -and -not $jobStatus.JobsRunning

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

# Job-start scriptblock — shared by Phase 1 and the late-arrival retry logic.
# Passed via $using:jobStartBlock in the parallel loop, and as -JobStartScript to
# Wait-ForUpdateJobs for late-arrival retries.
$jobStartBlock = {
    param($remoteTempPath)
    # Ensure temp directory exists
    if (-not (Test-Path $remoteTempPath)) {
        New-Item $remoteTempPath -ItemType Directory -Force | Out-Null
    }
    # Clear stale artifacts from any previous run to prevent false signals
    @("updatelog.csv", "update_completed.txt") | ForEach-Object {
        $p = "$remoteTempPath\$_"
        if (Test-Path $p) { Remove-Item $p -Force }
    }
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
    Invoke-WUJob -RunNow -Confirm:$false -Verbose -Script $updateScript
}

# Start updates on all computers in parallel — no throttle (each site is a separate network)
$jobResults = $computers | ForEach-Object -Parallel {
    $site          = $_.'Name'
    $IP            = $_.'IP'
    $cred          = $using:cred
    $rtPath        = $using:RemoteTempPath   # capture for use inside Invoke-Command

    # Pre-flight: quick TCP check on WinRM port before attempting full Invoke-Command.
    # Saves ~20s per offline machine (avoids the full WinRM connection timeout).
    $target = if ($IP) { $IP } else { $site }
    $preFlightOk = $false
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($target, 5985, $null, $null)
        $preFlightOk = $ar.AsyncWaitHandle.WaitOne(($using:WSManTimeoutSeconds) * 1000, $false) -and $tcp.Connected
        $tcp.Close()
    } catch {}

    if (-not $preFlightOk) {
        Write-Host "  [$site] Not reachable on port 5985 — skipping" -ForegroundColor DarkYellow
        return [PSCustomObject]@{ Site = $site; IP = $IP; Status = "Unreachable-PreFlight"; Error = "TCP 5985 not reachable" }
    }

    try {
        Write-Host "[$site] Starting update job..." -ForegroundColor Yellow

        Invoke-Command -ComputerName $IP -Credential $cred -ErrorAction Stop `
            -ArgumentList $rtPath -ScriptBlock ($using:jobStartBlock)

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
} -ThrottleLimit $computers.Count
 
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
                                        -WSManTimeoutSeconds $WSManTimeoutSeconds `
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
 
foreach ($computer in $computers) {
    $computerName = $computer.Name
    $computerIP = $computer.IP
    
    Write-Host "[$computerName] " -NoNewline
    
    # Initialize summary for this computer
    $summary = [PSCustomObject]@{
        ComputerName = $computerName
        IP = $computerIP
        Status = $monitoringResults.Status[$computerName]
        TotalUpdates = 0
        Installed = 0
        Failed = 0
        Skipped = 0
        InstalledWithErrors = 0
        RebootRequired = $false
        CollectionError = ""
    }
    
    # Skip log collection for computers whose job never started (Phase 1 failure)
    if ($summary.Status -eq "JobStartFailed") {
        $phase1Error = ($jobResults | Where-Object { $_.Site -eq $computerName } | Select-Object -First 1).Error
        $summary.CollectionError = if ($phase1Error) { "Job did not start: $phase1Error" } else { "Job did not start" }
        Write-Host "JobStartFailed — $($summary.CollectionError)" -ForegroundColor DarkYellow
        $computerSummary += $summary
        continue
    }

    # Only collect if job completed
    if ($monitoringResults.Completed[$computerName]) {
        try {
            # Use Copy-Item with PSSession to avoid double-hop authentication
            $tempLocalFile = "$env:TEMP\updatelog_${computerName}_${timestamp}.csv"

            # Create PSSession with retry for transient network failures
            $session = $null
            for ($retryAttempt = 0; $retryAttempt -le $MaxRetries; $retryAttempt++) {
                try {
                    $session = New-PSSession -ComputerName $computerIP -Credential $cred -ErrorAction Stop
                    break
                } catch {
                    if ($retryAttempt -lt $MaxRetries) {
                        Start-Sleep -Seconds (5 * ($retryAttempt + 1))
                    } else {
                        throw
                    }
                }
            }

            try {
                # Check if the update log exists on the remote machine
                $remoteLogPath = "$RemoteTempPath\updatelog.csv"
                $fileInfo = Invoke-Command -Session $session -ArgumentList $remoteLogPath -ScriptBlock {
                    param($logPath)
                    if (Test-Path $logPath) {
                        $file = Get-Item $logPath
                        @{ Exists = $true; Size = $file.Length; LastWriteTime = $file.LastWriteTime }
                    } else {
                        @{ Exists = $false }
                    }
                }

                if ($fileInfo.Exists) {
                    # Copy the file to local machine
                    Copy-Item -FromSession $session -Path $remoteLogPath -Destination $tempLocalFile -Force -ErrorAction Stop
                    
                    # Now read the local copy
                    $updates = Import-Csv $tempLocalFile -ErrorAction Stop
                    
                    # Check reboot status
                    $rebootStatus = $true
                    
                    # Process the updates with comprehensive parsing
                    if ($updates.Count -gt 0) {
                        foreach ($update in $updates) {
                            # Get values as strings for comparison
                            $resultValue = if ($update.Result) { $update.Result.ToString().Trim() } else { "" }
                            $statusValue = if ($update.Status) { $update.Status.ToString().Trim() } else { "" }
                            
                            # Determine actual status
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
                            # Check Result field for text values
                            elseif ($resultValue -eq "Failed" -or $resultValue -match "^Fail") {
                                $summary.Failed++
                                $actualStatus = "Failed"
                                $isFailure = $true
                            }
                            elseif ($resultValue -eq "Installed" -or $resultValue -match "^Install|^Success") {
                                $summary.Installed++
                                $actualStatus = "Installed"
                                $isSuccess = $true
                            }
                            elseif ($resultValue -match "Abort") {
                                $summary.Failed++
                                $actualStatus = "Aborted"
                                $isFailure = $true
                            }
                            elseif ($resultValue -match "Error") {
                                $summary.InstalledWithErrors++
                                $actualStatus = "InstalledWithErrors"
                            }
                            # Check numeric codes
                            elseif ($resultValue -eq "2" -or $resultValue -eq 2) {
                                $summary.Installed++
                                $actualStatus = "Installed"
                                $isSuccess = $true
                            }
                            elseif ($resultValue -eq "3" -or $resultValue -eq 3) {
                                $summary.InstalledWithErrors++
                                $actualStatus = "InstalledWithErrors"
                            }
                            elseif ($resultValue -eq "4" -or $resultValue -eq 4) {
                                $summary.Failed++
                                $actualStatus = "Failed"
                                $isFailure = $true
                            }
                            elseif ($resultValue -eq "5" -or $resultValue -eq 5) {
                                $summary.Failed++
                                $actualStatus = "Aborted"
                                $isFailure = $true
                            }
                            # If still unknown, check Status field
                            elseif ($statusValue -match "Fail") {
                                $summary.Failed++
                                $actualStatus = "Failed"
                                $isFailure = $true
                            }
                            elseif ($statusValue -match "Install|Success") {
                                $summary.Installed++
                                $actualStatus = "Installed"
                                $isSuccess = $true
                            }
                            
                            # Special case: Status is "Unknown" but Result is "Failed"
                            if ($statusValue -eq "Unknown" -and $resultValue -eq "Failed" -and !$isFailure) {
                                $summary.Failed++
                                $actualStatus = "Failed"
                                $isFailure = $true
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
                            }
                        }
 
                        # Total updates count includes all updates (including definition/Intel updates)
                        $summary.TotalUpdates = $updates.Count
                        $summary.RebootRequired = $rebootStatus
                        
                        Write-Host "Collected $($summary.TotalUpdates) updates " -NoNewline
                        if ($summary.Installed -gt 0) {
                            Write-Host "(✅ $($summary.Installed) installed" -NoNewline -ForegroundColor Green
                        }
                        if ($summary.Failed -gt 0) {
                            if ($summary.Installed -gt 0) {
                                Write-Host ", " -NoNewline
                            } else {
                                Write-Host "(" -NoNewline
                            }
                            Write-Host "❌ $($summary.Failed) failed" -NoNewline -ForegroundColor Red
                        }
                        Write-Host ")" -ForegroundColor Green
                        
                    } else {
                        Write-Host "CSV is empty - no updates were needed" -ForegroundColor Green
                        $summary.Status = "NoUpdatesNeeded"
                    }
 
                    # (No updates are filtered out here; filtering is handled client-side in the HTML report)
                    
                    # Clean up local temp file
                    Remove-Item $tempLocalFile -Force -ErrorAction SilentlyContinue

                } else {
                    Write-Host "Update log not found on remote" -ForegroundColor Yellow
                    $summary.Status = "NoLogFile"
                }

            } catch {
                Write-Host "Failed to collect: $($_.Exception.Message)" -ForegroundColor Red
                $summary.Status = "CollectionFailed"
                $summary.CollectionError = $_.Exception.Message
            } finally {
                # Always close the session, even when an error occurs
                if ($session) { Remove-PSSession $session -ErrorAction SilentlyContinue }
            }

        } catch {
            # Catches PSSession creation failures after all retries exhausted
            Write-Host "Could not connect to collect results: $($_.Exception.Message)" -ForegroundColor Red
            $summary.Status = "CollectionFailed"
            $summary.CollectionError = $_.Exception.Message
        }
    } else {
        Write-Host "Skipped (job did not complete)" -ForegroundColor Yellow
        $summary.Status = "Incomplete"
    }
    
    # Update the Status to "Completed" if we successfully collected data
    if ($summary.TotalUpdates -gt 0 -and $summary.Status -ne "CollectionFailed") {
        $summary.Status = "Completed"
    }
    
    $computerSummary += $summary
}
 
# Save collected data
$computerSummary | Export-Csv "$sessionReportPath\computer_summary.csv" -NoTypeInformation
if ($allUpdateData.Count -gt 0) {
    $allUpdateData | Export-Csv "$sessionReportPath\all_updates.csv" -NoTypeInformation
}

# Re-run CSV: computers that need another pass (real failures or did not complete)
$rerunList = $computerSummary | Where-Object {
    $_.Failed -gt 0 -or
    $_.Status -in @('Incomplete', 'CollectionFailed', 'Unreachable', 'Ignored-WSMAN', 'JobStartFailed')
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
 
# ============================================
# PHASE 4: GENERATE HTML REPORT (FIXED VERSION)
# ============================================
 
Write-Host "`n╔══ Phase 4: Generating HTML Report ══╗" -ForegroundColor Cyan
 
# Calculate overall statistics
$totalComputers = $computerSummary.Count
$totalCompleted = ($computerSummary | Where-Object { $_.Status -eq "Completed" }).Count
$totalUpdates = ($computerSummary | Measure-Object -Property TotalUpdates -Sum).Sum
$totalInstalled = ($computerSummary | Measure-Object -Property Installed -Sum).Sum
$totalFailed = ($computerSummary | Measure-Object -Property Failed -Sum).Sum
$computersNeedReboot = ($computerSummary | Where-Object { $_.RebootRequired }).Count
 
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
        $reason = if ($_.Failed -gt 0) {
            "$($_.Failed) failed update$(if($_.Failed -ne 1){'s'})"
        } else { $_.Status }
        $safeName = $_.ComputerName -replace '"', '\"' -replace '\\', '\\\\'
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
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        @keyframes gradientShift {
            0%, 100% { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
            50% { background: linear-gradient(135deg, #764ba2 0%, #667eea 100%); }
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
        }
        
        .header {
            background: white;
            border-radius: 20px;
            padding: 30px;
            margin-bottom: 30px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            animation: slideDown 0.5s ease-out;
        }
        
        @keyframes slideDown {
            from {
                opacity: 0;
                transform: translateY(-30px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        
        h1 {
            color: #1a202c;
            font-size: 2.5em;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 15px;
        }
        
        .timestamp {
            color: #718096;
            font-size: 0.9em;
            margin-bottom: 5px;
        }
        
        .session-id {
            color: #a0aec0;
            font-size: 0.85em;
            font-family: 'Courier New', monospace;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 20px;
            margin-top: 30px;
        }
        
        .stat-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
            transition: all 0.3s ease;
            animation: fadeIn 0.5s ease-out;
        }
        
        @keyframes fadeIn {
            from {
                opacity: 0;
                transform: scale(0.9);
            }
            to {
                opacity: 1;
                transform: scale(1);
            }
        }
        
        .stat-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 15px 35px rgba(0,0,0,0.15);
        }
        
        .stat-card.success {
            background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        }
        
        .stat-card.error {
            background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
        }
        
        .stat-card.warning {
            background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        }
        
        .stat-card.info {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        }
        
        .stat-value {
            font-size: 3em;
            font-weight: bold;
            margin-bottom: 5px;
            animation: countUp 1s ease-out;
        }
        
        @keyframes countUp {
            from {
                opacity: 0;
                transform: translateY(20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        
        .stat-label {
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
            opacity: 0.95;
        }
        
        /* Failure alert section */
        .failure-alert {
            background: white;
            border-radius: 15px;
            margin-bottom: 25px;
            padding: 0;
            box-shadow: 0 10px 25px rgba(239, 68, 68, 0.15);
            overflow: hidden;
            animation: slideIn 0.6s ease-out;
            border: 2px solid #fee2e2;
        }
        
        .connection-alert {
            background: white;
            border-radius: 15px;
            margin-bottom: 25px;
            padding: 0;
            box-shadow: 0 10px 25px rgba(236, 72, 153, 0.15);
            overflow: hidden;
            animation: slideIn 0.6s ease-out;
            border: 2px solid #fce7f3;
        }
        
        @keyframes slideIn {
            from {
                opacity: 0;
                transform: translateX(-30px);
            }
            to {
                opacity: 1;
                transform: translateX(0);
            }
        }
        
        .failure-alert-header {
            background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
            color: white;
            padding: 20px 25px;
            display: flex;
            align-items: center;
            gap: 15px;
        }
        
        .connection-alert-header {
            background: linear-gradient(135deg, #ec4899 0%, #be185d 100%);
            color: white;
            padding: 20px 25px;
            display: flex;
            align-items: center;
            gap: 15px;
        }
        
        .failure-alert-icon {
            font-size: 1.5em;
            animation: pulse 2s ease-in-out infinite;
        }
        
        @keyframes pulse {
            0%, 100% { transform: scale(1); }
            50% { transform: scale(1.1); }
        }
        
        .failure-alert-title {
            font-size: 1.3em;
            font-weight: 600;
        }
        
        .failure-alert-subtitle {
            font-size: 0.9em;
            opacity: 0.9;
            margin-left: auto;
        }
        
        .failure-list {
            padding: 20px 25px;
            background: #fef2f2;
        }
        
        .failure-item {
            background: white;
            border-left: 4px solid #ef4444;
            padding: 12px 20px;
            margin-bottom: 12px;
            border-radius: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            transition: all 0.3s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            cursor: pointer;
        }
        
        .failure-item:hover {
            transform: translateX(5px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }
        
        .failure-item:last-child {
            margin-bottom: 0;
        }
        
        .failure-computer-name {
            font-weight: 600;
            color: #1f2937;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .failure-computer-ip {
            color: #6b7280;
            font-size: 0.9em;
        }
        
        .failure-count {
            background: #ef4444;
            color: white;
            padding: 6px 14px;
            border-radius: 20px;
            font-weight: bold;
            font-size: 0.9em;
        }
        
        .failure-details {
            display: flex;
            align-items: center;
            gap: 15px;
        }
        
        .controls {
            background: white;
            border-radius: 15px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
            display: flex;
            gap: 15px;
            align-items: center;
            flex-wrap: wrap;
        }
        
        .btn {
            padding: 10px 20px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s ease;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }
        
        .btn-primary {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
            color: white;
        }
        
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(59, 130, 246, 0.3);
        }
        
        .btn-secondary {
            background: linear-gradient(135deg, #6b7280 0%, #4b5563 100%);
            color: white;
        }
        
        .btn-secondary:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(107, 114, 128, 0.3);
        }
        
        .btn-danger {
            background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);
            color: white;
        }
        
        .btn-danger:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(239, 68, 68, 0.3);
        }
        
        .search-box {
            flex: 1;
            min-width: 250px;
            padding: 10px 15px;
            border: 2px solid #e5e7eb;
            border-radius: 8px;
            font-size: 14px;
            transition: all 0.3s ease;
        }
        
        .search-box:focus {
            outline: none;
            border-color: #3b82f6;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1);
        }
        
        .computer-card {
            background: white;
            border-radius: 15px;
            margin-bottom: 20px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
            overflow: hidden;
            animation: slideUp 0.5s ease-out;
            transition: all 0.3s ease;
        }
        
        .computer-card.has-failures {
            border: 2px solid #fee2e2;
            box-shadow: 0 10px 25px rgba(239, 68, 68, 0.1);
        }
        
        .computer-card.unreachable {
            border: 2px solid #fce7f3;
            box-shadow: 0 10px 25px rgba(236, 72, 153, 0.1);
        }
        
        @keyframes slideUp {
            from {
                opacity: 0;
                transform: translateY(30px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        
        .computer-card:hover {
            box-shadow: 0 15px 35px rgba(0,0,0,0.15);
        }
        
        .computer-header {
            background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
            color: white;
            padding: 20px 25px;
            cursor: pointer;
            display: flex;
            justify-content: space-between;
            align-items: center;
            transition: all 0.3s ease;
        }
        
        .computer-header:hover {
            background: linear-gradient(135deg, #334155 0%, #475569 100%);
        }
        
        .computer-header.expanded {
            background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        }
        
        .computer-header.has-failures {
            background: linear-gradient(135deg, #991b1b 0%, #dc2626 100%);
        }
        
        .computer-header.has-failures:hover {
            background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%);
        }
        
        .computer-header.unreachable {
            background: linear-gradient(135deg, #be185d 0%, #ec4899 100%);
        }
        
        .computer-header.unreachable:hover {
            background: linear-gradient(135deg, #ec4899 0%, #f472b6 100%);
        }
        
        .computer-info {
            display: flex;
            align-items: center;
            gap: 25px;
            flex: 1;
        }
        
        .computer-name {
            font-size: 1.3em;
            font-weight: bold;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .computer-ip {
            color: #94a3b8;
            font-size: 0.9em;
        }
        
        .update-summary {
            display: flex;
            gap: 12px;
            align-items: center;
        }
        
        .summary-badge {
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 600;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }
        
        .badge-success {
            background: rgba(16, 185, 129, 0.2);
            color: #10b981;
        }
        
        .badge-error {
            background: rgba(239, 68, 68, 0.2);
            color: #ef4444;
        }
        
        .badge-warning {
            background: rgba(245, 158, 11, 0.2);
            color: #f59e0b;
        }
        
        .badge-info {
            background: rgba(59, 130, 246, 0.2);
            color: #3b82f6;
        }
        
        .expand-icon {
            font-size: 1.2em;
            transition: transform 0.3s ease;
        }
        
        .computer-header.expanded .expand-icon {
            transform: rotate(180deg);
        }
        
        .computer-content {
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.45s ease, padding 0.25s ease;
            background: #fafafa;
            padding: 0 20px; /* keep header/content spacing when collapsed */
        }
 
        /* When expanded, constrain height responsively and allow internal scrolling */
        .computer-content.expanded {
            /* Responsive max-height: at least 240px, ideally viewport minus header area, at most 720px */
            max-height: clamp(240px, calc(100vh - 360px), 720px);
            overflow-y: auto;
            padding: 20px; /* add breathing room when open */
        }
 
        /* Subtle, cross-browser scrollbar styling for modern browsers */
        .computer-content.expanded::-webkit-scrollbar {
            width: 12px;
        }
        .computer-content.expanded::-webkit-scrollbar-track {
            background: #f1f5f9;
            border-radius: 8px;
        }
        .computer-content.expanded::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 8px;
        }
        
        .update-table {
            width: 100%;
            border-collapse: collapse;
        }
        
        .update-table th {
            background: #f1f5f9;
            padding: 12px 15px;
            text-align: left;
            font-weight: 600;
            color: #475569;
            font-size: 0.85em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            border-bottom: 2px solid #e2e8f0;
        }
        
        .update-table td {
            padding: 12px 15px;
            border-bottom: 1px solid #e2e8f0;
            font-size: 0.9em;
            background: white;
        }
        
        .update-table tr:hover td {
            background: #f8fafc;
        }
        
        .update-table tr.failed-row {
            background-color: #fef2f2 !important;
            border-left: 3px solid #ef4444;
        }
        
        .update-table tr.failed-row:hover td {
            background-color: #fee2e2 !important;
        }
        
        .update-table tr.failed-row td:first-child {
            padding-left: 12px;
        }
        
        .kb-badge {
            background: #e2e8f0;
            padding: 4px 10px;
            border-radius: 6px;
            font-family: 'Courier New', monospace;
            font-size: 0.85em;
            font-weight: bold;
            color: #475569;
        }
        
        .status-icon {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            font-weight: 600;
        }
        
        .status-installed {
            color: #10b981;
        }
        
        .status-failed {
            color: #ef4444;
            font-weight: 700;
        }
        
        .status-warning {
            color: #f59e0b;
        }
        
        .reboot-warning {
            background: linear-gradient(135deg, #fef3c7 0%, #fed7aa 100%);
            border-left: 4px solid #f59e0b;
            padding: 12px 20px;
            margin: 20px;
            display: flex;
            align-items: center;
            gap: 12px;
            border-radius: 8px;
            font-weight: 500;
            color: #92400e;
        }
        
        .footer {
            text-align: center;
            margin-top: 50px;
            padding: 30px;
            color: white;
            opacity: 0.9;
        }
        
        .hidden {
            display: none !important;
        }
        
        @media (max-width: 768px) {
            .stats-grid {
                grid-template-columns: 1fr 1fr;
            }
            
            .computer-info {
                flex-direction: column;
                align-items: flex-start;
                gap: 10px;
            }
            
            .update-summary {
                flex-wrap: wrap;
            }
        }

        .rerun-alert {
            background: white;
            border-radius: 15px;
            margin-bottom: 25px;
            padding: 0;
            box-shadow: 0 10px 25px rgba(245, 158, 11, 0.15);
            overflow: hidden;
            animation: slideIn 0.6s ease-out;
            border: 2px solid #fde68a;
        }

        .rerun-alert-header {
            background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
            color: white;
            padding: 20px 25px;
            display: flex;
            align-items: center;
            gap: 15px;
        }

        .rerun-alert-subtitle {
            font-size: 0.9em;
            opacity: 0.9;
            margin-left: auto;
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .rerun-detail {
            padding: 14px 25px;
            background: #fffbeb;
            color: #92400e;
            font-size: 0.88em;
            line-height: 1.6;
        }

        .btn-rerun {
            background: white;
            color: #d97706;
            border: 2px solid white;
            padding: 6px 16px;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 700;
            font-size: 13px;
            white-space: nowrap;
        }

        .btn-rerun:hover {
            background: #fef3c7;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>
                <span style="font-size: 1.2em;">🔄</span>
                Windows Update Deployment Report
            </h1>
            <div class="timestamp">Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</div>
            <div class="session-id">Session ID: $timestamp</div>
            
            <div class="stats-grid">
                <div class="stat-card info">
                    <div class="stat-value">$totalComputers</div>
                    <div class="stat-label">Computers</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">$totalCompleted</div>
                    <div class="stat-label">Completed</div>
                </div>
                <div class="stat-card success">
                    <div class="stat-value">$totalInstalled</div>
                    <div class="stat-label">Installed</div>
                </div>
                <div class="stat-card error">
                    <div class="stat-value" id="stat-failed-value">$totalFailed</div>
                    <div class="stat-label">Failed</div>
                </div>
                <div class="stat-card warning">
                    <div class="stat-value">$computersNeedReboot</div>
                    <div class="stat-label">Need Reboot</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value" id="stat-total-value">$totalUpdates</div>
                    <div class="stat-label">Total Updates</div>
                </div>
            </div>
        </div>
"@
 
# Add failure alert section if there are any computers with failures
if ($computersWithFailures.Count -gt 0) {
    $html += @"
        
        <div class="failure-alert">
            <div class="failure-alert-header">
                <span class="failure-alert-icon">⚠️</span>
                <div>
                    <div class="failure-alert-title">Computers with Failed Updates</div>
                </div>
                <div class="failure-alert-subtitle">$($computersWithFailures.Count) computer$(if($computersWithFailures.Count -ne 1){'s'}) require$(if($computersWithFailures.Count -eq 1){'s'}) attention</div>
            </div>
            <div class="failure-list">
"@
    
    foreach ($failedComputer in $computersWithFailures) {
        $cleanName = $failedComputer.ComputerName -replace '[^a-zA-Z0-9]', '_'
        $safeComputerName = ConvertTo-HtmlEncoded $failedComputer.ComputerName
        $html += @"
                <div class="failure-item" data-computer="$cleanName" onclick="scrollToComputer('$cleanName')">
                    <div>
                        <div class="failure-computer-name">
                            <span>💻</span>
                            <span>$safeComputerName</span>
                        </div>
                        <div class="failure-computer-ip">$(ConvertTo-HtmlEncoded $failedComputer.IP)</div>
                    </div>
                    <div class="failure-details">
"@
        
        if ($failedComputer.Installed -gt 0) {
            $html += @"
                        <span class="summary-badge badge-success">✅ $($failedComputer.Installed) installed</span>
"@
        }
        
        $html += @"
                        <span class="failure-count">❌ <span class="failure-count-number">$($failedComputer.Failed)</span> failed</span>
                    </div>
                </div>
"@
    }
    
    $html += @"
            </div>
        </div>
"@
}

# Get computers with connection failures
$computersWithConnectionFailures = $computerSummary | Where-Object { $_.Status -eq "Ignored-WSMAN" -or $_.Status -eq "Incomplete" -or $_.Status -eq "CollectionFailed" } | Sort-Object ComputerName

# Add connection failure alert section if there are any computers with connection issues
if ($computersWithConnectionFailures.Count -gt 0) {
    $html += @"
        
        <div class="connection-alert">
            <div class="connection-alert-header">
                <span class="failure-alert-icon">📡</span>
                <div>
                    <div class="failure-alert-title">Computers Unable to Connect</div>
                </div>
                <div class="failure-alert-subtitle"><span id="connection-failure-count">$($computersWithConnectionFailures.Count)</span> computer$(if($computersWithConnectionFailures.Count -ne 1){'s'}) could not be reached</div>
            </div>
            <div class="failure-list">
"@
    
    foreach ($failedComputer in $computersWithConnectionFailures) {
        $cleanName = $failedComputer.ComputerName -replace '[^a-zA-Z0-9]', '_'
        $safeComputerName = ConvertTo-HtmlEncoded $failedComputer.ComputerName
        $statusReason = switch ($failedComputer.Status) {
            "Ignored-WSMAN" { "WSMan connectivity failed" }
            "Incomplete"    { "Job did not complete" }
            "CollectionFailed" { "Failed to collect results" }
            default         { "Unknown error" }
        }
        $html += @"
                <div class="failure-item" data-computer="$cleanName" onclick="scrollToComputer('$cleanName')">
                    <div>
                        <div class="failure-computer-name">
                            <span>📡</span>
                            <span>$safeComputerName</span>
                        </div>
                        <div class="failure-computer-ip">$(ConvertTo-HtmlEncoded $failedComputer.IP) — $statusReason</div>
                    </div>
                </div>
"@
    }
    
    $html += @"
            </div>
        </div>
"@
}

# Add re-run banner (hidden by JS if rerunComputers is empty)
$html += @"

        <div id="rerun-banner" class="rerun-alert" style="display:none">
            <div class="rerun-alert-header">
                <span class="failure-alert-icon">&#128260;</span>
                <div>
                    <div class="failure-alert-title">Computers Queued for Re-run</div>
                </div>
                <div class="rerun-alert-subtitle">
                    <span id="rerun-count"></span> computer(s) need another pass
                    <button class="btn-rerun" onclick="downloadRerunCSV()">&#11015; Download Re-run CSV</button>
                </div>
            </div>
            <div class="rerun-detail" id="rerun-detail"></div>
        </div>
"@

# Add controls section
$html += @"
        
        <div class="controls">
            <button class="btn btn-primary" onclick="expandAll()">
                <span>📂</span> Expand All
            </button>
            <button class="btn btn-primary" onclick="sortComputersByStatus()">
                <span>🔃</span> Sort by Status
            </button>
            <button class="btn btn-secondary" onclick="collapseAll()">
                <span>📁</span> Collapse All
            </button>
"@
 
# Add a button to jump to failures if they exist
if ($computersWithFailures.Count -gt 0) {
    $html += @"
            <button class="btn btn-danger" onclick="expandFailures()">
                <span>⚠️</span> Show Failures Only
            </button>
"@
}
 
$html += @"
            <input type="text" class="search-box" placeholder="🔍 Search computers or updates..." onkeyup="filterComputers(this.value)">
        </div>
        <div class="controls" style="margin-bottom: 10px;">
            <strong style="margin-right:10px;">Filter updates:</strong>
            <div style="display:flex;align-items:center;gap:8px;">
                <button type="button" onclick="document.getElementById('update-filters-body').classList.toggle('hidden')" style="padding:6px 10px;border-radius:8px;border:1px solid #e5e7eb;background:#fff;">Toggle filters</button>
                <div style="font-size:0.95em;color:#6b7280">Unique updates: <span id="unique-updates-count">$($uniqueUpdates.Keys.Count)</span> | Computers shown: <span id="filtered-computer-count">$($computerSummary.Count)</span></div>
                <div style="flex:1"></div>
                <label style="display:inline-flex;align-items:center;gap:6px;background:#fff;padding:6px 10px;border-radius:8px;border:1px solid #e5e7eb;">
                    <input type="checkbox" id="select-all-updates" onclick="(this.checked?selectAllUpdates():selectNoneUpdates())" checked /> Select all
                </label>
            </div>
            <div id="update-filters-body" style="margin-top:8px;">
                <div id="update-filters-failed" style="margin-bottom:8px;">
                    <strong style="display:block;margin-bottom:6px;color:#b91c1c">Failed updates</strong>
                    <div id="update-filters-failed-list" style="display:flex; gap:8px; flex-wrap:wrap;"></div>
                </div>
                <div id="update-filters-normal">
                    <strong style="display:block;margin-bottom:6px;color:#374151">Other updates</strong>
                    <div id="update-filters-normal-list" style="display:flex; gap:8px; flex-wrap:wrap;"></div>
                </div>
            </div>
        </div>
        
        <div id="computers-container">
"@
 
# Add each computer's data
foreach ($summary in $computerSummary | Sort-Object ComputerName) {
    $computerName = $summary.ComputerName
    $group = $computerGroups | Where-Object { $_.Name -eq $computerName }
    $computerData = if ($group) { $group.Group } else { @() }
    
    $cleanName = $computerName -replace '[^a-zA-Z0-9]', '_'
    $safeComputerName = ConvertTo-HtmlEncoded $computerName
    $hasFailures = $summary.Failed -gt 0
    $isUnreachable = $summary.Status -eq "Ignored-WSMAN" -or $summary.Status -eq "Incomplete" -or $summary.Status -eq "CollectionFailed"
    $statusBadge = switch ($summary.Status) {
        "Completed" { "✅" }
        "NoUpdatesNeeded" { "✔" }
        "Failed" { "❌" }
        "Incomplete" { "⏱️" }
        "Ignored-WSMAN" { "🚫" }
        "CollectionFailed" { "⚠️" }
        default { "❓" }
    }
    
    $html += @"
            <div class="computer-card$(if($hasFailures){' has-failures'})$(if($isUnreachable){' unreachable'})" data-computer="$computerName" data-status="$($summary.Status)" id="computer-$cleanName">
                <div class="computer-header$(if($hasFailures){' has-failures'})$(if($isUnreachable){' unreachable'})" onclick="toggleComputer('$cleanName')">
                    <div class="computer-info">
                        <div>
                            <div class="computer-name">
                                <span>💻</span>
                                <span>$safeComputerName</span>
                            </div>
                            <div class="computer-ip">$(ConvertTo-HtmlEncoded $summary.IP)</div>
                        </div>
                        <div class="update-summary">
                            <span class="summary-badge badge-info">
                                $statusBadge $($summary.Status)
                            </span>
"@
    
    if ($summary.TotalUpdates -gt 0) {
        $html += @"
                            <span class="summary-badge badge-info">
                                📦 $($summary.TotalUpdates) updates
                            </span>
"@
    }
    
    if ($summary.Installed -gt 0) {
        $html += @"
                            <span class="summary-badge badge-success">
                                ✅ $($summary.Installed) installed
                            </span>
"@
    }
    
    if ($summary.Failed -gt 0) {
        $html += @"
                            <span class="summary-badge badge-error">
                                ❌ $($summary.Failed) failed
                            </span>
"@
    }
    
    if ($summary.RebootRequired) {
        $html += @"
                            <span class="summary-badge badge-warning">
                                ⚠️ Reboot Required
                            </span>
"@
    }
    
    $html += @"
                        </div>
                    </div>
                    <div class="expand-icon">▼</div>
                </div>
                <div class="computer-content" id="content-$cleanName">
"@
    
    if ($summary.RebootRequired) {
        $html += @"
                    <div class="reboot-warning">
                        <span style="font-size: 1.2em;">⚠️</span>
                        <span>This computer requires a reboot to complete update installation.</span>
                    </div>
"@
    }
    
    if ($computerData.Count -gt 0) {
        $html += @"
                    <table class="update-table">
                        <thead>
                            <tr>
                                <th width="12%">KB Number</th>
                                <th width="50%">Update Title</th>
                                <th width="10%">Size</th>
                                <th width="15%">Status</th>
                                <th width="13%">Result Code</th>
                            </tr>
                        </thead>
                        <tbody>
"@
        
        foreach ($update in $computerData | Sort-Object KB) {
            # Enhanced status display that properly handles text results
            $resultText = if ($update.Result) { $update.Result.ToString() } else { "" }
            $statusText = if ($update.Status) { $update.Status.ToString() } else { "" }
            
            # Determine display and row styling
            $isFailedRow = $false
            $statusDisplay = ""
            
            # Check for Skipped (Defender definition updates) first
            if ($statusText -eq "Skipped") {
                $statusDisplay = '<span class="status-icon" style="color:#94a3b8;">&#9197; Skipped</span>'
            }
            # Check for failures
            elseif ($resultText -eq "Failed" -or $statusText -eq "Failed") {
                $statusDisplay = '<span class="status-icon status-failed">❌ Failed</span>'
                $isFailedRow = $true
            }
            elseif ($resultText -eq "Installed" -or $statusText -eq "Installed") {
                $statusDisplay = '<span class="status-icon status-installed">✅ Installed</span>'
            }
            elseif ($resultText -match "Abort" -or $statusText -match "Abort") {
                $statusDisplay = '<span class="status-icon status-failed">❌ Aborted</span>'
                $isFailedRow = $true
            }
            elseif ($statusText -match "Error") {
                $statusDisplay = '<span class="status-icon status-warning">⚠️ With Errors</span>'
            }
            else {
                # Default case for unknown
                $statusDisplay = '<span class="status-icon">❓ Unknown</span>'
                # If Result shows Failed but Status is Unknown, still mark as failed
                if ($resultText -eq "Failed") {
                    $isFailedRow = $true
                }
            }
            
            $kbDisplay = if ($update.KB -and $update.KB -ne "N/A") { 
                "<span class='kb-badge'>$($update.KB)</span>" 
            } else { 
                "<span style='color: #94a3b8;'>-</span>" 
            }
            
            # Add row with failed styling if needed
            $rowClass = if ($isFailedRow) { ' class="failed-row"' } else { '' }
            
            # assign a stable key per update row (prefer KB, fallback to sanitized Title)
            if ($update.KB -and $update.KB -ne '') {
                $uKey = $update.KB
            } else {
                $uKey = $update.Title -replace '\s+', '_' -replace '[^A-Za-z0-9_\-]', ''
            }
            $safeTitle = ConvertTo-HtmlEncoded $update.Title
            $html += @"
                            <tr$rowClass data-ukey='$uKey'>
                                <td>$kbDisplay</td>
                                <td>$safeTitle</td>
                                <td>$($update.Size)</td>
                                <td>$statusDisplay</td>
                                <td style="text-align: center; color: $(if($isFailedRow){'#ef4444'}else{'#6b7280'});">$resultText</td>
                            </tr>
"@
        }
        
        $html += @"
                        </tbody>
                    </table>
"@
    } elseif ($summary.Status -eq "NoUpdatesNeeded") {
        $html += @"
                    <div style="padding: 40px; text-align: center; color: #10b981;">
                        <div style="font-size: 3em;">✅</div>
                        <div style="font-size: 1.2em; margin-top: 10px;">System is up to date</div>
                        <div style="color: #6b7280; margin-top: 5px;">No updates were needed on this computer</div>
                    </div>
"@
    } elseif ($summary.CollectionError) {
        $safeError = ConvertTo-HtmlEncoded $summary.CollectionError
        $html += @"
                    <div style="padding: 20px; color: #ef4444;">
                        <strong>Error collecting results:</strong> $safeError
                    </div>
"@
    } else {
        $html += @"
                    <div style="padding: 20px; color: #6b7280; text-align: center;">
                        <em>No update data available</em>
                    </div>
"@
    }
    
    $html += @"
                </div>
            </div>
"@
}
 
$html += @"
        </div>
        
        <div class="footer">
            <p style="font-size: 1.1em; margin-bottom: 10px;">Windows Update Deployment Complete</p>
            <p style="opacity: 0.8;">Report Path: $sessionReportPath</p>
            <p style="opacity: 0.6; font-size: 0.9em;">Generated by Windows Update Script v2.0</p>
        </div>
    </div>
    
    <script>
        // Unique updates data (KB => Title)
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
        // Render update filter checkboxes
        function renderUpdateFilters() {
            const container = document.getElementById('update-filters-body');
            if (!container) return;
            Object.keys(uniqueUpdates).forEach(key => {
                const label = document.createElement('label');
                label.style.display = 'inline-flex';
                label.style.alignItems = 'center';
                label.style.gap = '6px';
                label.style.background = '#fff';
                label.style.padding = '6px 10px';
                label.style.borderRadius = '8px';
                label.style.border = '1px solid #e5e7eb';
 
                const cb = document.createElement('input');
                cb.type = 'checkbox';
                // Auto-uncheck Defender Security Intelligence updates by default
                const title = uniqueUpdates[key] || '';
                const isDefender = /Security Intelligence Update for Microsoft Defender Antivirus/i.test(title);
                cb.checked = !isDefender;
                cb.dataset.ukey = key;
                cb.addEventListener('change', updateFiltersFromUI);
 
                const span = document.createElement('span');
                span.textContent = uniqueUpdates[key].length > 60 ? uniqueUpdates[key].substring(0,57)+'...' : uniqueUpdates[key];
                span.title = uniqueUpdates[key];
 
                label.appendChild(cb);
                label.appendChild(span);
                // Decide whether this update has any failed rows in the current report
                const failedList = document.getElementById('update-filters-failed-list');
                const normalList = document.getElementById('update-filters-normal-list');
                const hasFailedInReport = document.querySelectorAll('tr.failed-row[data-ukey="' + key + '"]').length > 0;
                if (hasFailedInReport && failedList) {
                    failedList.appendChild(label);
                } else if (normalList) {
                    normalList.appendChild(label);
                } else {
                    container.appendChild(label);
                }
            });
 
            // Update the select-all control to reflect whether all filters are checked
            const allCheckboxes = Array.from(document.querySelectorAll('#update-filters-body input[type=checkbox]'));
            const allChecked = allCheckboxes.length === 0 ? false : allCheckboxes.every(cb => cb.checked);
            const selectAll = document.getElementById('select-all-updates');
            if (selectAll) selectAll.checked = allChecked;
 
            // Apply initial visibility based on the default checkbox state
            updateFiltersFromUI();
        }
 
        function updateFiltersFromUI() {
            // Build set of enabled keys
            const enabled = new Set();
            document.querySelectorAll('#update-filters-body input[type=checkbox]').forEach(cb => {
                if (cb.checked) enabled.add(cb.dataset.ukey);
            });

            // Show/hide rows
            document.querySelectorAll('tr[data-ukey]').forEach(row => {
                const key = row.getAttribute('data-ukey');
                if (enabled.has(key)) {
                    row.classList.remove('hidden');
                } else {
                    row.classList.add('hidden');
                }
            });

            // Recompute per-computer failure status: if a computer has no visible failed rows, remove the has-failures class
            // Also update the "X failed" badge in the card header to reflect only visible failures.
            document.querySelectorAll('.computer-card').forEach(card => {
                const header = card.querySelector('.computer-header');
                const visibleFailedCount = card.querySelectorAll('tr.failed-row:not(.hidden)').length;
                if (visibleFailedCount > 0) {
                    card.classList.add('has-failures');
                    if (header) header.classList.add('has-failures');
                } else {
                    card.classList.remove('has-failures');
                    if (header) header.classList.remove('has-failures');
                }

                // Update (or hide) the "X failed" badge in the header
                const failedBadge = card.querySelector('.update-summary .badge-error');
                if (failedBadge) {
                    if (visibleFailedCount > 0) {
                        failedBadge.classList.remove('hidden');
                        failedBadge.innerHTML = '&#10060; ' + visibleFailedCount + ' failed';
                    } else {
                        failedBadge.classList.add('hidden');
                    }
                }
            });

            // Hide computers that don't have any visible (non-hidden) rows
            let visibleComputerCount = 0;
            document.querySelectorAll('.computer-card').forEach(card => {
                const visibleRows = card.querySelectorAll('tr[data-ukey]:not(.hidden)').length > 0;
                if (visibleRows) {
                    card.classList.remove('hidden');
                    visibleComputerCount++;
                } else {
                    card.classList.add('hidden');
                }
            });
            
            // Update the filtered computer count
            const countDisplay = document.getElementById('filtered-computer-count');
            if (countDisplay) {
                countDisplay.textContent = visibleComputerCount;
            }

            // Update the top failure alert list to reflect current visible failed rows
            updateFailureAlertFromFilters();

            // Sync top stat cards with the filtered view
            const totalVisibleFailed  = document.querySelectorAll('tr.failed-row:not(.hidden)').length;
            const totalVisibleUpdates = document.querySelectorAll('tr[data-ukey]:not(.hidden)').length;
            const statFailed = document.getElementById('stat-failed-value');
            const statTotal  = document.getElementById('stat-total-value');
            if (statFailed) statFailed.textContent = totalVisibleFailed;
            if (statTotal)  statTotal.textContent  = totalVisibleUpdates;
        }
        
        function updateFailureAlertFromFilters() {
            // For each failure-item in the top alert, recalculate how many visible failed rows exist for that computer
            document.querySelectorAll('.failure-alert .failure-item').forEach(item => {
                const compId = item.getAttribute('data-computer');
                if (!compId) return;
                // match rows in the report for that computer that are failed and visible
                const visibleFailedRows = document.querySelectorAll('#computer-' + compId + ' tr.failed-row:not(.hidden)').length;
                const numberSpan = item.querySelector('.failure-count-number');
                if (visibleFailedRows > 0) {
                    item.classList.remove('hidden');
                    if (numberSpan) numberSpan.textContent = visibleFailedRows;
                } else {
                    // hide the failure item if there are no visible failed rows after filtering
                    item.classList.add('hidden');
                }
            });
            // Also update the summary count in the alert subtitle
            const visibleFailureItems = document.querySelectorAll('.failure-alert .failure-item:not(.hidden)').length;
            const subtitle = document.querySelector('.failure-alert-subtitle');
            if (subtitle) {
                if (visibleFailureItems === 1) subtitle.textContent = '1 computer requires attention';
                else subtitle.textContent = visibleFailureItems + ' computers require attention';
            }
            
            // Update connection failure count (always visible, not filtered by updates)
            const unreachableComputers = document.querySelectorAll('.computer-card.unreachable:not(.hidden)').length;
            const connFailureCount = document.getElementById('connection-failure-count');
            if (connFailureCount) {
                connFailureCount.textContent = unreachableComputers;
            }
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
 
        function toggleComputer(computerName) {
            const header = event.currentTarget;
            const content = document.getElementById('content-' + computerName);
 
            header.classList.toggle('expanded');
            content.classList.toggle('expanded');
 
            // When opening, ensure the content area is scrolled to top
            if (content.classList.contains('expanded')) {
                content.scrollTop = 0;
            }
        }
        
        function expandAll() {
            document.querySelectorAll('.computer-header').forEach(header => {
                header.classList.add('expanded');
            });
            document.querySelectorAll('.computer-content').forEach(content => {
                content.classList.add('expanded');
                // reset scroll so tables start at top
                content.scrollTop = 0;
            });
        }
        
        function collapseAll() {
            document.querySelectorAll('.computer-header').forEach(header => {
                header.classList.remove('expanded');
            });
            document.querySelectorAll('.computer-content').forEach(content => {
                content.classList.remove('expanded');
                // reset scroll position to top so next open shows top
                content.scrollTop = 0;
            });
            // remove any hidden filter when collapsing
            document.querySelectorAll('.computer-card').forEach(card => card.classList.remove('hidden'));
        }
        
        function expandFailures() {
            // First collapse all
            collapseAll();
            
            // Hide computers without failures
            document.querySelectorAll('.computer-card').forEach(card => {
                if (!card.classList.contains('has-failures')) {
                    card.classList.add('hidden');
                } else {
                    card.classList.remove('hidden');
                    // Expand the ones with failures
                    const cleanName = card.id.replace('computer-', '');
                    const header = card.querySelector('.computer-header');
                    const content = document.getElementById('content-' + cleanName);
                    if (header && content) {
                        header.classList.add('expanded');
                        content.classList.add('expanded');
                        content.scrollTop = 0;
                    }
                }
            });
        }
        
        function scrollToComputer(computerName) {
            // Show all computers first (in case they're filtered)
            document.querySelectorAll('.computer-card').forEach(card => {
                card.classList.remove('hidden');
            });
            
            // Find the computer card
            const computerCard = document.getElementById('computer-' + computerName);
            if (computerCard) {
                // Scroll to it
                computerCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
                
                // Expand it
                const header = computerCard.querySelector('.computer-header');
                const content = document.getElementById('content-' + computerName);
                if (header && content) {
                    header.classList.add('expanded');
                    content.classList.add('expanded');
                    // Ensure the content scroll starts at the top
                    content.scrollTop = 0;
 
                    // Add a highlight animation
                    computerCard.style.animation = 'none';
                    setTimeout(() => {
                        computerCard.style.animation = 'highlightCard 1.5s ease-out';
                    }, 100);
                }
            }
        }
        
        function filterComputers(searchText) {
            const searchLower = searchText.toLowerCase();
            const cards = document.querySelectorAll('.computer-card');
            
            cards.forEach(card => {
                const computerName = card.getAttribute('data-computer').toLowerCase();
                const content = card.textContent.toLowerCase();
                
                if (searchText === '' || computerName.includes(searchLower) || content.includes(searchLower)) {
                    card.classList.remove('hidden');
                } else {
                    card.classList.add('hidden');
                }
            });
        }
        
        // Add highlight animation style
        const style = document.createElement('style');
        style.textContent = '@keyframes highlightCard { 0% { box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.7); } 50% { box-shadow: 0 0 30px 10px rgba(239, 68, 68, 0.3); } 100% { box-shadow: 0 10px 25px rgba(0,0,0,0.1); } }';
        document.head.appendChild(style);
        
        // Animate elements on load
        document.addEventListener('DOMContentLoaded', function() {
            const cards = document.querySelectorAll('.computer-card');
            cards.forEach((card, index) => {
                card.style.animationDelay = (index * 0.1) + 's';
            });
            
            const statCards = document.querySelectorAll('.stat-card');
            statCards.forEach((card, index) => {
                card.style.animationDelay = (index * 0.1) + 's';
            });
            // render update filters and apply default visibility
            renderUpdateFilters();
            updateFiltersFromUI();

            // Show re-run banner if there are candidates
            if (rerunComputers.length > 0) {
                const banner = document.getElementById('rerun-banner');
                if (banner) banner.style.display = '';
                const cnt = document.getElementById('rerun-count');
                if (cnt) cnt.textContent = rerunComputers.length;
                const detail = document.getElementById('rerun-detail');
                if (detail) detail.textContent = rerunComputers.map(c => c.name + ' (' + c.reason + ')').join('  \u2022  ');
            }
        });
 
        // Sort toggle state
        let sortDescending = false;
 
        // Map statuses to a priority number for sorting (lower = higher priority)
        function statusPriority(status) {
            const map = {
                'Failed': 1,
                'Incomplete': 2,
                'NoUpdatesNeeded': 3,
                'Completed': 4,
                'Unknown': 5
            };
            return map[status] || 99;
        }
 
        function sortComputersByStatus() {
            const container = document.getElementById('computers-container');
            const cards = Array.from(container.querySelectorAll('.computer-card'));
 
            cards.sort((a, b) => {
                const aStatus = a.getAttribute('data-status') || 'Unknown';
                const bStatus = b.getAttribute('data-status') || 'Unknown';
                const diff = statusPriority(aStatus) - statusPriority(bStatus);
                return sortDescending ? -diff : diff;
            });
 
            // Re-append in new order
            cards.forEach(card => container.appendChild(card));
 
            // Toggle ordering for next click
            sortDescending = !sortDescending;
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
    </script>
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
