# CLAUDE.md — Windows Update Deployment

## Overview

Two PowerShell scripts for deploying Windows updates across remote fleets and correlating results against Qualys vulnerability scans. Both require **PowerShell 7+**.

## Scripts

### `Windows Update Script 12112025.ps1` (~2500 lines)
Parallel Windows Update deployment tool. Four phases:
1. **Phase 1** — Start `Invoke-WUJob` on all targets via `ForEach-Object -Parallel`
2. **Phase 2** — Smart monitoring loop (marker files, scheduled task state, process detection, log stability)
3. **Phase 3** — Batched result collection via `New-PSSession` + single `Invoke-Command`, post-install verification against `Get-HotFix` and WU API
4. **Phase 4** — Interactive HTML report generation

Key functions:
- `Save-WUCredential` / `Get-WUSavedCredential` / `Remove-WUSavedCredential` — Windows Credential Manager via P/Invoke
- `Test-S2WinRM` — WinRM reachability check
- `Wait-ForUpdateJobs` — Phase 2 monitoring engine (the big one)
- `ConvertTo-HtmlEncoded` — HTML report helper

### `Compare-VulnScanToUpdates.ps1` (~1100 lines)
Correlates Qualys XLSX against deployment results. Reads via Excel COM automation.

Key functions:
- `Get-NormalizedHostname` — DNS/NetBIOS hostname cleanup
- `Get-CumulativeProductFamily` — Classifies KBs as Windows/.NET/Other for supersession logic

## Architecture Notes

- **No modules or build pipeline** — both are standalone scripts, run directly
- **Remote artifacts live in `C:\temp`** on target machines (marker files, update logs)
- **Session folders** under `C:\WindowsUpdateReports\Session_YYYYMMDD_HHMMSS\` contain all CSVs and the HTML report
- **Credential Manager target**: `WindowsUpdateScript` (different from S2NetBox's `S2NetBox` target)
- **HTML report is a giant here-string** embedded in the script (~900 lines of HTML/CSS/JS) — edit carefully, it's all inline
- **Excel COM** is required for the correlation script (no ImportExcel module)

## Critical Patterns

### Phase 2 Monitoring Complexity
`Wait-ForUpdateJobs` is the most complex function. It uses cursor repositioning (not `Clear-Host`) to preserve Phase 1 output. It handles:
- Late-arrival retries (machines that come online after Phase 1)
- Stuck-job detection and re-triggering
- Running-state and unreachable-state stall detection (warn at 15 min, auto-release at 30 min)
- Keyboard input during sleep intervals (S=skip remaining, L=list machines)
- Session-aware timestamp validation to ignore stale artifacts from prior runs

### Post-Install Verification
Phase 3 doesn't just trust PSWindowsUpdate output — it cross-references `Get-HotFix` and the Windows Update API. False failures are auto-corrected when verification proves the update installed.

### Cumulative Update Supersession (Correlation Script)
Multiple detection strategies for cumulative coverage:
1. Global KB-to-date map built from all Qualys rows (title > results > solution for date extraction)
2. Hotfix metadata (`Description = "Security Update"` + `InstalledOn` date)
3. Session-verified coverage (hosts with `NoUpdatesNeeded`/`Completed` status)
4. Product family matching (Windows cumulative won't cover .NET KB)

Date regex validates month 1-12 to prevent CVE numbers (e.g., `CVE-2026-20805`) from being parsed as dates.

### WinRM / Remoting Gotchas
- Script uses `ForEach-Object -Parallel` for Phase 1 (known to deadlock with CIM — but this uses `Invoke-Command`, not CIM, so it's fine)
- Phase 3 uses batched `New-PSSession` + single `Invoke-Command` fan-out (not parallel foreach)
- `Copy-Item -FromSession` avoids double-hop auth issues
- TCP pre-check on port 5985 before WinRM calls to skip unreachable hosts early

### Log Preservation on Rerun
Same-day reruns archive previous logs (not delete). The HTML report shows diffs between runs. Previous-day archives are cleaned up automatically.

## CSV Format

Input CSV requires `Name` and `IP` columns:
```csv
Name,IP
Site-PC-01,10.0.1.100
```

## Testing

No automated tests. Manual testing against live remote machines. Use `-ReportOnly -SessionPath` to regenerate reports from existing session data without re-deploying.

## Common Modifications

- **Adding new status detection** in Phase 2 → modify `Wait-ForUpdateJobs`, update the status enum/display table
- **Changing HTML report** → search for the `$html = @"` here-string block near the end of the script
- **Adding Qualys column mappings** → header auto-detection is in the correlation script's Pass 1; fallback positions are hardcoded
- **New output sheets in correlation report** → Excel COM calls near the end of `Compare-VulnScanToUpdates.ps1`
