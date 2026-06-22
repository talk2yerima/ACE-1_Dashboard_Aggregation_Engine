# install_service.ps1
# Installs the RADET Dashboard Aggregation Engine as a Windows Service using NSSM.
# Double-click or run from any PowerShell window - auto-elevates to Admin.

$ServiceName = "RADET_Dashboard_Engine"
$DisplayName = "RADET Dashboard Aggregation Engine"
$Description = "Processes ACE-1 RADET XLSX files on a schedule and uploads DashboardSummary.csv to Azure Blob Storage."

$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Definition
$NodeCmd    = Get-Command node -ErrorAction SilentlyContinue
$NodeExe    = if ($NodeCmd) { $NodeCmd.Source } else { $null }
$Runner     = Join-Path $ScriptDir "service_runner.js"
$LogDir     = Join-Path $ScriptDir "logs"
$NssmDir    = Join-Path $ScriptDir "nssm"
$NssmExe    = Join-Path $NssmDir   "nssm.exe"

# --- Verify Administrator (setup.bat self-elevates before calling this) ---
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]"Administrator"
)
if (-not $isAdmin) {
    Write-Error "This script must run as Administrator. Run setup.bat which handles elevation automatically."
    Read-Host "Press Enter to exit"
    exit 1
}

# --- Check prerequisites ---
if (-not $NodeExe -or -not (Test-Path $NodeExe)) {
    Write-Error "Node.js not found in PATH. Install from https://nodejs.org (LTS recommended)."
    Read-Host "Press Enter to exit"
    exit 1
}

$distIndex = Join-Path $ScriptDir "dist\index.js"
if (-not (Test-Path $distIndex)) {
    Write-Error "dist\index.js not found. Run:  npm run build  first."
    Read-Host "Press Enter to exit"
    exit 1
}

if (-not (Test-Path (Join-Path $ScriptDir ".env"))) {
    Write-Error ".env not found. Copy .env.example to .env and fill in your values."
    Read-Host "Press Enter to exit"
    exit 1
}

# --- Ensure NSSM is present (bundled in ZIP or download fallback) ---
if (-not (Test-Path $NssmExe)) {
    Write-Host "NSSM not found locally - attempting download ..." -ForegroundColor Yellow
    Write-Host "  (If this VM has no internet, place nssm.exe at: $NssmExe)"
    $NssmZip     = Join-Path $env:TEMP "nssm.zip"
    $NssmExtract = Join-Path $env:TEMP "nssm_extract"
    try {
        Invoke-WebRequest -Uri "https://nssm.cc/release/nssm-2.24.zip" -OutFile $NssmZip -UseBasicParsing -TimeoutSec 30
        Expand-Archive -Path $NssmZip -DestinationPath $NssmExtract -Force
        New-Item -ItemType Directory -Force -Path $NssmDir | Out-Null
        Copy-Item "$NssmExtract\nssm-2.24\win64\nssm.exe" $NssmExe -Force
        Remove-Item $NssmZip, $NssmExtract -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "NSSM downloaded and ready." -ForegroundColor Green
    }
    catch {
        Write-Host ""
        Write-Host "ERROR: Could not download NSSM (no internet or firewall blocking nssm.cc)." -ForegroundColor Red
        Write-Host "  Fix: Download nssm-2.24.zip from https://nssm.cc on another machine," -ForegroundColor Yellow
        Write-Host "       extract win64\nssm.exe and place it at:" -ForegroundColor Yellow
        Write-Host "       $NssmExe" -ForegroundColor Cyan
        Read-Host "Press Enter to exit"
        exit 1
    }
} else {
    Write-Host "NSSM found: $NssmExe" -ForegroundColor Green
}

# --- Create logs dir (needed before short-path resolution) ---
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

# --- Resolve 8.3 short paths (fallback to quoted long path if 8.3 disabled) ---
function Get-ShortPath($p) {
    $s = & cmd /c "for %I in (`"$p`") do @echo %~sI" 2>$null
    $s = $s.Trim()
    # If 8.3 resolution failed (returned same long path or empty), use original
    if (-not $s -or $s -eq $p) { return $p }
    return $s
}

function Quote-IfSpaces($p) {
    if ($p -match ' ') { return "`"$p`"" }
    return $p
}

$NodeShort   = Get-ShortPath $NodeExe
$RunnerShort = Get-ShortPath $Runner
$DirShort    = Get-ShortPath $ScriptDir
$LogShort    = Get-ShortPath $LogDir

# Wrap in quotes if path still contains spaces (8.3 disabled on this drive)
$NodeArg    = Quote-IfSpaces $NodeShort
$RunnerArg  = Quote-IfSpaces $RunnerShort

Write-Host "Paths:"
Write-Host "  Node   : $NodeArg"
Write-Host "  Runner : $RunnerArg"
Write-Host "  Dir    : $DirShort"

# --- Remove any existing service (handles RUNNING, PAUSED, STOPPED states) ---
Write-Host "Checking for existing service '$ServiceName' ..."
$svcQuery = & sc.exe query $ServiceName 2>&1
if ("$svcQuery" -notmatch "does not exist") {
    Write-Host "Existing service found - removing ..." -ForegroundColor Yellow
    # Resume first if paused (nssm stop fails on PAUSED services)
    & sc.exe control $ServiceName RESUME 2>&1 | Out-Null
    Start-Sleep -Seconds 1
    & $NssmExe stop   $ServiceName confirm 2>&1 | Out-Null
    Start-Sleep -Seconds 3
    & $NssmExe remove $ServiceName confirm 2>&1 | Out-Null
    Start-Sleep -Seconds 2
    & sc.exe delete $ServiceName 2>&1 | Out-Null
    Start-Sleep -Seconds 3
    Write-Host "Removed." -ForegroundColor Yellow
} else {
    Write-Host "No existing service found." -ForegroundColor Green
}

# --- Install service ---
Write-Host "Installing service '$ServiceName' ..."
& $NssmExe install $ServiceName $NodeArg $RunnerArg

& $NssmExe set $ServiceName Application   $NodeArg
& $NssmExe set $ServiceName AppParameters $RunnerArg
& $NssmExe set $ServiceName AppDirectory  $DirShort
& $NssmExe set $ServiceName DisplayName   $DisplayName
& $NssmExe set $ServiceName Description   $Description
& $NssmExe set $ServiceName Start         SERVICE_AUTO_START

$StdoutLog = Quote-IfSpaces (Join-Path $LogShort "RADET_stdout.log")
$StderrLog = Quote-IfSpaces (Join-Path $LogShort "RADET_stderr.log")
& $NssmExe set $ServiceName AppStdout                    $StdoutLog
& $NssmExe set $ServiceName AppStderr                    $StderrLog
& $NssmExe set $ServiceName AppStdoutCreationDisposition 4
& $NssmExe set $ServiceName AppStderrCreationDisposition 4
& $NssmExe set $ServiceName AppRotateFiles               1
& $NssmExe set $ServiceName AppRotateOnline              1
& $NssmExe set $ServiceName AppRotateBytes               5242880
& $NssmExe set $ServiceName AppRestartDelay              30000
& $NssmExe set $ServiceName AppThrottle                  60000

# --- Start service ---
Write-Host ""
Write-Host "Starting service ..."
& $NssmExe start $ServiceName

# Wait up to 15 seconds for service to reach RUNNING
$status = ""
for ($i = 0; $i -lt 5; $i++) {
    Start-Sleep -Seconds 3
    $status = (& $NssmExe status $ServiceName).Trim()
    if ($status -eq "SERVICE_RUNNING") { break }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  $ServiceName" -ForegroundColor Green
Write-Host "  Status : $status"
Write-Host "  Logs   : $LogDir"
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Commands:"
Write-Host "  .\nssm\nssm.exe status  $ServiceName"
Write-Host "  .\nssm\nssm.exe stop    $ServiceName"
Write-Host "  .\nssm\nssm.exe start   $ServiceName"
Write-Host "  .\nssm\nssm.exe remove  $ServiceName confirm"
Write-Host ""
Write-Host "Logs:"
Write-Host "  Get-Content '$LogDir\RADET_stdout.log' -Tail 50 -Wait"
Write-Host ""
Write-Host "Schedule: 09:00  11:00  13:00  15:00  17:00  19:00  21:00  (daily)"
Write-Host "Startup : Runs immediately once when service starts/restarts"
Read-Host "Press Enter to close"
