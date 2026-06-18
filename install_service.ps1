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

# --- Auto-elevate to Administrator ---
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]"Administrator"
)
if (-not $isAdmin) {
    Write-Host "Relaunching as Administrator ..."
    $elevateArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Definition)`""
    Start-Process powershell -ArgumentList $elevateArgs -Verb RunAs
    exit
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

# --- Download NSSM if not present ---
if (-not (Test-Path $NssmExe)) {
    Write-Host "Downloading NSSM ..."
    $NssmZip     = Join-Path $env:TEMP "nssm.zip"
    $NssmExtract = Join-Path $env:TEMP "nssm_extract"
    try {
        Invoke-WebRequest -Uri "https://nssm.cc/release/nssm-2.24.zip" -OutFile $NssmZip -UseBasicParsing
        Expand-Archive -Path $NssmZip -DestinationPath $NssmExtract -Force
        New-Item -ItemType Directory -Force -Path $NssmDir | Out-Null
        Copy-Item "$NssmExtract\nssm-2.24\win64\nssm.exe" $NssmExe -Force
        Remove-Item $NssmZip, $NssmExtract -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "NSSM ready." -ForegroundColor Green
    }
    catch {
        Write-Error "NSSM download failed: $_"
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# --- Create logs dir (needed before short-path resolution) ---
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

# --- Resolve 8.3 short paths (no spaces - safe for NSSM) ---
function Get-ShortPath($p) {
    $s = & cmd /c "for %I in (`"$p`") do @echo %~sI" 2>$null
    return $s.Trim()
}
$NodeShort   = Get-ShortPath $NodeExe
$RunnerShort = Get-ShortPath $Runner
$DirShort    = Get-ShortPath $ScriptDir
$LogShort    = Get-ShortPath $LogDir

Write-Host "Paths resolved (8.3 short form):"
Write-Host "  Node   : $NodeShort"
Write-Host "  Runner : $RunnerShort"
Write-Host "  Dir    : $DirShort"

# --- Remove any existing service ---
Write-Host "Checking for existing service '$ServiceName' ..."
$svcQuery = & sc.exe query $ServiceName 2>&1
if ("$svcQuery" -notmatch "does not exist") {
    Write-Host "Existing service found - removing ..." -ForegroundColor Yellow
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
& $NssmExe install $ServiceName $NodeShort $RunnerShort

& $NssmExe set $ServiceName Application   $NodeShort
& $NssmExe set $ServiceName AppParameters $RunnerShort
& $NssmExe set $ServiceName AppDirectory  $DirShort
& $NssmExe set $ServiceName DisplayName   $DisplayName
& $NssmExe set $ServiceName Description   $Description
& $NssmExe set $ServiceName Start         SERVICE_AUTO_START

$StdoutLog = Join-Path $LogShort "RADET_stdout.log"
$StderrLog = Join-Path $LogShort "RADET_stderr.log"
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
Start-Sleep -Seconds 3
$status = & $NssmExe status $ServiceName

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
Write-Host "Schedule: runs every RUN_INTERVAL_HOURS hours (default: 6)"
Read-Host "Press Enter to close"
