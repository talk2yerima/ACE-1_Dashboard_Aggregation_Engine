@echo off
REM ============================================================
REM  setup.bat  -  Setup for RADET Dashboard Engine
REM  Double-click to install. Requires internet connection.
REM ============================================================

REM -- Self-elevate to Administrator if not already --
net session >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Requesting Administrator privileges...
    powershell -NoProfile -Command ^
        "Start-Process -FilePath cmd.exe -ArgumentList '/c cd /d \"%~dp0\" && \"%~f0\"' -Verb RunAs -Wait"
    exit /b
)

cd /d "%~dp0"

echo.
echo ============================================
echo  RADET Dashboard Aggregation Engine  -  Setup
echo ============================================
echo.

REM -- [1/4] Check Node.js --
node --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Node.js not found.
    echo.
    echo         Please install Node.js LTS from:
    echo           https://nodejs.org
    echo.
    echo         IMPORTANT: During install, check the box:
    echo           "Add to PATH"
    echo.
    echo         Then re-run this setup.
    pause & exit /b 1
)
echo [1/4] Node.js found:
node --version

REM -- [2/4] Install production dependencies from internet --
echo.
echo [2/4] Installing production packages from internet...
echo       (This downloads only what is needed to run - no dev tools)
call npm install --omit=dev --prefer-offline
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Package installation failed.
    echo         Make sure this computer has an internet connection.
    echo         Then re-run this setup.
    pause & exit /b 1
)
echo       Packages installed successfully.

REM -- [3/4] Create folders and verify .env --
echo.
echo [3/4] Preparing folders and config...
IF NOT EXIST "input"   mkdir input
IF NOT EXIST "outputs" mkdir outputs
IF NOT EXIST "logs"    mkdir logs
echo       input\  outputs\  logs\  ready.

IF NOT EXIST ".env" (
    copy .env.example .env >nul
    echo       .env created from .env.example
    echo.
    echo  *** ACTION REQUIRED ***
    echo  Open .env with Notepad and fill in your Azure values:
    echo    AZURE_STORAGE_CONNECTION_STRING
    echo    AZURE_STORAGE_CONTAINER
    echo.
    pause
) ELSE (
    echo       .env found - OK.
)

REM -- [4/4] Install the Windows Service --
echo.
echo [4/4] Installing RADET Dashboard Engine as a Windows Service...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_service.ps1"
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Service installation failed. See messages above.
    pause & exit /b 1
)

echo.
echo ============================================
echo  Setup complete!
echo ============================================
echo.
echo  Service name : RADET_Dashboard_Engine
echo  First run    : Immediately when service starts/restarts
echo  Schedule     : 09:00  11:00  13:00  15:00  17:00  19:00  21:00
echo  Logs         : %~dp0logs\RADET_stdout.log
echo.
echo  Useful commands (run from this folder):
echo    Check status : nssm\nssm.exe status RADET_Dashboard_Engine
echo    View logs    : powershell -Command "Get-Content logs\RADET_stdout.log -Tail 50 -Wait"
echo    Stop service : nssm\nssm.exe stop RADET_Dashboard_Engine
echo    Start service: nssm\nssm.exe start RADET_Dashboard_Engine
echo.
pause
