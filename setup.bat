@echo off
REM ============================================================
REM  setup.bat  -  First-time setup for RADET Dashboard Engine
REM  Double-click or run from Command Prompt.
REM ============================================================
echo.
echo ============================================
echo  RADET Dashboard Aggregation Engine  -  Setup
echo ============================================
echo.

REM -- 1. Check Node.js --
node --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Node.js not found.
    echo         Install Node.js LTS from https://nodejs.org
    echo         IMPORTANT: Check "Add to PATH" during install.
    pause & exit /b 1
)
echo [1/5] Node.js OK:
node --version

REM -- 2. Check npm --
npm --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] npm not found. Reinstall Node.js.
    pause & exit /b 1
)
echo       npm OK:
npm --version

REM -- 3. Install dependencies --
echo.
echo [2/5] Installing npm dependencies...
call npm install --omit=dev
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] npm install failed. Check your internet connection.
    pause & exit /b 1
)
echo       Dependencies installed.

REM -- 4. Build TypeScript --
echo.
echo [3/5] Building TypeScript...
call npm run build
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] TypeScript build failed. Check for compile errors above.
    pause & exit /b 1
)
echo       Build complete.

REM -- 5. Create folders and .env --
echo.
echo [4/5] Preparing folders and config...
IF NOT EXIST "input"   mkdir input
IF NOT EXIST "outputs" mkdir outputs
IF NOT EXIST "logs"    mkdir logs
echo       input\  outputs\  logs\  ready.

IF NOT EXIST ".env" (
    copy .env.example .env >nul
    echo       .env created from .env.example
    echo.
    echo  *** ACTION REQUIRED ***
    echo  Open .env and fill in your values before continuing:
    echo    AZURE_STORAGE_CONNECTION_STRING
    echo    AZURE_STORAGE_CONTAINER  (default: powerbi-datasource)
    echo    RUN_INTERVAL_HOURS       (default: 6)
    echo.
    echo  Also drop your ACE-1_Combined_RADET*.xlsx file into the input\ folder.
) ELSE (
    echo       .env already exists.
)

REM -- 6. Install the Windows Service --
echo.
echo [5/5] Installing RADET Dashboard Engine as a Windows Service...
echo       (This will open a UAC prompt to grant Administrator rights)
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_service.ps1"

echo.
echo ============================================
echo  Setup complete!
echo ============================================
echo.
echo  The service is now running as: RADET_Dashboard_Engine
echo  It will process the RADET file every RUN_INTERVAL_HOURS hours
echo  (default: 6 hours) and upload the output to Azure Blob Storage.
echo.
echo  To check status:
echo    nssm\nssm.exe status RADET_Dashboard_Engine
echo.
echo  To view live logs:
echo    powershell -Command "Get-Content logs\RADET_stdout.log -Tail 50 -Wait"
echo.
echo  To drop a new RADET file:
echo    Copy ACE-1_Combined_RADET*.xlsx into the input\ folder.
echo    The next scheduled run will pick it up automatically.
echo.
pause
