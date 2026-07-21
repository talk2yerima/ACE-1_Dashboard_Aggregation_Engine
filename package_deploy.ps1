# package_deploy.ps1  —  Build and zip the RADET Dashboard deployment package
# Run from the project root:  .\package_deploy.ps1
# Output: radet-dashboard-deploy.zip  (ready to copy to the VM)

$ErrorActionPreference = 'Stop'

$Root    = $PSScriptRoot
$ZipName = "radet-dashboard-deploy.zip"
$ZipPath = Join-Path $Root $ZipName

Write-Host "`n=== RADET Dashboard  -  Packaging for Deployment ===" -ForegroundColor Cyan

# 1. Compile TypeScript
Write-Host "`n[1/3] Compiling TypeScript..." -ForegroundColor Yellow
Push-Location $Root
npm run build
if ($LASTEXITCODE -ne 0) { Write-Error "Build failed."; exit 1 }
Write-Host "      Build OK." -ForegroundColor Green

# 2. Remove previous zip if any
if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }

# 3. Files/folders to include in the package
$Items = @(
    "dist",
    "config",
    "scripts",
    "nssm",
    "package.json",
    "package-lock.json",
    ".env.example",
    "setup.bat",
    "install_service.ps1",
    "service_runner.js",
    "DEPLOYMENT_GUIDE.txt"
)

# File extensions to exclude when zipping a folder (stray/editor temp files
# that shouldn't ship, e.g. config/indicators.yaml.tmp)
$ExcludeExtensions = @(".tmp")

Write-Host "`n[2/3] Creating deployment ZIP..." -ForegroundColor Yellow

Add-Type -AssemblyName System.IO.Compression.FileSystem

$Zip = [System.IO.Compression.ZipFile]::Open($ZipPath, 'Create')

function Add-ToZip {
    param($zip, $sourcePath, $entryPrefix)
    if (Test-Path $sourcePath -PathType Leaf) {
        $entry = $entryPrefix
        $stream = [System.IO.File]::OpenRead($sourcePath)
        $zipEntry = $zip.CreateEntry($entry)
        $zipStream = $zipEntry.Open()
        $stream.CopyTo($zipStream)
        $zipStream.Close(); $stream.Close()
    } elseif (Test-Path $sourcePath -PathType Container) {
        $files = Get-ChildItem $sourcePath -Recurse -File | Where-Object { $ExcludeExtensions -notcontains $_.Extension }
        foreach ($file in $files) {
            $relative = $file.FullName.Substring($sourcePath.Length).TrimStart('\','/')
            $entry = "$entryPrefix/$relative" -replace '\\','/'
            $stream = [System.IO.File]::OpenRead($file.FullName)
            $zipEntry = $zip.CreateEntry($entry)
            $zipStream = $zipEntry.Open()
            $stream.CopyTo($zipStream)
            $zipStream.Close(); $stream.Close()
        }
    }
}

foreach ($item in $Items) {
    $fullPath = Join-Path $Root $item
    if (-not (Test-Path $fullPath)) {
        Write-Host "      SKIP (not found): $item" -ForegroundColor DarkGray
        continue
    }
    Add-ToZip -zip $Zip -sourcePath $fullPath -entryPrefix $item
    Write-Host "      Added: $item"
}

$Zip.Dispose()
Pop-Location

# Done
$size = [math]::Round((Get-Item $ZipPath).Length / 1MB, 1)
Write-Host "`n[3/3] Done!" -ForegroundColor Green
Write-Host "`n  Package : $ZipPath"
Write-Host "  Size    : $size MB"
Write-Host "`n  Next steps:"
Write-Host "  1. Copy $ZipName to the target Windows VM"
Write-Host "  2. Extract to a folder, e.g.  C:\radet-dashboard\"
Write-Host "  3. Edit .env with your Azure connection strings"
Write-Host "  4. Double-click setup.bat (or right-click > Run as administrator)"
Write-Host ""
