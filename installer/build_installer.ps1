# ==============================================================================
# DocLayer Production Build & Installer Packaging Script
# ==============================================================================
param(
    [string]$Configuration = "Release",
    [string]$Version = "1.0.0"
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RootDir = Split-Path -Parent $ScriptDir
$OutputDir = Join-Path $ScriptDir "dist"

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " Building DocLayer v$Version ($Configuration)..." -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

if (Test-Path $OutputDir) {
    Remove-Item $OutputDir -Recurse -Force
}
New-Item -ItemType Directory -Path $OutputDir | Out-Null

$AppPublishDir = Join-Path $OutputDir "app"
$SetupPublishDir = Join-Path $OutputDir "setup"

# 1. Publish Main DocLayer Application
Write-Host "`n[1/3] Publishing DocLayer Main Application..." -ForegroundColor Yellow
dotnet publish "$RootDir\WordBarcodeStudio.csproj" `
    -c $Configuration `
    -r win-x64 `
    --self-contained false `
    -o $AppPublishDir

# 2. Publish DocLayer Setup Wizard
Write-Host "`n[2/3] Publishing DocLayer Setup Wizard..." -ForegroundColor Yellow
dotnet publish "$ScriptDir\DocLayer.Setup\DocLayer.Setup.csproj" `
    -c $Configuration `
    -r win-x64 `
    --self-contained false `
    -o $SetupPublishDir

# 3. Create Installer Bundle Directory
Write-Host "`n[3/3] Assembling Installer Bundle..." -ForegroundColor Yellow
$BundleDir = Join-Path $OutputDir "DocLayer-v$Version-Setup"
New-Item -ItemType Directory -Path $BundleDir | Out-Null

# Copy application files into the bundle
Copy-Item "$AppPublishDir\*" -Destination $BundleDir -Recurse -Force
# Copy Setup executable to root of bundle
Copy-Item "$SetupPublishDir\DocLayer.Setup.exe" -Destination "$BundleDir\DocLayer.Setup.exe" -Force
Copy-Item "$SetupPublishDir\DocLayer.Setup.dll" -Destination "$BundleDir\DocLayer.Setup.dll" -Force
Copy-Item "$SetupPublishDir\DocLayer.Setup.runtimeconfig.json" -Destination "$BundleDir\DocLayer.Setup.runtimeconfig.json" -Force

# 4. Create Release Zip Archive
$ZipPath = Join-Path $OutputDir "DocLayer-v$Version-Windows-Setup.zip"
Write-Host "`n[4/4] Creating Release ZIP Archive: $ZipPath" -ForegroundColor Yellow
if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
Compress-Archive -Path "$BundleDir\*" -DestinationPath $ZipPath -CompressionLevel Optimal

Write-Host "`n=====================================================" -ForegroundColor Green
Write-Host " Installer Bundle Created Successfully!" -ForegroundColor Green
Write-Host " Setup Executable: $BundleDir\DocLayer.Setup.exe" -ForegroundColor Green
Write-Host " Release Archive:  $ZipPath" -ForegroundColor Green
Write-Host "=====================================================" -ForegroundColor Green
