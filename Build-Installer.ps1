<#
.SYNOPSIS
    Build VectraConnect add-in and produce VectraConnect_Setup.exe

.DESCRIPTION
    1. Restores NuGet packages
    2. Builds the VSTO add-in (Release)
    3. Runs WiX candle + light to produce the MSI
    4. Runs WiX candle + light (Bundle) to produce the final Setup.exe
    5. Outputs everything to .\bin\Release\

.REQUIREMENTS
    - Visual Studio 2022 Build Tools (or full VS) with .NET desktop workload
    - WiX Toolset v3.11 installed  →  https://wixtoolset.org/releases/
    - vstor_redist.exe placed in   →  Installer\Prerequisites\vstor_redist.exe
      Download from: https://aka.ms/vstoruntime

.USAGE
    .\Build-Installer.ps1
    .\Build-Installer.ps1 -Configuration Debug
    .\Build-Installer.ps1 -SkipAddInBuild   # if DLL already built
#>

param(
    [string] $Configuration   = "Release",
    [switch] $SkipAddInBuild
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Paths ────────────────────────────────────────────────────────────────────

$Root         = $PSScriptRoot
$AddInDir     = Join-Path $Root "VectraConnect"
$InstallerDir = Join-Path $Root "Installer"
$BinDir       = Join-Path $Root "bin\$Configuration"
$ObjDir       = Join-Path $Root "Installer\obj\$Configuration"

# Locate MSBuild
$msbuild = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" `
    -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe 2>$null |
    Select-Object -First 1

if (-not $msbuild) {
    $msbuild = "MSBuild.exe"   # fall back to PATH
}

# Locate WiX tools
$wixDir = (Get-ItemProperty "HKLM:\SOFTWARE\WiX Toolset v3.11" -ErrorAction SilentlyContinue)?.InstallFolder
if (-not $wixDir) {
    $wixDir = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\WiX Toolset v3.11" -ErrorAction SilentlyContinue)?.InstallFolder
}
if (-not $wixDir) {
    # Try common install path
    $wixDir = "C:\Program Files (x86)\WiX Toolset v3.11\bin"
}
$candle = Join-Path $wixDir "candle.exe"
$light  = Join-Path $wixDir "light.exe"

if (-not (Test-Path $candle)) {
    Write-Error "WiX candle.exe not found at '$candle'. Install WiX Toolset v3.11 from https://wixtoolset.org/releases/"
    exit 1
}

# Check prereq file
$vstoRedist = Join-Path $InstallerDir "Prerequisites\vstor_redist.exe"
if (-not (Test-Path $vstoRedist)) {
    Write-Warning "vstor_redist.exe not found at '$vstoRedist'."
    Write-Warning "Download from https://aka.ms/vstoruntime and place it there."
    Write-Warning "Continuing build — the Setup.exe will fail at runtime without it."
}

New-Item -ItemType Directory -Force -Path $BinDir | Out-Null
New-Item -ItemType Directory -Force -Path $ObjDir | Out-Null

# ── Step 1: Build VSTO add-in ────────────────────────────────────────────────

if (-not $SkipAddInBuild) {
    Write-Host "`n=== Building VSTO add-in ($Configuration) ===" -ForegroundColor Cyan

    & $msbuild "$AddInDir\VectraConnect.csproj" `
        /t:Restore,Build `
        /p:Configuration=$Configuration `
        /p:Platform="Any CPU" `
        /p:OutputPath="$BinDir\AddIn\" `
        /v:minimal

    if ($LASTEXITCODE -ne 0) { Write-Error "Add-in build failed."; exit 1 }
    Write-Host "Add-in built OK" -ForegroundColor Green
}

$addInBin = Join-Path $BinDir "AddIn"

# ── Step 2: WiX candle — compile Product.wxs → MSI object ───────────────────

Write-Host "`n=== Compiling MSI (candle) ===" -ForegroundColor Cyan

$candleMsiArgs = @(
    "-nologo",
    "-arch", "x86",
    "-ext", "WixNetFxExtension",
    "-ext", "WixUtilExtension",
    "-ext", "WixUIExtension",
    "-dVectraConnect.TargetPath=$addInBin\VectraConnect.dll",
    "-dVectraConnect.TargetDir=$addInBin\",
    "-o", "$ObjDir\Product.wixobj",
    "$InstallerDir\Product.wxs"
)

& $candle @candleMsiArgs
if ($LASTEXITCODE -ne 0) { Write-Error "candle (MSI) failed."; exit 1 }

# ── Step 3: WiX light — link MSI ─────────────────────────────────────────────

Write-Host "`n=== Linking MSI (light) ===" -ForegroundColor Cyan

$msiPath = Join-Path $BinDir "VectraConnect.msi"

$lightMsiArgs = @(
    "-nologo",
    "-ext", "WixNetFxExtension",
    "-ext", "WixUtilExtension",
    "-ext", "WixUIExtension",
    "-cultures:en-US",
    "-loc", "$InstallerDir\Assets\en-US.wxl",
    "-out", $msiPath,
    "$ObjDir\Product.wixobj"
)

& $light @lightMsiArgs
if ($LASTEXITCODE -ne 0) { Write-Error "light (MSI) failed."; exit 1 }
Write-Host "MSI created: $msiPath" -ForegroundColor Green

# ── Step 4: WiX candle — compile Bundle.wxs → bootstrapper object ────────────

Write-Host "`n=== Compiling Setup.exe (candle bundle) ===" -ForegroundColor Cyan

$candleBundleArgs = @(
    "-nologo",
    "-arch", "x86",
    "-ext", "WixNetFxExtension",
    "-ext", "WixUtilExtension",
    "-ext", "WixBalExtension",
    "-dVectraConnectMsiPath=$msiPath",
    "-o", "$ObjDir\Bundle.wixobj",
    "$InstallerDir\Bundle.wxs"
)

& $candle @candleBundleArgs
if ($LASTEXITCODE -ne 0) { Write-Error "candle (Bundle) failed."; exit 1 }

# ── Step 5: WiX light — link bootstrapper EXE ────────────────────────────────

Write-Host "`n=== Linking Setup.exe (light) ===" -ForegroundColor Cyan

$exePath = Join-Path $BinDir "VectraConnect_Setup.exe"

$lightBundleArgs = @(
    "-nologo",
    "-ext", "WixNetFxExtension",
    "-ext", "WixUtilExtension",
    "-ext", "WixBalExtension",
    "-cultures:en-US",
    "-out", $exePath,
    "$ObjDir\Bundle.wixobj"
)

& $light @lightBundleArgs
if ($LASTEXITCODE -ne 0) { Write-Error "light (Bundle) failed."; exit 1 }

# ── Done ──────────────────────────────────────────────────────────────────────

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  BUILD COMPLETE" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Installer EXE : $exePath"
Write-Host "  MSI (internal): $msiPath"
Write-Host ""
Write-Host "  Share VectraConnect_Setup.exe with your planners." -ForegroundColor Yellow
Write-Host "  They double-click it — done. No admin rights needed." -ForegroundColor Yellow
Write-Host ""
