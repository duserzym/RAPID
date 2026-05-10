<#
.SYNOPSIS
    Builds the RapidPy Gaussmeter installer.

.DESCRIPTION
    1. Packages the Python app with PyInstaller (one-folder bundle).
    2. Compiles the Inno Setup installer script into a single .exe installer.

.REQUIREMENTS
    - PyInstaller  : pip install pyinstaller  (in the target conda/venv)
    - Inno Setup 6 : https://jrsoftware.org/isinfo.php
      Default install path expected at:
        C:\Program Files (x86)\Inno Setup 6\ISCC.exe
      Override with -InnoSetupPath if installed elsewhere.

.PARAMETER CondaEnv
    Name of the conda environment to use.  Defaults to "paleomag".

.PARAMETER InnoSetupPath
    Full path to ISCC.exe.  Defaults to the standard Inno Setup 6 location.

.PARAMETER SkipPyInstaller
    Skip the PyInstaller step (use existing dist\ output).

.PARAMETER SkipInnoSetup
    Skip the Inno Setup step (produce only the PyInstaller bundle).

.EXAMPLE
    # Full build
    .\installer\build_installer.ps1

    # Re-use an existing PyInstaller dist and just re-compile the installer
    .\installer\build_installer.ps1 -SkipPyInstaller

    # Use a different conda environment
    .\installer\build_installer.ps1 -CondaEnv myenv
#>

[CmdletBinding()]
param(
    [string] $CondaEnv      = "paleomag",
    [string] $InnoSetupPath = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
    [switch] $SkipPyInstaller,
    [switch] $SkipInnoSetup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot   = (Resolve-Path "$PSScriptRoot\..").Path
$SpecFile        = Join-Path $PSScriptRoot "rapid_gaussmeter.spec"
$DrvInstSpecFile = Join-Path $PSScriptRoot "install_fwbell_drivers.spec"
$IssFile    = Join-Path $PSScriptRoot "rapid_gaussmeter_installer.iss"
$DistDir    = Join-Path $RepoRoot "dist"
$BundleDir  = Join-Path $DistDir "RapidPy_Gaussmeter"
$DrvBundleDir = Join-Path $DistDir "install_fwbell_drivers"

Write-Host ""
Write-Host "==== RapidPy Gaussmeter Installer Build ====" -ForegroundColor Cyan
Write-Host "  Repo root          : $RepoRoot"
Write-Host "  App spec           : $SpecFile"
Write-Host "  Driver inst spec   : $DrvInstSpecFile"
Write-Host "  ISS file           : $IssFile"
Write-Host "  Output dir         : $DistDir"
Write-Host ""

# ── Step 1: PyInstaller ──────────────────────────────────────────────────────
if (-not $SkipPyInstaller) {
    Write-Host "-- Step 1: PyInstaller --" -ForegroundColor Yellow

    # Verify spec exists
    if (-not (Test-Path $SpecFile)) {
        Write-Error "Spec file not found: $SpecFile"
        exit 1
    }

    # Run PyInstaller inside the conda environment from the repo root so that
    # relative paths in the spec resolve correctly.
    Push-Location $RepoRoot
    try {
        conda run -n $CondaEnv --no-capture-output `
            pyinstaller $SpecFile --noconfirm --clean
        if ($LASTEXITCODE -ne 0) {
            Write-Error "PyInstaller failed with exit code $LASTEXITCODE"
            exit 1
        }
    } finally {
        Pop-Location
    }

    if (-not (Test-Path $BundleDir)) {
        Write-Error "Expected PyInstaller output not found: $BundleDir"
        exit 1
    }

    Write-Host "  Bundle created: $BundleDir" -ForegroundColor Green

    # Step 1b: Driver installer bundle
    Write-Host ""
    Write-Host "-- Step 1b: PyInstaller (driver installer) --" -ForegroundColor Yellow
    if (Test-Path $DrvInstSpecFile) {
        Push-Location $RepoRoot
        try {
            conda run -n $CondaEnv --no-capture-output `
                pyinstaller $DrvInstSpecFile --noconfirm --clean
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "Driver installer PyInstaller failed (exit $LASTEXITCODE) – continuing without it."
            } else {
                Write-Host "  Driver installer bundle created: $DrvBundleDir" -ForegroundColor Green
            }
        } finally {
            Pop-Location
        }
    } else {
        Write-Host "  Driver installer spec not found, skipping." -ForegroundColor DarkGray
    }
} else {
    Write-Host "-- Step 1: Skipped (PyInstaller) --" -ForegroundColor DarkGray
    if (-not (Test-Path $BundleDir)) {
        Write-Error "No existing bundle found at $BundleDir. Run without -SkipPyInstaller first."
        exit 1
    }
}

# ── Step 2: Inno Setup ───────────────────────────────────────────────────────
if (-not $SkipInnoSetup) {
    Write-Host ""
    Write-Host "-- Step 2: Inno Setup --" -ForegroundColor Yellow

    if (-not (Test-Path $InnoSetupPath)) {
        Write-Error "ISCC.exe not found at: $InnoSetupPath`nInstall Inno Setup 6 or pass -InnoSetupPath <path>"
        exit 1
    }

    if (-not (Test-Path $IssFile)) {
        Write-Error "ISS file not found: $IssFile"
        exit 1
    }

    & $InnoSetupPath $IssFile
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Inno Setup failed with exit code $LASTEXITCODE"
        exit 1
    }

    $InstallerDir = Join-Path $DistDir "installer"
    $InstallerExe = Get-ChildItem -Path $InstallerDir -Filter "*.exe" |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1

    if ($InstallerExe) {
        Write-Host "  Installer created: $($InstallerExe.FullName)" -ForegroundColor Green
    }
} else {
    Write-Host "-- Step 2: Skipped (Inno Setup) --" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "==== Build complete ====" -ForegroundColor Cyan
