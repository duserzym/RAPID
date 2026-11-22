param(
    [string]$Python = "",
    [ValidatePattern('^([0-9a-f]{40}|[0-9a-f]{64})?$')][string]$SourceRevision = ""
)
# Native tools (pip, PyInstaller) log to stderr; Windows PowerShell 5.1 turns that into
# terminating errors under "Stop". Failures are detected through $LASTEXITCODE instead.
$ErrorActionPreference = "Continue"
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $Python) { $Python = Join-Path $repoRoot '.venv\Scripts\python.exe' }
if (-not (Test-Path -LiteralPath $Python)) { throw "Python not found: $Python" }
$stage = Join-Path $repoRoot ('.tmp\portable-build-' + [guid]::NewGuid().ToString('N'))
$wheelDir = Join-Path $stage 'wheel'
$packageDir = Join-Path $stage 'packages'
& $Python -m pip wheel --no-deps --no-build-isolation --wheel-dir $wheelDir (Join-Path $repoRoot 'RapidPy')
if ($LASTEXITCODE -ne 0) { throw 'Release wheel build failed.' }
$wheels = @(Get-ChildItem -LiteralPath $wheelDir -Filter 'berkeley_rapidpy-*.whl')
if ($wheels.Count -ne 1) { throw 'Expected one RapidPy wheel from this build.' }
$wheel = $wheels[0]
& $Python -m pip install --no-deps --no-compile --upgrade --target $packageDir $wheel.FullName
if ($LASTEXITCODE -ne 0) { throw 'Release package staging failed.' }
$savedPackageDir = $env:RAPIDPY_RELEASE_PACKAGES
try {
    $env:RAPIDPY_RELEASE_PACKAGES = $packageDir
    & $Python -m PyInstaller --noconfirm --distpath (Join-Path $repoRoot 'dist') --workpath (Join-Path $stage 'build') (Join-Path $PSScriptRoot 'rapid_main.spec')
    if ($LASTEXITCODE -ne 0) { throw 'Portable app build failed.' }
} finally {
    $env:RAPIDPY_RELEASE_PACKAGES = $savedPackageDir
}
$ErrorActionPreference = 'Stop'
$bundleDir = Join-Path $repoRoot 'dist\RapidPyMain'
Copy-Item -LiteralPath (Join-Path $repoRoot 'docs\rapid-main-portable-release.md') -Destination (Join-Path $bundleDir 'OperatorGuide.md') -Force
Copy-Item -LiteralPath (Join-Path $repoRoot 'docs\rapid_hardware_acceptance_procedure_2026-08-29.md') -Destination (Join-Path $bundleDir 'StationAcceptance.md') -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'verify_rapid_main.ps1') -Destination (Join-Path $bundleDir 'Verify-RapidPy.ps1')
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'portable_bundle.psm1') -Destination $bundleDir
Import-Module (Join-Path $PSScriptRoot 'portable_bundle.psm1') -Force
$manifest = [ordered]@{
    schema = 'rapidpy.portable_build.v2'
    built_at_utc = [DateTimeOffset]::UtcNow.ToString('o')
    source_revision = $SourceRevision.ToLowerInvariant()
    wheel = $wheel.Name
    wheel_sha256 = Get-RapidPyFileHash -Path $wheel.FullName
    files = @(Get-RapidPyBundleInventory -Bundle $bundleDir)
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $bundleDir 'build_manifest.json') -Encoding utf8
Write-Output "Portable app: $(Join-Path $repoRoot 'dist\RapidPyMain\RapidPyMain.exe')"
