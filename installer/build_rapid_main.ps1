param([string]$Python = "")
# Native tools (pip, PyInstaller) log to stderr; Windows PowerShell 5.1 turns that into
# terminating errors under "Stop". Failures are detected through $LASTEXITCODE instead.
$ErrorActionPreference = "Continue"
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $Python) { $Python = Join-Path $repoRoot '.venv\Scripts\python.exe' }
if (-not (Test-Path -LiteralPath $Python)) { throw "Python not found: $Python" }
$wheelDir = Join-Path $repoRoot '.tmp\release-wheel'
$packageDir = Join-Path $repoRoot '.tmp\portable-packages'
& $Python -m pip wheel --no-deps --no-build-isolation --wheel-dir $wheelDir (Join-Path $repoRoot 'RapidPy')
if ($LASTEXITCODE -ne 0) { throw 'Release wheel build failed.' }
$wheel = Get-ChildItem -LiteralPath $wheelDir -Filter 'berkeley_rapidpy-*.whl' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
& $Python -m pip install --no-deps --no-compile --upgrade --target $packageDir $wheel.FullName
if ($LASTEXITCODE -ne 0) { throw 'Release package staging failed.' }
& $Python -m PyInstaller --noconfirm --distpath (Join-Path $repoRoot 'dist') --workpath (Join-Path $repoRoot 'build') (Join-Path $PSScriptRoot 'rapid_main.spec')
if ($LASTEXITCODE -ne 0) { throw 'Portable app build failed.' }
$bundleDir = Join-Path $repoRoot 'dist\RapidPyMain'
Copy-Item -LiteralPath (Join-Path $repoRoot 'docs\rapid-main-portable-release.md') -Destination (Join-Path $bundleDir 'OperatorGuide.md') -Force
$manifest = [ordered]@{
    schema = 'rapidpy.portable_build.v1'
    built_at_utc = [DateTimeOffset]::UtcNow.ToString('o')
    wheel = $wheel.Name
    wheel_sha256 = (Get-FileHash -LiteralPath $wheel.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    executables = @{}
}
foreach ($name in @('RapidPyMain.exe', 'RapidPyMainConsole.exe')) {
    $manifest.executables[$name] = (Get-FileHash -LiteralPath (Join-Path $bundleDir $name) -Algorithm SHA256).Hash.ToLowerInvariant()
}
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $bundleDir 'build_manifest.json') -Encoding utf8
Write-Output "Portable app: $(Join-Path $repoRoot 'dist\RapidPyMain\RapidPyMain.exe')"
