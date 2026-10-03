param([string]$Bundle = "")
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $Bundle) { $Bundle = Join-Path $repoRoot 'dist\RapidPyMain' }
$console = Join-Path ([System.IO.Path]::GetFullPath($Bundle)) 'RapidPyMainConsole.exe'
if (-not (Test-Path -LiteralPath $console)) { throw "Missing console entry point: $console" }
$checkRoot = Join-Path $repoRoot ('.tmp\release-check-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $checkRoot -Force | Out-Null
$savedConfig = $env:RAPID_CONFIG
$savedPythonPath = $env:PYTHONPATH
$savedPlatform = $env:QT_QPA_PLATFORM
try {
    $env:RAPID_CONFIG = Join-Path $checkRoot 'operator-test-config.json'
    $env:PYTHONPATH = ''
    $env:QT_QPA_PLATFORM = 'offscreen'
    Push-Location $checkRoot
    try {
        $startupJson = & $console --check-startup
        if ($LASTEXITCODE -ne 0) { throw 'Bundled startup check failed.' }
        $startup = ($startupJson -join "`n") | ConvertFrom-Json
        $smokeJson = & $console --smoke-test
        if ($LASTEXITCODE -ne 0) { throw 'Bundled UI smoke check failed.' }
        $smoke = ($smokeJson -join "`n") | ConvertFrom-Json
        if (-not $startup.ok -or -not $startup.packaged -or -not $smoke.ok -or -not $smoke.simulated) {
            throw 'Release checks did not verify an isolated packaged simulation.'
        }
        $report = [ordered]@{
            schema = 'rapidpy.portable_verification.v1'
            checked_at_utc = [DateTimeOffset]::UtcNow.ToString('o')
            executable = $console
            executable_sha256 = (Get-FileHash -LiteralPath $console -Algorithm SHA256).Hash.ToLowerInvariant()
            startup = $startup
            smoke = $smoke
        }
        $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $checkRoot 'verification.json') -Encoding utf8
        Write-Output "Portable startup and offline UI checks passed. Evidence: $checkRoot\verification.json"
    } finally { Pop-Location }
} finally {
    $env:RAPID_CONFIG = $savedConfig
    $env:PYTHONPATH = $savedPythonPath
    $env:QT_QPA_PLATFORM = $savedPlatform
}
