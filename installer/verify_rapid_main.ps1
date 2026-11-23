param(
    [string]$Bundle = '',
    [string]$EvidenceDirectory = '',
    [switch]$RequireFullInventory,
    [ValidateRange(10, 600)][int]$TimeoutSeconds = 120,
    [switch]$MeasurePerformance,
    [ValidateRange(1, 100)][int]$PerformanceSpecimens = 40,
    [ValidateRange(1, 60)][int]$PerformanceIdleSeconds = 5,
    [ValidateRange(15, 1800)][int]$PerformanceTimeoutSeconds = 300
)

$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'portable_bundle.psm1') -Force

if (-not $Bundle) {
    $Bundle = if ((Split-Path -Leaf $PSCommandPath) -eq 'Verify-RapidPy.ps1') {
        $PSScriptRoot
    } else {
        Join-Path (Split-Path -Parent $PSScriptRoot) 'dist\RapidPyMain'
    }
}
$Bundle = (Get-Item -LiteralPath $Bundle).FullName.TrimEnd('\', '/')
if (-not $EvidenceDirectory) {
    $EvidenceDirectory = Join-Path ([IO.Path]::GetTempPath()) 'RAPID portable verification'
}
$evidenceParent = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($EvidenceDirectory).TrimEnd('\', '/')
$bundlePrefix = $Bundle + [IO.Path]::DirectorySeparatorChar
if ($evidenceParent -eq $Bundle -or $evidenceParent.StartsWith($bundlePrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Verification evidence must be written outside the bundle.'
}
$checkRoot = Join-Path $evidenceParent ('release-check-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $checkRoot -Force | Out-Null
$console = Join-Path $Bundle 'RapidPyMainConsole.exe'
$windows = [Environment]::GetFolderPath([Environment+SpecialFolder]::Windows)
$system = Join-Path $windows 'System32'

function Measure-BundledProcess {
    param([Diagnostics.Process]$Process, [Diagnostics.Stopwatch]$Clock, [string]$ReadyFile, [int]$Timeout)

    $interval = 100
    $samples = 0
    $peakWorkingSet = [long]0
    $peakPrivate = [long]0
    $ready = $null
    while (-not $Process.WaitForExit($interval)) {
        $Process.Refresh()
        if ($Process.HasExited) { break }
        try {
            $workingSet = $Process.WorkingSet64
            $private = $Process.PrivateMemorySize64
        } catch [System.InvalidOperationException] {
            if ($Process.HasExited) { break }
            throw
        }
        $samples++
        $peakWorkingSet = [Math]::Max($peakWorkingSet, $workingSet)
        $peakPrivate = [Math]::Max($peakPrivate, $private)
        if ($null -eq $ready -and [IO.File]::Exists($ReadyFile)) {
            $observedPid = ''
            try { $observedPid = [IO.File]::ReadAllText($ReadyFile) } catch [IO.IOException] { }
            if ($observedPid -ceq [string]$Process.Id) {
                $ready = [ordered]@{
                    observed_launch_to_window_ms = [Math]::Round($Clock.Elapsed.TotalMilliseconds, 3)
                    working_set_bytes = $workingSet
                    private_bytes = $private
                }
            }
        }
        if ($Clock.Elapsed.TotalSeconds -ge $Timeout) { break }
    }
    [ordered]@{
        exited = $Process.HasExited
        sample_interval_ms = $interval
        samples = $samples
        window_ready = $ready
        peak_sampled_working_set_bytes = $peakWorkingSet
        peak_sampled_private_bytes = $peakPrivate
        logical_processors = [Environment]::ProcessorCount
    }
}

function Invoke-BundledCheck {
    param([string]$Name, [string]$Argument, [string]$Scale = '', [switch]$SamplePerformance)

    $profile = Join-Path $checkRoot $Name
    $roaming = Join-Path $profile 'AppData\Roaming'
    $local = Join-Path $profile 'AppData\Local'
    $scratch = Join-Path $local 'Temp'
    New-Item -ItemType Directory -Path $roaming, $scratch -Force | Out-Null
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $console
    $start.Arguments = $Argument
    $start.WorkingDirectory = $profile
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.StandardOutputEncoding = [Text.Encoding]::UTF8
    $start.StandardErrorEncoding = [Text.Encoding]::UTF8
    $start.EnvironmentVariables.Clear()
    $environment = @{
        SystemRoot = $windows
        WINDIR = $windows
        ComSpec = (Join-Path $system 'cmd.exe')
        PATH = "$system;$windows"
        TEMP = $scratch
        TMP = $scratch
        USERPROFILE = $profile
        HOME = $profile
        APPDATA = $roaming
        LOCALAPPDATA = $local
        RAPID_CONFIG = (Join-Path $profile 'config.json')
        RAPID_SAFETY_STATE = (Join-Path $profile 'hardware-safety.json')
        QT_QPA_PLATFORM = 'offscreen'
    }
    if ($Scale) { $environment.QT_SCALE_FACTOR = $Scale }
    $readyFile = Join-Path $profile 'window-ready'
    if ($SamplePerformance) { $environment.RAPID_PERFORMANCE_READY_FILE = $readyFile }
    foreach ($entry in $environment.GetEnumerator()) {
        $start.EnvironmentVariables[$entry.Key] = $entry.Value
    }
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $process = [Diagnostics.Process]::Start($start)
    try {
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $processSample = $null
        $limit = $TimeoutSeconds
        if ($SamplePerformance) {
            $limit = $PerformanceTimeoutSeconds
            $processSample = Measure-BundledProcess -Process $process -Clock $clock -ReadyFile $readyFile -Timeout $limit
            $exited = $processSample.exited
        } else {
            $exited = $process.WaitForExit($limit * 1000)
        }
        if (-not $exited) {
            if (-not $process.HasExited) { $process.Kill() }
            $process.WaitForExit()
        }
        $json = $stdout.GetAwaiter().GetResult()
        [IO.File]::WriteAllText((Join-Path $checkRoot "$Name.stdout.log"), $json)
        [IO.File]::WriteAllText((Join-Path $checkRoot "$Name.stderr.log"), $stderr.GetAwaiter().GetResult())
        if (-not $exited) { throw "Bundled $Name check timed out after $limit seconds." }
        if ($process.ExitCode -ne 0) { throw "Bundled $Name check failed with exit code $($process.ExitCode)." }
        $result = $json | ConvertFrom-Json
        if ($SamplePerformance) {
            if ($null -eq $processSample.window_ready -or $processSample.samples -lt 2) {
                throw 'The performance run did not yield enough live process samples.'
            }
            $result | Add-Member -NotePropertyName windows_process -NotePropertyValue $processSample
        }
        return $result
    } finally {
        try {
            if (-not $process.HasExited) {
                $process.Kill()
                $process.WaitForExit()
            }
        } finally {
        try {
            if (-not $process.HasExited) {
                $process.Kill()
                $process.WaitForExit()
            }
        } finally {
            $process.Dispose()
        }
        }
    }
}

$reportPath = Join-Path $checkRoot 'verification.json'
$report = [ordered]@{
    schema = 'rapidpy.portable_verification.v2'
    checked_at_utc = [DateTimeOffset]::UtcNow.ToString('o')
    scope = 'offline startup and UI with an isolated profile and Windows-only PATH'
    os_version = [Environment]::OSVersion.VersionString
    executable = $console
    passed = $false
}
try {
    $report.inventory = Test-RapidPyBundle -Bundle $Bundle -RequireFullInventory:$RequireFullInventory
    $report.executable_sha256 = Get-RapidPyFileHash -Path $console
    $report.startup = Invoke-BundledCheck -Name 'startup' -Argument '--check-startup'
    $startup = $report.startup
    if ($startup.schema -cne 'rapidpy.startup_environment.v1' -or $startup.ok -ne $true -or
        $startup.packaged -ne $true -or $startup.config_exists -ne $false -or
        $startup.config_path -ne (Join-Path $checkRoot 'startup\config.json') -or
        $startup.python_executable -ne $console -or
        -not $startup.icon_path.StartsWith($bundlePrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Startup evidence did not identify the bundled executable, assets and isolated configuration.'
    }

    $report.smoke = Invoke-BundledCheck -Name 'smoke' -Argument '--smoke-test'
    $smoke = $report.smoke
    if ($smoke.schema -cne 'rapidpy.release_smoke.v1' -or $smoke.ok -ne $true -or
        $smoke.simulated -ne $true -or $smoke.panels -lt 6) {
        throw 'UI checks did not verify a packaged No-Comm session.'
    }
    if ($report.inventory.source_revision -and $smoke.PSObject.Properties['software_version'] -and
        -not ([string]$smoke.software_version).EndsWith("+$($report.inventory.source_revision)", [StringComparison]::Ordinal)) {
        throw 'The packaged runtime version does not identify the candidate commit.'
    }
    foreach ($helper in @('af_tuner', 'data_viewer', 'gaussmeter_control', 'updown_control', 'vrm_logger', 'webcam_viewer')) {
        if ($smoke.helpers.$helper -ne $true) { throw "UI checks did not verify packaged helper: $helper" }
    }
    if ($report.inventory.full_inventory -and $smoke.transverse_probe -ne $true) {
        throw 'UI checks did not exercise the transverse probe scan.'
    }
    if ($smoke.compact_display_check -eq $true) {
        $report.compact_display = [ordered]@{}
        foreach ($profile in @(@{ Percent = 150; Factor = '1.5' }, @{ Percent = 200; Factor = '2' })) {
            $percent = $profile.Percent
            $display = Invoke-BundledCheck -Name "display-$percent" -Argument '--check-display' -Scale $profile.Factor
            $report.compact_display["scale_$percent"] = $display
            if ($display.schema -cne 'rapidpy.display_check.v1' -or $display.ok -ne $true -or
                $display.device_pixel_ratio -ne ($percent / 100.0) -or $display.registry_controls_visible -ne $true -or
                $display.dialog_in_work_area -ne $true -or $display.maximized_preserved -ne $true) {
                throw "The compact display check failed at ${percent}% scaling."
            }
        }
    }
    if ($MeasurePerformance) {
        $arguments = '--measure-performance --performance-specimens {0} --performance-idle-seconds {1} --performance-queue-timeout-seconds {2}' -f `
            $PerformanceSpecimens, $PerformanceIdleSeconds, $PerformanceTimeoutSeconds
        $performance = Invoke-BundledCheck -Name 'performance' -Argument $arguments -SamplePerformance
        $report.performance = $performance
        if ($performance.schema -cne 'rapidpy.performance_check.v1' -or $performance.ok -ne $true -or
            $performance.simulated -ne $true -or $performance.software_version -cne $smoke.software_version -or
            $performance.idle_seconds_requested -ne $PerformanceIdleSeconds -or
            $performance.idle.elapsed_ms -lt ($PerformanceIdleSeconds * 1000) -or
            $performance.queue.requested_specimens -ne $PerformanceSpecimens -or
            $performance.queue.published_specimens -ne $PerformanceSpecimens) {
            throw 'The performance report did not identify a completed workload for this candidate.'
        }
    }
    $after = Test-RapidPyBundle -Bundle $Bundle -RequireFullInventory:$RequireFullInventory
    if ($after.manifest_sha256 -cne $report.inventory.manifest_sha256) {
        throw 'The portable build manifest changed during verification.'
    }
    $report.passed = $true
    Write-Output "Portable qualification checks passed. Evidence: $reportPath"
} catch {
    $report.error = $_.Exception.Message
    throw "Portable verification failed. $($report.error) Evidence: $reportPath"
} finally {
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding utf8
}
