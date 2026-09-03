<#
.SYNOPSIS
    Compile a VB6 project to an EXE on this computer.

.DESCRIPTION
    Runs VB6.EXE /make elevated, because the IDE needs write access to the
    machine COM hive while it loads the project's licensed controls. See
    Start-VB6.ps1 for the full explanation of that requirement.

    Before compiling, the script writes a machine-local copy of the .vbp beside
    the original (<name>.localbuild.vbp) and applies fix-ups for references
    whose requested type-library version is not the one registered on this
    computer. The committed project file is never modified: what shipped in the
    repository stays the canonical, cross-machine definition, and every
    substitution is printed and recorded in the build receipt.

.PARAMETER Project
    Project to build. Defaults to the RAPID Paleomag project beside this script.

.PARAMETER OutDir
    Where the EXE is written. Defaults to <repo>\build\vb6.

.PARAMETER NoFixups
    Compile the project exactly as committed, with no reference substitution.

.PARAMETER NoElevate
    Do not self-elevate (the compile will usually fail; useful for diagnosis).

.PARAMETER PlanOnly
    Report the reference substitutions and unresolved components, then stop
    without compiling. Needs no elevation.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\VB6\Build-VB6Project.ps1
#>
[CmdletBinding()]
param(
    [string]$Project,
    [string]$OutDir,
    [switch]$NoFixups,
    [switch]$NoElevate,
    [switch]$PlanOnly,
    [switch]$KeepLocalProject,
    [int]$TimeoutSeconds = 300
)

$ErrorActionPreference = 'Stop'

$ideCandidates = @(
    'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE',
    'C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE',
    'C:\VB6\VB6.EXE'
)

function Test-Elevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Get-VB6EditionState {
    <#
        VB6 decides what it is allowed to do from the registration its setup
        program wrote. When the VB98 folder has simply been copied onto a
        machine, the binaries all run but the edition was never registered, and
        the compiler answers "No make available in the Working Model Edition".

        This reports the observable facts rather than guessing at a SKU.
    #>
    param([string]$IdePath)

    # Decisive: setup registered a ProductDir for Visual Basic, and the native
    # toolchain is on disk. Verified on this machine - a build succeeded with
    # exactly these two true and the softer markers below still absent.
    $blockers = [System.Collections.Generic.List[string]]::new()
    $notes = [System.Collections.Generic.List[string]]::new()

    $setupKey = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic'
    $productDir = $null
    if (Test-Path -LiteralPath $setupKey) {
        $productDir = (Get-ItemProperty -LiteralPath $setupKey -ErrorAction SilentlyContinue).ProductDir
    }
    if (-not $productDir) {
        $blockers.Add('Setup key "Microsoft Visual Basic" has no ProductDir: VB6 setup never ran, so /make will answer "No make available in the Working Model Edition".')
    }

    $linker = Join-Path (Split-Path -Parent $IdePath) 'LINK.EXE'
    $compiler = Join-Path (Split-Path -Parent $IdePath) 'C2.EXE'
    $toolchain = (Test-Path -LiteralPath $linker -PathType Leaf) -and (Test-Path -LiteralPath $compiler -PathType Leaf)
    if (-not $toolchain) {
        $blockers.Add('LINK.EXE / C2.EXE are missing: the native compiler toolchain is not installed.')
    }

    # Informational only. A working VS6 Enterprise install on this machine has
    # neither of these, so they must never gate a build.
    if (-not (Test-Path -LiteralPath 'HKCU:\SOFTWARE\Microsoft\VisualStudio\6.0')) {
        $notes.Add('no HKCU\SOFTWARE\Microsoft\VisualStudio\6.0 (per-user IDE key)')
    }
    $installed = $false
    foreach ($root in @(
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )) {
        foreach ($entry in (Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue)) {
            $name = (Get-ItemProperty -LiteralPath $entry.PSPath -ErrorAction SilentlyContinue).DisplayName
            if ($name -and ($name -match 'Visual Basic 6|Visual Studio 6\.0')) { $installed = $true }
        }
    }
    if (-not $installed) { $notes.Add('no Visual Studio 6.0 uninstall entry') }

    $canCompile = ($blockers.Count -eq 0)
    if ($canCompile) {
        $summary = 'registered, compiler available'
        if ($notes.Count -gt 0) { $summary += ' (' + ($notes -join '; ') + ')' }
    }
    else {
        $summary = 'not registered for compiling'
    }

    return [pscustomobject]@{
        CanCompile = $canCompile
        Summary    = $summary
        ProductDir = $productDir
        Toolchain  = $toolchain
        Installed  = $installed
        Reasons    = @($blockers)
        Notes      = @($notes)
    }
}

function Get-IdePath {
    foreach ($candidate in $ideCandidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return $null
}

function Get-RegisteredTypeLibVersions {
    <#
        Return the type-library versions for a GUID that are actually backed by
        a file on disk. A bare version key with no win32 payload is a leftover
        stub and will not satisfy the IDE.
    #>
    param([string]$Guid)

    $result = [System.Collections.Generic.List[object]]::new()
    $roots = @(
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\TypeLib',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node\TypeLib',
        'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\TypeLib'
    )
    foreach ($root in $roots) {
        $base = Join-Path $root $Guid
        if (-not (Test-Path -LiteralPath $base)) { continue }
        foreach ($versionKey in (Get-ChildItem -LiteralPath $base -ErrorAction SilentlyContinue)) {
            $version = $versionKey.PSChildName
            # VB6 is 32-bit, so only a win32 payload counts.
            $payload = ''
            $archKey = Join-Path $versionKey.PSPath '0\win32'
            if (Test-Path -LiteralPath $archKey) {
                $payload = (Get-ItemProperty -LiteralPath $archKey -ErrorAction SilentlyContinue).'(default)'
            }
            if (-not $payload) { continue }

            # A type-library path can carry a trailing resource index, as in
            # "C:\Windows\SysWOW64\vbscript.dll\3". Strip it before testing the file.
            $file = $payload
            if ($file -match '^(.*\.[A-Za-z0-9_]+)\\[0-9]+$') { $file = $Matches[1] }
            if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { continue }
            $result.Add([pscustomobject]@{ Version = $version; File = $file; Payload = $payload })
        }
    }
    return $result | Sort-Object -Property @{ Expression = { [version]('0.0.0.0'.Substring(0, 0) + $_.Version + '.0') } } -Descending -ErrorAction SilentlyContinue
}

function Compare-TypeLibVersion {
    param([string]$Left, [string]$Right)
    $l = 0.0; $r = 0.0
    [double]::TryParse($Left, [ref]$l) | Out-Null
    [double]::TryParse($Right, [ref]$r) | Out-Null
    if ($l -gt $r) { return 1 }
    if ($l -lt $r) { return -1 }
    return 0
}

function New-LocalProject {
    <#
        Copy the .vbp beside the original with machine-specific reference
        substitutions, and return the path plus the list of changes made.
    #>
    param([string]$SourceProject, [bool]$ApplyFixups)

    $dir = Split-Path -Parent $SourceProject
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($SourceProject)
    $target = Join-Path $dir ($stem + '.localbuild.vbp')

    $lines = [System.IO.File]::ReadAllLines($SourceProject)
    $changes = [System.Collections.Generic.List[object]]::new()
    $missing = [System.Collections.Generic.List[object]]::new()
    $output = [System.Collections.Generic.List[string]]::new()

    foreach ($line in $lines) {
        $newLine = $line

        # Object={GUID}#major.minor#0; FILE.OCX
        if ($line -match '^Object=\{([0-9A-Fa-f\-]+)\}#([0-9]+\.[0-9]+)#([0-9]+);\s*(.+)$') {
            $guid = '{' + $Matches[1] + '}'
            $requested = $Matches[2]
            $flag = $Matches[3]
            $file = $Matches[4]
            $available = @(Get-RegisteredTypeLibVersions -Guid $guid)
            if (@($available).Count -eq 0) {
                $missing.Add([pscustomobject]@{ Kind = 'Object'; File = $file; Guid = $guid; Requested = $requested })
            }
            elseif (-not ($available.Version -contains $requested)) {
                $best = $available[0].Version
                if ($ApplyFixups) {
                    $newLine = "Object={0}#{1}#{2}; {3}" -f $guid, $best, $flag, $file
                    $changes.Add([pscustomobject]@{
                        Kind = 'Object'; File = $file; Guid = $guid
                        Requested = $requested; Used = $best; Payload = $available[0].File
                    })
                }
                else {
                    $missing.Add([pscustomobject]@{ Kind = 'Object'; File = $file; Guid = $guid; Requested = $requested })
                }
            }
        }
        # Reference=*\G{GUID}#major.minor#0#path#description
        elseif ($line -match '^Reference=\*\\G\{([0-9A-Fa-f\-]+)\}#([0-9]+\.[0-9]+)#([0-9]+)#([^#]*)#(.*)$') {
            $guid = '{' + $Matches[1] + '}'
            $requested = $Matches[2]
            $flag = $Matches[3]
            $refPath = $Matches[4]
            $description = $Matches[5]
            $available = @(Get-RegisteredTypeLibVersions -Guid $guid)
            if (@($available).Count -eq 0) {
                $missing.Add([pscustomobject]@{
                    Kind = 'Reference'; File = [System.IO.Path]::GetFileName($refPath)
                    Guid = $guid; Requested = $requested; Description = $description
                })
            }
            elseif (-not ($available.Version -contains $requested)) {
                $best = $available[0].Version
                if ($ApplyFixups) {
                    $newLine = "Reference=*\G{0}#{1}#{2}#{3}#{4}" -f $guid, $best, $flag, $refPath, $description
                    $changes.Add([pscustomobject]@{
                        Kind = 'Reference'; File = [System.IO.Path]::GetFileName($refPath); Guid = $guid
                        Requested = $requested; Used = $best; Payload = $available[0].File
                    })
                }
                else {
                    $missing.Add([pscustomobject]@{
                        Kind = 'Reference'; File = [System.IO.Path]::GetFileName($refPath)
                        Guid = $guid; Requested = $requested; Description = $description
                    })
                }
            }
        }

        $output.Add($newLine)
    }

    [System.IO.File]::WriteAllLines($target, $output)
    return [pscustomobject]@{ Path = $target; Changes = $changes; Missing = $missing }
}

# ------------------------------------------------------------------ resolve
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $Project) { $Project = Join-Path $PSScriptRoot 'Paleomag v3.vbp' }
if (-not (Test-Path -LiteralPath $Project -PathType Leaf)) {
    throw "Project not found: $Project"
}
$Project = (Resolve-Path -LiteralPath $Project).Path
if (-not $OutDir) { $OutDir = Join-Path $repoRoot 'build\vb6' }
if (-not (Test-Path -LiteralPath $OutDir -PathType Container)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

$idePath = Get-IdePath
$elevated = Test-Elevated

Write-Host 'RAPID VB6 build'
Write-Host ("  IDE:      {0}" -f $(if ($idePath) { $idePath } else { 'NOT FOUND' }))
Write-Host ("  Project:  {0}" -f $Project)
Write-Host ("  Output:   {0}" -f $OutDir)
Write-Host ("  Elevated: {0}" -f $elevated)
Write-Host ''

if (-not $idePath) {
    Write-Host 'BLOCKED: the VB6 IDE was not found.' -ForegroundColor Red
    exit 1
}

$edition = Get-VB6EditionState -IdePath $idePath
Write-Host ("  Edition:  {0}" -f $edition.Summary)
if (-not $edition.CanCompile) {
    Write-Host ''
    Write-Host 'WARNING: this VB6 cannot produce an EXE yet.' -ForegroundColor Yellow
    foreach ($reason in $edition.Reasons) { Write-Host ("  - " + $reason) }
    Write-Host '  VB6.EXE /make will answer "No make available in the Working Model Edition".'
    Write-Host '  Run the setup program from your licensed VB6 media so the edition is registered.'
    Write-Host ''
}

if ((-not $elevated) -and (-not $NoElevate) -and (-not $PlanOnly)) {
    Write-Host 'Re-launching the build elevated. Approve the UAC prompt.' -ForegroundColor Cyan
    $arguments = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', ('"' + $PSCommandPath + '"'),
        '-Project', ('"' + $Project + '"'),
        '-OutDir', ('"' + $OutDir + '"')
    )
    if ($NoFixups) { $arguments += '-NoFixups' }
    if ($KeepLocalProject) { $arguments += '-KeepLocalProject' }
    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -Verb RunAs -PassThru -Wait
    exit $process.ExitCode
}

# ------------------------------------------------------------------ prepare
$local = New-LocalProject -SourceProject $Project -ApplyFixups (-not $NoFixups)

if (@($local.Changes).Count -gt 0) {
    Write-Host 'Machine-local reference substitutions (committed project unchanged):'
    foreach ($change in $local.Changes) {
        Write-Host ("  {0,-9} {1,-22} {2} -> {3}   [{4}]" -f `
            $change.Kind, $change.File, $change.Requested, $change.Used, $change.Payload)
    }
    Write-Host ''
}

if (@($local.Missing).Count -gt 0) {
    Write-Host 'References with no registered type library on this computer:' -ForegroundColor Yellow
    foreach ($item in $local.Missing) {
        Write-Host ("  {0,-9} {1,-22} wants {2}  {3}" -f $item.Kind, $item.File, $item.Requested, $item.Description)
    }
    Write-Host ''
    Write-Host 'The compile will fail on these until the component is installed and registered.' -ForegroundColor Yellow
    Write-Host ''
}

if ($PlanOnly) {
    if (-not $KeepLocalProject) {
        Remove-Item -LiteralPath $local.Path -Force -ErrorAction SilentlyContinue
    }
    if (@($local.Missing).Count -gt 0) { exit 1 }
    Write-Host 'Plan is clean: every reference resolves on this computer.' -ForegroundColor Green
    exit 0
}

# ------------------------------------------------------------------ compile
$logPath = Join-Path $OutDir 'vb6-build.log'
if (Test-Path -LiteralPath $logPath) { Remove-Item -LiteralPath $logPath -Force }

$makeArgs = @(
    '/make', ('"' + $local.Path + '"'),
    '/outdir', ('"' + $OutDir + '"'),
    '/out', ('"' + $logPath + '"')
)
# VB6 rewrites the project workspace (.vbw) whenever it opens a project, and a
# /make run truncates it. That file is committed, so snapshot it and put it back.
$workspacePath = [System.IO.Path]::ChangeExtension($Project, '.vbw')
$workspaceBackup = $null
if (Test-Path -LiteralPath $workspacePath -PathType Leaf) {
    $workspaceBackup = [System.IO.File]::ReadAllBytes($workspacePath)
}

Write-Host ("Running: VB6.EXE {0}" -f ($makeArgs -join ' '))
$started = Get-Date
$process = Start-Process -FilePath $idePath -ArgumentList $makeArgs `
    -WorkingDirectory (Split-Path -Parent $Project) -PassThru

# /make normally runs headless, but a component it cannot load makes the IDE
# raise a modal dialog that nobody is there to dismiss. Do not leave an
# elevated VB6 stuck on the desktop.
$timedOut = $false
if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
    $timedOut = $true
    Write-Host ("VB6 did not finish within {0}s - it is most likely showing a modal error dialog." -f $TimeoutSeconds) -ForegroundColor Yellow
    try { $process.Kill() } catch { }
    Start-Sleep -Seconds 1
}
$elapsed = (Get-Date) - $started
$exitCode = -1
if (-not $timedOut) { $exitCode = $process.ExitCode }
Write-Host ("VB6 exited with code {0} after {1:n1}s" -f $exitCode, $elapsed.TotalSeconds)
Write-Host ''

if ($null -ne $workspaceBackup) {
    $current = $null
    if (Test-Path -LiteralPath $workspacePath -PathType Leaf) {
        $current = [System.IO.File]::ReadAllBytes($workspacePath)
    }
    if (($null -eq $current) -or ($current.Length -ne $workspaceBackup.Length)) {
        [System.IO.File]::WriteAllBytes($workspacePath, $workspaceBackup)
        Write-Host ("Restored the committed workspace file that VB6 rewrote: {0}" -f (Split-Path -Leaf $workspacePath))
    }
}

# The build copy leaves its own workspace behind; it is not wanted either.
$localWorkspace = [System.IO.Path]::ChangeExtension($local.Path, '.vbw')
if (Test-Path -LiteralPath $localWorkspace -PathType Leaf) {
    Remove-Item -LiteralPath $localWorkspace -Force -ErrorAction SilentlyContinue
}

$logText = ''
if (Test-Path -LiteralPath $logPath -PathType Leaf) {
    # [string] strips the PSPath/PSDrive note properties Get-Content attaches;
    # without it ConvertTo-Json serialises every drive on the machine.
    $logText = [string]([System.IO.File]::ReadAllText($logPath))
    if ($logText.Trim()) {
        Write-Host '--- VB6 build log ---'
        Write-Host $logText.Trim()
        Write-Host '---------------------'
        Write-Host ''
    }
}

# ------------------------------------------------------------------ verify
$exeName = ''
foreach ($line in [System.IO.File]::ReadLines($Project)) {
    if ($line -like 'ExeName32=*') { $exeName = $line.Substring(10).Trim('"'); break }
}
$exePath = ''
if ($exeName) { $exePath = Join-Path $OutDir $exeName }

$succeeded = $false
if ($exePath -and (Test-Path -LiteralPath $exePath -PathType Leaf)) {
    $exeItem = Get-Item -LiteralPath $exePath
    if ($exeItem.LastWriteTime -ge $started.AddSeconds(-5)) {
        $succeeded = $true
        Write-Host ('BUILT: ' + $exePath) -ForegroundColor Green
        Write-Host ("  size:     {0:n0} bytes" -f $exeItem.Length)
        Write-Host ("  modified: {0}" -f $exeItem.LastWriteTime)
        $info = $exeItem.VersionInfo
        if ($info.FileVersion) { Write-Host ("  version:  {0}" -f $info.FileVersion) }
    }
}

if (-not $succeeded) {
    Write-Host 'BUILD FAILED: no fresh EXE was produced.' -ForegroundColor Red
    if ($logText -match 'Working Model Edition') {
        Write-Host ''
        Write-Host 'Cause: the VB6 compiler is disabled because the edition was never registered.' -ForegroundColor Yellow
        Write-Host 'This is not a project problem. Run the setup program from your licensed VB6'
        Write-Host 'media, then re-run this script. Nothing in the project needs to change.'
    }
    elseif (-not $logText.Trim()) {
        Write-Host 'VB6 wrote no error log, which usually means it could not load the project at all.'
    }
}

# ------------------------------------------------------------------ receipt
$receipt = [pscustomobject]@{
    schema         = 'rapid.vb6.build-receipt.v1'
    builtAtIso     = (Get-Date).ToString('o')
    machine        = $env:COMPUTERNAME
    user           = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    elevated       = Test-Elevated
    ide            = $idePath
    ideVersion     = (Get-Item -LiteralPath $idePath).VersionInfo.FileVersion
    editionSummary = $edition.Summary
    editionCanCompile = $edition.CanCompile
    editionReasons = @($edition.Reasons)
    sourceProject  = $Project
    builtProject   = $local.Path
    outputDir      = $OutDir
    exePath        = $exePath
    succeeded      = $succeeded
    exitCode       = $exitCode
    timedOut       = $timedOut
    substitutions  = @($local.Changes)
    unresolved     = @($local.Missing)
    buildLog       = $logText
}
$receiptPath = Join-Path $OutDir 'vb6-build-receipt.json'
$receipt | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $receiptPath -Encoding UTF8
Write-Host ''
Write-Host ('Receipt: ' + $receiptPath)

if (-not $KeepLocalProject) {
    Remove-Item -LiteralPath $local.Path -Force -ErrorAction SilentlyContinue
}

if ($succeeded) { exit 0 }
exit 1
