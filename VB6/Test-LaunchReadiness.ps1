[CmdletBinding()]
param(
    [switch]$Launch,
    [switch]$RunAndExit
)

$ErrorActionPreference = 'Stop'

$projectPath = Join-Path $PSScriptRoot 'Paleomag v3.vbp'
$repoRoot = Split-Path -Parent $PSScriptRoot
# Build-VB6Project.ps1 writes to build\vb6, so look there first; the two paths
# beside the project are where a hand-run IDE build lands.
$compiledCandidates = @(
    (Join-Path $repoRoot 'build\vb6\PALEOMAG2013.exe'),
    (Join-Path $PSScriptRoot 'PALEOMAG2013.exe'),
    (Join-Path $PSScriptRoot 'PALEOMAG.exe')
)
$ideCandidates = @(
    'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE',
    'C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE',
    'C:\VB6\VB6.EXE'
)

$runtimePath = 'C:\Windows\SysWOW64\MSVBVM60.DLL'
$dependencyNames = @(
    'MSSTDFMT.DLL',
    'msscript.ocx',
    'vbSendMail_v3.0.dll',
    'MSDERUN.DLL',
    'comdlg32.ocx',
    'comctl32.ocx',
    'mscomm32.ocx',
    'MSFLXGRD.OCX',
    'mshflxgd.ocx',
    'MSCHRT20.OCX',
    'TABCTL32.OCX',
    'comct332.ocx',
    'RICHTX32.OCX',
    'MSCOMCTL.OCX'
)
$typeLibraryIds = @{
    'MSSTDFMT.DLL'      = '{6B263850-900B-11D0-9484-00A0C91110ED}'
    'msscript.ocx'      = '{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}'
    'vbSendMail_v3.0.dll' = '{332B82D3-3ED6-11D4-B1B5-00105AA5CCFF}'
    'MSDERUN.DLL'       = '{3D5C6BF0-69A3-11D0-B393-00A0C9055D8E}'
    'comdlg32.ocx'      = '{F9043C88-F6F2-101A-A3C9-08002B2F49FB}'
    'comctl32.ocx'      = '{6B7E6392-850A-101B-AFC0-4210102A8DA7}'
    'mscomm32.ocx'      = '{648A5603-2C6E-101B-82B6-000000000014}'
    'MSFLXGRD.OCX'      = '{5E9E78A0-531B-11CF-91F6-C2863C385E30}'
    'mshflxgd.ocx'      = '{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}'
    'MSCHRT20.OCX'      = '{65E121D4-0C60-11D2-A9FC-0000F8754DA1}'
    'TABCTL32.OCX'      = '{BDC217C8-ED16-11CD-956C-0000C04E4C0A}'
    'comct332.ocx'      = '{38911DA0-E448-11D0-84A3-00DD01104159}'
    'RICHTX32.OCX'      = '{3B7C8863-D78F-101B-B9B5-04021C009402}'
    'MSCOMCTL.OCX'      = '{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}'
}
$dependencyRoots = @(
    'C:\Windows\SysWOW64',
    'C:\Windows\System32',
    'C:\Program Files (x86)\Common Files\designer',
    'C:\Program Files\Common Files\designer'
)

function Find-FirstExistingPath {
    param([string[]]$Candidates)

    foreach ($candidate in $Candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }
    return $null
}

function Find-Dependency {
    param([string]$Name)

    foreach ($root in $dependencyRoots) {
        $candidate = Join-Path $root $Name
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }
    return $null
}

function Test-Elevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Test-MachineComHiveWritable {
    <#
        VB6 writes the project's own type library and the licence keys for
        licensed controls into the machine COM hive while a project loads. A
        standard-user token cannot, and the IDE reports
        "Error accessing the system registry".
    #>
    foreach ($sub in @('SOFTWARE\Classes', 'SOFTWARE\Classes\Licenses', 'SOFTWARE\WOW6432Node\Classes\CLSID')) {
        try {
            $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($sub, $true)
            if ($null -eq $key) { continue }
            $key.Close()
        }
        catch { return $false }
    }
    return $true
}

function Get-CompilerRegistrationState {
    <#
        Report whether VB6 was installed by its setup program. A copied VB98
        folder runs the IDE but leaves the edition unregistered, and /make then
        answers "No make available in the Working Model Edition".
    #>
    $reasons = [System.Collections.Generic.List[string]]::new()

    # A ProductDir plus the native toolchain is what actually enables /make.
    # The per-user key and the uninstall entry are not reliable: a VS6
    # Enterprise install that compiles fine on this machine has neither.
    $setupKey = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic'
    $productDir = $null
    if (Test-Path -LiteralPath $setupKey) {
        $productDir = (Get-ItemProperty -LiteralPath $setupKey -ErrorAction SilentlyContinue).ProductDir
    }
    if (-not $productDir) { $reasons.Add('setup key has no ProductDir (VB6 setup never ran)') }

    $vbDir = $productDir
    if (-not $vbDir) { $vbDir = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98' }
    foreach ($tool in @('LINK.EXE', 'C2.EXE')) {
        if (-not (Test-Path -LiteralPath (Join-Path $vbDir $tool) -PathType Leaf)) {
            $reasons.Add("$tool is missing from $vbDir")
        }
    }

    return [pscustomobject]@{ Registered = ($reasons.Count -eq 0); Reasons = @($reasons) }
}

function Test-TypeLibraryRegistration {
    param([string]$TypeLibraryId)

    $roots = @(
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\WOW6432Node\TypeLib',
        'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WOW6432Node\TypeLib',
        'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\TypeLib',
        'Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\TypeLib'
    )
    foreach ($root in $roots) {
        if (Test-Path -LiteralPath (Join-Path $root $TypeLibraryId)) {
            return $true
        }
    }
    return $false
}

$compiledPath = Find-FirstExistingPath -Candidates $compiledCandidates
$idePath = Find-FirstExistingPath -Candidates $ideCandidates
$missingDependencies = [System.Collections.Generic.List[string]]::new()
$unregisteredDependencies = [System.Collections.Generic.List[string]]::new()

Write-Host 'RAPID VB6 launch readiness'
Write-Host "Project:  $projectPath"
Write-Host "Runtime:  $runtimePath"
Write-Host "Compiled: $(if ($compiledPath) { $compiledPath } else { 'not found' })"
Write-Host "VB6 IDE:  $(if ($idePath) { $idePath } else { 'not found' })"

$elevated = Test-Elevated
$comHiveWritable = Test-MachineComHiveWritable
$compilerState = Get-CompilerRegistrationState
Write-Host "Elevated: $elevated"
Write-Host "COM hive: $(if ($comHiveWritable) { 'writable (design-time registration will succeed)' } else { 'NOT writable by this token' })"
Write-Host "Compiler: $(if ($compilerState.Registered) { 'edition registered' } else { 'edition NOT registered (' + ($compilerState.Reasons -join '; ') + ')' })"
Write-Host ''

foreach ($name in $dependencyNames) {
    $path = Find-Dependency -Name $name
    $registered = Test-TypeLibraryRegistration -TypeLibraryId $typeLibraryIds[$name]
    if ($path -and $registered) {
        Write-Host ("[OK]           {0} -> {1}" -f $name, $path)
    }
    elseif (-not $path) {
        $missingDependencies.Add($name)
        Write-Host ("[MISSING FILE] {0}" -f $name)
    }
    else {
        $unregisteredDependencies.Add($name)
        Write-Host ("[UNREGISTERED] {0} -> {1}" -f $name, $path)
    }
}

$blockers = [System.Collections.Generic.List[string]]::new()
if (-not (Test-Path -LiteralPath $projectPath -PathType Leaf)) {
    $blockers.Add("Project file is missing: $projectPath")
}
if (-not (Test-Path -LiteralPath $runtimePath -PathType Leaf)) {
    $blockers.Add("32-bit VB6 runtime is missing: $runtimePath")
}
if (-not $compiledPath -and -not $idePath) {
    $blockers.Add('Neither PALEOMAG2013.exe nor the VB6 IDE/compiler is installed.')
}
if ($missingDependencies.Count -gt 0) {
    $blockers.Add("Missing VB6/ActiveX dependencies: $($missingDependencies -join ', ')")
}
if ($unregisteredDependencies.Count -gt 0) {
    $blockers.Add("Unregistered 32-bit VB6/ActiveX dependencies: $($unregisteredDependencies -join ', ')")
}

$warnings = [System.Collections.Generic.List[string]]::new()
if (-not $comHiveWritable) {
    $warnings.Add('This token cannot write the machine COM hive, so opening the project in the IDE will fail with "Error accessing the system registry". Use VB6\Start-VB6.ps1, which runs the IDE elevated.')
}
if (-not $compilerState.Registered) {
    $warnings.Add("The VB6 compiler is not enabled ($($compilerState.Reasons -join '; ')), so /make answers 'No make available in the Working Model Edition'. Run the setup program from your licensed VB6 media.")
}

Write-Host ''
if ($blockers.Count -gt 0) {
    Write-Host 'BLOCKED: Paleomag cannot launch on this computer yet.' -ForegroundColor Red
    foreach ($blocker in $blockers) {
        Write-Host " - $blocker"
    }
    if ($warnings.Count -gt 0) {
        Write-Host ''
        Write-Host 'Environment warnings:' -ForegroundColor Yellow
        foreach ($warning in $warnings) { Write-Host " - $warning" }
    }
    Write-Host ''
    Write-Host 'Install the VB6 IDE (to run source) or copy a known-good PALEOMAG2013.exe, then install/register the missing 32-bit dependencies with the 32-bit regsvr32 at C:\Windows\SysWOW64\regsvr32.exe.'
    exit 1
}

if ($warnings.Count -gt 0) {
    Write-Host 'Environment warnings:' -ForegroundColor Yellow
    foreach ($warning in $warnings) { Write-Host " - $warning" }
    Write-Host ''
}

Write-Host 'READY: Required launch files are present.' -ForegroundColor Green
if (-not $Launch) {
    Write-Host 'Re-run with -Launch to start Paleomag.'
    exit 0
}

if ($compiledPath) {
    Start-Process -FilePath $compiledPath -WorkingDirectory $PSScriptRoot
    exit 0
}

$arguments = @()
if ($RunAndExit) {
    $arguments += '/runexit'
}
$arguments += $projectPath
Start-Process -FilePath $idePath -ArgumentList $arguments -WorkingDirectory $PSScriptRoot
