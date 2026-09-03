<#
.SYNOPSIS
    Open a VB6 project without "Error accessing the system registry".

.DESCRIPTION
    On Windows Vista and later the VB6 IDE needs write access to the machine
    COM registration hive at design time. It registers the project's own type
    library and the licence keys for licensed controls (MSCOMCTL, MSCOMM32,
    MSFLXGRD and friends) under:

        HKLM\SOFTWARE\Classes
        HKLM\SOFTWARE\Classes\Licenses
        HKLM\SOFTWARE\WOW6432Node\Classes

    A standard user token cannot write there, so VB6 reports
    "Error accessing the system registry" while loading the project's controls.

    The supported fix is to run the IDE elevated. This script does that: it
    diagnoses the access, self-elevates through the normal UAC consent prompt,
    and starts VB6 on the project you asked for.

    It deliberately does NOT change the permissions on HKLM\SOFTWARE\Classes.
    Granting a normal user write access to the machine COM hive would let any
    process running as that user hijack COM registrations for every account on
    the computer. Elevation gets the same result for the few minutes the IDE
    needs it, and gives it back when you close the IDE.

.PARAMETER Project
    Project path, or a wildcard matched against the cache written by
    Find-VB6Projects.ps1. Defaults to the RAPID Paleomag project beside this
    script.

.PARAMETER Diagnose
    Report the registry access state and exit without launching anything.

.PARAMETER Shortcut
    Create a desktop shortcut that opens the project with the elevation flag
    already set, so a double-click gives one UAC prompt and a working IDE.

.PARAMETER NoElevate
    Do not self-elevate. Useful to demonstrate the failure, or when the IDE is
    already running elevated.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\VB6\Start-VB6.ps1 -Diagnose

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\VB6\Start-VB6.ps1
#>
[CmdletBinding()]
param(
    [string]$Project,
    [switch]$Diagnose,
    [switch]$Shortcut,
    [switch]$NoElevate,
    [switch]$RunAndExit
)

$ErrorActionPreference = 'Stop'

$ideCandidates = @(
    'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE',
    'C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE',
    'C:\VB6\VB6.EXE'
)

# The keys VB6 opens for write while loading a project's licensed controls.
$requiredWriteKeys = @(
    @{ Hive = 'HKLM'; Sub = 'SOFTWARE\Classes' },
    @{ Hive = 'HKLM'; Sub = 'SOFTWARE\Classes\CLSID' },
    @{ Hive = 'HKLM'; Sub = 'SOFTWARE\Classes\Licenses' },
    @{ Hive = 'HKLM'; Sub = 'SOFTWARE\WOW6432Node\Classes\CLSID' }
)

function Test-Elevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Test-KeyWritable {
    param([string]$Hive, [string]$Sub)

    $root = switch ($Hive) {
        'HKLM' { [Microsoft.Win32.Registry]::LocalMachine }
        'HKCU' { [Microsoft.Win32.Registry]::CurrentUser }
        'HKCR' { [Microsoft.Win32.Registry]::ClassesRoot }
        default { $null }
    }
    if ($null -eq $root) { return 'unknown-hive' }

    try {
        $key = $root.OpenSubKey($Sub, $true)
        if ($null -eq $key) { return 'missing' }
        $key.Close()
        return 'writable'
    }
    catch [System.Security.SecurityException] { return 'denied' }
    catch { return 'error' }
}

function Get-IdePath {
    foreach ($candidate in $ideCandidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return $null
}

function Resolve-ProjectPath {
    param([string]$Requested)

    if ($Requested -and (Test-Path -LiteralPath $Requested -PathType Leaf)) {
        return (Resolve-Path -LiteralPath $Requested).Path
    }

    $default = Join-Path $PSScriptRoot 'Paleomag v3.vbp'
    if (-not $Requested) {
        if (Test-Path -LiteralPath $default -PathType Leaf) { return $default }
    }

    $finder = Join-Path $PSScriptRoot 'Find-VB6Projects.ps1'
    if (Test-Path -LiteralPath $finder -PathType Leaf) {
        $matches = @(& $finder -AsObject -Name $Requested)
        if (@($matches).Count -ge 1) { return $matches[0].FullName }
    }

    if (Test-Path -LiteralPath $default -PathType Leaf) { return $default }
    throw "No VB6 project matched '$Requested'. Run Find-VB6Projects.ps1 to list what is on this computer."
}

function Set-ShortcutRunAsAdmin {
    param([string]$LinkPath)

    # Byte 21 of a .lnk header carries the flags; 0x20 is "run as administrator".
    $bytes = [System.IO.File]::ReadAllBytes($LinkPath)
    $bytes[0x15] = $bytes[0x15] -bor 0x20
    [System.IO.File]::WriteAllBytes($LinkPath, $bytes)
}

# ------------------------------------------------------------------ diagnose
$idePath = Get-IdePath
$elevated = Test-Elevated

Write-Host 'RAPID VB6 launcher'
Write-Host ("  IDE:      {0}" -f $(if ($idePath) { $idePath } else { 'NOT FOUND' }))
Write-Host ("  User:     {0}" -f [Security.Principal.WindowsIdentity]::GetCurrent().Name)
Write-Host ("  Elevated: {0}" -f $elevated)
Write-Host ''
Write-Host 'Machine COM hive write access (what VB6 needs at design time):'

$denied = [System.Collections.Generic.List[string]]::new()
foreach ($entry in $requiredWriteKeys) {
    $state = Test-KeyWritable -Hive $entry.Hive -Sub $entry.Sub
    $label = "$($entry.Hive)\$($entry.Sub)"
    switch ($state) {
        'writable' { Write-Host ("  [OK]      {0}" -f $label) }
        'missing'  { Write-Host ("  [ABSENT]  {0}" -f $label) }
        default {
            Write-Host ("  [DENIED]  {0}" -f $label) -ForegroundColor Yellow
            $denied.Add($label)
        }
    }
}

Write-Host ''
if ($denied.Count -gt 0 -and -not $elevated) {
    Write-Host 'Diagnosis: this token cannot write the machine COM hive, which is exactly' -ForegroundColor Yellow
    Write-Host 'what produces "Error accessing the system registry" when the project loads.' -ForegroundColor Yellow
    Write-Host 'Fix: run the IDE elevated (this script does that for you).'
}
elseif ($elevated) {
    Write-Host 'Diagnosis: elevated token, the machine COM hive is writable. VB6 will not' -ForegroundColor Green
    Write-Host 'raise "Error accessing the system registry" from this session.' -ForegroundColor Green
}
else {
    Write-Host 'Diagnosis: the machine COM hive is already writable by this token.'
}

if ($Diagnose) { exit 0 }

if (-not $idePath) {
    Write-Host ''
    Write-Host 'BLOCKED: the VB6 IDE was not found in any known location.' -ForegroundColor Red
    exit 1
}

$projectPath = Resolve-ProjectPath -Requested $Project
$projectDir = Split-Path -Parent $projectPath
Write-Host ''
Write-Host ("Project:  {0}" -f $projectPath)

# ------------------------------------------------------------------ shortcut
if ($Shortcut) {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $linkPath = Join-Path $desktop ('VB6 - ' + [System.IO.Path]::GetFileNameWithoutExtension($projectPath) + '.lnk')
    $shell = New-Object -ComObject WScript.Shell
    $link = $shell.CreateShortcut($linkPath)
    $link.TargetPath = $idePath
    $link.Arguments = '"' + $projectPath + '"'
    $link.WorkingDirectory = $projectDir
    $link.IconLocation = $idePath + ',0'
    $link.Description = 'Open ' + [System.IO.Path]::GetFileName($projectPath) + ' in VB6 as administrator'
    $link.Save()
    Set-ShortcutRunAsAdmin -LinkPath $linkPath
    Write-Host ("Shortcut: {0}  (elevation flag set)" -f $linkPath) -ForegroundColor Green
    exit 0
}

# ------------------------------------------------------------------ elevate
if ((-not $elevated) -and (-not $NoElevate)) {
    Write-Host ''
    Write-Host 'Re-launching this script elevated. Approve the UAC prompt.' -ForegroundColor Cyan
    $arguments = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', ('"' + $PSCommandPath + '"'),
        '-Project', ('"' + $projectPath + '"')
    )
    if ($RunAndExit) { $arguments += '-RunAndExit' }
    try {
        Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -Verb RunAs | Out-Null
        exit 0
    }
    catch {
        Write-Host 'UAC consent was declined; VB6 was not started.' -ForegroundColor Red
        Write-Host 'Without elevation the project will fail with "Error accessing the system registry".'
        exit 1
    }
}

# ------------------------------------------------------------------ launch
$vbArgs = @()
if ($RunAndExit) { $vbArgs += '/runexit' }
$vbArgs += ('"' + $projectPath + '"')

Write-Host 'Starting VB6...'
Start-Process -FilePath $idePath -ArgumentList $vbArgs -WorkingDirectory $projectDir
Write-Host 'VB6 started.' -ForegroundColor Green
