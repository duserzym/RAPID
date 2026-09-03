<#
.SYNOPSIS
    Install and register a legacy 32-bit VB6/ActiveX dependency.

.DESCRIPTION
    Copies a legacy COM component from an archive (a vendor package, the
    original licensed media, or the RAPID legacy archive) into the 32-bit
    system directory and registers it with the 32-bit regsvr32.

    The script refuses to touch anything it has not verified first:

      * the source must exist and be a 32-bit (i386) PE image;
      * when -ExpectedTypeLibGuid is supplied, that GUID must actually be
        present in the binary, so a same-named but different component cannot
        be registered by mistake;
      * an existing target file is backed up before it is replaced.

    Registration is machine-wide, so the script elevates. It reports the
    type-library versions that appear afterwards, which is what the VB6
    project reference has to match.

.PARAMETER Source
    The component to install.

.PARAMETER TargetName
    File name to install as. Defaults to the source file name. Use this when a
    project references the component under a different name.

.PARAMETER TargetDir
    Install directory. Defaults to the 32-bit system directory (SysWOW64 on
    64-bit Windows).

.PARAMETER ExpectedTypeLibGuid
    Type-library GUID that must be present in the binary, e.g.
    '{332B82D3-3ED6-11D4-B1B5-00105AA5CCFF}'.

.PARAMETER WhatIfOnly
    Verify the source and report what would happen, changing nothing.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\VB6\Install-VB6Dependency.ps1 `
        -Source 'F:\Paleomag2013\vbSendMail\vbSendMail.dll' `
        -TargetName 'vbSendMail_v3.0.dll' `
        -ExpectedTypeLibGuid '{332B82D3-3ED6-11D4-B1B5-00105AA5CCFF}'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Source,
    [string]$TargetName,
    [string]$TargetDir,
    [string]$ExpectedTypeLibGuid,
    [switch]$WhatIfOnly,
    [switch]$NoElevate
)

$ErrorActionPreference = 'Stop'

function Test-Elevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

function Get-PeMachine {
    param([string]$FilePath)

    $stream = [System.IO.File]::OpenRead($FilePath)
    try {
        $reader = New-Object System.IO.BinaryReader($stream)
        $null = $stream.Seek(0x3C, 'Begin')
        $peOffset = $reader.ReadInt32()
        if ($peOffset -le 0 -or $peOffset -ge $stream.Length) { return 'not-a-pe' }
        $null = $stream.Seek($peOffset, 'Begin')
        if ($reader.ReadUInt32() -ne 0x00004550) { return 'not-a-pe' }
        $machine = $reader.ReadUInt16()
        switch ($machine) {
            0x014C { return 'i386' }
            0x8664 { return 'x64' }
            0x01C4 { return 'arm' }
            0xAA64 { return 'arm64' }
            default { return ('0x{0:X}' -f $machine) }
        }
    }
    finally { $stream.Dispose() }
}

function Test-GuidInBinary {
    param([string]$FilePath, [string]$Guid)

    $parsed = [Guid]::Parse($Guid)
    $needle = $parsed.ToByteArray()   # little-endian COM layout, as stored in a PE
    $bytes = [System.IO.File]::ReadAllBytes($FilePath)
    $limit = $bytes.Length - $needle.Length
    for ($i = 0; $i -le $limit; $i++) {
        if ($bytes[$i] -ne $needle[0]) { continue }
        $match = $true
        for ($j = 1; $j -lt $needle.Length; $j++) {
            if ($bytes[$i + $j] -ne $needle[$j]) { $match = $false; break }
        }
        if ($match) { return $true }
    }
    return $false
}

function Get-TypeLibVersions {
    param([string]$Guid)

    $versions = [System.Collections.Generic.List[string]]::new()
    foreach ($root in @(
        'HKLM:\SOFTWARE\Classes\TypeLib',
        'HKLM:\SOFTWARE\Classes\WOW6432Node\TypeLib'
    )) {
        $base = Join-Path $root $Guid
        if (-not (Test-Path -LiteralPath $base)) { continue }
        foreach ($key in (Get-ChildItem -LiteralPath $base -ErrorAction SilentlyContinue)) {
            $win32 = Join-Path $key.PSPath '0\win32'
            $payload = ''
            if (Test-Path -LiteralPath $win32) {
                $payload = (Get-ItemProperty -LiteralPath $win32 -ErrorAction SilentlyContinue).'(default)'
            }
            $versions.Add(("{0} -> {1}" -f $key.PSChildName, $(if ($payload) { $payload } else { '(no win32 payload)' })))
        }
    }
    return @($versions | Select-Object -Unique)
}

# ------------------------------------------------------------------- verify
if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) {
    throw "Source component not found: $Source"
}
$Source = (Resolve-Path -LiteralPath $Source).Path
if (-not $TargetName) { $TargetName = [System.IO.Path]::GetFileName($Source) }
if (-not $TargetDir) {
    $TargetDir = Join-Path $env:SystemRoot 'SysWOW64'
    if (-not (Test-Path -LiteralPath $TargetDir -PathType Container)) {
        $TargetDir = Join-Path $env:SystemRoot 'System32'
    }
}
$targetPath = Join-Path $TargetDir $TargetName

$sourceInfo = Get-Item -LiteralPath $Source
$machine = Get-PeMachine -FilePath $Source
$versionInfo = $sourceInfo.VersionInfo

Write-Host 'RAPID legacy component install'
Write-Host ("  Source:   {0}" -f $Source)
Write-Host ("  Size:     {0:n0} bytes" -f $sourceInfo.Length)
Write-Host ("  Modified: {0}" -f $sourceInfo.LastWriteTime)
Write-Host ("  Version:  {0}" -f $versionInfo.FileVersion)
Write-Host ("  Company:  {0}" -f $versionInfo.CompanyName)
Write-Host ("  Machine:  {0}" -f $machine)
Write-Host ("  Target:   {0}" -f $targetPath)
Write-Host ''

if ($machine -ne 'i386') {
    Write-Host "REFUSED: VB6 needs a 32-bit component; this image is '$machine'." -ForegroundColor Red
    exit 1
}

if ($ExpectedTypeLibGuid) {
    if (Test-GuidInBinary -FilePath $Source -Guid $ExpectedTypeLibGuid) {
        Write-Host ("  [OK] type-library GUID {0} is present in the binary." -f $ExpectedTypeLibGuid) -ForegroundColor Green
    }
    else {
        Write-Host ("REFUSED: {0} is not present in this binary. Wrong component." -f $ExpectedTypeLibGuid) -ForegroundColor Red
        exit 1
    }
    Write-Host ''
}

if ($WhatIfOnly) {
    Write-Host 'Verification only; nothing was changed.'
    exit 0
}

# ------------------------------------------------------------------ elevate
if ((-not (Test-Elevated)) -and (-not $NoElevate)) {
    Write-Host 'Copying into the system directory and registering a COM server needs elevation.'
    Write-Host 'Re-launching elevated. Approve the UAC prompt.' -ForegroundColor Cyan
    $arguments = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', ('"' + $PSCommandPath + '"'),
        '-Source', ('"' + $Source + '"'),
        '-TargetName', ('"' + $TargetName + '"'),
        '-TargetDir', ('"' + $TargetDir + '"')
    )
    if ($ExpectedTypeLibGuid) { $arguments += @('-ExpectedTypeLibGuid', ('"' + $ExpectedTypeLibGuid + '"')) }
    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $arguments -Verb RunAs -PassThru -Wait
    exit $process.ExitCode
}

# ------------------------------------------------------------------ install
if (Test-Path -LiteralPath $targetPath -PathType Leaf) {
    $backup = $targetPath + '.bak-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
    Copy-Item -LiteralPath $targetPath -Destination $backup -Force
    Write-Host ("  Existing file backed up to {0}" -f $backup)
}

Copy-Item -LiteralPath $Source -Destination $targetPath -Force
Write-Host ("  Copied to {0}" -f $targetPath)

$regsvr = Join-Path $env:SystemRoot 'SysWOW64\regsvr32.exe'
if (-not (Test-Path -LiteralPath $regsvr -PathType Leaf)) {
    $regsvr = Join-Path $env:SystemRoot 'System32\regsvr32.exe'
}
Write-Host ("  Registering with {0}" -f $regsvr)
$proc = Start-Process -FilePath $regsvr -ArgumentList @('/s', ('"' + $targetPath + '"')) -PassThru -Wait
Write-Host ("  regsvr32 exit code: {0}" -f $proc.ExitCode)
Write-Host ''

if ($proc.ExitCode -ne 0) {
    Write-Host 'REGISTRATION FAILED.' -ForegroundColor Red
    Write-Host 'Run without /s to see the dialog, or check that the component''s own dependencies are present.'
    exit 1
}

if ($ExpectedTypeLibGuid) {
    $versions = Get-TypeLibVersions -Guid $ExpectedTypeLibGuid
    if (@($versions).Count -eq 0) {
        Write-Host 'Registered, but no type library appeared under that GUID.' -ForegroundColor Yellow
        exit 1
    }
    Write-Host 'Registered type-library versions:' -ForegroundColor Green
    foreach ($version in $versions) { Write-Host ("  " + $version) }
    Write-Host ''
    Write-Host 'If the project references a different version, Build-VB6Project.ps1 substitutes'
    Write-Host 'the registered one into its machine-local copy of the .vbp automatically.'
}

Write-Host ''
Write-Host 'DONE.' -ForegroundColor Green
