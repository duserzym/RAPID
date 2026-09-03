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

    After replacing an OCX it also deletes the sibling .oca type-information
    cache, which otherwise still describes the control that was there before.

    It also reads the type-library version embedded in the component before
    installing anything, so you can tell whether a candidate file will satisfy
    a project reference. That is the number the .vbp has to match, and it is
    not the same as the file version: MSCOMCTL.OCX 6.01.9782 and 6.01.9786 both
    embed type library 2.0, while the post-MS12-027 builds (6.1.98.x) embed 2.2.

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

.EXAMPLE
    # Check what type-library version a candidate control provides, changing nothing.
    powershell -ExecutionPolicy Bypass -File .\VB6\Install-VB6Dependency.ps1 `
        -Source 'C:\Downloads\MSCOMCTL.OCX' -WhatIfOnly
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

function Get-EmbeddedTypeLib {
    <#
        Read the type-library GUID and version out of a component without
        registering it (REGKIND_NONE). VB6 components are 32-bit, so the probe
        runs in the 32-bit PowerShell host regardless of who called us.

        This is the number a .vbp reference has to match. It is independent of
        the file version.
    #>
    param([string]$FilePath)

    $probeScript = Join-Path ([System.IO.Path]::GetTempPath()) ('rapid-tlbprobe-' + [Guid]::NewGuid().ToString('N') + '.ps1')
    $body = @'
$src = @"
using System;
using System.Runtime.InteropServices;
using CT = System.Runtime.InteropServices.ComTypes;
public static class TlbProbe {
    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void LoadTypeLibEx(string file, int regKind, out CT.ITypeLib tlb);
    public static string Describe(string file) {
        CT.ITypeLib tlb;
        LoadTypeLibEx(file, 2, out tlb);
        IntPtr p;
        tlb.GetLibAttr(out p);
        try {
            CT.TYPELIBATTR a = (CT.TYPELIBATTR)Marshal.PtrToStructure(p, typeof(CT.TYPELIBATTR));
            string name, doc, help; int ctx;
            tlb.GetDocumentation(-1, out name, out doc, out ctx, out help);
            return a.guid.ToString("B").ToUpper() + "|" + a.wMajorVerNum + "." + a.wMinorVerNum + "|" + name + "|" + doc;
        } finally { tlb.ReleaseTLibAttr(p); }
    }
}
"@
Add-Type -TypeDefinition $src -Language CSharp -IgnoreWarnings
try { [TlbProbe]::Describe($args[0]) } catch { "ERROR|" + $_.Exception.Message }
'@
    Set-Content -LiteralPath $probeScript -Value $body -Encoding UTF8

    $host32 = Join-Path $env:SystemRoot 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $host32 -PathType Leaf)) {
        $host32 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    }
    try {
        $raw = & $host32 -NoProfile -ExecutionPolicy Bypass -File $probeScript $FilePath 2>&1 | Select-Object -Last 1
    }
    catch { $raw = 'ERROR|' + $_.Exception.Message }
    finally { Remove-Item -LiteralPath $probeScript -Force -ErrorAction SilentlyContinue }

    $text = [string]$raw
    if ((-not $text) -or $text.StartsWith('ERROR|')) {
        return [pscustomobject]@{ Ok = $false; Detail = $text -replace '^ERROR\|', '' }
    }
    $parts = $text.Split('|')
    if ($parts.Count -lt 4) { return [pscustomobject]@{ Ok = $false; Detail = $text } }
    return [pscustomobject]@{
        Ok = $true; Guid = $parts[0]; Version = $parts[1]; Name = $parts[2]; Description = $parts[3]
    }
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

$embedded = Get-EmbeddedTypeLib -FilePath $Source
if ($embedded.Ok) {
    Write-Host ("  Type library: {0} version {1}  ({2})" -f $embedded.Name, $embedded.Version, $embedded.Description)
    Write-Host ("                {0}" -f $embedded.Guid)
    Write-Host '  A .vbp reference must ask for exactly this version.'
}
else {
    Write-Host ("  Type library: could not be read ({0})" -f $embedded.Detail) -ForegroundColor Yellow
}
Write-Host ''

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

# VB6 caches an OCX's extended type information in a sibling .oca file. After
# the control is replaced that cache describes the old type library, and the
# IDE will happily keep using it. Remove it so VB6 regenerates.
$ocaPath = [System.IO.Path]::ChangeExtension($targetPath, '.oca')
if (Test-Path -LiteralPath $ocaPath -PathType Leaf) {
    Remove-Item -LiteralPath $ocaPath -Force -ErrorAction SilentlyContinue
    Write-Host ("  Removed stale type-info cache {0}" -f (Split-Path -Leaf $ocaPath))
}

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
