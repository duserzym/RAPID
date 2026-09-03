<#
.SYNOPSIS
    Find Visual Basic 6 projects (.vbp) on this computer and cache the results.

.DESCRIPTION
    Walks a set of likely roots with a skip list so a full sweep does not crawl
    Windows, Program Files, package caches, or version-control internals. The
    RAPID Paleomag project is always ranked first when it is found.

    Results are cached at %LOCALAPPDATA%\RAPID\vb6-projects.json so
    Start-VB6.ps1 can resolve a project by name without rescanning.

.PARAMETER Path
    Extra roots to scan, in addition to the defaults.

.PARAMETER Deep
    Scan every fixed drive from its root instead of the default shortlist.

.PARAMETER Refresh
    Ignore the cache and rescan.

.PARAMETER Name
    Return only projects whose file name or folder matches this wildcard.

.PARAMETER AsObject
    Emit objects instead of a formatted table (for scripting).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\VB6\Find-VB6Projects.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\VB6\Find-VB6Projects.ps1 -Deep -Refresh
#>
[CmdletBinding()]
param(
    [string[]]$Path,
    [switch]$Deep,
    [switch]$Refresh,
    [string]$Name,
    [switch]$AsObject,
    [int]$MaxDepth = 12
)

$ErrorActionPreference = 'Stop'

$cacheDir = Join-Path $env:LOCALAPPDATA 'RAPID'
$cachePath = Join-Path $cacheDir 'vb6-projects.json'

# Directories that never hold a hand-written VB6 project but cost minutes to walk.
$skipNames = @(
    '$recycle.bin', 'system volume information', 'windows', 'winsxs',
    'node_modules', '.git', '.svn', '.hg', '__pycache__', '.venv', 'venv',
    'site-packages', 'appdata', 'onedrivetemp', 'packages', '.gradle', '.nuget',
    'obj', 'bin.cache', '.cache', '.vs', 'temp', 'tmp'
)

function Get-DefaultRoots {
    $roots = [System.Collections.Generic.List[string]]::new()

    # The repository this script ships in, then the usual places a lab keeps code.
    $repoRoot = Split-Path -Parent $PSScriptRoot
    if ($repoRoot) { $roots.Add($repoRoot) }

    foreach ($candidate in @(
        $env:USERPROFILE,
        (Join-Path $env:USERPROFILE 'Documents'),
        (Join-Path $env:USERPROFILE 'Desktop'),
        (Join-Path $env:USERPROFILE 'Downloads'),
        'C:\Paleomag',
        'C:\RAPID',
        'C:\Source',
        'C:\Projects',
        'C:\VB6'
    )) {
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Container)) {
            $roots.Add($candidate)
        }
    }

    # Every fixed drive except the system drive, which the shortlist already covers.
    foreach ($drive in [System.IO.DriveInfo]::GetDrives()) {
        if ($drive.DriveType -ne [System.IO.DriveType]::Fixed) { continue }
        if (-not $drive.IsReady) { continue }
        if ($Deep -or ($drive.Name -ne $env:SystemDrive + '\')) {
            $roots.Add($drive.RootDirectory.FullName)
        }
    }

    return $roots | Select-Object -Unique
}

function Find-ProjectFiles {
    param([string[]]$Roots, [int]$Depth)

    $found = [System.Collections.Generic.List[string]]::new()
    $seenDirs = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)

    foreach ($root in $Roots) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }

        $stack = [System.Collections.Generic.Stack[object]]::new()
        $stack.Push([pscustomobject]@{ Dir = (Resolve-Path -LiteralPath $root).Path; Level = 0 })

        while ($stack.Count -gt 0) {
            $node = $stack.Pop()
            if (-not $seenDirs.Add($node.Dir)) { continue }

            try {
                foreach ($file in [System.IO.Directory]::EnumerateFiles($node.Dir, '*.vbp')) {
                    $found.Add($file)
                }
            }
            catch { continue }

            if ($node.Level -ge $Depth) { continue }

            try {
                foreach ($child in [System.IO.Directory]::EnumerateDirectories($node.Dir)) {
                    $leaf = [System.IO.Path]::GetFileName($child)
                    if ($skipNames -contains $leaf.ToLowerInvariant()) { continue }
                    if ($leaf.StartsWith('.') -and $leaf.Length -gt 1) { continue }
                    $stack.Push([pscustomobject]@{ Dir = $child; Level = $node.Level + 1 })
                }
            }
            catch { continue }
        }
    }

    return $found | Select-Object -Unique
}

function Get-ProjectDetail {
    param([string]$ProjectPath)

    $item = Get-Item -LiteralPath $ProjectPath
    $exeName = ''
    $projectType = ''
    $startup = ''
    $referenceCount = 0
    $objectCount = 0

    try {
        foreach ($line in [System.IO.File]::ReadLines($ProjectPath)) {
            if ($line -like 'ExeName32=*') { $exeName = $line.Substring(10).Trim('"') }
            elseif ($line -like 'Type=*') { $projectType = $line.Substring(5).Trim() }
            elseif ($line -like 'Startup=*') { $startup = $line.Substring(8).Trim('"') }
            elseif ($line -like 'Reference=*') { $referenceCount++ }
            elseif ($line -like 'Object=*') { $objectCount++ }
        }
    }
    catch { }

    $isRapid = $false
    if ($exeName -match 'PALEOMAG') { $isRapid = $true }
    if ($item.Name -like 'Paleomag*') { $isRapid = $true }

    return [pscustomobject]@{
        Name           = $item.Name
        Directory      = $item.DirectoryName
        FullName       = $item.FullName
        ExeName        = $exeName
        ProjectType    = $projectType
        Startup        = $startup
        References     = $referenceCount
        Controls       = $objectCount
        SizeBytes      = $item.Length
        LastWriteTime  = $item.LastWriteTime
        IsRapidPaleomag = $isRapid
    }
}

# ---------------------------------------------------------------- resolve list
$projects = $null

if ((-not $Refresh) -and (Test-Path -LiteralPath $cachePath -PathType Leaf)) {
    try {
        $cached = Get-Content -LiteralPath $cachePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $stillThere = @($cached.projects | Where-Object { Test-Path -LiteralPath $_.FullName -PathType Leaf })
        if ($stillThere.Count -gt 0) {
            $projects = $stillThere
            Write-Verbose ("Loaded {0} project(s) from cache {1}" -f $projects.Count, $cachePath)
        }
    }
    catch {
        Write-Verbose "Cache unreadable, rescanning."
    }
}

if (-not $projects) {
    $roots = Get-DefaultRoots
    if ($Path) { $roots = @($Path) + @($roots) | Select-Object -Unique }

    Write-Verbose ("Scanning {0} root(s) to depth {1}" -f @($roots).Count, $MaxDepth)
    $files = Find-ProjectFiles -Roots $roots -Depth $MaxDepth
    $projects = @($files | ForEach-Object { Get-ProjectDetail -ProjectPath $_ })

    if (-not (Test-Path -LiteralPath $cacheDir -PathType Container)) {
        New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null
    }
    $payload = [pscustomobject]@{
        schema    = 'rapid.vb6.project-cache.v1'
        scannedAt = (Get-Date).ToString('o')
        roots     = @($roots)
        projects  = @($projects)
    }
    $payload | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $cachePath -Encoding UTF8
}

# RAPID first, then most recently modified.
$projects = @($projects | Sort-Object -Property @{ Expression = 'IsRapidPaleomag'; Descending = $true },
                                                @{ Expression = 'LastWriteTime'; Descending = $true })

if ($Name) {
    $projects = @($projects | Where-Object { $_.Name -like $Name -or $_.Directory -like $Name -or $_.FullName -like $Name })
}

if ($AsObject) {
    $projects
    return
}

if (@($projects).Count -eq 0) {
    Write-Host 'No .vbp projects found.' -ForegroundColor Yellow
    Write-Host 'Try -Deep to sweep every fixed drive, or -Path to add a root.'
    exit 1
}

Write-Host ''
Write-Host ("Found {0} VB6 project(s).  Cache: {1}" -f @($projects).Count, $cachePath)
Write-Host ''
$index = 0
foreach ($project in $projects) {
    $index++
    $marker = '   '
    if ($project.IsRapidPaleomag) { $marker = ' * ' }
    Write-Host ("{0}{1,2}. {2}" -f $marker, $index, $project.FullName)
    Write-Host ("       exe={0}  type={1}  refs={2}  controls={3}  modified={4:yyyy-MM-dd}" -f `
        $(if ($project.ExeName) { $project.ExeName } else { '(unset)' }),
        $(if ($project.ProjectType) { $project.ProjectType } else { '?' }),
        $project.References, $project.Controls, $project.LastWriteTime)
}
Write-Host ''
Write-Host '  * = RAPID Paleomag project'
Write-Host ''
Write-Host 'Open one with:'
Write-Host '  powershell -ExecutionPolicy Bypass -File .\VB6\Start-VB6.ps1'
Write-Host '  powershell -ExecutionPolicy Bypass -File .\VB6\Start-VB6.ps1 -Project "<path or wildcard>"'
