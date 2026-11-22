function Get-RapidPyFileHash {
    param([Parameter(Mandatory)][string]$Path)

    $stream = [IO.File]::OpenRead($Path)
    try {
        $hash = [Security.Cryptography.SHA256]::Create()
        try { [BitConverter]::ToString($hash.ComputeHash($stream)).Replace('-', '').ToLowerInvariant() }
        finally { $hash.Dispose() }
    } finally {
        $stream.Dispose()
    }
}

function Get-RapidPyBundleInventory {
    param([Parameter(Mandatory)][string]$Bundle)

    $root = Get-Item -LiteralPath $Bundle -ErrorAction Stop
    if (-not $root.PSIsContainer -or ($root.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        throw 'The portable bundle must be an ordinary directory.'
    }
    $prefix = $root.FullName.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
    $pending = New-Object 'System.Collections.Generic.Stack[System.IO.DirectoryInfo]'
    $pending.Push($root)
    $files = while ($pending.Count -gt 0) {
        $directory = $pending.Pop()
        foreach ($item in Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction Stop) {
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Portable bundle contains a reparse point: $($item.FullName)"
            }
            if ($item.PSIsContainer) {
                $pending.Push($item)
            } else {
                $relative = $item.FullName.Substring($prefix.Length).Replace('\', '/')
                if ($relative -ceq 'build_manifest.json') { continue }
                [ordered]@{
                    path = $relative
                    size_bytes = $item.Length
                    sha256 = Get-RapidPyFileHash -Path $item.FullName
                }
            }
        }
    }
    $files | Sort-Object -Property { $_.path }
}

function Test-RapidPyBundle {
    param([Parameter(Mandatory)][string]$Bundle, [switch]$RequireFullInventory)

    $manifestPath = Join-Path $Bundle 'build_manifest.json'
    $manifestFile = Get-Item -LiteralPath $manifestPath -ErrorAction Stop
    if ($manifestFile.PSIsContainer -or ($manifestFile.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
        throw 'The portable build manifest must be an ordinary file.'
    }
    $manifestHash = Get-RapidPyFileHash -Path $manifestFile.FullName
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
    $executables = @('RapidPyMain.exe', 'RapidPyMainConsole.exe')
    if ($manifest.schema -ceq 'rapidpy.portable_build.v1' -and -not $RequireFullInventory) {
        foreach ($name in $executables) {
            $actual = Get-RapidPyFileHash -Path (Join-Path $Bundle $name)
            if ($actual -ne $manifest.executables.$name) { throw "Legacy manifest executable differs: $name" }
        }
        return [ordered]@{ full_inventory = $false; files_checked = 2; manifest_sha256 = $manifestHash }
    }
    if ($manifest.schema -cne 'rapidpy.portable_build.v2' -or $manifest.files -isnot [array]) {
        throw 'A version 2 manifest is required for whole-bundle verification.'
    }

    $expected = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([StringComparer]::Ordinal)
    foreach ($entry in $manifest.files) {
        if ($entry.path -isnot [string] -or [string]::IsNullOrEmpty($entry.path) -or
            $entry.sha256 -cnotmatch '^[0-9a-f]{64}$' -or
            ($entry.size_bytes -isnot [int] -and $entry.size_bytes -isnot [long]) -or
            $entry.size_bytes -lt 0 -or $expected.ContainsKey($entry.path)) {
            throw 'The portable file inventory contains an invalid or duplicate entry.'
        }
        $expected.Add($entry.path, $entry)
    }
    foreach ($name in $executables + @('OperatorGuide.md', 'Verify-RapidPy.ps1', 'portable_bundle.psm1')) {
        if (-not $expected.ContainsKey($name)) { throw "Required portable file is missing from the manifest: $name" }
    }
    if (-not ($expected.Keys | Where-Object { $_.StartsWith('_internal/', [StringComparison]::Ordinal) } | Select-Object -First 1)) {
        throw 'The portable manifest has no bundled runtime files.'
    }

    $actualFiles = @(Get-RapidPyBundleInventory -Bundle $Bundle)
    foreach ($file in $actualFiles) {
        if (-not $expected.ContainsKey($file.path)) { throw "Unexpected portable file: $($file.path)" }
        $entry = $expected[$file.path]
        if ($file.sha256 -cne $entry.sha256 -or $file.size_bytes -ne $entry.size_bytes) {
            throw "Portable file differs from its manifest: $($file.path)"
        }
        $expected.Remove($file.path) | Out-Null
    }
    if ($expected.Count -gt 0) { throw "Missing portable files: $($expected.Keys -join ', ')" }
    return [ordered]@{
        full_inventory = $true
        files_checked = $actualFiles.Count
        manifest_sha256 = $manifestHash
        source_revision = $manifest.source_revision
    }
}

Export-ModuleMember -Function Get-RapidPyFileHash, Get-RapidPyBundleInventory, Test-RapidPyBundle
