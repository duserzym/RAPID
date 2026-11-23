function Get-VB6SourceFingerprint {
    param([string]$Project)

    $projectDirectory = Split-Path -Parent $Project
    $names = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    [void]$names.Add((Split-Path -Leaf $Project))
    foreach ($line in [IO.File]::ReadLines($Project)) {
        if ($line -match '^(?:Form|UserControl|UserDocument|Designer|PropertyPage|ResFile32)=(.+)$' -or
            $line -match '^(?:Class|Module)=[^;]+;\s*(.+)$') {
            $name = $Matches[1].Trim().Trim('"')
            [void]$names.Add($name)
            if ([IO.Path]::GetExtension($name) -in @('.frm', '.ctl', '.dsr', '.dob', '.pag')) {
                $source = [IO.File]::ReadAllText((Join-Path $projectDirectory $name))
                foreach ($resource in [regex]::Matches($source, '(?i)"([^"\r\n]+\.(?:frx|ctx|dsx|dox|pgx))":[0-9a-f]+')) {
                    [void]$names.Add([IO.Path]::Combine([IO.Path]::GetDirectoryName($name), $resource.Groups[1].Value))
                }
            }
        }
    }

    [string[]]$orderedNames = @($names)
    [Array]::Sort($orderedNames, [StringComparer]::OrdinalIgnoreCase)
    $entries = foreach ($name in $orderedNames) {
        $hash = Get-FileHash -LiteralPath (Join-Path $projectDirectory $name) -Algorithm SHA256
        "$($hash.Hash) $name"
    }
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes(($entries -join "`n"))
        [BitConverter]::ToString($sha256.ComputeHash($bytes)).Replace('-', '')
    } finally {
        $sha256.Dispose()
    }
}
