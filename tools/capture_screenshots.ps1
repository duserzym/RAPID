<#
.SYNOPSIS
    Launches each RapidPy app briefly and captures a screenshot.

.DESCRIPTION
    Starts each EXE from dist/, waits for the window to appear, captures
    the window to docs/images/, then closes the process.
    Requires: Add-Type from System.Windows.Forms and System.Drawing.

.USAGE
    From repo root (as administrator if UAC blocks screen capture):
        .\tools\capture_screenshots.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
}
"@

$RepoRoot  = Split-Path $PSScriptRoot -Parent
$DistDir   = Join-Path $RepoRoot "dist"
$ImagesDir = Join-Path $RepoRoot "docs\images"

function Capture-AppWindow {
    param(
        [string]$ExePath,
        [string]$OutputFile,
        [string]$WindowTitle,
        [int]$WaitSeconds = 5
    )

    if (-not (Test-Path $ExePath)) {
        Write-Warning "EXE not found: $ExePath — skipping"
        return
    }

    Write-Host "Launching: $ExePath" -ForegroundColor Cyan
    $proc = Start-Process -FilePath $ExePath -PassThru

    # Wait for window to appear
    $deadline = (Get-Date).AddSeconds($WaitSeconds + 10)
    $hwnd = [IntPtr]::Zero
    while ((Get-Date) -lt $deadline -and $hwnd -eq [IntPtr]::Zero) {
        Start-Sleep -Milliseconds 500
        $proc.Refresh()
        if ($proc.MainWindowHandle -ne [IntPtr]::Zero) {
            $hwnd = $proc.MainWindowHandle
        }
    }

    if ($hwnd -eq [IntPtr]::Zero) {
        Write-Warning "Window not found for $ExePath"
        $proc | Stop-Process -Force -ErrorAction SilentlyContinue
        return
    }

    # Give the UI time to fully render
    Start-Sleep -Seconds $WaitSeconds

    [Win32]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 300

    # Get window bounds
    $rect = New-Object Win32+RECT
    [Win32]::GetWindowRect($hwnd, [ref]$rect) | Out-Null
    $width  = $rect.Right  - $rect.Left
    $height = $rect.Bottom - $rect.Top

    if ($width -le 0 -or $height -le 0) {
        Write-Warning "Invalid window bounds for $ExePath"
        $proc | Stop-Process -Force -ErrorAction SilentlyContinue
        return
    }

    # Capture the window region
    $bmp = New-Object System.Drawing.Bitmap($width, $height)
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    $gfx.CopyFromScreen($rect.Left, $rect.Top, 0, 0,
        [System.Drawing.Size]::new($width, $height),
        [System.Drawing.CopyPixelOperation]::SourceCopy)
    $gfx.Dispose()

    $bmp.Save($OutputFile, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host "  Saved: $OutputFile" -ForegroundColor Green

    $proc | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
}

# ── Capture each app ─────────────────────────────────────────────────────────

Capture-AppWindow `
    -ExePath    (Join-Path $DistDir "RapidPy_Gaussmeter.exe") `
    -OutputFile (Join-Path $ImagesDir "gaussmeter_app.png") `
    -WindowTitle "RapidPy Gaussmeter" `
    -WaitSeconds 4

Capture-AppWindow `
    -ExePath    (Join-Path $DistDir "RapidPyVRM.exe") `
    -OutputFile (Join-Path $ImagesDir "vrm_logger_app.png") `
    -WindowTitle "RapidPy VRM" `
    -WaitSeconds 4

Capture-AppWindow `
    -ExePath    (Join-Path $DistDir "RapidPyADWin.exe") `
    -OutputFile (Join-Path $ImagesDir "adwin_comms_app.png") `
    -WindowTitle "ADwin Communication Tester" `
    -WaitSeconds 5

Capture-AppWindow `
    -ExePath    (Join-Path $DistDir "RapidPyCOMMapper.exe") `
    -OutputFile (Join-Path $ImagesDir "com_port_mapper_app.png") `
    -WindowTitle "RapidPy COM Port Mapper" `
    -WaitSeconds 4

Write-Host ""
Write-Host "Screenshot capture complete." -ForegroundColor Green
Write-Host "Images saved to: $ImagesDir"
