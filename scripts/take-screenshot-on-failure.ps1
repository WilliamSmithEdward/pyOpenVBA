#!/usr/bin/env pwsh
<#
.SYNOPSIS
Takes a screenshot of the desktop, minimizing PowerShell windows first.
Useful for capturing Excel state when macros hang or fail.

.PARAMETER OutputPath
The file path where the screenshot should be saved.

.EXAMPLE
./scripts/take-screenshot-on-failure.ps1 -OutputPath ./screenshots/failure.png
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$OutputPath
)

# Ensure output directory exists
$outputDir = Split-Path -Parent $OutputPath
if (!(Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# Win32 API for window manipulation
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

# Minimize all PowerShell windows
Write-Host "Minimizing PowerShell windows..." -ForegroundColor Cyan
$processes = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 }

foreach ($proc in $processes) {
    if ($proc.ProcessName -like "*pwsh*" -or $proc.ProcessName -like "*powershell*") {
        try {
            [Win32]::ShowWindow($proc.MainWindowHandle, 6) # SW_MINIMIZE = 6
            Write-Host "Minimized: $($proc.ProcessName)" -ForegroundColor Green
        }
        catch {
            # Continue silently
        }
    }
}

Write-Host "Waiting 2 seconds for windows to minimize..." -ForegroundColor Cyan
Start-Sleep -Seconds 2

# Create temp script for screenshot
$tempScript = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.ps1'
$screenshotCode = @"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
`$bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
`$screenshot = New-Object System.Drawing.Bitmap(`$bounds.Width, `$bounds.Height)
`$graphics = [System.Drawing.Graphics]::FromImage(`$screenshot)
`$graphics.CopyFromScreen(`$bounds.Location, [System.Drawing.Point]::Empty, `$bounds.Size)
`$screenshot.Save('$OutputPath')
`$graphics.Dispose()
`$screenshot.Dispose()
"@

$screenshotCode | Out-File -FilePath $tempScript -Encoding UTF8

# Run screenshot in hidden PowerShell process
Start-Process pwsh -ArgumentList "-NoProfile", "-Command", ". '$tempScript'" -WindowStyle Hidden -Wait

# Clean up temp script
Remove-Item -Path $tempScript -Force -ErrorAction SilentlyContinue

Write-Host "Screenshot saved to: $OutputPath" -ForegroundColor Green



