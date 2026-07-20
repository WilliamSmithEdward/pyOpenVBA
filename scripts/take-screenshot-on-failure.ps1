#!/usr/bin/env pwsh
<#
.SYNOPSIS
Takes a screenshot of the desktop, minimizing the PowerShell window first.
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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Minimize PowerShell window to see Excel
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

$pwsh = Get-Process -Id $PID
[Win32]::ShowWindow($pwsh.MainWindowHandle, 6) # SW_MINIMIZE = 6

Start-Sleep -Seconds 1

$screenshot = New-Object System.Drawing.Bitmap(
    [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width,
    [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height
)
$graphics = [System.Drawing.Graphics]::FromImage($screenshot)
$graphics.CopyFromScreen(
    [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Location,
    [System.Drawing.Point]::Empty,
    [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Size
)

# Ensure output directory exists
$outputDir = Split-Path -Parent $OutputPath
if (!(Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$screenshot.Save($OutputPath)
$graphics.Dispose()
$screenshot.Dispose()

Write-Host "Screenshot saved to: $OutputPath"
