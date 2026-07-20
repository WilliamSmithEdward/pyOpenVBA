# This function should minimize the window using the Win32 API
# It receives a string with the first part of the window title
# and minimizes the window if it finds a match among the open windows
function Minimize-Window {
    param (
        [string]$windowTitlePart
    )

    Write-Host "Starting window minimization process..." -ForegroundColor Cyan
    Write-Host "Looking for windows with title containing: '$windowTitlePart'" -ForegroundColor Cyan

    # Add the necessary Win32 API function
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
"@

    # Get all open windows
    $windows = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 }
    
    Write-Host "Found $($windows.Count) windows with valid handles" -ForegroundColor Gray
    
    $matchFound = $false
    
    foreach ($window in $windows) {
        Write-Verbose "Checking window: $($window.MainWindowTitle) (Handle: $($window.MainWindowHandle))"
        
        if ($window.MainWindowTitle -like "*$windowTitlePart*") {
            Write-Host "Match found! Window title: $($window.MainWindowTitle)" -ForegroundColor Green
            
            # Minimize the window
            try {
                # SW_MINIMIZE = 6
                [Win32]::ShowWindow($window.MainWindowHandle, 6)
                Write-Host "Successfully minimized window: $($window.MainWindowTitle)" -ForegroundColor Green
                $matchFound = $true
            }
            catch {
                Write-Host "Error minimizing window: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    # If no match was found, return a list of all window handles
    if (-not $matchFound) {
        Write-Host "No windows found matching the pattern: '$windowTitlePart'" -ForegroundColor Yellow
        Write-Host "Here's a list of all available windows:" -ForegroundColor Yellow
        
        $windowList = $windows | ForEach-Object {
            [PSCustomObject]@{
                Title = $_.MainWindowTitle
                Handle = $_.MainWindowHandle
                ProcessName = $_.ProcessName
                Id = $_.Id
            }
        }
        
        $windowList | Format-Table -AutoSize
        return $windowList
    }
}
