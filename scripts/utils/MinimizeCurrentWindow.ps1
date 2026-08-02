# Minimize the current PowerShell window
function Minimize-CurrentWindow {
    Add-Type -TypeDefinition @'
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
'@
    
    $pwsh = Get-Process -Id $PID
    # SW_MINIMIZE = 6
    [Win32]::ShowWindow($pwsh.MainWindowHandle, 6)
    Write-Host "Current PowerShell window minimized" -ForegroundColor Green
}
