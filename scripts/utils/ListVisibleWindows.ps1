function Get-VisibleWindows {
    $windows = Get-Process |
        Where-Object { $_.MainWindowHandle -ne 0 -and -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle) } |
        Sort-Object ProcessName, MainWindowTitle |
        Select-Object ProcessName, Id, MainWindowHandle, MainWindowTitle

    return $windows
}

function Write-VisibleWindowsReport {
    param (
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [string]$Header = "Visible windows"
    )

    $outputDir = Split-Path -Parent $OutputPath
    if (!(Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $windows = Get-VisibleWindows

    $lines = @()
    $lines += "$Header"
    $lines += "Timestamp: $timestamp"
    $lines += "Count: $($windows.Count)"
    $lines += ""

    foreach ($w in $windows) {
        $lines += "Process=$($w.ProcessName) | PID=$($w.Id) | Handle=$($w.MainWindowHandle) | Title=$($w.MainWindowTitle)"
    }

    Set-Content -Path $OutputPath -Value $lines -Encoding UTF8

    Write-Host "Visible window list saved to: $OutputPath" -ForegroundColor Green
}
