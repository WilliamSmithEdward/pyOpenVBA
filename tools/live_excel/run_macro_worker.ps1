# Worker: open one workbook in Excel over COM and run one macro under a
# popup watcher.  Never call this directly from an interactive session --
# use run_macro.ps1, which enforces a hard deadline and kills the
# processes recorded in the PID manifest if this worker hangs.
#
# Popup-aware pattern adapted from the ROneCOne project's live test
# harness (same author, MIT): PID manifest written before any risky COM
# call, watcher dismisses and logs VBE modals, cleanup in finally.
param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$MacroName,
    [Parameter(Mandatory = $true)][string]$ProcessInfoPath,
    [Parameter(Mandatory = $true)][string]$DialogLogPath,
    [Parameter(Mandatory = $true)][string]$StopPath
)

$ErrorActionPreference = "Stop"
$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedProcessInfo = [System.IO.Path]::GetFullPath($ProcessInfoPath)
$excel = $null
$workbook = $null
$watcher = $null
$excelProcessId = 0
$stage = "create Excel application"
$runOutcome = ""

function Write-ProcessOwnership {
    [ordered]@{
        worker_process_id = $PID
        excel_process_id = $excelProcessId
        watcher_process_id = if ($null -eq $watcher) { 0 } else { $watcher.Id }
    } | ConvertTo-Json -Compress | Set-Content -LiteralPath $resolvedProcessInfo
}

try {
    Remove-Item -LiteralPath $StopPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $DialogLogPath -Force -ErrorAction SilentlyContinue

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class PyOpenVbaExcelProcess
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
}
'@
    [uint32]$ownedExcelProcessId = 0
    [void][PyOpenVbaExcelProcess]::GetWindowThreadProcessId(
        [IntPtr]$excel.Hwnd, [ref]$ownedExcelProcessId)
    $excelProcessId = [int]$ownedExcelProcessId
    Write-ProcessOwnership

    $watcherScript = Join-Path $PSScriptRoot "watch_vbe_dialogs.ps1"
    $watcherArguments = @(
        "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", "`"$watcherScript`"",
        "-ExcelProcessId", $excelProcessId,
        "-LogPath", "`"$DialogLogPath`"",
        "-StopPath", "`"$StopPath`"",
        "-TimeoutSeconds", 75,
        "-DismissKnownDialogs",
        "-TerminateOnBreakMode"
    )
    $watcher = Start-Process `
        -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -ArgumentList $watcherArguments -WindowStyle Hidden -PassThru
    Write-ProcessOwnership

    $stage = "open workbook"
    $workbook = $excel.Workbooks.Open($resolvedWorkbook, 0, $false)

    $stage = "run macro"
    $qualified = "'" + $workbook.Name.Replace("'", "''") + "'!" + $MacroName
    try {
        $excel.Run($qualified) | Out-Null
        $runOutcome = "run-ok"
    }
    catch {
        $message = $_.Exception.Message
        if ($null -ne $_.Exception.InnerException) {
            $message = $_.Exception.InnerException.Message
        }
        $runOutcome = "run-error: " + ($message -replace "\s+", " ").Trim()
    }
}
catch {
    $runOutcome = "stage '$stage' failed: " + $_.Exception.Message
}
finally {
    Set-Content -LiteralPath $StopPath -Value "stop" -ErrorAction SilentlyContinue
    if ($null -ne $watcher) {
        if (-not $watcher.WaitForExit(3000)) {
            Stop-Process -Id $watcher.Id -Force -ErrorAction SilentlyContinue
        }
        $watcher.Dispose()
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($excelProcessId -gt 0) {
        Stop-Process -Id $excelProcessId -Force -ErrorAction SilentlyContinue
    }
}

# Collect every modal the watcher observed (each entry carries the
# dialog text -- for a compile error this is the exact VBE message).
$dialogText = @()
if (Test-Path -LiteralPath $DialogLogPath) {
    foreach ($line in Get-Content -LiteralPath $DialogLogPath) {
        try {
            $record = $line | ConvertFrom-Json
            if ($record.class_name -eq "#32770" -or $record.dismissal_action -ne "none") {
                $text = @($record.child_text | Where-Object { $_ }) -join " | "
                $dialogText += ("[{0}] {1} => {2}" -f $record.title, $text, $record.dismissal_action)
            }
        } catch {}
    }
}

[pscustomobject]@{
    workbook = (Split-Path -Leaf $resolvedWorkbook)
    macro = $MacroName
    outcome = $runOutcome
    popups = $dialogText
} | ConvertTo-Json -Compress
