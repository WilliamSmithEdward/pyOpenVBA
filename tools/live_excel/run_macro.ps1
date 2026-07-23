# Bounded live-Excel macro runner.
#
#     powershell -ExecutionPolicy Bypass -File tools\live_excel\run_macro.ps1 `
#         -WorkbookPath build\gate.xlsm -MacroName RunGate
#
# Spawns run_macro_worker.ps1 in a separate process, enforces a hard
# wall-clock deadline, and kills every task-owned process (worker,
# Excel, dialog watcher) recorded in the PID manifest if the deadline
# passes.  The worker's final stdout line is a JSON object:
#
#     {"workbook": ..., "macro": ..., "outcome": "run-ok", "popups": []}
#
# Any VBE modal (for example a compile error) is dismissed by the
# watcher, logged verbatim into "popups", and surfaces as a non
# "run-ok" outcome -- the caller never deadlocks on an Office dialog.
param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$MacroName,
    [ValidateRange(10, 600)][int]$TimeoutSeconds = 80
)

$ErrorActionPreference = "Stop"

# Start-Process rejects environments carrying both Path and PATH.
$taskPath = [Environment]::GetEnvironmentVariable("Path")
if (-not [string]::IsNullOrWhiteSpace($taskPath)) {
    [Environment]::SetEnvironmentVariable("PATH", $null, [EnvironmentVariableTarget]::Process)
    [Environment]::SetEnvironmentVariable("Path", $taskPath, [EnvironmentVariableTarget]::Process)
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$workDir = Split-Path -Parent $resolvedWorkbook
$stem = [IO.Path]::GetFileNameWithoutExtension($resolvedWorkbook)
$workerOut = Join-Path $workDir "$stem.worker.stdout.log"
$workerErr = Join-Path $workDir "$stem.worker.stderr.log"
$processInfo = Join-Path $workDir "$stem.processes.json"
$dialogLog = Join-Path $workDir "$stem.dialogs.jsonl"
$stopPath = Join-Path $workDir "$stem.watcher.stop"
Remove-Item -LiteralPath $workerOut, $workerErr, $processInfo, $dialogLog, $stopPath `
    -Force -ErrorAction SilentlyContinue

function Stop-TaskProcess {
    param([Nullable[int]]$ProcessId)
    if ($null -eq $ProcessId -or $ProcessId -le 0) { return }
    Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
}

$workerScript = Join-Path $PSScriptRoot "run_macro_worker.ps1"
$arguments = @(
    "-NoProfile", "-ExecutionPolicy", "Bypass",
    "-File", "`"$workerScript`"",
    "-WorkbookPath", "`"$resolvedWorkbook`"",
    "-MacroName", "`"$MacroName`"",
    "-ProcessInfoPath", "`"$processInfo`"",
    "-DialogLogPath", "`"$dialogLog`"",
    "-StopPath", "`"$stopPath`""
)
$worker = Start-Process `
    -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
    -ArgumentList $arguments -WindowStyle Hidden `
    -RedirectStandardOutput $workerOut -RedirectStandardError $workerErr -PassThru

$exitCode = 0
if (-not $worker.WaitForExit($TimeoutSeconds * 1000)) {
    $owned = $null
    if (Test-Path -LiteralPath $processInfo) {
        try { $owned = Get-Content -Raw -LiteralPath $processInfo | ConvertFrom-Json } catch {}
    }
    if ($null -ne $owned) {
        Stop-TaskProcess -ProcessId $owned.excel_process_id
        Stop-TaskProcess -ProcessId $owned.watcher_process_id
    }
    Stop-TaskProcess -ProcessId $worker.Id
    [Console]::Error.WriteLine(
        "TIMEOUT after ${TimeoutSeconds}s; task-owned processes terminated")
    $exitCode = 3
}
else {
    if (Test-Path -LiteralPath $workerOut) {
        Get-Content -LiteralPath $workerOut | Where-Object { $_ } | ForEach-Object { $_ }
    }
    if ((Test-Path -LiteralPath $workerErr) -and (Get-Item $workerErr).Length -gt 0) {
        Get-Content -LiteralPath $workerErr | ForEach-Object {
            [Console]::Error.WriteLine($_)
        }
    }
}
$worker.Dispose()

# Safety sweep: no task-owned process may outlive the run.
if (Test-Path -LiteralPath $processInfo) {
    try {
        $owned = Get-Content -Raw -LiteralPath $processInfo | ConvertFrom-Json
        Stop-TaskProcess -ProcessId $owned.excel_process_id
        Stop-TaskProcess -ProcessId $owned.watcher_process_id
    } catch {}
}
exit $exitCode
