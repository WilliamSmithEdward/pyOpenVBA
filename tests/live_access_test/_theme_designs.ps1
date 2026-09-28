# Build designs_aptos.accdb: a new database, which Access gives Office's
# 2023 theme with its first form, holding one form and one report with a
# control of every type Access makes without asking for anything (no object
# frame, ActiveX control or chart).  Its control-defaults objects are what
# the library's, captured on Office's 2007 theme, become in the 2023 theme.
#
# Usage: .\_theme_designs.ps1 -Dest <directory>

param([Parameter(Mandatory=$true)][string]$Dest)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $Dest | Out-Null
$path = Join-Path (Resolve-Path $Dest).Path "designs_aptos.accdb"
if (Test-Path $path) { Remove-Item -Force $path }
$types = @(100, 101, 102, 104, 105, 106, 107, 108, 109, 110, 111, 112, 118, 122, 123, 126, 128, 129, 134)

$access = New-Object -ComObject Access.Application
$access.Visible = $false
try {
    $access.NewCurrentDatabase($path, 12)   # acNewDatabaseFormatAccess2007
    $access.DoCmd.SetWarnings($false)
    $form = $access.CreateForm()
    $name = $form.Name
    $i = 0
    foreach ($t in $types) {
        $i++
        try { [void]$access.CreateControl($name, $t, 0, "", "", 200, 200 + 400 * $i, 1400, 300) }
        catch { Write-Host "form: type $t refused" }
    }
    $access.DoCmd.Close(2, $name, 1)            # acForm, acSaveYes
    $report = $access.CreateReport()
    $name = $report.Name
    $i = 0
    foreach ($t in $types) {
        $i++
        try { [void]$access.CreateReportControl($name, $t, 0, "", "", 200, 200 + 400 * $i, 1400, 300) }
        catch { Write-Host "report: type $t refused" }
    }
    $access.DoCmd.Close(3, $name, 1)            # acReport, acSaveYes
    $access.CloseCurrentDatabase()
}
finally {
    $access.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Host "wrote $path"
