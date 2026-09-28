# Build folders_past_nine.accdb: the shipped blank template after Access
# has filled every storage container past folder 9 (GitHub issue #34).
#
# Twelve modules go in one Access session each, saved with
# acCmdCompileAndSaveAllModules, because that is what makes Access write
# `Modules/PropData` entries for them, two-digit folders included.  Then
# one more session adds eleven forms, eleven reports and eleven macros,
# which takes `Forms`, `Reports` and `Scripts` to folder 10.
#
# Usage: .\_folders_past_nine.ps1 -Dest <out.accdb>

param([Parameter(Mandatory=$true)][string]$Dest)

$ErrorActionPreference = "Stop"
$repo = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$template = Join-Path $repo "src\pyopenvba\_templates\blank_files\blank_database.accdb"
$moduleOps = Join-Path $repo "docs\research\access_write\module_ops.ps1"

Copy-Item -Force $template $Dest
$target = (Resolve-Path $Dest).Path

for ($i = 2; $i -le 13; $i++) {
    & $moduleOps -Target $target -Op add -Name "Mod$i" | Out-Null
}

$macroText = [System.IO.Path]::GetTempFileName()
$macroLines = @('Version =196611', 'ColumnsShown =0', 'Begin', '    Action ="Beep"', 'End')
[System.IO.File]::WriteAllText($macroText, ($macroLines -join "`r`n") + "`r`n", [System.Text.Encoding]::ASCII)

$access = New-Object -ComObject Access.Application
$access.Visible = $false
try {
    $access.OpenCurrentDatabase($target)
    for ($i = 1; $i -le 11; $i++) {
        $form = $access.CreateForm()
        $access.DoCmd.Close(2, $form.Name, 1)      # acForm, acSaveYes
    }
    for ($i = 1; $i -le 11; $i++) {
        $report = $access.CreateReport()
        $access.DoCmd.Close(3, $report.Name, 1)    # acReport, acSaveYes
    }
    for ($i = 1; $i -le 11; $i++) {
        $access.LoadFromText(4, "Macro$i", $macroText)  # acMacro
    }
    $access.CloseCurrentDatabase()
}
finally {
    $access.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Remove-Item -Force $macroText
}
Write-Host "wrote $target"
