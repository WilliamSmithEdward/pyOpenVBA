# Build the databases that show what Access writes for a database's first
# VBA project, all from no_project.accdb, a database Access made and never
# gave code:
#
#   first_module.accdb           Access gave it a first module
#   first_module_reopened.accdb  first_module.accdb opened once more, nothing done
#   first_form.accdb             Access gave it a first form
#   first_macro.accdb            Access gave it a first macro
#
# A project is named after the file it is made in, up to the first dot, so
# each is made under its own name.  Access is driven over COM alone:
# pyvbaharness would inject modules of its own into the project it runs
# code in, and those take storage folders and object ids while Access works.
#
# Usage: .\_first_project.ps1 -Dest <directory>

param([Parameter(Mandatory=$true)][string]$Dest)

$ErrorActionPreference = "Stop"
$repo = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$moduleOps = Join-Path $repo "docs\research\access_write\module_ops.ps1"
New-Item -ItemType Directory -Force -Path $Dest | Out-Null
$folder = (Resolve-Path $Dest).Path

function In-Access([string]$Path, [scriptblock]$Work) {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    try {
        $access.OpenCurrentDatabase($Path)
        & $Work $access
        $access.CloseCurrentDatabase()
    }
    finally {
        $access.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

$bare = Join-Path $folder "no_project.accdb"
if (Test-Path $bare) { Remove-Item -Force $bare }
$access = New-Object -ComObject Access.Application
$access.Visible = $false
try {
    $access.NewCurrentDatabase($bare, 12)   # acNewDatabaseFormatAccess2007
    $access.CloseCurrentDatabase()
}
finally {
    $access.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

$first = Join-Path $folder "first_module.accdb"
Copy-Item -Force $bare $first
& $moduleOps -Target $first -Op add -Name Module1 | Out-Null

$reopened = Join-Path $folder "first_module_reopened.accdb"
Copy-Item -Force $first $reopened
In-Access $reopened { param($a) }

$form = Join-Path $folder "first_form.accdb"
Copy-Item -Force $bare $form
In-Access $form { param($a) $f = $a.CreateForm(); $a.DoCmd.Close(2, $f.Name, 1) }   # acForm, acSaveYes

$macroText = [System.IO.Path]::GetTempFileName()
$macroLines = @('Version =196611', 'ColumnsShown =0', 'Begin', '    Action ="Beep"', 'End')
[System.IO.File]::WriteAllText($macroText, ($macroLines -join "`r`n") + "`r`n", [System.Text.Encoding]::ASCII)
$macro = Join-Path $folder "first_macro.accdb"
Copy-Item -Force $bare $macro
In-Access $macro { param($a) $a.LoadFromText(4, "Macro1", $macroText) }   # acMacro
Remove-Item -Force $macroText

Write-Host "wrote $folder"
