# Drive Access COM to perform a single VBA-tree mutation against a copy
# of the fixture .accdb, save it, and emit a sibling file for byte-diffing.
#
# Usage:
#   .\_com_mutate.ps1 -Source <baseline.accdb> -Dest <out.accdb> -Op <op>
#       [-ModuleName <name>] [-NewName <name>] [-Body <verbatim>]
#
# Ops:
#   add_empty       Add VBComponent of type vbext_ct_StdModule with name -ModuleName
#   add_with_body   Add module -ModuleName whose code body is -Body (literal text)
#   rename          Rename -ModuleName -> -NewName
#   remove          Remove module -ModuleName
#   touch_source    Replace -ModuleName's full body with -Body
#
# COM oracle is then run on the output for parity checks (separate script).

param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Dest,
    [Parameter(Mandatory=$true)][ValidateSet("add_empty","add_with_body","rename","remove","touch_source")][string]$Op,
    [string]$ModuleName = "",
    [string]$NewName = "",
    [string]$Body = ""
)

$ErrorActionPreference = "Stop"
Copy-Item -Path $Source -Destination $Dest -Force
$absDest = (Resolve-Path $Dest).Path

$access = New-Object -ComObject Access.Application
$access.Visible = $false

try {
    $access.OpenCurrentDatabase($absDest)
    $proj = $access.VBE.ActiveVBProject
    $comps = $proj.VBComponents

    switch ($Op) {
        "add_empty" {
            $c = $comps.Add(1)  # vbext_ct_StdModule
            $c.Name = $ModuleName
        }
        "add_with_body" {
            $c = $comps.Add(1)
            $c.Name = $ModuleName
            if ($Body -ne "") {
                $c.CodeModule.AddFromString($Body)
            }
        }
        "rename" {
            $c = $comps.Item($ModuleName)
            $c.Name = $NewName
        }
        "remove" {
            $c = $comps.Item($ModuleName)
            $comps.Remove($c)
        }
        "touch_source" {
            $c = $comps.Item($ModuleName)
            $cm = $c.CodeModule
            if ($cm.CountOfLines -gt 0) {
                $cm.DeleteLines(1, $cm.CountOfLines)
            }
            if ($Body -ne "") {
                $cm.AddFromString($Body)
            }
        }
    }

    # Force a save through Access (DoCmd.RunCommand acCmdCompileAndSaveAllModules = 126)
    try { $access.RunCommand(126) } catch { }
    $access.CloseCurrentDatabase()
}
finally {
    $access.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host "ok -> $absDest"
