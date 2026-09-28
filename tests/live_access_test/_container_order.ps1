# Have Access delete and rename objects in copies of folders_past_nine.accdb,
# one change per copy and per session, so the order Access writes a
# container's `\x03DirData` and `PropData` in after each can be compared
# with the library's (tests/test_access_vba.py records what these wrote).
# Access is driven over COM alone, and never on the fixture itself: opening
# a database is enough for Access to rewrite parts of it.
#
# Usage: .\_container_order.ps1 -Dest <directory>

param([Parameter(Mandatory=$true)][string]$Dest)

$ErrorActionPreference = "Stop"
$source = Join-Path $PSScriptRoot "folders_past_nine.accdb"
New-Item -ItemType Directory -Force -Path $Dest | Out-Null
$folder = (Resolve-Path $Dest).Path

function In-Access([string]$Name, [scriptblock]$Work) {
    $path = Join-Path $folder $Name
    Copy-Item -Force $source $path
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    try {
        $access.OpenCurrentDatabase($path)
        & $Work $access
        $access.DoCmd.RunCommand(126)   # acCmdCompileAndSaveAllModules
        $access.CloseCurrentDatabase()
    }
    finally {
        $access.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

In-Access "module_deleted.accdb" { param($a) $p = $a.VBE.ActiveVBProject; $p.VBComponents.Remove($p.VBComponents("Mod7")) }
In-Access "module_renamed.accdb" { param($a) $a.VBE.ActiveVBProject.VBComponents("Mod3").Name = "Renamed3" }
In-Access "form_deleted.accdb" { param($a) $a.DoCmd.DeleteObject(2, "Form5") }        # acForm
In-Access "report_renamed.accdb" { param($a) $a.DoCmd.Rename("Summary", 3, "Report2") }   # acReport
Write-Host "wrote $folder"
