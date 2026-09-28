# Build modules_added.accdb: the shipped blank template after Access added
# One, Two and Three, deleted Two and added Four, one session each, which
# shows the storage folder and object id Access gives each module.
#
# Access is driven over COM alone.  pyvbaharness would inject three
# modules of its own into the project it runs code in, and while they
# exist they hold the lowest folders and the next ids.
#
# Usage: .\_module_allocation.ps1 -Dest <out.accdb>

param([Parameter(Mandatory=$true)][string]$Dest)

$ErrorActionPreference = "Stop"
$repo = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$template = Join-Path $repo "src\pyopenvba\_templates\blank_files\blank_database.accdb"
$moduleOps = Join-Path $repo "docs\research\access_write\module_ops.ps1"

Copy-Item -Force $template $Dest
$target = (Resolve-Path $Dest).Path
foreach ($step in @(@("add", "One"), @("add", "Two"), @("add", "Three"), @("delete", "Two"), @("add", "Four"))) {
    & $moduleOps -Target $target -Op $step[0] -Name $step[1] | Out-Null
}
Write-Host "wrote $target"
