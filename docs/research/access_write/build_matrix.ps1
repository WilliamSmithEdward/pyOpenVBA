# Dev-only: compile the VBA construct matrix with Access itself, so the
# differential gate has Microsoft-generated p-code to compare against.
#
#   powershell -File build_matrix.ps1 -Target out.accdb -Source construct_matrix.bas
#
# Adds a Microsoft Scripting Runtime reference so the matrix can exercise
# an external type library alongside the built-in ones.
param(
    [Parameter(Mandatory=$true)][string]$Target,
    [Parameter(Mandatory=$true)][string]$Source
)
$ErrorActionPreference = "Stop"
if (Test-Path -LiteralPath $Target) { Remove-Item -LiteralPath $Target -Force }
$Source = (Resolve-Path -LiteralPath $Source).Path
$app = New-Object -ComObject Access.Application
try {
    $app.Visible = $false
    $app.NewCurrentDatabase($Target)
    $app.References.AddFromFile("C:\Windows\System32\scrrun.dll") | Out-Null
    $app.VBE.ActiveVBProject.VBComponents.Import($Source) | Out-Null
    # 126 = Compile And Save All Modules; this is what persists p-code.
    $app.DoCmd.RunCommand(126)
    Write-Host "compiled and saved $Target"
    $app.CloseCurrentDatabase()
} finally {
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    [System.GC]::Collect() | Out-Null
}
