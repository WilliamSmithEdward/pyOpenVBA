# Read-only: open a database and report what its VBA project looks like,
# without adding anything to it.
param([Parameter(Mandatory=$true)][string]$Target)
$ErrorActionPreference = "Stop"
$app = New-Object -ComObject Access.Application
try {
    $app.Visible = $false
    $app.OpenCurrentDatabase($Target)
    $names = @()
    foreach ($c in $app.VBE.ActiveVBProject.VBComponents) { $names += $c.Name }
    $all = @()
    for ($i = 0; $i -lt $app.CurrentProject.AllModules.Count; $i++) {
        $all += $app.CurrentProject.AllModules.Item($i).Name
    }
    [Console]::Out.Write("vbe=" + ($names -join ",") + " access=" + ($all -join ","))
    $app.CloseCurrentDatabase()
} catch {
    [Console]::Out.Write("ERR " + $_.Exception.Message)
    exit 1
} finally {
    $app.Quit(2)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
