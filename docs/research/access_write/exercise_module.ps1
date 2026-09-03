# Dev-only: put a procedure into a module and run it, reporting each step,
# so a created module can be exercised without the harness injecting one.
param([Parameter(Mandatory=$true)][string]$Target,
      [Parameter(Mandatory=$true)][string]$Module)
$ErrorActionPreference = "Stop"
$app = New-Object -ComObject Access.Application
$steps = @()
try {
    $app.Visible = $false
    $app.OpenCurrentDatabase($Target)
    $steps += "opened"
    $c = $app.VBE.ActiveVBProject.VBComponents.Item($Module)
    $steps += "found=" + $c.Name
    $c.CodeModule.AddFromString(
        "Public Function ProbeX() As Variant" + [char]13 + [char]10 +
        "    ProbeX = 777" + [char]13 + [char]10 + "End Function")
    $steps += "added lines=" + $c.CodeModule.CountOfLines
    $steps += "ran=" + $app.Run("ProbeX")
    $app.CloseCurrentDatabase()
} catch {
    $steps += "ERR " + $_.Exception.Message
} finally {
    [Console]::Out.Write($steps -join " | ")
    $app.Quit(2)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
