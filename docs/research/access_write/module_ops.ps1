# Dev-only: perform one module-level operation with Access, so its effect
# on the file can be diffed. Operations: add, rename, delete.
#
#   powershell -File module_ops.ps1 -Target db.accdb -Op add    -Name Gamma
#   powershell -File module_ops.ps1 -Target db.accdb -Op rename -Name Alpha -NewName Beta
#   powershell -File module_ops.ps1 -Target db.accdb -Op delete -Name Alpha
param(
    [Parameter(Mandatory=$true)][string]$Target,
    [Parameter(Mandatory=$true)][string]$Op,
    [Parameter(Mandatory=$true)][string]$Name,
    [string]$NewName
)
$ErrorActionPreference = "Stop"
$app = New-Object -ComObject Access.Application
try {
    $app.Visible = $false
    $app.OpenCurrentDatabase($Target)
    $vbp = $app.VBE.ActiveVBProject
    switch ($Op) {
        "add" {
            $c = $vbp.VBComponents.Add(1)          # 1 = vbext_ct_StdModule
            $c.Name = $Name
            $c.CodeModule.AddFromString(
                "Public Function " + $Name + "Go() As Variant" + [char]13 + [char]10 +
                "    " + $Name + "Go = 1" + [char]13 + [char]10 +
                "End Function")
        }
        "rename" { $vbp.VBComponents($Name).Name = $NewName }
        "delete" { $vbp.VBComponents.Remove($vbp.VBComponents($Name)) }
        default  { throw "unknown op $Op" }
    }
    # 126 = acCmdCompileAndSaveAllModules; the save half is what persists.
    $app.DoCmd.RunCommand(126)
    Write-Host "$Op ok"
    $app.CloseCurrentDatabase()
} finally {
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    [System.GC]::Collect() | Out-Null
}
