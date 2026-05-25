# Access COM oracle for testing only.
# Dumps every VBA module from an .accdb to plain text files in $OutDir.
# Used to verify pure-Python pyopenvba.access output.

param(
    [Parameter(Mandatory=$true)][string]$Accdb,
    [Parameter(Mandatory=$true)][string]$OutDir
)

$ErrorActionPreference = "Stop"
$Accdb = (Resolve-Path -LiteralPath $Accdb).Path
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

$app = New-Object -ComObject Access.Application
try {
    $app.Visible = $false
    $app.OpenCurrentDatabase($Accdb)
    $vbp = $app.VBE.ActiveVBProject

    $manifest = @()
    foreach ($comp in $vbp.VBComponents) {
        $cm = $comp.CodeModule
        $lineCount = $cm.CountOfLines
        if ($lineCount -gt 0) {
            $src = $cm.Lines(1, $lineCount)
        } else {
            $src = ""
        }
        $safe = $comp.Name -replace '[^A-Za-z0-9_.-]', '_'
        $outFile = Join-Path $OutDir "$safe.bas.txt"
        [System.IO.File]::WriteAllText($outFile, $src, [System.Text.UTF8Encoding]::new($false))
        $manifest += [pscustomobject]@{
            Name = $comp.Name
            Type = [int]$comp.Type
            Lines = $lineCount
            File = (Split-Path -Leaf $outFile)
        }
        Write-Host "Exported $($comp.Name) (type=$([int]$comp.Type), lines=$lineCount) -> $(Split-Path -Leaf $outFile)"
    }
    $manifest | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $OutDir "manifest.json") -Encoding UTF8
} finally {
    try { $app.CloseCurrentDatabase() } catch {}
    $app.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
    [System.GC]::Collect() | Out-Null
}
