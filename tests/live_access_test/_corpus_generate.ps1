# Generate a controlled RE corpus of .accdb files for VBA p-code analysis.
#
# Strategy: start from a freshly-created empty .accdb baseline, then apply
# ONE parameterized VBA-tree mutation per output file. Pairwise byte-diffing
# of these samples is the primary tool for reverse-engineering Access's
# undocumented project p-code wire format.
#
# Output layout:
#   tests/live_access_test/re_corpus/
#     baseline_empty.accdb              -- freshly created, no VBA project
#     baseline_empty_proj.accdb         -- VBA project initialized, no modules
#     samples/<id>__<descr>.accdb       -- baseline_empty_proj + one mutation
#     samples/<id>__<descr>.bas         -- the source authored into that module
#     INDEX.tsv                         -- id, descr, op, args, source-hash
#
# Usage:
#   .\_corpus_generate.ps1 [-Force]
#
# `-Force` regenerates baselines and samples even if they already exist.

param([switch]$Force)

$ErrorActionPreference = "Stop"

$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Corpus = Join-Path $Here "re_corpus"
$Samples = Join-Path $Corpus "samples"
$Baseline = Join-Path $Corpus "baseline_empty.accdb"
$BaselineProj = Join-Path $Corpus "baseline_empty_proj.accdb"
$Index = Join-Path $Corpus "INDEX.tsv"

New-Item -ItemType Directory -Force -Path $Corpus, $Samples | Out-Null

function New-EmptyAccdb {
    param([string]$Path)
    if ((Test-Path $Path) -and -not $Force) { return }
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    try {
        # NewCurrentDatabase(filepath, fileformat=12 = Access 2007)
        $access.NewCurrentDatabase($Path, 12)
        $access.CloseCurrentDatabase()
    } finally {
        $access.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function New-EmptyProj {
    param([string]$SourcePath, [string]$DestPath)
    if ((Test-Path $DestPath) -and -not $Force) { return }
    Copy-Item -Path $SourcePath -Destination $DestPath -Force
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    try {
        $access.OpenCurrentDatabase($DestPath)
        # Touch the VBE to force project initialization. The project
        # already exists implicitly on .accdb open in modern Access.
        $proj = $access.VBE.ActiveVBProject
        $null = $proj.VBComponents.Count
        try { $access.RunCommand(126) } catch { }   # acCmdCompileAndSaveAllModules
        $access.CloseCurrentDatabase()
    } finally {
        $access.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Add-Sample {
    param(
        [string]$Id,
        [string]$Descr,
        [string]$ModuleName,
        [int]$Kind,            # 1 = vbext_ct_StdModule, 2 = vbext_ct_ClassModule
        [string]$Body
    )
    $Slug = ("{0}__{1}" -f $Id, ($Descr -replace '[^A-Za-z0-9._-]+','_'))
    $Out = Join-Path $Samples "$Slug.accdb"
    $BasOut = Join-Path $Samples "$Slug.bas"
    if ((Test-Path $Out) -and -not $Force) { return }
    Copy-Item -Path $BaselineProj -Destination $Out -Force
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    try {
        $access.OpenCurrentDatabase($Out)
        $proj = $access.VBE.ActiveVBProject
        $c = $proj.VBComponents.Add($Kind)
        $c.Name = $ModuleName
        if ($Body -ne $null -and $Body -ne "") {
            $c.CodeModule.AddFromString($Body)
        }
        try { $access.RunCommand(126) } catch { }
        $access.CloseCurrentDatabase()
        Set-Content -Path $BasOut -Value $Body -Encoding UTF8 -NoNewline
        Add-Content -Path $Index -Value ("{0}`t{1}`t{2}`t{3}`tkind={4}" -f $Id, $Descr, $ModuleName, $Body.Length, $Kind)
        Write-Host "  sample $Slug ($($Body.Length) source bytes)"
    } finally {
        $access.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ---------------------------------------------------------------------------
# 1. Baselines.
# ---------------------------------------------------------------------------
Write-Host "[1/3] baseline_empty.accdb"
New-EmptyAccdb $Baseline

Write-Host "[2/3] baseline_empty_proj.accdb (VBE project touched)"
New-EmptyProj $Baseline $BaselineProj

# ---------------------------------------------------------------------------
# 2. Samples. Each .accdb differs from baseline_empty_proj.accdb by exactly
#    one VBComponent insertion with the listed parameters. Sample IDs are
#    fixed and stable -- editing this list = appending new IDs.
# ---------------------------------------------------------------------------
Write-Host "[3/3] samples"
if (-not (Test-Path $Index) -or $Force) {
    Set-Content -Path $Index -Value "id`tdescr`tmodule_name`tsource_bytes`textras"
}

# 010..019: empty module with varying NAME lengths and characters
Add-Sample -Id "010" -Descr "empty_StdModule_M"        -ModuleName "M"        -Kind 1 -Body ""
Add-Sample -Id "011" -Descr "empty_StdModule_AB"       -ModuleName "AB"       -Kind 1 -Body ""
Add-Sample -Id "012" -Descr "empty_StdModule_ABC"      -ModuleName "ABC"      -Kind 1 -Body ""
Add-Sample -Id "013" -Descr "empty_StdModule_Mod1"     -ModuleName "Mod1"     -Kind 1 -Body ""
Add-Sample -Id "014" -Descr "empty_StdModule_Module1"  -ModuleName "Module1"  -Kind 1 -Body ""
Add-Sample -Id "015" -Descr "empty_StdModule_LongName" -ModuleName "ThisIsALongModuleName" -Kind 1 -Body ""
Add-Sample -Id "016" -Descr "empty_StdModule_UnicodeA" -ModuleName "M_a"      -Kind 1 -Body ""

# 020..029: empty CLASS module variants (vs std)
Add-Sample -Id "020" -Descr "empty_ClassModule_C"      -ModuleName "C"        -Kind 2 -Body ""
Add-Sample -Id "021" -Descr "empty_ClassModule_Class1" -ModuleName "Class1"   -Kind 2 -Body ""

# 030..039: single Sub body, varying SUB name only (same caller-empty body)
Add-Sample -Id "030" -Descr "sub_A_empty"              -ModuleName "M"        -Kind 1 -Body "Sub A()`r`nEnd Sub`r`n"
Add-Sample -Id "031" -Descr "sub_B_empty"              -ModuleName "M"        -Kind 1 -Body "Sub B()`r`nEnd Sub`r`n"
Add-Sample -Id "032" -Descr "sub_AB_empty"             -ModuleName "M"        -Kind 1 -Body "Sub AB()`r`nEnd Sub`r`n"
Add-Sample -Id "033" -Descr "sub_LongName_empty"       -ModuleName "M"        -Kind 1 -Body "Sub ThisIsALongName()`r`nEnd Sub`r`n"

# 040..049: single Sub body with ONE statement (build up the opcode vocab)
Add-Sample -Id "040" -Descr "sub_msgbox_hello"         -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    MsgBox `"hello`"`r`nEnd Sub`r`n"
Add-Sample -Id "041" -Descr "sub_msgbox_world"         -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    MsgBox `"world`"`r`nEnd Sub`r`n"
Add-Sample -Id "042" -Descr "sub_msgbox_long"          -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    MsgBox `"a much longer literal that is clearly not stored inline`"`r`nEnd Sub`r`n"
Add-Sample -Id "043" -Descr "sub_msgbox_two"           -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    MsgBox `"one`"`r`n    MsgBox `"two`"`r`nEnd Sub`r`n"
Add-Sample -Id "044" -Descr "sub_dim_int"              -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    Dim x As Integer`r`nEnd Sub`r`n"
Add-Sample -Id "045" -Descr "sub_dim_long"             -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    Dim x As Long`r`nEnd Sub`r`n"
Add-Sample -Id "046" -Descr "sub_dim_string"           -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    Dim x As String`r`nEnd Sub`r`n"
Add-Sample -Id "047" -Descr "sub_let_int"              -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    Dim x As Integer`r`n    x = 7`r`nEnd Sub`r`n"
Add-Sample -Id "048" -Descr "sub_let_int_42"           -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    Dim x As Integer`r`n    x = 42`r`nEnd Sub`r`n"
Add-Sample -Id "049" -Descr "sub_comment_only"         -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    ' a comment`r`nEnd Sub`r`n"

# 050..059: control flow
Add-Sample -Id "050" -Descr "sub_if_true"              -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    If True Then MsgBox `"y`"`r`nEnd Sub`r`n"
Add-Sample -Id "051" -Descr "sub_for_1_to_3"           -ModuleName "M"        -Kind 1 -Body "Sub A()`r`n    Dim i As Integer`r`n    For i = 1 To 3`r`n    Next i`r`nEnd Sub`r`n"

Write-Host ""
Write-Host "corpus generated -> $Samples"
Write-Host "  baseline: $Baseline"
Write-Host "  baseline+proj: $BaselineProj"
Write-Host "  index: $Index"
