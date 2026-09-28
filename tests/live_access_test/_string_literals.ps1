# Build string_literals.accdb: a database whose compiled modules hold string
# literals of every shape the reader has to decode -- plain, empty, with a
# doubled quote, with a character past ASCII, of odd and even length,
# several on one line, and one in a class module -- and a comment that only
# looks like one.  Access compiles and saves the modules, so each module's
# stream carries the p-code with its LitStr records.
#
# Access is driven over COM alone: pyvbaharness would inject modules of its
# own into the project it runs code in.
#
# Usage: .\_string_literals.ps1 -Dest <directory>

param([Parameter(Mandatory=$true)][string]$Dest)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $Dest | Out-Null
$path = Join-Path (Resolve-Path $Dest).Path "string_literals.accdb"
if (Test-Path $path) { Remove-Item -Force $path }

$crlf = [string][char]13 + [char]10
$e_acute = [string][char]0x00E9
$quote = [string][char]34
$module = @(
    "Public Function Greeting() As String",
    "    Greeting = ${quote}Hello${quote}",
    "End Function",
    "Public Function Nothing_() As String",
    "    Nothing_ = ${quote}${quote}",
    "End Function",
    "Public Function Quoted() As String",
    "    Quoted = ${quote}say ${quote}${quote}hi${quote}${quote}${quote}",
    "End Function",
    "Public Function Accented() As String",
    "    Accented = ${quote}caf${e_acute}${quote}",
    "End Function",
    "Public Function Joined() As String",
    "    Joined = ${quote}a${quote} & ${quote}bc${quote} & ${quote}def${quote}",
    "End Function",
    "Public Sub Remark()",
    "    ' ${quote}not a literal${quote}",
    "End Sub",
    "Public Sub Show()",
    "    MsgBox ${quote}hello${quote}",
    "End Sub",
    "Public Function AsVariant() As Variant",
    "    AsVariant = ${quote}variant${quote}",
    "End Function"
) -join $crlf
$class = @(
    "Public Function Name_() As String",
    "    Name_ = ${quote}from the class${quote}",
    "End Function"
) -join $crlf

$access = New-Object -ComObject Access.Application
$access.Visible = $false
try {
    $access.NewCurrentDatabase($path, 12)   # acNewDatabaseFormatAccess2007
    $project = $access.VBE.ActiveVBProject
    $standard = $project.VBComponents.Add(1)   # vbext_ct_StdModule
    $standard.Name = "Literals"
    $standard.CodeModule.AddFromString($module)
    $classModule = $project.VBComponents.Add(2)   # vbext_ct_ClassModule
    $classModule.Name = "Holder"
    $classModule.CodeModule.AddFromString($class)
    $access.DoCmd.RunCommand(126)   # acCmdCompileAndSaveAllModules
    $access.CloseCurrentDatabase()
}
finally {
    $access.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Write-Host "wrote $path"
