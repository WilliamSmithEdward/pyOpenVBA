# Ask Access what colour each theme slot takes at every whole tint and
# shade: a command button's BackColor, read after setting its
# BackThemeColorIndex and BackTint or BackShade in design view.  Access
# works on a scratch copy of -Source, never on -Source itself.
#
# Writes one line per answer: slot,tint|shade,percent,BackColor (Access's
# BGR long).
#
# Usage: .\_theme_tints.ps1 -Source <database> -Out <file.csv>

param([Parameter(Mandatory=$true)][string]$Source, [Parameter(Mandatory=$true)][string]$Out)

$ErrorActionPreference = "Stop"
$work = Join-Path ([System.IO.Path]::GetTempPath()) ("theme_tints_" + [guid]::NewGuid().ToString("N") + ".accdb")
Copy-Item -Force $Source $work
$access = New-Object -ComObject Access.Application
$access.Visible = $false
$lines = New-Object System.Collections.Generic.List[string]
try {
    $access.OpenCurrentDatabase($work)
    $form = $access.CreateForm()
    $button = $access.CreateControl($form.Name, 104, 0, "", "", 100, 100, 1000, 400)   # acCommandButton
    for ($slot = 0; $slot -le 11; $slot++) {
        for ($value = 0; $value -le 100; $value++) {
            $button.BackThemeColorIndex = $slot
            $button.BackShade = 100
            $button.BackTint = $value
            $lines.Add("$slot,tint,$value,$($button.BackColor)")
            $button.BackTint = 100
            $button.BackShade = $value
            $lines.Add("$slot,shade,$value,$($button.BackColor)")
        }
    }
    $access.DoCmd.Close(2, $form.Name, 2)   # acForm, acSaveNo
    $access.CloseCurrentDatabase()
}
finally {
    $access.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Remove-Item -Force $work -ErrorAction SilentlyContinue
}
[System.IO.File]::WriteAllLines($Out, $lines)
Write-Host "wrote $($lines.Count) answers to $Out"
