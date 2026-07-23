param(
    [Parameter(Mandatory = $true)]
    [int]$ExcelProcessId,

    [Parameter(Mandatory = $true)]
    [string]$LogPath,

    [Parameter(Mandatory = $true)]
    [string]$StopPath,

    [int]$TimeoutSeconds = 120,

    [switch]$DismissKnownDialogs,

    [switch]$TerminateOnBreakMode
)

$ErrorActionPreference = "Stop"

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public sealed class ROneCOneChildWindow
{
    public long Handle { get; set; }
    public string ClassName { get; set; }
    public string Text { get; set; }
}

public sealed class ROneCOneTopWindow
{
    public long Handle { get; set; }
    public string ClassName { get; set; }
    public string Title { get; set; }
    public List<ROneCOneChildWindow> Children { get; set; }
}

public static class ROneCOneWindowProbe
{
    private const uint BM_CLICK = 0x00F5;
    private const uint WM_COMMAND = 0x0111;
    private const uint WM_CLOSE = 0x0010;
    private const int IDOK = 1;

    private delegate bool EnumWindowsProc(IntPtr handle, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(
        IntPtr parent,
        EnumWindowsProc callback,
        IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr handle, StringBuilder text, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr handle, StringBuilder text, int maxCount);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr handle);

    [DllImport("user32.dll")]
    private static extern IntPtr SendMessage(
        IntPtr handle,
        uint message,
        IntPtr wordParameter,
        IntPtr longParameter);

    public static List<ROneCOneTopWindow> Snapshot(uint processId)
    {
        var result = new List<ROneCOneTopWindow>();
        EnumWindows(delegate(IntPtr handle, IntPtr parameter)
        {
            uint owner;
            GetWindowThreadProcessId(handle, out owner);
            if (owner != processId || !IsWindowVisible(handle))
            {
                return true;
            }

            var record = new ROneCOneTopWindow
            {
                Handle = handle.ToInt64(),
                ClassName = ReadClassName(handle),
                Title = ReadText(handle),
                Children = new List<ROneCOneChildWindow>()
            };

            EnumChildWindows(handle, delegate(IntPtr child, IntPtr childParameter)
            {
                if (IsWindowVisible(child))
                {
                    record.Children.Add(new ROneCOneChildWindow
                    {
                        Handle = child.ToInt64(),
                        ClassName = ReadClassName(child),
                        Text = ReadText(child)
                    });
                }
                return true;
            }, IntPtr.Zero);

            result.Add(record);
            return true;
        }, IntPtr.Zero);
        return result;
    }

    public static uint ProcessIdForWindow(long windowHandle)
    {
        uint processId;
        GetWindowThreadProcessId(new IntPtr(windowHandle), out processId);
        return processId;
    }

    public static string DismissSafely(long dialogHandle)
    {
        IntPtr dialog = new IntPtr(dialogHandle);
        var buttons = new List<KeyValuePair<IntPtr, string>>();

        EnumChildWindows(dialog, delegate(IntPtr child, IntPtr parameter)
        {
            string className = ReadClassName(child);
            string text = ReadText(child).Replace("&", "").Trim();
            if (className == "Button" && text.Length > 0)
            {
                buttons.Add(new KeyValuePair<IntPtr, string>(child, text));
            }
            return true;
        }, IntPtr.Zero);

        string[] conservative = new[] { "Cancel", "Close" };
        foreach (string choice in conservative)
        {
            foreach (var button in buttons)
            {
                if (string.Equals(button.Value, choice, StringComparison.OrdinalIgnoreCase))
                {
                    SendMessage(button.Key, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                    return "click:" + choice;
                }
            }
        }

        var okButtons = buttons.FindAll(button =>
            string.Equals(button.Value, "OK", StringComparison.OrdinalIgnoreCase));
        var decisionButtons = buttons.FindAll(button =>
            !string.Equals(button.Value, "Help", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(button.Value, "Details", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(button.Value, "Copy", StringComparison.OrdinalIgnoreCase));
        if (okButtons.Count == 1 && decisionButtons.Count == 1)
        {
            SendMessage(okButtons[0].Key, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
            return "click:OK";
        }

        SendMessage(dialog, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        return "close-window";
    }

    private static string ReadText(IntPtr handle)
    {
        var buffer = new StringBuilder(4096);
        GetWindowText(handle, buffer, buffer.Capacity);
        return buffer.ToString();
    }

    private static string ReadClassName(IntPtr handle)
    {
        var buffer = new StringBuilder(256);
        GetClassName(handle, buffer, buffer.Capacity);
        return buffer.ToString();
    }
}
'@

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

function Get-AutomationSurface {
    param([long]$Handle)

    $records = @()
    try {
        $root = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$Handle)
        if ($null -eq $root) {
            return $records
        }
        $elements = $root.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition)
        $limit = [Math]::Min($elements.Count, 100)
        for ($index = 0; $index -lt $limit; $index++) {
            $element = $elements.Item($index)
            try {
                $name = [string]$element.Current.Name
                $controlType = [string]$element.Current.ControlType.ProgrammaticName
                $automationId = [string]$element.Current.AutomationId
                $selectedText = ""
                if ($controlType -eq "ControlType.Document") {
                    $textPatternObject = $null
                    if ($element.TryGetCurrentPattern(
                        [System.Windows.Automation.TextPattern]::Pattern,
                        [ref]$textPatternObject)) {
                        $selection = ([System.Windows.Automation.TextPattern]$textPatternObject).GetSelection()
                        if ($selection.Count -gt 0) {
                            $selectedText = [string]$selection.Item(0).GetText(500)
                        }
                    }
                }
                if ($name.Length -gt 0 -or $automationId.Length -gt 0) {
                    $records += [ordered]@{
                        name = $name
                        control_type = $controlType
                        automation_id = $automationId
                        enabled = [bool]$element.Current.IsEnabled
                        selected_text = $selectedText
                    }
                }
            }
            catch {
                continue
            }
        }
    }
    catch {
        $records += [ordered]@{
            inspection_error = $_.Exception.Message
        }
    }
    return $records
}

function Get-OwnedVbeSelection {
    param([int]$ExpectedProcessId)

    $application = $null
    $vbe = $null
    $codePane = $null
    $codeModule = $null
    try {
        $application = [Runtime.InteropServices.Marshal]::GetActiveObject(
            "Excel.Application")
        $actualProcessId = [ROneCOneWindowProbe]::ProcessIdForWindow(
            [long]$application.Hwnd)
        if ($actualProcessId -ne $ExpectedProcessId) {
            return ""
        }
        $vbe = $application.VBE
        $codePane = $vbe.ActiveCodePane
        if ($null -eq $codePane) {
            return ""
        }
        [int]$startLine = 0
        [int]$startColumn = 0
        [int]$endLine = 0
        [int]$endColumn = 0
        $codePane.GetSelection(
            [ref]$startLine,
            [ref]$startColumn,
            [ref]$endLine,
            [ref]$endColumn)
        $codeModule = $codePane.CodeModule
        $lineText = [string]$codeModule.Lines($startLine, 1)
        return "line=$startLine; columns=$startColumn-$endColumn; code=$lineText"
    }
    catch {
        return "selection-error=$($_.Exception.Message)"
    }
    finally {
        foreach ($comObject in @($codeModule, $codePane, $vbe, $application)) {
            if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject(
                $comObject)) {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
            }
        }
    }
}

$resolvedLog = [System.IO.Path]::GetFullPath($LogPath)
$resolvedStop = [System.IO.Path]::GetFullPath($StopPath)
$logDirectory = Split-Path -Parent $resolvedLog
[System.IO.Directory]::CreateDirectory($logDirectory) | Out-Null
$seen = @{}
$deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
$popupObserved = $false

while ([DateTime]::UtcNow -lt $deadline) {
    if (Test-Path -LiteralPath $resolvedStop) {
        break
    }
    if ($null -eq (Get-Process -Id $ExcelProcessId -ErrorAction SilentlyContinue)) {
        break
    }

    $windows = [ROneCOneWindowProbe]::Snapshot([uint32]$ExcelProcessId)
    foreach ($window in $windows) {
        if ($window.ClassName -eq "XLMAIN") {
            continue
        }
        if ($window.Title.Length -eq 0 -and $window.Children.Count -eq 0) {
            continue
        }

        $childText = @(
            $window.Children |
                Where-Object { $_.Text.Length -gt 0 } |
                ForEach-Object { $_.Text }
        )
        $automationSurface = @(Get-AutomationSurface -Handle $window.Handle)
        $automationText = @($automationSurface | ForEach-Object { $_.name })
        $automationSelection = @(
            $automationSurface | ForEach-Object { $_.selected_text }
        )
        $fingerprint = "$($window.Handle)|$($childText -join '|')|" +
            "$($automationText -join '|')|$($automationSelection -join '|')"
        if (-not $seen.ContainsKey($fingerprint)) {
            $dismissalAction = "none"
            $hasWin32Button = @(
                $window.Children | Where-Object { $_.ClassName -eq "Button" }
            ).Count -gt 0
            $hasAutomationButton = @(
                $automationSurface | Where-Object {
                    $_.control_type -eq "ControlType.Button"
                }
            ).Count -gt 0
            $isVbeMainWindow = $window.ClassName -eq "wndclass_desked_gsk"
            $looksModal = -not $isVbeMainWindow -and (
                $window.ClassName -eq "#32770" -or `
                $hasWin32Button -or $hasAutomationButton)
            if ($DismissKnownDialogs -and $looksModal) {
                $dismissalAction = [ROneCOneWindowProbe]::DismissSafely($window.Handle)
                $popupObserved = $true
            }
            $record = [ordered]@{
                observed_at = [DateTime]::UtcNow.ToString("o")
                process_id = $ExcelProcessId
                handle = $window.Handle
                class_name = $window.ClassName
                title = $window.Title
                child_text = $childText
                children = $window.Children
                automation = $automationSurface
                dismissal_action = $dismissalAction
            }
            Add-Content -LiteralPath $resolvedLog -Value ($record | ConvertTo-Json -Compress -Depth 5)
            $seen[$fingerprint] = $true
        }

        if ($TerminateOnBreakMode -and $popupObserved -and $window.Title -match "\[break\]") {
            $vbeSelection = Get-OwnedVbeSelection -ExpectedProcessId $ExcelProcessId
            $breakRecord = [ordered]@{
                observed_at = [DateTime]::UtcNow.ToString("o")
                process_id = $ExcelProcessId
                handle = $window.Handle
                class_name = $window.ClassName
                title = $window.Title
                child_text = @($vbeSelection)
                children = @()
                automation = @()
                dismissal_action = "terminate-task-excel-after-break"
            }
            Add-Content -LiteralPath $resolvedLog -Value ($breakRecord | ConvertTo-Json -Compress -Depth 5)
            Stop-Process -Id $ExcelProcessId -Force -ErrorAction SilentlyContinue
            break
        }
    }

    Start-Sleep -Milliseconds 100
}
