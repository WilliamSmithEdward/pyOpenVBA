"""Build a form whose every property holds a value nothing else holds, so
that matching a record to a text line by value is unambiguous."""

import shutil
import pathlib
import sys
from pathlib import Path

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[3] / "src"))

import pyvbaharness  # noqa: E402

#: Where the database and its text export are written.
HERE = Path(".").resolve()
TEMPLATE = (
    pathlib.Path(__file__).resolve().parents[3]
    / "src"
    / "pyopenvba"
    / "_templates"
    / "blank_files"
    / "blank_database.accdb"
)

KINDS = [
    ("Label", 100),
    ("TextBox", 109),
    ("CommandButton", 104),
    ("ToggleButton", 122),
    ("OptionButton", 105),
    ("CheckBox", 106),
    ("OptionGroup", 107),
    ("ListBox", 110),
    ("ComboBox", 111),
    ("Rectangle", 101),
    ("Line", 102),
    ("Image", 103),
    ("BoundObjectFrame", 108),
    ("ObjectFrame", 114),
    ("Subform", 112),
    ("Tab", 123),
]

VBA = """
Public Function BuildRich() As String
    Dim f As Object, c As Object, i As Long
    Dim kinds As Variant, names As Variant
    kinds = Array({codes})
    names = Array({names})
    Set f = CreateForm()
    For i = LBound(kinds) To UBound(kinds)
        On Error Resume Next
        Err.Clear
        Set c = CreateControl(f.Name, kinds(i), 0, "", "", 211, 307 + i * 400, 1409, 313)
        If Err.Number <> 0 Then GoTo NextOne
        c.Name = "X" & names(i)
        SetOne c, "ControlTipText", "tip-" & names(i)
        SetOne c, "Tag", "tag-" & names(i)
        SetOne c, "StatusBarText", "bar-" & names(i)
        SetOne c, "ValidationText", "val-" & names(i)
        SetOne c, "ValidationRule", ">0"
        SetOne c, "Caption", "cap-" & names(i)
        SetOne c, "ControlSource", "src-" & names(i)
        SetOne c, "RowSource", "row-" & names(i)
        SetOne c, "DefaultValue", "def-" & names(i)
        SetOne c, "Format", "General Number"
        SetOne c, "InputMask", "0000\\-0000"
        SetOne c, "ColumnWidths", "1417;2835"
        SetOne c, "FontName", "Consolas"
        SetOne c, "FontSize", 17
        SetOne c, "FontWeight", 700
        SetOne c, "FontItalic", True
        SetOne c, "FontUnderline", True
        SetOne c, "ForeColor", 3355443
        SetOne c, "BackColor", 13421772
        SetOne c, "BorderColor", 6710886
        SetOne c, "GridlineColor", 10066329
        SetOne c, "HoverColor", 5592405
        SetOne c, "PressedColor", 11184810
        SetOne c, "BorderWidth", 5
        SetOne c, "BorderStyle", 3
        SetOne c, "GridlineStyle", 2
        SetOne c, "SpecialEffect", 4
        SetOne c, "TextAlign", 2
        SetOne c, "ScrollBars", 2
        SetOne c, "DecimalPlaces", 7
        SetOne c, "ColumnCount", 3
        SetOne c, "ColumnHeads", True
        SetOne c, "ListRows", 9
        SetOne c, "ListWidth", 2531
        SetOne c, "LimitToList", False
        SetOne c, "LineSpacing", 23
        SetOne c, "LeftMargin", 31
        SetOne c, "TopMargin", 37
        SetOne c, "RightMargin", 41
        SetOne c, "BottomMargin", 43
        SetOne c, "CanGrow", True
        SetOne c, "CanShrink", True
        SetOne c, "HideDuplicates", True
        SetOne c, "DisplayWhen", 1
        SetOne c, "Enabled", False
        SetOne c, "Locked", True
        SetOne c, "TabStop", False
        SetOne c, "Visible", True
        SetOne c, "AddColon", False
        SetOne c, "AutoLabel", False
        SetOne c, "LabelAlign", 1
        SetOne c, "LabelX", 53
        SetOne c, "LabelY", 59
        SetOne c, "IMEMode", 1
        SetOne c, "IMESentenceMode", 3
NextOne:
        On Error GoTo 0
    Next i
    On Error Resume Next
    f.Caption = "cap-Form"
    f.RecordSelectors = False
    f.NavigationButtons = False
    f.NavigationCaption = "nav-Form"
    f.AutoCenter = True
    f.PopUp = True
    f.Modal = True
    f.CloseButton = False
    f.MinMaxButtons = 1
    f.ControlBox = False
    f.GridX = 19
    f.GridY = 29
    f.RecordSource = "src-Form"
    f.Filter = "1=1"
    f.OrderBy = "ord-Form"
    On Error GoTo 0
    DoCmd.Save acForm, f.Name
    BuildRich = f.Name
    DoCmd.Close acForm, f.Name, acSaveYes
End Function

Private Sub SetOne(ByVal c As Object, ByVal prop As String, ByVal value As Variant)
    On Error Resume Next
    c.Properties(prop) = value
    On Error GoTo 0
End Sub

Public Function DumpText(ByVal name As String, ByVal target As String) As String
    Application.SaveAsText acForm, name, target
    DumpText = "ok"
End Function
"""


def main() -> None:
    target = HERE / "rich.accdb"
    shutil.copy(TEMPLATE, target)
    source = VBA.format(
        codes=", ".join(str(code) for _, code in KINDS),
        names=", ".join(f'"{name}"' for name, _ in KINDS),
    )
    text = HERE / "rich.txt"
    with pyvbaharness.AccessSession() as access:
        access.open_document(target, read_only=False)
        built = access.run_vba(source, proc="BuildRich", timeout=600.0)
        print("build:", built.outcome, built.value)
        if getattr(built, "dialogs", None):
            print("dialogs:", [getattr(d, "message", d) for d in built.dialogs])
        dumped = access.run_vba(
            source, proc="DumpText", args=(str(built.value), str(text)), timeout=300.0
        )
        print("saveastext:", dumped.outcome, dumped.value)
    print("text:", text.exists() and text.stat().st_size)


if __name__ == "__main__":
    main()
