"""
Bake the nested-UserForm test fixture, using Excel itself.

The designer reader has to be tested against streams MSForms actually
wrote.  The interesting structure is nesting, and no hand-built file
would get it right: a Frame's children live in a storage of their own,
and a MultiPage owns a hidden TabStrip site plus one storage per Page.
So this drives live Excel through pyvbaharness and saves as .xlsm.

The control set is chosen to cover every shape the reader distinguishes:
a plain leaf control, a Frame with children, a MultiPage with two Pages
each holding a control, and a leaf that comes after the containers so the
site ordering is not accidentally sorted.

The option buttons share a GroupName and two controls carry a tip on
purpose.  Those set the GroupID and ControlTipTextData bits, which is
what leaves a site's DataBlock ending off a 4-byte boundary -- the only
case where the alignment before the name actually has to move, and so
the only case that can catch getting it wrong.

Run from the repo root, on Windows with desktop Excel installed:
    python scripts/bake_form_fixture.py

Dev-only: pyvbaharness is a test-time oracle and is never a runtime
dependency of pyOpenVBA.
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

OUT = ROOT / "tests" / "live_excel_testing" / "nested_form.xlsm"

# xlOpenXMLWorkbookMacroEnabled, and vbext_ct_MSForm.
XL_OPEN_XML_WORKBOOK_MACRO_ENABLED = 52
VBEXT_CT_MSFORM = 3

HARNESS_PREFIX = "PyVba"

BUILD_FORM = f"""\
Sub Main()
    Dim vbc As Object, d As Object, c As Object, p As Object
    Set vbc = ActiveWorkbook.VBProject.VBComponents.Add({VBEXT_CT_MSFORM})
    vbc.Name = "FrmNested"
    Set d = vbc.Designer
    d.Caption = "Nested Fixture"

    Set c = d.Controls.Add("Forms.Label.1", "TopLabel")
    c.Caption = "Top level"
    c.Left = 6: c.Top = 6

    Set c = d.Controls.Add("Forms.Frame.1", "GroupBox")
    c.Caption = "A frame"
    c.Left = 6: c.Top = 30: c.Width = 200: c.Height = 80
    Set p = c.Controls.Add("Forms.OptionButton.1", "OptOne")
    p.Caption = "One": p.Left = 6: p.Top = 12: p.GroupName = "Choice"
    Set p = c.Controls.Add("Forms.OptionButton.1", "OptTwo")
    p.Caption = "Two": p.Left = 6: p.Top = 34: p.GroupName = "Choice"
    Set p = c.Controls.Add("Forms.TextBox.1", "InnerText")
    p.Left = 6: p.Top = 56
    p.ControlTipText = "Type here"

    Set c = d.Controls.Add("Forms.MultiPage.1", "Pages")
    c.Left = 6: c.Top = 120: c.Width = 200: c.Height = 90
    Set p = c.Pages(0).Controls.Add("Forms.CheckBox.1", "PageOneCheck")
    p.Caption = "On page one": p.Left = 6: p.Top = 6
    Set p = c.Pages(1).Controls.Add("Forms.CommandButton.1", "PageTwoButton")
    p.Caption = "On page two": p.Left = 6: p.Top = 6

    Set c = d.Controls.Add("Forms.ListBox.1", "Picker")
    c.Left = 220: c.Top = 30: c.Width = 120: c.Height = 60
    c.RowSource = "Sheet1!A1:A3"
    c.ControlSource = "Sheet1!C1"

    Set c = d.Controls.Add("Forms.CommandButton.1", "CloseButton")
    c.Caption = "Close": c.Left = 6: c.Top = 220
    c.ControlTipText = "Closes the form"
End Sub
"""

CODE_BEHIND = (
    "Option Explicit\r\n\r\n"
    "Public Sub ShowNested()\r\n"
    "    FrmNested.Show\r\n"
    "End Sub\r\n"
)

SAVE_AS = """\
Sub Main()
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs "{path}", {fmt}
    Application.DisplayAlerts = True
End Sub
"""


def main() -> int:
    import pyvbaharness

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.unlink(missing_ok=True)
    session = pyvbaharness.ExcelSession()
    try:
        session.new_document()
        if not session.run_vba(BUILD_FORM, timeout=180.0).ok:
            raise SystemExit("Excel refused to build the form")
        session.add_module("FormHost", CODE_BEHIND)
        saver = SAVE_AS.format(
            path=OUT.resolve(), fmt=XL_OPEN_XML_WORKBOOK_MACRO_ENABLED
        )
        if not session.run_vba(saver, timeout=180.0).ok:
            raise SystemExit("Excel refused to save the workbook")
    finally:
        try:
            session.close()
        except Exception:
            pass
    if not OUT.exists():
        raise SystemExit(f"Excel wrote no file to {OUT}")
    strip_harness_modules(OUT)
    print(f"baked {OUT.relative_to(ROOT)} ({OUT.stat().st_size} bytes)")
    return 0


def strip_harness_modules(path: Path) -> None:
    """Drop the harness's runner modules, then prove the form survived.

    The harness cannot remove them itself -- the code doing the removing
    is one of them -- so the library's own write path does it, and the
    reopen below is what makes that safe to rely on.
    """
    from pyopenvba.excel import ExcelFile

    with ExcelFile(path) as workbook:
        project = workbook.vba_project()
        for name in workbook.module_names():
            if name.startswith(HARNESS_PREFIX):
                project.delete_module(name)
        workbook.save()
    with ExcelFile(path) as workbook:
        left = [n for n in workbook.module_names() if n.startswith(HARNESS_PREFIX)]
        if left:
            raise SystemExit(f"harness modules survived the strip: {left}")
        forms = workbook.forms()
        if len(forms) != 1 or forms[0].name != "FrmNested":
            raise SystemExit(f"expected one form named FrmNested, got {forms}")
        names = {c.name for c in forms[0].walk()}
        expected = {
            "TopLabel", "GroupBox", "OptOne", "OptTwo", "InnerText",
            "Pages", "Page1", "Page2", "PageOneCheck", "PageTwoButton",
            "Picker", "CloseButton",
        }
        missing = expected - names
        if missing:
            raise SystemExit(f"controls lost in the strip: {sorted(missing)}")


if __name__ == "__main__":
    raise SystemExit(main())
