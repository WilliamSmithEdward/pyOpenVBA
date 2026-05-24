"""
inject_xlsb_with_class_demo.py
--------------------------------
Open the live .xlsb fixture, replace Module1's body, add a fresh
DataModel class module, and write the result to demo/output/ for
live verification in Excel (Alt+F11).
"""

from pathlib import Path

from pyopenvba import ExcelFile
from pyopenvba.vba import VBAModuleKind

SRC = Path(__file__).parent.parent / "tests" / "live_excel_testing" / "test_macro_workbook.xlsb"
OUT = Path(__file__).parent / "output" / "injected_with_class.xlsb"

# Universal VBA class-module CLSID; required as VB_Base for plain class modules.
_CLASS_VB_BASE = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"

# ---------------------------------------------------------------------------
# Module1  (standard — body-only replacement)
# ---------------------------------------------------------------------------
MODULE1_SOURCE = """\
Option Explicit

Sub RunDemo()
    Dim d As New DataModel
    d.Tag   = "xlsb-demo"
    d.Score = 99
    MsgBox d.Describe(), vbInformation, "DataModel (xlsb)"
End Sub
"""

# ---------------------------------------------------------------------------
# DataModel  (class, kind=other — newly added)
# ---------------------------------------------------------------------------
DATAMODEL_HEADER = (
    'Attribute VB_Name = "DataModel"\r\n'
    f'Attribute VB_Base = "{_CLASS_VB_BASE}"\r\n'
    "Attribute VB_GlobalNameSpace = False\r\n"
    "Attribute VB_Creatable = False\r\n"
    "Attribute VB_PredeclaredId = False\r\n"
    "Attribute VB_Exposed = False\r\n"
    "Attribute VB_TemplateDerived = False\r\n"
    "Attribute VB_Customizable = False\r\n"
)
DATAMODEL_BODY = """\
Option Explicit

Private mTag   As String
Private mScore As Long

Property Get Tag() As String
    Tag = mTag
End Property

Property Let Tag(s As String)
    mTag = s
End Property

Property Get Score() As Long
    Score = mScore
End Property

Property Let Score(n As Long)
    mScore = n
End Property

Function Describe() As String
    Describe = "[" & mTag & "] score=" & mScore
End Function

Private Sub Class_Initialize()
    mTag   = "unset"
    mScore = 0
End Sub
"""


def main() -> None:
    import shutil

    OUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SRC, OUT)

    with ExcelFile(OUT) as wb:
        print(f"Before : {wb.module_names()}")
        wb.set_module("Module1", MODULE1_SOURCE)
        wb.vba_project().add_module(
            "DataModel",
            DATAMODEL_HEADER + DATAMODEL_BODY,
            kind=VBAModuleKind.other,
        )
        wb.save()

    # Verify round-trip
    with ExcelFile(OUT) as wb:
        names = wb.module_names()
        mod1  = wb.get_module("Module1")
        dm    = wb.get_module("DataModel")

    print(f"Written : {OUT}")
    print(f"  modules  : {names}")
    print(f"  Module1  : {len(mod1.splitlines())} lines  (standard)")
    print(f"  DataModel: {len(dm.splitlines())} lines  (class)")


if __name__ == "__main__":
    main()
