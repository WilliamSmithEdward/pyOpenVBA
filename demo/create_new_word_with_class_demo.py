"""
create_new_word_with_class_demo.py
----------------------------------
Build a brand-new .docm via WordFile.create_new(), inject a standard
module + a freshly-added class module, and write the result to
demo/output/new_word_with_class.docm for live verification in Word
(Alt+F11).
"""

from pathlib import Path

from pyopenvba import WordFile
from pyopenvba.vba import VBAModuleKind

OUT = Path(__file__).parent / "output" / "new_word_with_class.docm"

# Universal VBA class-module CLSID; required as VB_Base for plain class modules.
_CLASS_VB_BASE = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"

# ---------------------------------------------------------------------------
# Module1  (standard)
# ---------------------------------------------------------------------------
MODULE1_SOURCE = """\
Option Explicit

Sub RunDemo()
    Dim d As New DataModel
    d.Tag   = "demo"
    d.Score = 42
    MsgBox d.Describe(), vbInformation, "DataModel"
End Sub
"""

# ---------------------------------------------------------------------------
# DataModel  (class, kind=other)
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
    OUT.parent.mkdir(parents=True, exist_ok=True)

    with WordFile.create_new(OUT) as doc:
        doc.set_module("Module1", MODULE1_SOURCE)
        doc.vba_project().add_module(
            "DataModel",
            DATAMODEL_HEADER + DATAMODEL_BODY,
            kind=VBAModuleKind.other,
        )
        doc.save()

    # Verify round-trip
    with WordFile(OUT) as doc:
        names = doc.module_names()
        mod1  = doc.get_module("Module1")
        dm    = doc.get_module("DataModel")

    print(f"Written: {OUT}")
    print(f"  modules  : {names}")
    print(f"  Module1  : {len(mod1.splitlines())} lines  (standard)")
    print(f"  DataModel: {len(dm.splitlines())} lines  (class, kind=other)")
    assert "RunDemo"          in mod1
    assert "Class_Initialize" in dm
    assert "VB_Base"          in dm
    print("  All assertions passed.")


if __name__ == "__main__":
    main()
