"""
create_new_with_class_demo.py
-----------------------------
Build a brand-new .pptm from scratch via PowerPointFile.create_new(),
inject a standard module + a freshly-added class module, and write the
result to demo/output/new_with_class.pptm for live verification in
PowerPoint (Alt+F11).
"""

from pathlib import Path

from pyopenvba import PowerPointFile
from pyopenvba.vba import VBAModuleKind

OUT = Path(__file__).parent / "output" / "new_with_class.pptm"

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
# DataModel  (class, kind=other — header synthesized automatically)
# ---------------------------------------------------------------------------
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

    with PowerPointFile.create_new(OUT) as prs:
        prs.set_module("Module1", MODULE1_SOURCE)
        prs.vba_project().add_module(
            "DataModel",
            DATAMODEL_BODY,
            kind=VBAModuleKind.other,
        )
        prs.save()

    # Verify round-trip
    with PowerPointFile(OUT) as prs:
        names = prs.module_names()
        mod1  = prs.get_module("Module1")
        dm    = prs.get_module("DataModel")

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
