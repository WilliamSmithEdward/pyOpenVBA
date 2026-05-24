"""
inject_all_module_types.py
--------------------------
Open Presentation1.pptm (which already has Module1 + UserForm1), inject
real code into every existing module, add a new Class module, and write
the result to demo/output/all_types_from_presentation1.pptm.
"""

from pathlib import Path

from pyopenvba import PowerPointFile
from pyopenvba.vba import VBAModuleKind

SRC  = Path(__file__).parent.parent / "tests" / "live_powerpoint_testing" / "Presentation1.pptm"
OUT  = Path(__file__).parent / "output" / "all_types_from_presentation1.pptm"

# ---------------------------------------------------------------------------
# Module1  (standard module)
# ---------------------------------------------------------------------------
MODULE1_CODE = """\
Option Explicit

' ---- Slide utilities (injected by pyOpenVBA) ----

Function SlideCount() As Integer
    SlideCount = ActivePresentation.Slides.Count
End Function

Function SlideTitle(idx As Integer) As String
    On Error Resume Next
    SlideTitle = ActivePresentation.Slides(idx).Name
    On Error GoTo 0
End Function

Sub ShowStats()
    Dim msg As String
    msg = "Slides: " & SlideCount() & Chr(10)
    msg = msg & "First slide: " & SlideTitle(1)
    MsgBox msg, vbInformation, "Presentation Stats"
End Sub

Sub RunAll()
    Dim m As New DataModel
    m.Tag   = "demo"
    m.Score = 42
    MsgBox m.Describe(), vbInformation, "DataModel"
    ShowStats
End Sub
"""

# ---------------------------------------------------------------------------
# DataModel  (new Class module added programmatically)
# ---------------------------------------------------------------------------
DATAMODEL_HEADER = (
    'Attribute VB_Name = "DataModel"\r\n'
    'Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"\r\n'
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

# ---------------------------------------------------------------------------
# UserForm1  (body-only; VB_Base header is preserved automatically)
# ---------------------------------------------------------------------------
USERFORM1_BODY = """\
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "pyOpenVBA Demo"
End Sub

Private Sub UserForm_Activate()
    MsgBox "Form activated! Slides: " & ActivePresentation.Slides.Count, _
           vbInformation, Me.Caption
End Sub
"""


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)

    with PowerPointFile(SRC) as prs:
        proj = prs.vba_project()

        # 1. Standard module — body-only inject
        prs.set_module("Module1", MODULE1_CODE)

        # 2. UserForm — body-only inject (VB_Base is preserved automatically)
        prs.set_module("UserForm1", USERFORM1_BODY)

        # 3. New Class module — add_module requires full header for non-standard kinds
        proj.add_module(
            "DataModel",
            DATAMODEL_HEADER + DATAMODEL_BODY,
            kind=VBAModuleKind.other,
        )

        prs.save(OUT)

    # Verify round-trip
    with PowerPointFile(OUT) as prs:
        names = prs.module_names()
        mod1  = prs.get_module("Module1")
        dm    = prs.get_module("DataModel")
        uf    = prs.get_module("UserForm1")

    print(f"Written: {OUT}")
    print(f"  modules  : {names}")
    print(f"  Module1  : {len(mod1.splitlines())} lines  (standard)")
    print(f"  DataModel: {len(dm.splitlines())} lines  (class, kind=other)")
    print(f"  UserForm1: {len(uf.splitlines())} lines  (form,  kind=other)")
    uf_src = uf
    assert "VB_Base"             in uf_src, "UserForm1 VB_Base lost"
    assert "UserForm_Initialize" in uf_src
    assert "SlideCount"          in mod1
    assert "Class_Initialize"    in dm
    print("  All assertions passed.")


if __name__ == "__main__":
    main()
