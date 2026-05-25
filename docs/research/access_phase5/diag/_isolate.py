"""Generate three test files to isolate which mutation breaks PowerPoint."""
from pathlib import Path
from pyopenvba import PowerPointFile
from pyopenvba.vba import VBAModuleKind

SRC  = Path("tests/live_powerpoint_testing/Presentation1.pptm")
OUT  = Path("demo/output")
OUT.mkdir(parents=True, exist_ok=True)

# A: Only Module1 edited
with PowerPointFile(SRC) as p:
    p.set_module("Module1", 'Attribute VB_Name = "Module1"\r\nOption Explicit\r\nSub Hello()\r\n  MsgBox "A"\r\nEnd Sub\r\n')
    p.save(OUT / "A_module1_only.pptm")

# B: Module1 + UserForm1 body
with PowerPointFile(SRC) as p:
    p.set_module("Module1", 'Attribute VB_Name = "Module1"\r\nOption Explicit\r\nSub Hello()\r\n  MsgBox "B"\r\nEnd Sub\r\n')
    p.set_module("UserForm1",
                 "Option Explicit\r\n\r\n"
                 "Private Sub UserForm_Initialize()\r\n"
                 "    Me.Caption = \"B form\"\r\n"
                 "End Sub\r\n")
    p.save(OUT / "B_module1_and_userform1.pptm")

# C: Module1 + add DataModel class (no UserForm change)
with PowerPointFile(SRC) as p:
    p.set_module("Module1", 'Attribute VB_Name = "Module1"\r\nOption Explicit\r\nSub Hello()\r\n  Dim d As New DataModel\r\n  d.Tag = "x"\r\n  MsgBox d.Tag\r\nEnd Sub\r\n')
    p.vba_project().add_module(
        "DataModel",
        "VERSION 1.0 CLASS\r\n"
        "BEGIN\r\n"
        "  MultiUse = -1  'True\r\n"
        "END\r\n"
        'Attribute VB_Name = "DataModel"\r\n'
        "Attribute VB_GlobalNameSpace = False\r\n"
        "Attribute VB_Creatable = False\r\n"
        "Attribute VB_PredeclaredId = False\r\n"
        "Attribute VB_Exposed = False\r\n"
        "Option Explicit\r\n\r\n"
        "Public Tag As String\r\n",
        kind=VBAModuleKind.other,
    )
    p.save(OUT / "C_module1_plus_datamodel.pptm")

print("Generated A, B, C in demo/output/")
