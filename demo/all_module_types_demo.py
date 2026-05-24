"""
all_module_types_demo.py
------------------------
Write two Office files that exercise every VBA module kind pyOpenVBA
supports — one Word document and one PowerPoint presentation.

Open them in Office 365 to confirm the injected code is intact and
runnable from the VBA IDE (Alt+F11).

    python demo/all_module_types_demo.py
    python demo/all_module_types_demo.py --out-dir my_output
"""

import argparse
from pathlib import Path

from pyopenvba import WordFile, PowerPointFile


# ---------------------------------------------------------------------------
# Word: Module1 (standard) + ThisDocument (document / other)
# ---------------------------------------------------------------------------

_WORD_MODULE1 = (
    "Option Explicit\r\n"
    "\r\n"
    "' ---- Utility functions injected by pyOpenVBA ----\r\n"
    "\r\n"
    "Function DoubleIt(n As Long) As Long\r\n"
    "    DoubleIt = n * 2\r\n"
    "End Function\r\n"
    "\r\n"
    "Function Cube(n As Long) As Long\r\n"
    "    Cube = n * n * n\r\n"
    "End Function\r\n"
    "\r\n"
    "Function Factorial(n As Long) As Long\r\n"
    "    If n <= 1 Then\r\n"
    "        Factorial = 1\r\n"
    "    Else\r\n"
    "        Factorial = n * Factorial(n - 1)\r\n"
    "    End If\r\n"
    "End Function\r\n"
    "\r\n"
    "Sub RunDemo()\r\n"
    "    MsgBox \"DoubleIt(7)=\" & DoubleIt(7) & Chr(10) &\r\n"
    "           \"Cube(4)=\" & Cube(4) & Chr(10) &\r\n"
    "           \"Factorial(6)=\" & Factorial(6)\r\n"
    "End Sub\r\n"
)

_WORD_THIS_DOCUMENT_BODY = (
    "' ---- Document event handlers injected by pyOpenVBA ----\r\n"
    "\r\n"
    "Private Sub Document_Open()\r\n"
    "    MsgBox \"Document opened!  pyOpenVBA wrote this handler.\"\r\n"
    "End Sub\r\n"
    "\r\n"
    "Private Sub Document_Close()\r\n"
    "    MsgBox \"Document closing.  Goodbye!\"\r\n"
    "End Sub\r\n"
    "\r\n"
    "Private Sub Document_New()\r\n"
    "    MsgBox \"New document created.\"\r\n"
    "End Sub\r\n"
)


def create_word_all_modules(out: Path) -> None:
    with WordFile.create_new(out) as doc:
        doc.set_module("Module1",      _WORD_MODULE1)
        doc.set_module("ThisDocument", _WORD_THIS_DOCUMENT_BODY)
        doc.save()
    # verify round-trip
    with WordFile(out) as doc:
        names = doc.module_names()
        mod1 = doc.get_module("Module1")
        this = doc.get_module("ThisDocument")
    assert "DoubleIt"       in mod1,  "Module1 body missing"
    assert "Document_Open"  in this,  "ThisDocument body missing"
    assert "VB_Base"        in this,  "ThisDocument VB_Base lost"
    print(f"  [word]  {out}")
    print(f"          modules : {names}")
    print(f"          Module1 : {len(mod1.splitlines())} lines  (standard)")
    print(f"          ThisDocument: {len(this.splitlines())} lines  "
          f"(other/document, VB_Base preserved)")


# ---------------------------------------------------------------------------
# PowerPoint: Module1 (standard)
# ---------------------------------------------------------------------------

_PPT_MODULE1 = (
    "Option Explicit\r\n"
    "\r\n"
    "' ---- Presentation-level utilities injected by pyOpenVBA ----\r\n"
    "\r\n"
    "Sub ShowSlideCount()\r\n"
    "    MsgBox \"Slides in this presentation: \" & _\r\n"
    "           ActivePresentation.Slides.Count\r\n"
    "End Sub\r\n"
    "\r\n"
    "Function SlideTitle(idx As Integer) As String\r\n"
    "    SlideTitle = ActivePresentation.Slides(idx).Name\r\n"
    "End Function\r\n"
    "\r\n"
    "Sub RunDemo()\r\n"
    "    MsgBox \"Slide count: \" & ActivePresentation.Slides.Count\r\n"
    "End Sub\r\n"
)


def create_ppt_all_modules(out: Path) -> None:
    with PowerPointFile.create_new(out) as prs:
        prs.set_module("Module1", _PPT_MODULE1)
        prs.save()
    with PowerPointFile(out) as prs:
        names = prs.module_names()
        mod1  = prs.get_module("Module1")
    assert "ShowSlideCount" in mod1, "Module1 body missing"
    print(f"  [pptm]  {out}")
    print(f"          modules : {names}")
    print(f"          Module1 : {len(mod1.splitlines())} lines  (standard)")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create Office files with all VBA module types populated."
    )
    parser.add_argument(
        "--out-dir",
        default=str(Path(__file__).parent / "output"),
        help="Output directory (default: demo/output/)",
    )
    args = parser.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"Writing files to: {out_dir.resolve()}\n")

    create_word_all_modules(out_dir / "all_modules_word.docm")
    print()
    create_ppt_all_modules(out_dir / "all_modules_ppt.pptm")
    print("\nDone. Open the files in Office 365 and press Alt+F11 to inspect the VBA.")


if __name__ == "__main__":
    main()
