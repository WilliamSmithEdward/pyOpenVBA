"""Live Excel compile-and-run gate (opt-in).

Every other test verifies structure; this one verifies the thing users
actually experience: a workbook built by pyOpenVBA must COMPILE AND RUN
in real Excel.  GitHub issue #1 shipped precisely because "opens without
a repair prompt" was the strongest live check: the broken class module
opened fine and only failed when the compiler reached ``New Class1``.

The gate builds a workbook whose macro instantiates a class added from
VBE-export-form source (the normalizer's hardest input), runs it through
``tools/live_excel/run_macro.ps1`` -- a popup-aware bounded harness that
dismisses and captures any VBE modal instead of deadlocking -- and
requires a clean run plus the sentinel file the macro writes.

Opt-in: set ``RUN_LIVE_EXCEL=1`` on a Windows machine with desktop Excel
installed.  Skipped everywhere else (including CI).
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

from pyopenvba import ExcelFile
from pyopenvba.vba import VBAModuleKind

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_EXCEL") != "1" or sys.platform != "win32",
    reason="live Excel gate: set RUN_LIVE_EXCEL=1 on Windows with Excel installed",
)

_RUNNER = Path(__file__).parent.parent / "tools" / "live_excel" / "run_macro.ps1"

_MODULE1 = (
    'Attribute VB_Name = "Module1"\r\n'
    "\r\n"
    "Sub RunGate()\r\n"
    "    Dim probe As New GateClass\r\n"
    '    probe.Tag = "gate-ok"\r\n'
    "    Dim f As Integer\r\n"
    "    f = FreeFile\r\n"
    '    Open ThisWorkbook.Path & "\\gate_sentinel.txt" For Output As #f\r\n'
    "    Print #f, probe.Describe()\r\n"
    "    Close #f\r\n"
    "End Sub\r\n"
)

# Deliberately VBE-export form: VERSION preamble present, VB_Base absent.
# Both defects were individually fatal in live Excel before the
# normalize_class_source() fix.
_GATECLASS_EXPORT_FORM = (
    "VERSION 1.0 CLASS\r\n"
    "BEGIN\r\n"
    "  MultiUse = -1  'True\r\n"
    "END\r\n"
    'Attribute VB_Name = "GateClass"\r\n'
    "Attribute VB_Exposed = False\r\n"
    "\r\n"
    "Private mTag As String\r\n"
    "\r\n"
    "Public Property Let Tag(v As String)\r\n"
    "    mTag = v\r\n"
    "End Property\r\n"
    "\r\n"
    "Public Function Describe() As String\r\n"
    '    Describe = "GateClass: " & mTag\r\n'
    "End Function\r\n"
)


def test_export_form_class_compiles_and_runs_in_live_excel(
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "live_gate.xlsm"
    with ExcelFile.create_new(workbook) as wb:
        wb.set_module("Module1", _MODULE1)
        wb.vba_project().add_module(
            "GateClass", _GATECLASS_EXPORT_FORM, kind=VBAModuleKind.other
        )
        wb.save()

    completed = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(_RUNNER),
            "-WorkbookPath",
            str(workbook),
            "-MacroName",
            "RunGate",
        ],
        capture_output=True,
        text=True,
        timeout=180,
    )
    assert completed.returncode == 0, (
        f"runner failed: rc={completed.returncode}\n"
        f"stdout={completed.stdout}\nstderr={completed.stderr}"
    )
    report = json.loads(completed.stdout.strip().splitlines()[-1])
    assert report["outcome"] == "run-ok", report
    assert report["popups"] in ([], None), report

    sentinel = tmp_path / "gate_sentinel.txt"
    assert sentinel.exists(), "macro ran but wrote no sentinel file"
    assert sentinel.read_text().strip() == "GateClass: gate-ok"
