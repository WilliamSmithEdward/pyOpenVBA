"""Live gate: a project add_vba_project gave a file is one its application opens and runs (opt-in).

What only the applications can answer: that a .xlsm, .docm or .pptm
saved with no VBA, given a project and a module here, opens without a
repair prompt, holds the components its application would have made,
with Excel's code names on its sheets, and runs the module's code.

Opt-in per application: RUN_LIVE_EXCEL=1, RUN_LIVE_WORD=1 or
RUN_LIVE_POWERPOINT=1, on Windows with it installed. Skipped everywhere
else, including CI.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, PowerPointFile, WordFile

FIXTURES = Path(__file__).parent / "fixtures" / "first_macro"
RECORD: dict[str, dict[str, Any]] = json.loads((FIXTURES / "first_macro.json").read_text(encoding="utf-8"))
#: Each application: the switch that opts in, the library's file type, the harness session, and how its VBA
#: opens a file, runs one of its macros and closes it.
HOSTS: dict[str, tuple[str, Any, str, str, str, str]] = {
    "excel": ("RUN_LIVE_EXCEL", ExcelFile, "ExcelSession", 'Workbooks.Open("{path}")',
              "Application.Run(\"'\" & d.Name & \"'!Module1.Answer\")", "Close False"),
    "word": ("RUN_LIVE_WORD", WordFile, "WordSession", 'Documents.Open("{path}")',
             'Application.Run("Module1.Answer")', "Close SaveChanges:=0"),
    "powerpoint": ("RUN_LIVE_POWERPOINT", PowerPointFile, "PowerPointSession",
                   'Presentations.Open("{path}", 0, 0, 0)', 'Application.Run(d.Name & "!Module1.Answer")', "Close"),
}
ANSWER = "Public Function Answer() As Long\r\n    Answer = 42\r\nEnd Function\r\n"
QUIET = {"excel": "Application.DisplayAlerts = False", "word": "Application.DisplayAlerts = 0",
         "powerpoint": "Application.DisplayAlerts = 1"}


def _param(host: str) -> Any:
    switch = HOSTS[host][0]
    return pytest.param(host, id=host, marks=pytest.mark.skipif(
        os.environ.get(switch) != "1" or sys.platform != "win32",
        reason=f"live add-project gate: set {switch}=1 on Windows with the application installed"))


@pytest.mark.parametrize("host", [_param(host) for host in HOSTS])
def test_the_application_runs_the_project_added_here(host: str, tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    _, opener, session_name, opening, running, closing = HOSTS[host]
    record = RECORD[host]
    source, out = tmp_path / f"in.{record['suffix']}", tmp_path / f"out.{record['suffix']}"
    shutil.copyfile(FIXTURES / f"{host}_before.{record['suffix']}", source)
    with opener(source) as office_file:
        office_file.add_vba_project().add_module("Module1", ANSWER)
        office_file.save(out)
    names = ('Dim s As Object\nFor Each s In d.Sheets\nout = out & s.Name & "=" & s.CodeName & ";"\nNext\n'
             if host == "excel" else "")
    code = (f"Public Function Probe() As String\nDim d As Object, c As Object, out As String\n{QUIET[host]}\n"
            f"Set d = {opening.format(path=out)}\n"
            'For Each c In d.VBProject.VBComponents\nout = out & c.Name & ":" & c.Type & ";"\nNext\n'
            f'out = out & "#"\n{names}out = out & "#" & {running}\nd.{closing}\nProbe = out\nEnd Function\n')
    with getattr(harness, session_name)(harness.HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        result = office.run_vba(code, "Probe", timeout=180.0)
    assert result.ok, f"{result.outcome}: {result.message}"
    components, code_names, answer = str(result.value).split("#")
    assert [one.split(":") for one in components.split(";") if one] == record["module"]["components"]
    assert [one.split("=") for one in code_names.split(";") if one] == record["module"]["code_names"]
    assert answer == "42"
