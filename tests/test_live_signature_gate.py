"""Live gate: a file whose signature the library dropped opens unsigned (opt-in).

What only the applications can answer: that a .xlsm, .docm or .pptm the
library saved after changing its code, taking out the signature parts
beside vbaProject.bin, opens without complaint, reads VBASigned False and
holds the change. The fixtures' signature parts are placeholders, which
Word and PowerPoint refuse to open (tests/fixtures/signature_parts/), so
only the library's save goes to the application.

Opt-in per application: RUN_LIVE_EXCEL=1, RUN_LIVE_WORD=1 or
RUN_LIVE_POWERPOINT=1, on Windows with it installed. Skipped everywhere
else, including CI.
"""

from __future__ import annotations

import json
import os
import sys
import warnings
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, PowerPointFile, WordFile

FIXTURES = Path(__file__).parent / "fixtures" / "signature_parts"
RECORD: dict[str, dict[str, Any]] = json.loads((FIXTURES / "signature_parts.json").read_text(encoding="utf-8"))
#: Each application: the switch that opts in, the library's file type, the harness session, and how its VBA opens
#: and closes a file with no alert in the way.
HOSTS: dict[str, tuple[str, Any, str, str, str, str]] = {
    "excel": ("RUN_LIVE_EXCEL", ExcelFile, "ExcelSession", 'Workbooks.Open("{path}")', "Close False",
              "Application.DisplayAlerts = False"),
    "word": ("RUN_LIVE_WORD", WordFile, "WordSession", 'Documents.Open("{path}")', "Close SaveChanges:=0",
             "Application.DisplayAlerts = 0"),
    "powerpoint": ("RUN_LIVE_POWERPOINT", PowerPointFile, "PowerPointSession",
                   'Presentations.Open("{path}", 0, 0, 0)', "Close", "Application.DisplayAlerts = 1"),
}
EDIT = "Public Function Answer() As Long\r\n    Answer = 42\r\nEnd Function\r\n"


def _param(host: str) -> Any:
    switch = HOSTS[host][0]
    return pytest.param(host, id=host, marks=pytest.mark.skipif(
        os.environ.get(switch) != "1" or sys.platform != "win32",
        reason=f"live signature gate: set {switch}=1 on Windows with the application installed"))


@pytest.mark.parametrize("host", [_param(host) for host in HOSTS])
def test_the_application_opens_the_save_unsigned(host: str, tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    _, opener, session_name, opening, closing, quiet = HOSTS[host]
    record = RECORD[host]
    out = tmp_path / record["file"]
    with opener(FIXTURES / record["file"]) as office_file, warnings.catch_warnings():
        warnings.simplefilter("ignore")
        office_file.set_module(record["module"], office_file.get_module(record["module"]) + EDIT)
        office_file.save(out)
    code = (f"Public Function Probe() As String\nDim d As Object, m As Object\n{quiet}\n"
            f"Set d = {opening.format(path=out)}\n"
            f'Set m = d.VBProject.VBComponents("{record["module"]}").CodeModule\n'
            'Probe = CStr(CBool(d.VBASigned)) & "|" & m.Lines(1, m.CountOfLines)\n'
            f"d.{closing}\nEnd Function\n")
    with getattr(harness, session_name)(harness.HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        result = office.run_vba(code, "Probe", timeout=180.0)
    assert result.ok, f"{result.outcome}: {result.message}"
    signed, _, source = str(result.value).partition("|")
    assert signed == "False"
    assert "Answer = 42" in source
