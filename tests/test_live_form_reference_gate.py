"""Live gate: code naming MSForms types compiles in a form add_form made (opt-in).

What only the editor's own compiler can answer (issue #29): that a file
given a UserForm here, and code that names Microsoft Forms types -- the
form held As MSForms.UserForm, a KeyPress handler taking KeyAscii As
MSForms.ReturnInteger -- compiles, which needs the reference add_form
declares. The applications' compile check runs through pyvbaharness, as
xlide_mcp's does.

Opt-in per application: RUN_LIVE_EXCEL=1, RUN_LIVE_WORD=1 or
RUN_LIVE_POWERPOINT=1, on Windows with it installed. Skipped everywhere
else, including CI.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, PowerPointFile, WordFile

#: Each application: the switch that opts in, the library's file type, the extension and the harness session.
HOSTS: dict[str, tuple[str, Any, str, str]] = {
    "excel": ("RUN_LIVE_EXCEL", ExcelFile, "xlsm", "ExcelSession"),
    "word": ("RUN_LIVE_WORD", WordFile, "docm", "WordSession"),
    "powerpoint": ("RUN_LIVE_POWERPOINT", PowerPointFile, "pptm", "PowerPointSession"),
}
MODULE = ("Public Sub ShowIt()\r\n    Dim f As MSForms.UserForm\r\n    Set f = Wizard\r\n"
          "    Dim c As MSForms.Control\r\n    For Each c In Wizard.Controls\r\n    Next\r\nEnd Sub\r\n")
HANDLER = "Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)\r\n    KeyAscii = 0\r\nEnd Sub\r\n"


def _param(host: str) -> Any:
    switch = HOSTS[host][0]
    return pytest.param(host, id=host, marks=pytest.mark.skipif(
        os.environ.get(switch) != "1" or sys.platform != "win32",
        reason=f"live form reference gate: set {switch}=1 on Windows with the application installed"))


@pytest.mark.parametrize("host", [_param(host) for host in HOSTS])
def test_code_naming_msforms_types_compiles(host: str, tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    _, opener, suffix, session_name = HOSTS[host]
    target = tmp_path / f"forms.{suffix}"
    with opener.create_new(target) as office_file:
        office_file.add_form("Wizard").add_control("CommandButton", "Go", left=12, top=12)
        office_file.set_module("Wizard", HANDLER)
        office_file.vba_project().add_module("Shows", MODULE)
        office_file.save()
    with getattr(harness, session_name)(harness.HarnessConfig(lock_wait_s=1200.0)) as office:
        office.open_document(target)
        compiled = office.compile_project()
    assert compiled.ok, f"{compiled.outcome}: {compiled.message} {compiled.dialog}"
