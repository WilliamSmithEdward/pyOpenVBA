"""Live PowerPoint gate for the UserForm designer writer (opt-in).

The designer streams are host-independent -- the same ``f`` and ``o``
inside whichever ``vbaProject.bin`` the container holds -- but "should
be" is not a check.  This runs the composed-form case against a second
host, from a file pyOpenVBA created itself, so nothing in it came from
bytes any Office application wrote.

Opt-in: set ``RUN_LIVE_POWERPOINT=1`` on a Windows machine with desktop
PowerPoint and ``pyvbaharness`` installed.  Skipped everywhere else,
including CI.  ``pyvbaharness`` is a test-time oracle only; pyOpenVBA
never uses COM.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

from pyopenvba import PowerPointFile

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_LIVE_POWERPOINT") != "1" or sys.platform != "win32",
    reason=(
        "live PowerPoint gate: set RUN_LIVE_POWERPOINT=1 on Windows with "
        "PowerPoint installed"
    ),
)

_TIMEOUT = 60.0
# PowerPoint reaches its VBA project through the presentation, not through
# Application.VBE.ActivePresentation as Excel-shaped code expects.
_COMPONENTS = "Application.Presentations(1).VBProject.VBComponents"


def test_a_composed_form_loads_in_powerpoint(tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")

    target = tmp_path / "composed.pptm"
    with PowerPointFile.create_new(target) as presentation:
        form = presentation.add_form(
            "Wizard", caption="Deck setup", width=300, height=200
        )
        form.add_control("Label", "Prompt", left=12, top=12, width=200)
        form.add_control("TextBox", "Answer", left=12, top=40, width=200)
        form.add_control("Frame", "Group", left=12, top=80, width=200, height=70)
        form.add_control("OptionButton", "First", container="Group", left=6, top=12)
        form.add_control("CommandButton", "Ok", left=12, top=160)
        presentation.set_module(
            "Wizard", "Private Sub Ok_Click()\r\n    Me.Hide\r\nEnd Sub\r\n"
        )
        presentation.save()

    checks: list[tuple[str, object]] = [
        (f'{_COMPONENTS}("Wizard").Name', "Wizard"),
        (f'{_COMPONENTS}("Wizard").Type', 3),          # vbext_ct_MSForm
        (f'{_COMPONENTS}("Wizard").Designer.Caption', "Deck setup"),
        (f'{_COMPONENTS}("Wizard").Designer.Controls.Count', 5),
        (f'TypeName({_COMPONENTS}("Wizard").Designer.Controls("Answer"))', "TextBox"),
        (f'{_COMPONENTS}("Wizard").Designer.Controls("First").Parent.Name', "Group"),
        (f'{_COMPONENTS}("Wizard").CodeModule.CountOfLines > 0', True),
    ]
    session = harness.PowerPointSession()
    try:
        session.open_document(str(target.resolve()), read_only=False, timeout=120.0)
        wrong: list[str] = []
        for expression, expected in checks:
            got = session.eval(expression, timeout=_TIMEOUT)
            if got != expected:
                wrong.append(f"{expression}: expected {expected!r}, got {got!r}")
        assert not wrong, "; ".join(wrong)
    finally:
        try:
            session.close()
        except Exception:
            pass
