"""No committed fixture names a folder on the machine that measured it.

A probe's JSON records the paths it opened and saved, and Excel records the
folder it saved a workbook in as x15ac:absPath in xl/workbook.xml. Probes mask
the first and scripts/fixture_workbook.py strips the second. A VBA project's
own bytes are not looked at here: some recorded projects still name folders.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import zipfile
from collections.abc import Iterable
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
#: A drive, then Users or GitHub, spelled as JSON escapes a backslash or with a slash.
MACHINE_PATH = re.compile(rb"(?i)\b[A-Z]:(?:\\\\|/)+(?:Users|GitHub)(?:\\\\|/)")


def committed(folder: str) -> Iterable[Path]:
    """The files git holds under a folder. A file git ignores is on one machine only, so it is no fixture. Without git, as in an
    unpacked source distribution, every file there."""
    git = shutil.which("git")
    listed = subprocess.run([git, "ls-files", "-z", "--", folder], cwd=ROOT, capture_output=True, check=False) if git else None
    if listed is None or listed.returncode != 0:
        return (path for path in (ROOT / folder).rglob("*") if path.is_file())
    return (path for path in (ROOT / name.decode() for name in listed.stdout.split(b"\0") if name) if path.is_file())


def test_no_fixture_json_names_a_machine_path() -> None:
    found = [path.relative_to(ROOT).as_posix() for path in committed("tests/fixtures") if path.suffix.lower() == ".json"
             and MACHINE_PATH.search(path.read_bytes())]
    assert found == []


def test_no_workbook_records_the_folder_it_was_saved_in() -> None:
    found: list[str] = []
    for folder in ("tests", "demo"):
        for path in committed(folder):
            if path.suffix.lower() not in (".xlsx", ".xlsm"):
                continue
            with zipfile.ZipFile(path) as package:
                if "xl/workbook.xml" in package.namelist() and b"x15ac:absPath" in package.read("xl/workbook.xml"):
                    found.append(path.relative_to(ROOT).as_posix())
    assert found == []
