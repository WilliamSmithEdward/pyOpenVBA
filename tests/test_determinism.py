"""Byte-determinism of save() across process boundaries.

pending_adds / pending_deletes are sets; Python randomizes string hash
order per process, so iterating them unsorted made multi-add saves emit
PROJECT declarations (and CFB stream allocation) in a different order on
each run.  save() now iterates them sorted.  The subprocess test drives
two different PYTHONHASHSEED values to prove the output bytes no longer
depend on hash order.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from pyopenvba import ExcelFile
from pyopenvba.vba import VBAModuleKind

_BUILDER = """
import sys
from pyopenvba import ExcelFile
from pyopenvba.vba import VBAModuleKind

out = sys.argv[1]
with ExcelFile.create_new(out) as wb:
    project = wb.vba_project()
    for name in ["Zed", "Alpha", "Mid", "Beta"]:
        project.add_module(
            name, "Sub " + name + "()\\r\\nEnd Sub\\r\\n", kind=VBAModuleKind.standard
        )
    wb.save()
"""


def _build_in_subprocess(target: Path, hash_seed: str) -> bytes:
    env = dict(os.environ)
    env["PYTHONHASHSEED"] = hash_seed
    subprocess.run(
        [sys.executable, "-c", _BUILDER, str(target)],
        check=True,
        env=env,
        capture_output=True,
    )
    return target.read_bytes()


def test_multi_add_save_bytes_do_not_depend_on_hash_seed(tmp_path: Path) -> None:
    first = _build_in_subprocess(tmp_path / "a.xlsm", "1")
    second = _build_in_subprocess(tmp_path / "b.xlsm", "271828")
    assert first == second


def test_multi_add_project_declarations_are_sorted(tmp_path: Path) -> None:
    target = tmp_path / "book.xlsm"
    with ExcelFile.create_new(target) as wb:
        project = wb.vba_project()
        for name in ["Zed", "Alpha", "Mid"]:
            project.add_module(
                name, f"Sub {name}()\r\nEnd Sub\r\n", kind=VBAModuleKind.standard
            )
        wb.save()
    with ExcelFile(target) as wb:
        cfb = wb._get_cfb()  # pyright: ignore[reportPrivateUsage]
        text = cfb.get_stream("PROJECT").decode("cp1252")
    added = [
        line.split("=", 1)[1]
        for line in text.splitlines()
        if line.startswith("Module=") and line != "Module=Module1"
    ]
    assert added == sorted(added)
