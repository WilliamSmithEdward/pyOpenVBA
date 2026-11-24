"""scripts/build_excel_checklist.py stops, with a message and a non-zero exit status, where it has no type-library dump
to read, and finds the pyVBAReference checkout beside a worktree's main checkout when given no path."""

from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path
from types import ModuleType

import pytest

SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "build_excel_checklist.py"


def _builder() -> ModuleType:
    spec = importlib.util.spec_from_file_location("build_excel_checklist", SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _dump(folder: Path) -> Path:
    (folder / "reference" / "excel" / "json").mkdir(parents=True)
    return folder


def test_a_path_with_no_dump_fails_with_a_message_and_writes_nothing(tmp_path: Path) -> None:
    checklist = SCRIPT.parents[1] / "docs" / "excel_checklist.csv"
    before = checklist.read_bytes()
    missing = tmp_path / "no-such-checkout"
    run = subprocess.run([sys.executable, str(SCRIPT), str(missing)], capture_output=True, text=True, timeout=120,
                         check=False)
    assert run.returncode == 1
    assert f"no Excel type-library dump under {missing}" in run.stderr
    assert checklist.read_bytes() == before


def test_no_path_and_no_checkout_beside_fails_with_a_message(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    builder = _builder()
    monkeypatch.setattr(builder, "DEFAULT_REFERENCES", (tmp_path / "a" / "pyVBAReference",))
    with pytest.raises(SystemExit, match="give its path"):
        builder.reference_from([])


def test_no_path_takes_the_nearest_checkout_with_a_dump(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    builder = _builder()
    near, far = tmp_path / "near" / "pyVBAReference", _dump(tmp_path / "far" / "pyVBAReference")
    monkeypatch.setattr(builder, "DEFAULT_REFERENCES", (near, far))
    assert builder.reference_from([]) == far
    _dump(near)
    assert builder.reference_from([]) == near


def test_a_worktree_looks_beside_its_main_checkout() -> None:
    builder = _builder()
    root = SCRIPT.parents[1]
    assert root.parent / "pyVBAReference" in builder.DEFAULT_REFERENCES
    assert all(one.name == "pyVBAReference" for one in builder.DEFAULT_REFERENCES)
