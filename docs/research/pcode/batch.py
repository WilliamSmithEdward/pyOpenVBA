"""Compile many source variants through ONE Excel session.

Differential analysis needs dozens of near-identical modules compiled by
the real engine; starting an Excel instance per variant dominates the
runtime, so every variant is written first and then driven through a
single session.

Dev-only: requires Windows, desktop Excel and ``pyvbaharness``. Nothing
in ``src/pyopenvba`` imports this.
"""
from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))

from pyopenvba import ExcelFile
from pyopenvba.vba import VBAModuleKind


def scratch_dir() -> Path:
    """Where compiled variants are written.

    ``PYOPENVBA_PCODE_SCRATCH`` overrides; otherwise a per-user temp
    directory, so nothing here depends on one machine's layout.
    """
    target = Path(os.environ.get("PYOPENVBA_PCODE_SCRATCH")
                  or Path(tempfile.gettempdir()) / "pyopenvba-pcode")
    target.mkdir(parents=True, exist_ok=True)
    return target


def compile_many(variants: dict[str, str], *, name: str = "M",
                 target: Path | None = None) -> dict[str, bytes]:
    """Compile ``tag -> module source`` and return ``tag -> p-code bytes``.

    Each source must be a complete module body including its
    ``Attribute VB_Name`` line (see :func:`hdr`). The returned bytes are
    the module stream's compiled prefix, which is what the disassembler
    consumes.
    """
    from pyvbaharness import ExcelSession

    target = target or scratch_dir()
    paths: dict[str, Path] = {}
    for tag, body in variants.items():
        path = target / f"v_{tag}.xlsm"
        if path.exists():
            path.unlink()
        with ExcelFile.create_new(path) as workbook:
            project = workbook.vba_project()
            if name in [m.name for m in project.modules]:
                workbook.set_module(name, body)
            else:
                project.add_module(name, body, kind=VBAModuleKind.standard)
            workbook.save()
        paths[tag] = path

    session = ExcelSession()
    try:
        for path in paths.values():
            session.open_document(path, read_only=False)
            session.compile_project(watch_seconds=20)
            session.save_as(path)
    finally:
        session.close()

    out: dict[str, bytes] = {}
    for tag, path in paths.items():
        with ExcelFile(path) as workbook:
            out[tag] = bytes(workbook.vba_project().get_module(name).prefix_bytes)
    return out


def hdr(body: str, name: str = "M") -> str:
    """Prefix a module body with the ``Attribute VB_Name`` line."""
    return f'Attribute VB_Name = "{name}"\r\n' + body
