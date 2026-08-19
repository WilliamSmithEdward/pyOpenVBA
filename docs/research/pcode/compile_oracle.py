"""Compile oracle: VBA source -> Excel-compiled p-code bytes.

pyOpenVBA writes the source (pure Python), Excel compiles and saves,
pyOpenVBA reads the compiled module stream back. Excel is the reference
implementation the assembler is validated against. COM is DEV-ONLY here;
nothing in ``src/pyopenvba`` imports this module.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from batch import scratch_dir

from pyopenvba import ExcelFile
from pyopenvba.vba import VBAModuleKind


def compile_source(body: str, *, name: str = "M", tag: str = "t",
                   session=None, target: Path | None = None) -> bytes:
    """Compiled module-stream bytes (prefix, including the CAFE p-code).

    ``body`` is a complete module source including its
    ``Attribute VB_Name`` line.
    """
    workbook_path = (target or scratch_dir()) / f"oracle_{tag}.xlsm"
    if workbook_path.exists():
        workbook_path.unlink()
    with ExcelFile.create_new(workbook_path) as workbook:
        project = workbook.vba_project()
        if name in [m.name for m in project.modules]:
            workbook.set_module(name, body)
        else:
            project.add_module(name, body, kind=VBAModuleKind.standard)
        workbook.save()
    _excel_compile_and_save(workbook_path, session)
    with ExcelFile(workbook_path) as workbook:
        return bytes(workbook.vba_project().get_module(name).prefix_bytes)


def _excel_compile_and_save(path: Path, session=None) -> None:
    from pyvbaharness import ExcelSession

    own = session is None
    session = session or ExcelSession()
    try:
        session.open_document(path, read_only=False)
        session.compile_project(watch_seconds=20)
        session.save_as(path)
    finally:
        if own:
            session.close()
