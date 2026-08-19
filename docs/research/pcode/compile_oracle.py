"""Compile oracle: VBA source -> Excel-compiled p-code bytes.

pyOpenVBA writes the source (pure Python), Excel compiles and saves,
pyOpenVBA reads the compiled module stream back. Excel is the reference
implementation the assembler is validated against. COM is DEV-ONLY here.
"""
from __future__ import annotations
import shutil, sys
from pathlib import Path

sys.path.insert(0, "F:/GitHub/pyOpenVBA/src")
from pyopenvba import ExcelFile
from pyopenvba.vba import VBAModuleKind

SC = Path(r"C:\Users\William\AppData\Local\Temp\claude\F--GitHub-pyOpenVBA\0d1bb132-3785-4d59-92a0-31a36a2948cc\scratchpad\pcode")

def compile_source(body: str, *, name: str = "M", tag: str = "t",
                   session=None) -> bytes:
    """Return the compiled module-stream bytes (prefix incl. CAFE p-code)
    Excel produces for ``body`` in a standard module called ``name``."""
    wb_path = SC / f"oracle_{tag}.xlsm"
    if wb_path.exists():
        wb_path.unlink()
    with ExcelFile.create_new(wb_path) as wb:
        proj = wb.vba_project()
        if name in [m.name for m in proj.modules]:
            wb.set_module(name, body)
        else:
            proj.add_module(name, body, kind=VBAModuleKind.standard)
        wb.save()
    _excel_compile_and_save(wb_path, session)
    with ExcelFile(wb_path) as wb:
        m = wb.vba_project().get_module(name)
        return bytes(m.prefix_bytes)

def _excel_compile_and_save(path: Path, session=None) -> None:
    from pyvbaharness import ExcelSession
    own = session is None
    s = ExcelSession() if own else session
    try:
        s.open_document(path, read_only=False)
        s.compile_project(watch_seconds=20)
        s.save_as(path)
    finally:
        if own:
            s.close()
