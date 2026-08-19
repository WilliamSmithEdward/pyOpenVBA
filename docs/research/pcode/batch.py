"""Compile many source variants through ONE Excel session (fast)."""
from __future__ import annotations
import sys
from pathlib import Path
sys.path.insert(0,"F:/GitHub/pyOpenVBA/src")
from pyopenvba import ExcelFile
from pyopenvba.vba import VBAModuleKind
SC = Path(r"C:\Users\William\AppData\Local\Temp\claude\F--GitHub-pyOpenVBA\0d1bb132-3785-4d59-92a0-31a36a2948cc\scratchpad\pcode")

def compile_many(variants: dict[str,str], *, name="M") -> dict[str,bytes]:
    """variants: tag -> module body (full source w/ Attribute line).
    Returns tag -> compiled module-stream prefix bytes."""
    from pyvbaharness import ExcelSession
    paths={}
    for tag, body in variants.items():
        p = SC/f"v_{tag}.xlsm"
        if p.exists(): p.unlink()
        with ExcelFile(p if False else p) if False else ExcelFile.create_new(p) as wb:
            proj=wb.vba_project()
            if name in [m.name for m in proj.modules]: wb.set_module(name, body)
            else: proj.add_module(name, body, kind=VBAModuleKind.standard)
            wb.save()
        paths[tag]=p
    out={}
    s=ExcelSession()
    try:
        for tag,p in paths.items():
            s.open_document(p, read_only=False)
            s.compile_project(watch_seconds=20)
            s.save_as(p)
    finally:
        s.close()
    for tag,p in paths.items():
        with ExcelFile(p) as wb:
            out[tag]=bytes(wb.vba_project().get_module(name).prefix_bytes)
    return out

def hdr(body_lines: str, name="M") -> str:
    return f'Attribute VB_Name = "{name}"\r\n' + body_lines
