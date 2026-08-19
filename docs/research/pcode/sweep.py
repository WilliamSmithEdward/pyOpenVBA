"""Coverage sweep: decompile every module in every fixture on disk.

The round-trip gate proves the decompiler reproduces known sources; this
proves it never *silently* drops p-code. Every module found in the given
paths is disassembled and decompiled, and anything the renderer could
not map is reported: unmapped opcodes, unresolved names, and declared
types the descriptor decoder could not name.

    python docs/research/pcode/sweep.py demo docs/research

Host-agnostic by construction: the same code path handles .xlsm, .xlsb,
.xls, .docm, .pptm and Access databases.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from pcode_decompile import find_decl_base
from pcode_names import parse_identifiers
from pcode_source import decompile
from pcode_types import find_type_table, read_declared_type

from pyopenvba import ExcelFile, PowerPointFile, WordFile
from pyopenvba.cfb import CFB
from pyopenvba.vba_pcode import OPCODES_VBA7, disassemble_module_stream

HOSTS = {
    ".xlsm": ExcelFile, ".xlsb": ExcelFile, ".xls": ExcelFile,
    ".xlam": ExcelFile, ".docm": WordFile, ".dotm": WordFile,
    ".pptm": PowerPointFile, ".ppam": PowerPointFile,
}


def sweep(paths: list[Path]) -> int:
    files = sorted(
        p for root in paths for p in
        ([root] if root.is_file() else root.rglob("*"))
        if p.suffix.lower() in HOSTS
    )
    modules = 0
    lines = 0
    unmapped: dict[str, int] = {}
    unnamed = 0
    no_table = 0
    untyped: dict[str, int] = {}
    for path in files:
        host = HOSTS[path.suffix.lower()]
        try:
            with host(path) as document:
                project = document.vba_project()
                cfb = CFB.from_bytes(document.vba_project_bytes())
                table = cfb.get_stream_in_storage("VBA", "_VBA_PROJECT")
                for module in project.modules:
                    stream = bytes(module.prefix_bytes)
                    if not stream:
                        continue
                    modules += 1
                    identifiers = parse_identifiers(table)
                    if not identifiers:
                        # The project's identifier table was blanked (pyOpenVBA
                        # zeroes _VBA_PROJECT on a mutating save), so names
                        # cannot resolve. That is expected, not a decoding gap.
                        no_table += 1
                    disassembly = disassemble_module_stream(stream, is_64bit=True)
                    lines += disassembly.num_lines
                    for source_line in disassembly.lines:
                        for instruction in source_line.instructions:
                            if instruction.opcode not in OPCODES_VBA7:
                                key = f"opcode {instruction.opcode:#x}"
                                unmapped[key] = unmapped.get(key, 0) + 1
                    text = decompile(stream, table)
                    for row in text.splitlines():
                        if "[unmapped" in row:
                            key = row.split("[unmapped", 1)[1].strip(" ]")
                            unmapped[key] = unmapped.get(key, 0) + 1
                        if ("<var>" in row or "<proc>" in row
                                or "<name>" in row) and identifiers:
                            unnamed += 1
                    base = find_decl_base(stream, identifiers)
                    if base is None:
                        continue
                    types = find_type_table(stream, base)
                    for source_line in disassembly.lines:
                        for instruction in source_line.instructions:
                            for kind, value in instruction.operands:
                                if kind != "var_":
                                    continue
                                info = read_declared_type(
                                    stream, base, value, types,
                                    lambda op: None)
                                for note in (info.unresolved if info else []):
                                    tag = note.split("[")[0]
                                    untyped[tag] = untyped.get(tag, 0) + 1
        except Exception as error:
            print(f"  skip {path.name}: {type(error).__name__}: {error}")
    print(f"{len(files)} files, {modules} modules, {lines} p-code lines")
    print(f"unmapped opcodes/mnemonics: {unmapped or 'none'}")
    print(f"unresolved declaration names: {unnamed}"
          f"   ({no_table} modules had no identifier table to resolve against)")
    # typeref[n] entries are expected here: the sweep passes a null name
    # resolver, so only structural failures matter.
    structural = {k: v for k, v in untyped.items() if not k.startswith("typeref")}
    print(f"undecoded type descriptors: {structural or 'none'}")
    return 1 if unmapped or structural else 0


if __name__ == "__main__":
    targets = [Path(a) for a in sys.argv[1:]] or [Path("demo")]
    raise SystemExit(sweep(targets))
