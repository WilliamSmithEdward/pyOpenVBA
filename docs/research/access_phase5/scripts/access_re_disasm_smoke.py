"""Smoke-test the new disassembler against the corpus."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pyopenvba.access import AccessFile
from pyopenvba.vba_pcode import disassemble_module_stream

CORPUS = Path(__file__).resolve().parents[1] / "tests" / "live_access_test" / "re_corpus" / "samples"

for path in sorted(CORPUS.glob("040*.accdb")) + sorted(CORPUS.glob("044*.accdb")) + sorted(CORPUS.glob("030*.accdb")):
    db = AccessFile(path)
    streams = db.find_module_streams()
    if not streams:
        continue
    print(f"\n=== {path.name} (cafe={streams[0].cafe_offset:#x}) ===")
    mod = disassemble_module_stream(streams[0].raw)
    print(f"  num_lines={mod.num_lines}")
    for line in mod.lines:
        print(f"  Line #{line.line_no} @{line.start_offset:#x} len={line.byte_length}:")
        if 0 <= line.start_offset < len(streams[0].raw) and line.byte_length > 0:
            chunk = streams[0].raw[line.start_offset:line.start_offset + line.byte_length]
            print(f"    bytes: {chunk.hex(' ')}")
        for ins in line.instructions:
            parts = [ins.mnemonic]
            if ins.op_type:
                parts.append(f"(op_type=0x{ins.op_type:X})")
            for at, v in ins.operands:
                parts.append(f"{at}{v:#x}")
            if ins.payload is not None:
                try:
                    s = ins.payload.decode("latin-1")
                except Exception:
                    s = ins.payload.hex()
                parts.append(f'"{s}"')
            print("    " + " ".join(parts))
