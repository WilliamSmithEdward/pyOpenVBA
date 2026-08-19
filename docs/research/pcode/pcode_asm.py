"""Prototype VBA7 p-code ASSEMBLER (inverse of vba_pcode.disassemble).

Milestone 1: byte-exact instruction encoder. Given the decoded fields
of a PCodeInstruction, re-emit the exact on-disk bytes. Verified by
round-tripping real compiled modules: disassemble -> re-encode each
line -> assert identical to the original line bytes.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
from pyopenvba.vba_pcode import _translate_opcode


# Inverse opcode map for 32-bit hosts (canonical -> raw). For 64-bit, identity.
def _canonical_to_raw(opcode: int, is_64bit: bool) -> int:
    if is_64bit:
        return opcode
    # invert _translate_opcode by searching (small table, fine)
    for raw in range(300):
        if _translate_opcode(raw, False) == opcode:
            return raw
    return opcode

def encode_instruction(ins, *, is_64bit: bool = True) -> bytes:
    """Re-emit one PCodeInstruction to its on-disk bytes.

    Uses raw_word verbatim for the header (preserves op_type and the
    exact raw opcode), then re-emits operands and varg payload per the
    opcode table -- the same widths the disassembler read.
    """
    out = bytearray()
    out += int(ins.raw_word).to_bytes(2, "little")
    for at, v in ins.operands:
        width = 2 if at in ("name", "0x", "imp_") else 4
        out += int(v).to_bytes(width, "little")
    if ins.payload is not None:
        out += len(ins.payload).to_bytes(2, "little")
        out += ins.payload
        if len(ins.payload) & 1:
            out += b"\x00"
    return bytes(out)

def encode_line(instructions, *, is_64bit: bool = True) -> bytes:
    out = bytearray()
    for ins in instructions:
        out += encode_instruction(ins, is_64bit=is_64bit)
    return bytes(out)
