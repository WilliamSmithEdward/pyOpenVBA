"""VBA7 p-code disassembler -- dependency-free.

This module decodes the canonical Microsoft VBA7 p-code stream
embedded in every Office-VBA module stream (and, as Phase 4d RE
established, also stored verbatim in every Access database alongside
its ``rU@`` execodes form).

The bytecode layout is the one defined by [MS-OVBA] section 2.3.4.3:

* binary metadata + tables (declaration / indirect / object) at the
  head of the module stream,
* a ``0xCAFE`` magic word,
* a ``<u16 numLines>`` count,
* ``numLines`` 12-byte line records (one per source line),
* a ``numLines * 12 + 10``-byte gap, then
* the per-line p-code instructions themselves -- each a ``<u16>``
  encoding ``opcode`` (low 10 bits) and ``op_type`` (high 6 bits),
  followed by 0..3 operand words depending on the mnemonic.

This implementation targets VBA7 (Office 2010+; Access 2010+) 32-bit
hosts -- which covers every Access database we have observed. The
264-entry opcode table is factual data about Microsoft's file format,
reverse-engineered and widely published (see e.g. ``pcodedmp`` by
Vesselin Bontchev). All parser code in this module is an independent
implementation.

Public API:

* :class:`PCodeInstruction` -- one decoded instruction.
* :class:`PCodeLine` -- one source-line's worth of instructions.
* :class:`DisassembledModule` -- the full module's decoded p-code.
* :func:`disassemble_module_stream` -- bytes -> DisassembledModule.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Sequence

__all__ = [
    "PCodeInstruction",
    "PCodeLine",
    "PCodeProcedure",
    "DisassembledModule",
    "disassemble_module_stream",
    "find_cafe_offset",
    "OPCODES_VBA7",
]


# ----------------------------------------------------------------------
# VBA7 opcode table.
#
# Each entry: (mnemonic, arg_types, has_varg_payload).
#
# arg_types contains zero or more of:
#   "name"     -- u16 identifier-table index
#   "0x"       -- u16 raw hex literal
#   "imp_"     -- u16 import / object reference (or Open-modes bitfield)
#   "context_" -- u32 (or u32+u32 on 64-bit)
#   "func_"    -- u32 offset into indirect table (Sub/Function decl)
#   "var_"     -- u32 offset into indirect table (Dim decl)
#   "rec_"     -- u32 offset into indirect table (Type decl)
#   "type_"    -- u32 offset into indirect table (As-type)
#
# has_varg_payload: if True, an additional ``<u16 length>`` follows the
# fixed arguments and is followed by ``length`` bytes of payload
# (string literal, line continuation, etc., padded to even length).
#
# Source: independently re-implemented from the published VBA7 format;
# matches the table used by `pcodedmp` and other public tools.
# ----------------------------------------------------------------------

# (mnemonic, args, varg)
_OP = tuple[str, tuple[str, ...], bool]

OPCODES_VBA7: dict[int, _OP] = {
    0: ("Imp", (), False),
    1: ("Eqv", (), False),
    2: ("Xor", (), False),
    3: ("Or", (), False),
    4: ("And", (), False),
    5: ("Eq", (), False),
    6: ("Ne", (), False),
    7: ("Le", (), False),
    8: ("Ge", (), False),
    9: ("Lt", (), False),
    10: ("Gt", (), False),
    11: ("Add", (), False),
    12: ("Sub", (), False),
    13: ("Mod", (), False),
    14: ("IDiv", (), False),
    15: ("Mul", (), False),
    16: ("Div", (), False),
    17: ("Concat", (), False),
    18: ("Like", (), False),
    19: ("Pwr", (), False),
    20: ("Is", (), False),
    21: ("Not", (), False),
    22: ("UMi", (), False),
    23: ("FnAbs", (), False),
    24: ("FnFix", (), False),
    25: ("FnInt", (), False),
    26: ("FnSgn", (), False),
    27: ("FnLen", (), False),
    28: ("FnLenB", (), False),
    29: ("Paren", (), False),
    30: ("Sharp", (), False),
    31: ("LdLHS", ("name",), False),
    32: ("Ld", ("name",), False),
    33: ("MemLd", ("name",), False),
    34: ("DictLd", ("name",), False),
    35: ("IndexLd", ("0x",), False),
    36: ("ArgsLd", ("name", "0x"), False),
    37: ("ArgsMemLd", ("name", "0x"), False),
    38: ("ArgsDictLd", ("name", "0x"), False),
    39: ("St", ("name",), False),
    40: ("MemSt", ("name",), False),
    41: ("DictSt", ("name",), False),
    42: ("IndexSt", ("0x",), False),
    43: ("ArgsSt", ("name", "0x"), False),
    44: ("ArgsMemSt", ("name", "0x"), False),
    45: ("ArgsDictSt", ("name", "0x"), False),
    46: ("Set", ("name",), False),
    47: ("Memset", ("name",), False),
    48: ("Dictset", ("name",), False),
    49: ("Indexset", ("0x",), False),
    50: ("ArgsSet", ("name", "0x"), False),
    51: ("ArgsMemSet", ("name", "0x"), False),
    52: ("ArgsDictSet", ("name", "0x"), False),
    53: ("MemLdWith", ("name",), False),
    54: ("DictLdWith", ("name",), False),
    55: ("ArgsMemLdWith", ("name", "0x"), False),
    56: ("ArgsDictLdWith", ("name", "0x"), False),
    57: ("MemStWith", ("name",), False),
    58: ("DictStWith", ("name",), False),
    59: ("ArgsMemStWith", ("name", "0x"), False),
    60: ("ArgsDictStWith", ("name", "0x"), False),
    61: ("MemSetWith", ("name",), False),
    62: ("DictSetWith", ("name",), False),
    63: ("ArgsMemSetWith", ("name", "0x"), False),
    64: ("ArgsDictSetWith", ("name", "0x"), False),
    65: ("ArgsCall", ("name", "0x"), False),
    66: ("ArgsMemCall", ("name", "0x"), False),
    67: ("ArgsMemCallWith", ("name", "0x"), False),
    68: ("ArgsArray", ("name", "0x"), False),
    69: ("Assert", (), False),
    70: ("BoS", ("0x",), False),
    71: ("BoSImplicit", (), False),
    72: ("BoL", (), False),
    73: ("LdAddressOf", ("name",), False),
    74: ("MemAddressOf", ("name",), False),
    75: ("Case", (), False),
    76: ("CaseTo", (), False),
    77: ("CaseGt", (), False),
    78: ("CaseLt", (), False),
    79: ("CaseGe", (), False),
    80: ("CaseLe", (), False),
    81: ("CaseNe", (), False),
    82: ("CaseEq", (), False),
    83: ("CaseElse", (), False),
    84: ("CaseDone", (), False),
    85: ("Circle", ("0x",), False),
    86: ("Close", ("0x",), False),
    87: ("CloseAll", (), False),
    88: ("Coerce", (), False),
    89: ("CoerceVar", (), False),
    90: ("Context", ("context_",), False),
    91: ("Debug", (), False),
    92: ("DefType", ("0x", "0x"), False),
    93: ("Dim", (), False),
    94: ("DimImplicit", (), False),
    95: ("Do", (), False),
    96: ("DoEvents", (), False),
    97: ("DoUntil", (), False),
    98: ("DoWhile", (), False),
    99: ("Else", (), False),
    100: ("ElseBlock", (), False),
    101: ("ElseIfBlock", (), False),
    102: ("ElseIfTypeBlock", ("imp_",), False),
    103: ("End", (), False),
    104: ("EndContext", (), False),
    105: ("EndFunc", (), False),
    106: ("EndIf", (), False),
    107: ("EndIfBlock", (), False),
    108: ("EndImmediate", (), False),
    109: ("EndProp", (), False),
    110: ("EndSelect", (), False),
    111: ("EndSub", (), False),
    112: ("EndType", (), False),
    113: ("EndWith", (), False),
    114: ("Erase", ("0x",), False),
    115: ("Error", (), False),
    116: ("EventDecl", ("func_",), False),
    117: ("RaiseEvent", ("name", "0x"), False),
    118: ("ArgsMemRaiseEvent", ("name", "0x"), False),
    119: ("ArgsMemRaiseEventWith", ("name", "0x"), False),
    120: ("ExitDo", (), False),
    121: ("ExitFor", (), False),
    122: ("ExitFunc", (), False),
    123: ("ExitProp", (), False),
    124: ("ExitSub", (), False),
    125: ("FnCurDir", (), False),
    126: ("FnDir", (), False),
    127: ("Empty0", (), False),
    128: ("Empty1", (), False),
    129: ("FnError", (), False),
    130: ("FnFormat", (), False),
    131: ("FnFreeFile", (), False),
    132: ("FnInStr", (), False),
    133: ("FnInStr3", (), False),
    134: ("FnInStr4", (), False),
    135: ("FnInStrB", (), False),
    136: ("FnInStrB3", (), False),
    137: ("FnInStrB4", (), False),
    138: ("FnLBound", ("0x",), False),
    139: ("FnMid", (), False),
    140: ("FnMidB", (), False),
    141: ("FnStrComp", (), False),
    142: ("FnStrComp3", (), False),
    143: ("FnStringVar", (), False),
    144: ("FnStringStr", (), False),
    145: ("FnUBound", ("0x",), False),
    146: ("For", (), False),
    147: ("ForEach", (), False),
    148: ("ForEachAs", ("imp_",), False),
    149: ("ForStep", (), False),
    150: ("FuncDefn", ("func_",), False),
    151: ("FuncDefnSave", ("func_",), False),
    152: ("GetRec", (), False),
    153: ("GoSub", ("name",), False),
    154: ("GoTo", ("name",), False),
    155: ("If", (), False),
    156: ("IfBlock", (), False),
    157: ("TypeOf", ("imp_",), False),
    158: ("IfTypeBlock", ("imp_",), False),
    159: ("Implements", ("0x", "0x", "0x", "0x"), False),
    160: ("Input", (), False),
    161: ("InputDone", (), False),
    162: ("InputItem", (), False),
    163: ("Label", ("name",), False),
    164: ("Let", (), False),
    165: ("Line", ("0x",), False),
    166: ("LineCont", (), True),
    167: ("LineInput", (), False),
    168: ("LineNum", ("name",), False),
    169: ("LitCy", ("0x", "0x", "0x", "0x"), False),
    170: ("LitDate", ("0x", "0x", "0x", "0x"), False),
    171: ("LitDefault", (), False),
    172: ("LitDI2", ("0x",), False),
    173: ("LitDI4", ("0x", "0x"), False),
    174: ("LitDI8", ("0x", "0x", "0x", "0x"), False),
    175: ("LitHI2", ("0x",), False),
    176: ("LitHI4", ("0x", "0x"), False),
    177: ("LitHI8", ("0x", "0x", "0x", "0x"), False),
    178: ("LitNothing", (), False),
    179: ("LitOI2", ("0x",), False),
    180: ("LitOI4", ("0x", "0x"), False),
    181: ("LitOI8", ("0x", "0x", "0x", "0x"), False),
    182: ("LitR4", ("0x", "0x"), False),
    183: ("LitR8", ("0x", "0x", "0x", "0x"), False),
    184: ("LitSmallI2", (), False),
    185: ("LitStr", (), True),
    186: ("LitVarSpecial", (), False),
    187: ("Lock", (), False),
    188: ("Loop", (), False),
    189: ("LoopUntil", (), False),
    190: ("LoopWhile", (), False),
    191: ("LSet", (), False),
    192: ("Me", (), False),
    193: ("MeImplicit", (), False),
    194: ("MemRedim", ("name", "0x", "type_"), False),
    195: ("MemRedimWith", ("name", "0x", "type_"), False),
    196: ("MemRedimAs", ("name", "0x", "type_"), False),
    197: ("MemRedimAsWith", ("name", "0x", "type_"), False),
    198: ("Mid", (), False),
    199: ("MidB", (), False),
    200: ("Name", (), False),
    201: ("New", ("imp_",), False),
    202: ("Next", (), False),
    203: ("NextVar", (), False),
    204: ("OnError", ("name",), False),
    205: ("OnGosub", (), True),
    206: ("OnGoto", (), True),
    207: ("Open", ("0x",), False),
    208: ("Option", (), False),
    209: ("OptionBase", (), False),
    210: ("ParamByVal", (), False),
    211: ("ParamOmitted", (), False),
    212: ("ParamNamed", ("name",), False),
    213: ("PrintChan", (), False),
    214: ("PrintComma", (), False),
    215: ("PrintEoS", (), False),
    216: ("PrintItemComma", (), False),
    217: ("PrintItemNL", (), False),
    218: ("PrintItemSemi", (), False),
    219: ("PrintNL", (), False),
    220: ("PrintObj", (), False),
    221: ("PrintSemi", (), False),
    222: ("PrintSpc", (), False),
    223: ("PrintTab", (), False),
    224: ("PrintTabComma", (), False),
    225: ("PSet", ("0x",), False),
    226: ("PutRec", (), False),
    227: ("QuoteRem", ("0x",), True),
    228: ("Redim", ("name", "0x", "type_"), False),
    229: ("RedimAs", ("name", "0x", "type_"), False),
    230: ("Reparse", (), True),
    231: ("Rem", (), True),
    232: ("Resume", ("name",), False),
    233: ("Return", (), False),
    234: ("RSet", (), False),
    235: ("Scale", ("0x",), False),
    236: ("Seek", (), False),
    237: ("SelectCase", (), False),
    238: ("SelectIs", ("imp_",), False),
    239: ("SelectType", (), False),
    240: ("SetStmt", (), False),
    241: ("Stack", ("0x", "0x"), False),
    242: ("Stop", (), False),
    243: ("Type", ("rec_",), False),
    244: ("Unlock", (), False),
    245: ("VarDefn", ("var_",), False),
    246: ("Wend", (), False),
    247: ("While", (), False),
    248: ("With", (), False),
    249: ("WriteChan", (), False),
    250: ("ConstFuncExpr", (), False),
    251: ("LbConst", ("name",), False),
    252: ("LbIf", (), False),
    253: ("LbElse", (), False),
    254: ("LbElseIf", (), False),
    255: ("LbEndIf", (), False),
    256: ("LbMark", (), False),
    257: ("EndForVariable", (), False),
    258: ("StartForVariable", (), False),
    259: ("NewRedim", (), False),
    260: ("StartWithExpr", (), False),
    261: ("SetOrSt", ("name",), False),
    262: ("EndEnum", (), False),
    263: ("Illegal", (), False),
}


# Encoding scheme: a u16 instruction header packs opcode (low 10 bits)
# and op_type (high 6 bits).
#
# VBA7 64-bit (Access 2010+ on 64-bit Office, which is what every
# sample in our test corpus uses): raw opcode == canonical opcode, no
# remap needed.
#
# VBA7 32-bit / VBA6: opcodes above 173 are remapped by small offsets
# to make room for 64-bit-only mnemonics. We provide both translations
# and pick by the ``is_64bit`` flag.
def _translate_opcode(raw: int, is_64bit: bool) -> int:
    """Translate a raw 10-bit on-disk opcode to its canonical VBA7
    table index.

    On 64-bit hosts, raw == canonical (identity). On 32-bit VBA6/VBA7
    hosts, small +1/+2/+3 offsets are applied above raw 173."""
    if is_64bit:
        return raw
    if raw <= 173:
        return raw
    if raw <= 175:
        return raw + 1
    if raw <= 178:
        return raw + 2
    return raw + 3


@dataclass(frozen=True)
class PCodeInstruction:
    """A single decoded p-code instruction.

    Attributes:
        offset: Byte offset within the per-line p-code region.
        raw_word: The original u16 instruction header (opcode + op_type).
        opcode: Canonical VBA7 opcode index (see :data:`OPCODES_VBA7`).
        op_type: Upper-6-bit op_type flags from the header.
        mnemonic: Human-readable mnemonic (e.g. ``"LitStr"``).
        operands: Tuple of (arg_type, raw_value) pairs.
        payload: Variable-length payload (for varg opcodes such as
            ``LitStr``); ``None`` otherwise.
    """

    offset: int
    raw_word: int
    opcode: int
    op_type: int
    mnemonic: str
    operands: tuple[tuple[str, int], ...]
    payload: bytes | None

    def format(self) -> str:
        """Format this instruction as a single-line listing string.

        Operand width is fixed per arg type: ``name`` / ``0x`` /
        ``imp_`` -> 4 hex digits (u16); ``func_`` / ``var_`` /
        ``context_`` / ``rec_`` / ``type_`` -> 8 hex digits (u32).
        ``payload`` bytes from varg opcodes are rendered as a
        latin-1-decoded quoted literal when printable, else hex.
        """
        parts = [self.mnemonic]
        if self.op_type:
            parts.append(f"(op_type=0x{self.op_type:02X})")
        for at, v in self.operands:
            width = 4 if at in ("name", "0x", "imp_") else 8
            parts.append(f"{at}{v:0{width}X}")
        if self.payload is not None:
            try:
                txt = self.payload.decode("latin-1")
                if all(32 <= ord(c) < 127 for c in txt):
                    parts.append(f'"{txt}"')
                else:
                    parts.append(f"<{self.payload.hex()}>")
            except Exception:  # pragma: no cover - latin-1 cannot fail
                parts.append(f"<{self.payload.hex()}>")
        return " ".join(parts)


@dataclass(frozen=True)
class PCodeLine:
    """All instructions for one VBA source line."""

    line_no: int
    start_offset: int
    byte_length: int
    instructions: tuple[PCodeInstruction, ...]


@dataclass(frozen=True)
class PCodeProcedure:
    """A grouped Sub / Function / Property body.

    Attributes:
        kind: Mnemonic that opened the procedure -- one of
            ``"FuncDefn"``, ``"PropertyGet"``, ``"PropertyLet"``,
            ``"PropertySet"``.
        instructions: All instructions from the opener through the
            matching ``EndSub`` / ``EndFunction`` / ``EndProperty``
            (inclusive). Empty if the procedure was never terminated.
    """

    kind: str
    instructions: tuple[PCodeInstruction, ...]


@dataclass(frozen=True)
class DisassembledModule:
    """The full disassembled p-code for one VBA module."""

    cafe_offset: int
    num_lines: int
    lines: tuple[PCodeLine, ...] = field(default_factory=tuple)

    def iter_instructions(self) -> "list[PCodeInstruction]":
        """Flatten every instruction across every line, in source order."""
        return [ins for line in self.lines for ins in line.instructions]

    def iter_procedures(self) -> "list[PCodeProcedure]":
        """Group instructions into procedures.

        A procedure starts with a ``FuncDefn`` / ``PropertyGet`` /
        ``PropertyLet`` / ``PropertySet`` instruction and ends with
        the next ``EndSub`` / ``EndFunction`` / ``EndProperty``. The
        terminating instruction is included in the procedure's body.

        Top-level instructions outside any procedure (e.g.
        ``Option Compare Database`` -> ``Option``) are NOT included
        in any procedure -- request them via :meth:`iter_instructions`
        if needed.
        """
        starts = {
            "FuncDefn",
            "PropertyGet",
            "PropertyLet",
            "PropertySet",
        }
        ends = {"EndSub", "EndFunction", "EndProperty"}
        procs: list[PCodeProcedure] = []
        current: list[PCodeInstruction] | None = None
        kind: str = ""
        for ins in self.iter_instructions():
            if current is None:
                if ins.mnemonic in starts:
                    current = [ins]
                    kind = ins.mnemonic
                continue
            current.append(ins)
            if ins.mnemonic in ends:
                procs.append(
                    PCodeProcedure(
                        kind=kind,
                        instructions=tuple(current),
                    )
                )
                current = None
                kind = ""
        if current is not None:
            # Procedure body not properly terminated -- emit as-is so
            # the caller can see the truncation rather than dropping it.
            procs.append(
                PCodeProcedure(kind=kind, instructions=tuple(current))
            )
        return procs

    def to_listing(self) -> str:
        """Render the disassembly as a ``.lst``-style multi-line text
        dump suitable for human inspection or diffing.

        Format::

            ; DisassembledModule cafe_offset=0xNNN num_lines=N
            ; Line  K  off=0xNNN  len=N
                <mnemonic> <operands>
                <mnemonic> <operands>
            ; Line  K+1  (source-only)

        Source-only lines (``Attribute``, ``Option``, blank) carry
        no p-code; they are emitted as a single comment to preserve
        source-line numbering.
        """
        out: list[str] = []
        out.append(
            f"; DisassembledModule cafe_offset=0x{self.cafe_offset:X} "
            f"num_lines={self.num_lines}"
        )
        for line in self.lines:
            if not line.instructions:
                out.append(
                    f"; Line {line.line_no:3d}  (source-only, "
                    f"len={line.byte_length})"
                )
                continue
            out.append(
                f"; Line {line.line_no:3d}  off=0x{line.start_offset:X}  "
                f"len={line.byte_length}"
            )
            for ins in line.instructions:
                out.append(f"    {ins.format()}")
        return "\n".join(out)


def find_cafe_offset(data: bytes) -> int:
    """Return the offset of the ``0xCAFE`` magic word in ``data``, or
    ``-1`` if not present. CAFE is stored in little-endian byte order
    on the wire (``FE CA``)."""
    return data.find(b"\xfe\xca")


def _decode_line(
    data: bytes, start: int, length: int, is_64bit: bool
) -> tuple[PCodeInstruction, ...]:
    """Decode all p-code instructions for a single source line."""
    out: list[PCodeInstruction] = []
    off = start
    end = start + length
    while off < end:
        if off + 2 > len(data):
            break
        word = int.from_bytes(data[off:off + 2], "little")
        instr_offset = off
        off += 2
        raw_opcode = word & 0x03FF
        op_type = (word & ~0x03FF) >> 10
        opcode = _translate_opcode(raw_opcode, is_64bit)
        op = OPCODES_VBA7.get(opcode)
        if op is None:
            out.append(
                PCodeInstruction(
                    offset=instr_offset,
                    raw_word=word,
                    opcode=opcode,
                    op_type=op_type,
                    mnemonic=f"Unknown_{opcode:#x}",
                    operands=(),
                    payload=None,
                )
            )
            break
        mnemonic, arg_types, has_varg = op
        operands: list[tuple[str, int]] = []
        for arg in arg_types:
            if arg in ("name", "0x", "imp_"):
                if off + 2 > end:
                    break
                value = int.from_bytes(data[off:off + 2], "little")
                off += 2
                operands.append((arg, value))
            elif arg in ("func_", "var_", "rec_", "type_", "context_"):
                if off + 4 > end:
                    break
                value = int.from_bytes(data[off:off + 4], "little")
                off += 4
                operands.append((arg, value))
            else:  # pragma: no cover -- table is closed
                raise ValueError(f"unknown arg type {arg!r}")
        payload: bytes | None = None
        if has_varg:
            if off + 2 > end:
                break
            wlen = int.from_bytes(data[off:off + 2], "little")
            off += 2
            payload = bytes(data[off:off + wlen])
            off += wlen
            if wlen & 1:
                off += 1
        out.append(
            PCodeInstruction(
                offset=instr_offset,
                raw_word=word,
                opcode=opcode,
                op_type=op_type,
                mnemonic=mnemonic,
                operands=tuple(operands),
                payload=payload,
            )
        )
    return tuple(out)


# Sentinel line_offset meaning "this source line has no p-code"
# (Attribute lines, Option-only lines, blank lines, etc.). On 32-bit
# hosts this is u32 0xFFFFFFFF; treat it as a tombstone.
_LINE_OFFSET_NONE = 0xFFFFFFFF


def disassemble_module_stream(
    data: bytes | memoryview | Sequence[int],
    *,
    is_64bit: bool = True,
) -> DisassembledModule:
    """Decode the canonical VBA7 p-code embedded in a module stream.

    ``data`` should be the raw bytes of the module stream (the entire
    LVAL row from :class:`pyopenvba.access.AccessVBAModuleStream`, or
    the equivalent OLE stream from a non-Access Office host). The
    function locates the ``0xCAFE`` magic word, parses the
    ``numLines`` field and the per-line record table, and decodes
    each source line's p-code into :class:`PCodeInstruction` records.

    Identifier / declaration / indirect / object tables are NOT
    resolved here; instruction operands carry raw u16 / u32 values
    (e.g. an identifier-table index). A future revision will resolve
    these against the dir-stream and project-info tables.
    """
    buf = bytes(data)
    cafe = find_cafe_offset(buf)
    if cafe < 0:
        return DisassembledModule(cafe_offset=-1, num_lines=0, lines=())
    # CAFE is at cafe..cafe+2. pcodedmp skips the next 2 bytes before
    # reading numLines.
    off = cafe + 2 + 2
    if off + 2 > len(buf):
        return DisassembledModule(cafe_offset=cafe, num_lines=0, lines=())
    num_lines = int.from_bytes(buf[off:off + 2], "little")
    off += 2
    # Line record table: each record is 12 bytes
    # (4 skip, u16 lineLength, 2 skip, u32 lineOffset). After the
    # table there is a 10-byte gap, then the per-line p-code starts.
    line_records: list[tuple[int, int]] = []  # (length, offset)
    rec_off = off
    for _ in range(num_lines):
        if rec_off + 12 > len(buf):
            break
        line_length = int.from_bytes(buf[rec_off + 4:rec_off + 6], "little")
        line_offset = int.from_bytes(buf[rec_off + 8:rec_off + 12], "little")
        line_records.append((line_length, line_offset))
        rec_off += 12
    pcode_start = off + num_lines * 12 + 10
    lines_out: list[PCodeLine] = []
    for i, (line_length, line_offset) in enumerate(line_records):
        if line_offset == _LINE_OFFSET_NONE or line_length <= 0:
            # Source-only line (Attribute, Option, blank, etc.); no
            # p-code emitted.
            lines_out.append(
                PCodeLine(
                    line_no=i,
                    start_offset=-1,
                    byte_length=line_length,
                    instructions=(),
                )
            )
            continue
        line_start = pcode_start + line_offset
        instructions = _decode_line(buf, line_start, line_length, is_64bit)
        lines_out.append(
            PCodeLine(
                line_no=i,
                start_offset=line_start,
                byte_length=line_length,
                instructions=instructions,
            )
        )
    return DisassembledModule(
        cafe_offset=cafe, num_lines=num_lines, lines=tuple(lines_out)
    )
