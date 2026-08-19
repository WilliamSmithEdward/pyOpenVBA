"""VBA source reconstruction from compiled p-code.

P-code is a stack machine, so expressions reconstruct by simulation:
``Ld b | Ld c | LitDI2 2 | Mul | Add | St a`` pops its way back to
``a = b + c * 2``. Control-flow opcodes (``IfBlock``, ``For``,
``DoWhile``, ``SelectCase``, ...) drive block structure and indentation.

Host-agnostic: consumes only the module stream plus the project
identifier table, both identical across Excel / Word / PowerPoint /
Access.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from pcode_decompile import (
    describe_type,
    find_decl_base,
    read_signature,
    resolve_decl,
    resolve_type,
)
from pcode_names import parse_identifiers, resolve_name
from pcode_types import find_type_table

from pyopenvba.vba_pcode import disassemble_module_stream

# opcode -> (VBA operator, precedence). Higher binds tighter.
BINARY_OPS: dict[str, tuple[str, int]] = {
    "Imp": ("Imp", 1), "Eqv": ("Eqv", 2), "Xor": ("Xor", 3),
    "Or": ("Or", 4), "And": ("And", 5),
    "Eq": ("=", 6), "Ne": ("<>", 6), "Lt": ("<", 6), "Gt": (">", 6),
    "Le": ("<=", 6), "Ge": (">=", 6), "Is": ("Is", 6), "Like": ("Like", 6),
    "Concat": ("&", 7),
    "Add": ("+", 8), "Sub": ("-", 8),
    "Mod": ("Mod", 9), "IDiv": ("\\", 10),
    "Mul": ("*", 11), "Div": ("/", 11),
    "Pwr": ("^", 13),
}
UNARY_OPS: dict[str, str] = {"Not": "Not ", "UMi": "-"}

# The declaration keyword is carried in the Dim opcode's op_type, and
# whether an entry is a constant in the VarDefn's op_type. That also
# disambiguates a Const's value from an array's bounds, which otherwise
# look identical (both are literals pushed before VarDefn).
DIM_KEYWORDS: dict[int, str] = {
    0x00: "Dim", 0x01: "Const", 0x08: "Public", 0x10: "Private",
    0x20: "Static",
}
VARDEFN_CONST = 0x02

# Literal opcodes whose value is carried in operands or payload. The
# operands are u16 words, little-endian least-significant first.
_LIT_INT = {"LitDI2": (1, 10), "LitDI4": (2, 10), "LitDI8": (4, 10),
            "LitHI2": (1, 16), "LitHI4": (2, 16), "LitHI8": (4, 16),
            "LitOI2": (1, 8), "LitOI4": (2, 8), "LitOI8": (4, 8)}

# LitVarSpecial carries the value in its op_type.
LIT_SPECIAL: dict[int, str] = {0: "False", 1: "True", 2: "Null", 3: "Empty"}

# Intrinsics the compiler emits as a dedicated opcode rather than a call.
# value = (source name, argument count).
INTRINSICS: dict[str, tuple[str, int]] = {
    "FnAbs": ("Abs", 1), "FnFix": ("Fix", 1), "FnInt": ("Int", 1),
    "FnSgn": ("Sgn", 1), "FnLen": ("Len", 1), "FnLenB": ("LenB", 1),
    "FnCurDir": ("CurDir", 1), "FnDir": ("Dir", 1), "FnError": ("Error", 1),
    "FnFormat": ("Format", 2), "FnFreeFile": ("FreeFile", 1),
    "FnInStr": ("InStr", 2), "FnInStr3": ("InStr", 3), "FnInStr4": ("InStr", 4),
    "FnInStrB": ("InStrB", 2), "FnInStrB3": ("InStrB", 3),
    "FnInStrB4": ("InStrB", 4), "FnMid": ("Mid", 2), "FnMidB": ("MidB", 2),
    "FnStrComp": ("StrComp", 2), "FnStrComp3": ("StrComp", 3),
    "FnStringVar": ("String", 2), "FnStringStr": ("String", 2),
}

# FuncDefn op_type. Bit 0x02 marks a value-returning procedure
# (Function, Property Get); bit 0x04 marks an explicit Public keyword.
FUNC_RETURNS_VALUE = 0x02
FUNC_EXPLICIT_PUBLIC = 0x04
# The FuncDefn record's own visibility byte; bit 0x02 = effectively public.
FUNC_VISIBILITY_OFFSET = 0x51
FUNC_PUBLIC_FLAG = 0x02

# Type op_type: bit 0x01 = a visibility keyword was written, bit 0x02 =
# Enum rather than Type. The record's u16 at +16 is 1 for Public.
TYPE_KEYWORD_WRITTEN = 0x01
TYPE_VISIBILITY_OFFSET = 16

# Every declaration record is introduced by a u16 tag in the two bytes
# before it; bit 0x20 of its low byte marks an explicit As clause, which
# is the only thing separating "Dim x" from "Dim x As Variant".
DECL_TAG_HAS_AS = 0x20

# ArgsCall op_type 0x10 marks a bare call statement; without it the
# source used the Call keyword.
CALL_WITHOUT_KEYWORD = 0x10

# Coerce op_type -> conversion function. Mostly VARTYPE codes, with
# CVar at 0 and CLngLng at 0x0D rather than VT_I8.
COERCIONS: dict[int, str] = {
    0x00: "CVar", 0x02: "CInt", 0x03: "CLng", 0x04: "CSng", 0x05: "CDbl",
    0x06: "CCur", 0x07: "CDate", 0x08: "CStr", 0x0B: "CBool", 0x0D: "CLngLng",
    0x11: "CByte",
}

# Open ... For <mode>; the mode is the Open opcode's operand.
OPEN_MODES: dict[int, str] = {
    1: "Input", 2: "Output", 4: "Append", 8: "Binary", 16: "Random",
}

# On Error / Resume forms are selected by op_type.
ONERROR_FORMS: dict[int, str] = {1: "On Error Resume Next",
                                 2: "On Error GoTo 0"}
RESUME_FORMS: dict[int, str] = {1: "Resume Next", 8: "Resume"}

# Option statement forms (op_type). The argument is not carried in the
# p-code, so Option Base 1 and Option Compare Text render by keyword.
OPTION_FORMS: dict[int, str] = {
    1: "Option Base 1", 2: "Option Compare Text", 4: "Option Explicit",
    5: "Option Private Module",
}

# DefType op_type is a VARTYPE; its two operands are a 64-bit bitmap of
# the letters the statement covers (bit 0 = A).
DEFTYPE_NAMES: dict[int, str] = {
    0x02: "DefInt", 0x03: "DefLng", 0x04: "DefSng", 0x05: "DefDbl",
    0x06: "DefCur", 0x07: "DefDate", 0x08: "DefStr", 0x09: "DefObj",
    0x0B: "DefBool", 0x0C: "DefVar", 0x0D: "DefLngLng", 0x0E: "DefDec",
    0x11: "DefByte",
}


def _render_float(mnemonic: str, operands) -> str:
    """Render LitR4 / LitR8 / LitCy / LitDate from their u16 words."""
    import datetime
    import struct
    raw = b"".join(v.to_bytes(2, "little") for _, v in operands)
    if mnemonic == "LitR4":
        return repr(struct.unpack("<f", raw[:4])[0])
    if mnemonic == "LitCy":
        return repr(int.from_bytes(raw[:8], "little", signed=True) / 10000)
    value = struct.unpack("<d", raw[:8])[0]
    if mnemonic == "LitR8":
        text = repr(value)
        return text[:-2] if text.endswith(".0") else text
    # OLE automation date: days since 1899-12-30.
    stamp = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=value)
    if stamp.time() == datetime.time(0):
        return "#" + stamp.strftime("%m/%d/%Y") + "#"
    return "#" + stamp.strftime("%m/%d/%Y %I:%M:%S %p") + "#"


def _flush(stmt: str | None, inline_if: dict | None,
           line_stmts: list[str]) -> None:
    """File a completed statement under the line it belongs to.

    A source line can hold several statements -- separated by ``:`` or
    inside the single-line ``If ... Then ... Else ...`` form -- so each
    one is filed as it completes and the line is rejoined at the end.
    """
    if stmt is None:
        return
    if inline_if is not None:
        segment = (inline_if["else"] if inline_if["else"] is not None
                   else inline_if["then"])
        segment.append(stmt)
    else:
        line_stmts.append(stmt)


class _Frame:
    """One reconstructed statement's evaluation stack."""

    def __init__(self) -> None:
        self.stack: list[str] = []

    def push(self, s: str) -> None:
        self.stack.append(s)

    def pop(self) -> str:
        return self.stack.pop() if self.stack else "<?>"


def decompile(module_stream: bytes, vba_project_stream: bytes,
              *, is_64bit: bool = True) -> str:
    """Reconstruct VBA source from a compiled module stream."""
    ids = parse_identifiers(vba_project_stream)
    dis = disassemble_module_stream(module_stream, is_64bit=is_64bit)
    dbase = find_decl_base(module_stream, ids, is_64bit=is_64bit)
    # Named types (UDT, Enum, type-library class) resolve through the
    # module's own type-reference table; see pcode_types.
    ttab = find_type_table(module_stream, dbase) if dbase is not None else []
    def resolver(operand: int) -> str | None:
        return resolve_name(operand, ids)

    def record_byte(operand: int, offset: int) -> int | None:
        """A byte inside the declaration record at DECL_BASE + operand."""
        if dbase is None:
            return None
        p = dbase + operand + offset
        return module_stream[p] if 0 <= p < len(module_stream) else None

    def record_tag(operand: int) -> int:
        """The u16 tag introducing a declaration record."""
        if dbase is None:
            return 0
        p = dbase + operand - 2
        if p < 0 or p + 2 > len(module_stream):
            return 0
        return int.from_bytes(module_stream[p:p + 2], "little")

    def proc_prefix(op) -> str:
        """"Public " / "Private " / "" for a FuncDefn."""
        fo = next((v for a, v in op.operands if a == "func_"), None)
        if fo is None:
            return ""
        if op.op_type & FUNC_EXPLICIT_PUBLIC:
            return "Public "
        flags = record_byte(fo, FUNC_VISIBILITY_OFFSET)
        if flags is not None and not flags & FUNC_PUBLIC_FLAG:
            return "Private "
        return ""

    def nm(ins) -> str:
        for a, v in ins.operands:
            if a == "name":
                return resolve_name(v, ids) or f"var{v:04X}"
        return "<?>"

    def dname(ins) -> str | None:
        for a, v in ins.operands:
            if a in ("func_", "var_", "rec_"):
                return resolve_decl(module_stream, v, dbase, ids)
        return None

    def dtype(ins) -> str | None:
        for a, v in ins.operands:
            if a == "var_":
                return resolve_type(module_stream, v, dbase, ttab, resolver)
        return None

    def dtype_full(ins):
        for a, v in ins.operands:
            if a == "var_":
                return describe_type(module_stream, v, dbase, ttab, resolver)
        return None

    # Pair each FuncDefn with the End* that closes it, so Sub /
    # Function / Property and the return type can be rendered.
    flat = [i for line in dis.lines for i in line.instructions]
    local_offsets = frozenset(
        v for i in flat if i.mnemonic.startswith("VarDefn")
        for a, v in i.operands if a == "var_"
    )
    proc_kind: dict[int, str] = {}
    # A Type opcode opens either a user-defined type or an Enum; only the
    # closing opcode (EndType vs EndEnum) tells them apart.
    rec_kind: dict[int, str] = {}
    rec_pending: list[int] = []
    for i in flat:
        if i.mnemonic == "Type":
            for a, v in i.operands:
                if a == "rec_":
                    rec_pending.append(v)
        elif i.mnemonic in ("EndType", "EndEnum") and rec_pending:
            rec_kind[rec_pending.pop(0)] = i.mnemonic
    pending: list[int] = []
    for i in flat:
        if i.mnemonic in ("FuncDefn", "FuncDefnSave"):
            for a, v in i.operands:
                if a == "func_":
                    pending.append(v)
        elif i.mnemonic in ("EndSub", "EndFunc", "EndFunction",
                            "EndProp") and pending:
            proc_kind[pending.pop(0)] = i.mnemonic

    out: list[str] = []
    in_case = [False]
    indent = 0
    pending_for: list[str] = []          # for-variable names awaiting For
    channel_kw: list = [None]            # Print/Write/Debug channel in flight
    input_chan: list = [None]
    input_items: list[str] = []

    def emit(text: str, delta_before: int = 0, delta_after: int = 0) -> None:
        nonlocal indent
        indent = max(0, indent + delta_before)
        out.append("    " * indent + text)
        indent = max(0, indent + delta_after)

    for line in dis.lines:
        ins = list(line.instructions)
        if not ins:
            # A source line with no p-code is a blank line (or a comment
            # the compiler dropped); keeping it preserves the layout.
            out.append("")
            continue
        f = _Frame()
        decls: list[str] = []
        dim_kw = ["Dim"]
        dim_implicit = [False]
        stmt: str | None = None
        trailing: list[str | None] = [None]
        # A source line can hold several statements: colon-separated, or
        # the single-line If ... Then ... Else ... form. They are
        # accumulated here and rejoined when the line ends.
        line_stmts: list[str] = []
        inline_if: dict | None = None
        i = 0

        while i < len(ins):
            op = ins[i]
            m = op.mnemonic

            if m in ("FuncDefn", "FuncDefnSave"):
                fo = next((v for a, v in op.operands if a == "func_"), None)
                end = proc_kind.get(fo, "EndSub") if fo is not None else "EndSub"
                kw = {"EndProp": "Property", "EndFunc": "Function",
                      "EndFunction": "Function"}.get(end, "Sub")
                if fo is not None:
                    pname, params, ret = read_signature(
                        module_stream, fo, dbase, ids,
                        is_function=(kw != "Sub"),
                        local_offsets=local_offsets, type_table=ttab,
                        defaults=list(f.stack))
                    f.stack.clear()
                else:
                    pname, params, ret = None, [], None
                if kw == "Property":
                    # Get returns a value, Let and Set do not; Let and Set
                    # are indistinguishable in p-code, so an object-typed
                    # final parameter is what picks Set.
                    if op.op_type & FUNC_RETURNS_VALUE:
                        kw = "Property Get"
                    elif params and params[-1][1] in ("Object", None):
                        kw = "Property Set"
                    else:
                        kw = "Property Let"
                arglist = ", ".join(
                    f"{a} As {b}" if b else a for a, b in params)
                sig = (f"{proc_prefix(op)}{kw} "
                       f"{pname or dname(op) or '<proc>'}({arglist})")
                if ret and not kw.startswith(("Sub", "Property Let",
                                              "Property Set")):
                    sig += f" As {ret}"
                out.append(sig)
                indent = 1
            elif m == "EndSub":
                out.append("End Sub")
                indent = 0
            elif m in ("EndFunction", "EndFunc"):
                out.append("End Function")
                indent = 0
            elif m == "EndProp":
                out.append("End Property")
                indent = 0
            elif m in ("Dim", "DimImplicit"):
                dim_kw[0] = DIM_KEYWORDS.get(op.op_type, "Dim")
                dim_implicit[0] = (m == "DimImplicit")
            elif m.startswith("VarDefn"):
                vo = next((v for a, v in op.operands if a == "var_"), None)
                vn = dname(op) or "<var>"
                info = dtype_full(op)
                vt = info.render() if info else None
                if vo is not None and not record_tag(vo) & DECL_TAG_HAS_AS:
                    # No As clause in the source. The record still holds a
                    # type -- Variant by default, or whatever a DefType
                    # statement assigned to the initial letter.
                    vt = None
                pending = list(f.stack)
                f.stack.clear()
                if info is not None and info.string_length is not None:
                    # The length was pushed as a literal; it is already
                    # part of the rendered type.
                    pending = pending[:-1]
                shape = ""
                if op.op_type == VARDEFN_CONST:
                    pass
                elif len(pending) == 2:
                    shape = f"({pending[0]} To {pending[1]})"
                    pending = []
                elif len(pending) == 1:
                    shape = f"({pending[0]})"
                    pending = []
                elif info is not None and info.array:
                    # Dynamic array: no bounds were pushed, but the
                    # descriptor still records that it is an array.
                    shape = "()"
                text = vn + shape
                if vt:
                    text += f" As {vt}"
                if op.op_type == VARDEFN_CONST and pending:
                    text += f" = {pending[-1]}"
                decls.append(text)
            elif m == "Ld" or m == "LdLHS":
                f.push(nm(op))
            elif m in _LIT_INT:
                words, radix = _LIT_INT[m]
                value = 0
                for k, (_, v) in enumerate(op.operands[:words]):
                    value |= v << (16 * k)
                bits = 16 * words
                if value >= 1 << (bits - 1):
                    value -= 1 << bits
                if radix == 16:
                    f.push(f"&H{value:X}")
                elif radix == 8:
                    f.push(f"&O{value:o}")
                else:
                    f.push(str(value))
            elif m == "LitSmallI2":
                f.push(str(op.op_type))
            elif m in ("LitR4", "LitR8", "LitCy", "LitDate"):
                f.push(_render_float(m, op.operands))
            elif m == "LitStr":
                f.push('"' + (op.payload or b"").decode("latin-1") + '"')
            elif m == "LitNothing":
                f.push("Nothing")
            elif m == "LitDefault":
                f.push("")
            elif m == "LitVarSpecial":
                f.push(LIT_SPECIAL.get(op.op_type, "Empty"))
            elif m in INTRINSICS:
                fname, argc = INTRINSICS[m]
                args = [f.pop() for _ in range(min(argc, len(f.stack)))][::-1]
                f.push(f"{fname}({', '.join(args)})")
            elif m in ("FnLBound", "FnUBound"):
                fname = "LBound" if m == "FnLBound" else "UBound"
                dim = next((v for a, v in op.operands if a == "0x"), 0)
                target = f.pop()
                f.push(f"{fname}({target})" if not dim
                       else f"{fname}({target}, {dim + 1})")
            elif m in BINARY_OPS:
                sym, _ = BINARY_OPS[m]
                rhs = f.pop()
                lhs = f.pop()
                f.push(f"{lhs} {sym} {rhs}")
            elif m in UNARY_OPS:
                f.push(UNARY_OPS[m] + f.pop())
            elif m == "Paren":
                f.push("(" + f.pop() + ")")
            elif m in ("St", "SetOrSt"):
                stmt = f"{nm(op)} = {f.pop()}"
            elif m == "SetStmt":
                pass                     # marker; the Set opcode assigns
            elif m == "Set":
                stmt = f"Set {nm(op)} = {f.pop()}"
            elif m == "MemSet":
                obj = f.pop() if f.stack else "<?>"
                stmt = f"Set {obj}.{nm(op)} = {f.pop()}"
            elif m == "SetWith" or m == "MemSetWith":
                stmt = f"Set .{nm(op)} = {f.pop()}"
            elif m in ("ArgsCall", "ArgsMemCall", "ArgsMemCallWith"):
                argc = 0
                for a, v in op.operands:
                    if a == "0x":
                        argc = v
                args = [f.pop() for _ in range(min(argc, len(f.stack)))][::-1]
                callee = nm(op)
                if m == "ArgsMemCallWith":
                    callee = "." + callee
                elif m == "ArgsMemCall" and f.stack:
                    callee = f"{f.pop()}.{callee}"
                call = f"{callee}({', '.join(args)})" if args else callee
                # statement position if nothing consumes it
                if i == len(ins) - 1:
                    if op.op_type & CALL_WITHOUT_KEYWORD:
                        stmt = (f"{callee} {', '.join(args)}".rstrip()
                                if args else callee)
                    else:
                        stmt = f"Call {call}"
                else:
                    f.push(call)
            elif m in ("ArgsSt", "ArgsMemSt", "ArgsDictSt"):
                argc = next((v for a, v in op.operands if a == "0x"), 0)
                args = [f.pop() for _ in range(min(argc, len(f.stack)))][::-1]
                value = f.pop()
                target = nm(op)
                if m == "ArgsMemSt" and f.stack:
                    target = f"{f.pop()}.{target}"
                stmt = f"{target}({', '.join(args)}) = {value}"
            elif m == "MemSt":
                obj = f.pop() if f.stack else "<?>"
                stmt = f"{obj}.{nm(op)} = {f.pop()}"
            elif m == "MemStWith":
                stmt = f".{nm(op)} = {f.pop()}"
            elif m == "MemLd":
                f.push(f"{f.pop()}.{nm(op)}")
            elif m == "MemLdWith":
                f.push(f".{nm(op)}")
            elif m in ("ArgsMemStWith", "ArgsDictStWith"):
                argc = next((v for a, v in op.operands if a == "0x"), 0)
                args = [f.pop() for _ in range(min(argc, len(f.stack)))][::-1]
                stmt = f".{nm(op)}({', '.join(args)}) = {f.pop()}"
            elif m in ("ArgsMemSetWith", "ArgsDictSetWith"):
                argc = next((v for a, v in op.operands if a == "0x"), 0)
                args = [f.pop() for _ in range(min(argc, len(f.stack)))][::-1]
                stmt = f"Set .{nm(op)}({', '.join(args)}) = {f.pop()}"
            elif m in ("DictLd", "DictLdWith"):
                base = "" if m.endswith("With") else f.pop()
                f.push(f"{base}!{nm(op)}")
            elif m in ("ArgsMemLdWith", "ArgsDictLdWith"):
                argc = next((v for a, v in op.operands if a == "0x"), 0)
                args = [f.pop() for _ in range(min(argc, len(f.stack)))][::-1]
                f.push(f".{nm(op)}({', '.join(args)})")
            elif m == "New":
                imp = next((v for a, v in op.operands if a == "imp_"), 0)
                entry = ttab[imp // 8] if 0 <= imp // 8 < len(ttab) else None
                cls = resolver(entry.name_operand) if entry else None
                f.push(f"New {cls or '<class>'}")
            elif m == "ForEach":
                seq = f.pop()
                var = pending_for.pop() if pending_for else "<v>"
                emit(f"For Each {var} In {seq}", 0, 1)
                stmt = None
            elif m == "Coerce":
                f.push(f"{COERCIONS.get(op.op_type, 'CVar')}({f.pop()})")
            elif m == "CoerceVar":
                f.push(f"CVErr({f.pop()})")
            elif m == "Sharp":
                f.push("#" + f.pop())
            elif m == "Open":
                mode = next((v for a, v in op.operands if a == "0x"), 0)
                items = [x for x in f.stack if x != ""]
                f.stack.clear()
                path = items[0] if items else "<?>"
                chan = items[1] if len(items) > 1 else "<?>"
                stmt = (f"Open {path} For {OPEN_MODES.get(mode, mode)} "
                        f"As {chan}")
            elif m in ("Close", "CloseAll"):
                chans = list(f.stack)
                f.stack.clear()
                stmt = ("Close " + ", ".join(chans)).strip() if chans else "Close"
            elif m in ("PrintChan", "WriteChan"):
                channel_kw[0] = ("Print" if m == "PrintChan" else "Write",
                                 f.pop())
            elif m == "Debug":
                channel_kw[0] = ("Debug", None)
            elif m == "PrintObj":
                pass
            elif m in ("PrintItemNL", "PrintItemComma", "PrintItemSemiColon",
                       "PrintNL"):
                items = list(f.stack)
                f.stack.clear()
                kw, chan = channel_kw[0] or ("Print", None)
                head = "Debug.Print" if kw == "Debug" else kw
                target = f"{head} {chan}," if chan else head
                stmt = (f"{target} " + ", ".join(items)).strip()
                channel_kw[0] = None
            elif m == "Assert":
                stmt = f"Debug.Assert {f.pop()}"
            elif m == "LineInput":
                target = f.pop()
                chan = f.pop()
                stmt = f"Line Input #{chan}, {target}"
            elif m == "Input":
                input_chan[0] = f.pop()
            elif m == "InputItem":
                input_items.append(f.pop())
            elif m == "InputDone":
                stmt = f"Input {input_chan[0]}, " + ", ".join(input_items)
                input_chan[0] = None
                input_items.clear()
            elif m == "Name":
                new = f.pop()
                old = f.pop()
                stmt = f"Name {old} As {new}"
            elif m == "Mid":
                args = [f.pop() for _ in range(min(2, len(f.stack)))][::-1]
                target = f.pop() if f.stack else "<?>"
                value = f.pop() if f.stack else "<?>"
                stmt = f"Mid({target}, {', '.join(args)}) = {value}"
            elif m in ("LSet", "RSet"):
                target = f.pop()
                value = f.pop()
                stmt = f"{m} {target} = {value}"
            elif m == "Return":
                stmt = "Return"
            elif m == "Error":
                stmt = f"Error {f.pop()}"
            elif m == "Option":
                stmt = OPTION_FORMS.get(op.op_type, f"Option {op.op_type:#x}")
            elif m == "DefType":
                letters = 0
                for k, (_, v) in enumerate(op.operands):
                    letters |= v << (16 * k)
                spans = []
                run = None
                for bit in range(27):
                    on = bit < 26 and bool(letters >> bit & 1)
                    if on and run is None:
                        run = bit
                    elif not on and run is not None:
                        a, b = chr(65 + run), chr(65 + bit - 1)
                        spans.append(a if a == b else f"{a}-{b}")
                        run = None
                stmt = (f"{DEFTYPE_NAMES.get(op.op_type, 'DefVar')} "
                        + ", ".join(spans))
            elif m == "LbIf":
                emit(f"#If {f.pop()} Then", 0, 0)
            elif m == "LbElseIf":
                emit(f"#ElseIf {f.pop()} Then", 0, 0)
            elif m == "LbElse":
                emit("#Else", 0, 0)
            elif m == "LbEndIf":
                emit("#End If", 0, 0)
            elif m == "LbMark" or m == "StartWithExpr":
                pass
            elif m == "With":
                emit(f"With {f.pop()}", 0, 1)
            elif m == "EndWith":
                emit("End With", -1, 0)
            elif m in ("Redim", "RedimAs", "NewRedim"):
                argc = next((v for a, v in op.operands if a == "0x"), 0)
                args = [f.pop() for _ in range(min(argc, len(f.stack)))][::-1]
                if m != "NewRedim":
                    stmt = f"ReDim {nm(op)}({', '.join(args)})"
            elif m == "OnError":
                stmt = ONERROR_FORMS.get(op.op_type) or f"On Error GoTo {nm(op)}"
            elif m == "Label":
                out.append(f"{nm(op)}:")
            elif m == "Resume":
                stmt = RESUME_FORMS.get(op.op_type) or f"Resume {nm(op)}"
            elif m == "GoTo":
                stmt = f"GoTo {nm(op)}"
            elif m == "GoSub":
                stmt = f"GoSub {nm(op)}"
            elif m == "Type":
                ro = next((v for a, v in op.operands if a == "rec_"), None)
                kw = "Enum" if rec_kind.get(ro) == "EndEnum" else "Type"
                prefix = ""
                if op.op_type & TYPE_KEYWORD_WRITTEN and ro is not None:
                    flag = record_byte(ro, TYPE_VISIBILITY_OFFSET)
                    prefix = "Public " if flag else "Private "
                out.append(f"{prefix}{kw} {dname(op) or '<name>'}")
                indent = 1
            elif m == "EndType":
                out.append("End Type")
                indent = 0
            elif m == "EndEnum":
                out.append("End Enum")
                indent = 0
            elif m == "Stop":
                stmt = "Stop"
            elif m == "DoEvents":
                stmt = "DoEvents"
            elif m == "Erase":
                stmt = f"Erase {f.pop()}"
            elif m in ("ArgsLd", "ArgsMemLd", "IndexLd"):
                argc = 0
                for a, v in op.operands:
                    if a == "0x":
                        argc = v
                args = [f.pop() for _ in range(min(argc, len(f.stack)))][::-1]
                base = nm(op) if m != "IndexLd" else f.pop()
                f.push(f"{base}({', '.join(args)})")
            elif m == "StartForVariable":
                pass
            elif m == "EndForVariable":
                pending_for.append(f.pop())
            elif m in ("For", "ForStep"):
                step = f.pop() if m == "ForStep" else None
                to = f.pop()
                frm = f.pop()
                var = pending_for.pop() if pending_for else "<i>"
                text = f"For {var} = {frm} To {to}"
                if step is not None:
                    text += f" Step {step}"
                emit(text, 0, 1)
                stmt = None
            elif m in ("Next", "NextVar"):
                var = pending_for.pop() if pending_for else ""
                emit(f"Next {var}".rstrip(), -1, 0)
            elif m == "IfBlock":
                emit(f"If {f.pop()} Then", 0, 1)
            elif m == "ElseIfBlock":
                emit(f"ElseIf {f.pop()} Then", -1, 1)
            elif m == "ElseBlock":
                emit("Else", -1, 1)
            elif m == "EndIfBlock":
                emit("End If", -1, 0)
            elif m == "If":
                _flush(stmt, inline_if, line_stmts)
                stmt = None
                inline_if = {"cond": f.pop(), "then": [], "else": None}
            elif m == "Else":
                _flush(stmt, inline_if, line_stmts)
                stmt = None
                if inline_if is not None:
                    inline_if["else"] = []
            elif m == "EndIf":
                _flush(stmt, inline_if, line_stmts)
                stmt = None
                if inline_if is not None:
                    text = f"If {inline_if['cond']} Then " + ": ".join(
                        inline_if["then"])
                    if inline_if["else"]:
                        text += " Else " + ": ".join(inline_if["else"])
                    inline_if = None
                    stmt = text
            elif m in ("DoWhile", "While"):
                kw = "Do While" if m == "DoWhile" else "While"
                emit(f"{kw} {f.pop()}", 0, 1)
            elif m == "DoUntil":
                emit(f"Do Until {f.pop()}", 0, 1)
            elif m == "Do":
                emit("Do", 0, 1)
            elif m == "Loop":
                emit("Loop", -1, 0)
            elif m == "LoopWhile":
                emit(f"Loop While {f.pop()}", -1, 0)
            elif m == "LoopUntil":
                emit(f"Loop Until {f.pop()}", -1, 0)
            elif m == "Wend":
                emit("Wend", -1, 0)
            elif m == "SelectCase":
                emit(f"Select Case {f.pop()}", 0, 1)
                in_case[0] = False
            elif m == "Case":
                pass
            elif m == "CaseEq":
                f.push(f.pop())
            elif m == "CaseDone":
                vals = list(f.stack)
                f.stack.clear()
                emit("Case " + ", ".join(vals), -1 if in_case[0] else 0, 1)
                in_case[0] = True
            elif m == "CaseElse":
                emit("Case Else", -1 if in_case[0] else 0, 1)
                in_case[0] = True
            elif m == "EndSelect":
                emit("End Select", -2 if in_case[0] else -1, 0)
                in_case[0] = False
            elif m in ("ExitSub", "ExitFunc", "ExitFor", "ExitDo"):
                stmt = {"ExitSub": "Exit Sub", "ExitFunc": "Exit Function",
                        "ExitFor": "Exit For", "ExitDo": "Exit Do"}[m]
            elif m == "Reparse":
                # Constructs the p-code compiler does not encode (Friend,
                # Tab, Spc, ...) are kept as the original source text.
                stmt = (op.payload or b"").decode("latin-1").strip()
            elif m == "QuoteRem":
                text = "'" + (op.payload or b"").decode("latin-1")
                if stmt is None and not decls and i == 0:
                    stmt = text
                else:
                    trailing[0] = text
            elif m == "Rem":
                stmt = "Rem" + (op.payload or b"").decode("latin-1")
            elif m in ("BoS", "BoSImplicit"):
                # statement separator: a source ":" or an If body
                _flush(stmt, inline_if, line_stmts)
                stmt = None
            elif m in ("BoL", "EndContext", "Context", "OptionBase",
                       "DimImplicit", "NewRedim", "LineCont", "PSetDefault",
                       "ConstFuncExpr"):
                pass
            else:
                stmt = f"' [unmapped {m}]"
            i += 1

        if stmt is None and f.stack and not decls and not any(
            x.mnemonic in ("Dim", "DimImplicit") for x in ins
        ):
            leftover = " ".join(f.stack)
            if leftover.strip():
                stmt = leftover
        if decls:
            # Type/Enum members carry no declaration keyword.
            stmt = (", ".join(decls) if dim_implicit[0]
                    else f"{dim_kw[0]} " + ", ".join(decls))
        if stmt:
            line_stmts.append(stmt)
        joined = ": ".join(line_stmts)
        if trailing[0]:
            joined = f"{joined}    {trailing[0]}" if joined else trailing[0]
        if joined:
            emit(joined)
    return "\n".join(out)
