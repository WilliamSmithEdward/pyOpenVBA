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

from pcode_decompile import (  # noqa: E402
    find_decl_base, read_signature, resolve_decl, resolve_type,
)
from pcode_names import parse_identifiers, resolve_name  # noqa: E402
from pyopenvba.vba_pcode import disassemble_module_stream  # noqa: E402

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

# Literal opcodes whose value is carried in operands or payload.
_LIT_INT = {"LitDI2", "LitHI2", "LitDI4", "LitHI4", "LitDI8", "LitHI8",
            "LitOI2", "LitOI4", "LitOI8"}


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
                return resolve_type(module_stream, v, dbase)
        return None

    # Pair each FuncDefn with the End* that closes it, so Sub /
    # Function / Property and the return type can be rendered.
    flat = [i for l in dis.lines for i in l.instructions]
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
        elif i.mnemonic in ("EndType", "EndEnum"):
            if rec_pending:
                rec_kind[rec_pending.pop(0)] = i.mnemonic
    pending: list[int] = []
    for i in flat:
        if i.mnemonic in ("FuncDefn", "FuncDefnSave"):
            for a, v in i.operands:
                if a == "func_":
                    pending.append(v)
        elif i.mnemonic in ("EndSub", "EndFunc", "EndFunction", "EndProp"):
            if pending:
                proc_kind[pending.pop(0)] = i.mnemonic

    out: list[str] = []
    in_case = [False]
    indent = 1
    pending_for: list[str] = []          # for-variable names awaiting For

    def emit(text: str, delta_before: int = 0, delta_after: int = 0) -> None:
        nonlocal indent
        indent = max(0, indent + delta_before)
        out.append("    " * indent + text)
        indent = max(0, indent + delta_after)

    for line in dis.lines:
        ins = list(line.instructions)
        if not ins:
            continue
        f = _Frame()
        decls: list[str] = []
        dim_kw = ["Dim"]
        dim_implicit = [False]
        stmt: str | None = None
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
                        local_offsets=local_offsets)
                else:
                    pname, params, ret = None, [], None
                arglist = ", ".join(
                    f"{a} As {b}" if b else a for a, b in params)
                sig = f"{kw} {pname or dname(op) or '<proc>'}({arglist})"
                if ret and kw != "Sub":
                    sig += f" As {ret}"
                out.append(sig)
                indent = 1
            elif m == "EndSub":
                out.append("End Sub"); indent = 1
            elif m in ("EndFunction", "EndFunc"):
                out.append("End Function"); indent = 1
            elif m == "EndProp":
                out.append("End Property"); indent = 1
            elif m in ("Dim", "DimImplicit"):
                dim_kw[0] = DIM_KEYWORDS.get(op.op_type, "Dim")
                dim_implicit[0] = (m == "DimImplicit")
            elif m.startswith("VarDefn"):
                vn = dname(op) or "<var>"
                vt = dtype(op)
                pending = list(f.stack); f.stack.clear()
                if op.op_type == VARDEFN_CONST:
                    text = f"{vn} As {vt}" if vt else vn
                    if pending:
                        text += f" = {pending[-1]}"
                elif len(pending) == 2:
                    text = f"{vn}({pending[0]} To {pending[1]})"
                    if vt: text += f" As {vt}"
                elif len(pending) == 1:
                    text = f"{vn}({pending[0]})"
                    if vt: text += f" As {vt}"
                else:
                    text = f"{vn} As {vt}" if vt else vn
                decls.append(text)
            elif m == "Ld" or m == "LdLHS":
                f.push(nm(op))
            elif m in _LIT_INT:
                vals = [v for a, v in op.operands]
                f.push(str(vals[0] if vals else 0))
            elif m == "LitStr":
                f.push('"' + (op.payload or b"").decode("latin-1") + '"')
            elif m == "LitNothing":
                f.push("Nothing")
            elif m == "LitDefault":
                f.push("")
            elif m == "LitVarSpecial":
                f.push("Empty")
            elif m in BINARY_OPS:
                sym, _ = BINARY_OPS[m]
                rhs = f.pop(); lhs = f.pop()
                joiner = f" {sym} " if sym.isalpha() or len(sym) > 1 else f" {sym} "
                f.push(f"{lhs}{joiner}{rhs}")
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
                call = f"{callee}({', '.join(args)})" if args else callee
                # statement position if nothing consumes it
                if i == len(ins) - 1:
                    stmt = f"{callee} {', '.join(args)}".rstrip() if args else callee
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
            elif m in ("ArgsMemLdWith", "ArgsDictLdWith"):
                argc = next((v for a, v in op.operands if a == "0x"), 0)
                args = [f.pop() for _ in range(min(argc, len(f.stack)))][::-1]
                f.push(f".{nm(op)}({', '.join(args)})")
            elif m == "StartWithExpr":
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
                tgt = nm(op)
                stmt = "On Error Resume Next" if tgt in ("<?>",) else f"On Error GoTo {tgt}"
            elif m == "Label":
                emit(f"{nm(op)}:", -1, 1) if False else out.append(f"{nm(op)}:")
            elif m == "Resume":
                tgt = nm(op)
                stmt = "Resume Next" if not tgt or tgt.startswith("var") else f"Resume {tgt}"
            elif m == "GoTo":
                stmt = f"GoTo {nm(op)}"
            elif m == "GoSub":
                stmt = f"GoSub {nm(op)}"
            elif m == "Type":
                ro = next((v for a, v in op.operands if a == "rec_"), None)
                kw = "Enum" if rec_kind.get(ro) == "EndEnum" else "Type"
                out.append(f"{kw} {dname(op) or '<name>'}"); indent = 1
            elif m == "EndType":
                out.append("End Type"); indent = 1
            elif m == "EndEnum":
                out.append("End Enum"); indent = 1
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
                to = f.pop(); frm = f.pop()
                var = pending_for.pop() if pending_for else "<i>"
                text = f"For {var} = {frm} To {to}"
                if step is not None:
                    text += f" Step {step}"
                emit(text, 0, 1); stmt = None
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
                stmt = f"If {f.pop()} Then"
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
                vals = list(f.stack); f.stack.clear()
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
            elif m in ("QuoteRem", "Rem"):
                stmt = "'" + (op.payload or b"").decode("latin-1")
            elif m in ("BoS", "BoSImplicit", "BoL", "Coerce", "CoerceVar",
                       "LitSmallI2", "EndContext", "Context", "Option",
                       "OptionBase", "DimImplicit", "NewRedim"):
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
            emit(stmt)
    return "\n".join(out)
