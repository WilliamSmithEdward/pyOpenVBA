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

from pcode_decompile import find_decl_base, resolve_decl, resolve_type  # noqa: E402
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
            if a in ("func_", "var_"):
                return resolve_decl(module_stream, v, dbase, ids)
        return None

    def dtype(ins) -> str | None:
        for a, v in ins.operands:
            if a == "var_":
                return resolve_type(module_stream, v, dbase)
        return None

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
        stmt: str | None = None
        i = 0
        while i < len(ins):
            op = ins[i]
            m = op.mnemonic

            if m in ("FuncDefn", "FuncDefnSave"):
                out.append(f"Sub {dname(op) or '<proc>'}()")
                indent = 1
            elif m == "EndSub":
                out.append("End Sub"); indent = 1
            elif m in ("EndFunction", "EndFunc"):
                out.append("End Function"); indent = 1
            elif m == "EndProp":
                out.append("End Property"); indent = 1
            elif m == "Dim":
                pass                       # handled by VarDefn below
            elif m.startswith("VarDefn"):
                vn = dname(op) or "<var>"
                vt = dtype(op)
                f.push(f"{vn} As {vt}" if vt else vn)
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
                rhs = f.pop(); lhs = f.pop() if f.stack else "<?>"
                stmt = f"Set {lhs} = {rhs}"
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

        if stmt is None and f.stack and not any(
            x.mnemonic in ("Dim",) for x in ins
        ):
            leftover = " ".join(f.stack)
            if leftover.strip():
                stmt = leftover
        if any(x.mnemonic == "Dim" for x in ins):
            stmt = "Dim " + ", ".join(f.stack) if f.stack else None
        if stmt:
            emit(stmt)
    return "\n".join(out)
