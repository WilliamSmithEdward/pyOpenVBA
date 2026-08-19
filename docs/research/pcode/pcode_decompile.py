"""P-code decompiler: compiled VBA7 bytecode -> readable listing/source.

Host-agnostic: consumes the module-stream p-code region (CAFE) plus the
project identifier table, both of which are identical in layout across
Excel, Word, PowerPoint, and Access.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "src"))
from pcode_names import Identifier, parse_identifiers, resolve_name
from pcode_types import (
    DeclaredType,
    TypeRefEntry,
    find_type_table,
    read_declared_type,
)

from pyopenvba.vba_pcode import disassemble_module_stream

# Procedure layout, relative to DECL_BASE + func_operand:
#   +0x00   the procedure's own name entry
#   +0x2C   a type field of the same shape a variable's is (VARTYPE at
#           +14, discriminator at +16), holding the return type
#   +0x36   u32 offset of the first parameter, 0xFFFFFFFF when there are
#           none -- which is what separates parameters from locals, since
#           both live in the same slot region
# and, relative to a parameter's own record:
#   +0x16   u32 offset of the next parameter, 0xFFFFFFFF at the end
#   +0x1A   u16 passing mode and flags
PROC_RETURN_FIELD = 0x2C
PROC_PARAM_HEAD = 0x36
PARAM_NEXT_OFFSET = 0x16
PARAM_MODE_OFFSET = 0x1A
NO_LINK = 0xFFFFFFFF

# Passing mode (low byte of the mode word): 0x04 ByVal, 0x02 an explicit
# ByRef keyword; neither means ByRef written without the keyword.
PARAM_BYVAL = 0x04
PARAM_BYREF_KEYWORD = 0x02
# Flags (high byte): 0x02 Optional, 0x04 the parameter has a default.
PARAM_OPTIONAL = 0x02
PARAM_HAS_DEFAULT = 0x04

# Retained for callers that still walk the slot region directly.
PROC_PARAM_OFFSET = 0x58
PROC_PARAM_STRIDE = 0x20

# Declared types are decoded by pcode_types: a record either carries a
# plain VARTYPE byte or points at an indirect type descriptor (arrays,
# fixed-length strings, UDTs, Enums, type-library classes).
TYPE_FIELD_OFFSET = 14
VARTYPE_FLAG_CONST = 0x40


def describe_type(module_stream: bytes, operand: int, base: int | None,
                  type_table: list[TypeRefEntry] | None = None,
                  resolver=None) -> DeclaredType | None:
    """Decode a declaration operand's full declared type."""
    return read_declared_type(module_stream, base, operand, type_table, resolver)


def resolve_type(module_stream: bytes, operand: int, base: int | None,
                 type_table: list[TypeRefEntry] | None = None,
                 resolver=None) -> str | None:
    """Resolve a ``var_`` operand to its declared VBA type name.

    Without ``type_table``/``resolver`` only intrinsic types resolve;
    supply both (see :func:`pcode_types.find_type_table`) to name
    user-defined types, Enums and type-library classes as well.
    """
    described = describe_type(module_stream, operand, base, type_table, resolver)
    return described.render() if described else None


# The byte after the VARTYPE carries the parameter passing mode:
# 0x00 = ByVal, 0x01 = ByRef (VBA's default). Explicit vs implicit
# ByRef is not distinguished -- both compile to 0x01.
BYREF_FLAG_OFFSET = TYPE_FIELD_OFFSET + 1

def is_byval(module_stream: bytes, operand: int, base: int | None) -> bool:
    """True when a parameter is passed ByVal."""
    if base is None:
        return False
    p = base + operand + BYREF_FLAG_OFFSET
    if p >= len(module_stream):
        return False
    return module_stream[p] == 0x00


def is_const_type(module_stream: bytes, operand: int, base: int | None) -> bool:
    """True when the declared-type byte carries the constant flag."""
    if base is None:
        return False
    p = base + operand + TYPE_FIELD_OFFSET
    if p >= len(module_stream):
        return False
    return bool(module_stream[p] & VARTYPE_FLAG_CONST)

# Instructions whose name operand is the *callee/target* identifier.
_CALLISH = {"ArgsCall","ArgsMemCall","ArgsMemCallWith","ArgsLd","ArgsMemLd"}

def find_decl_base(module_stream: bytes, table: list[Identifier],
                   *, is_64bit: bool = True) -> int | None:
    """Derive the module's declaration-table base.

    ``func_`` / ``var_`` / ``rec_`` operands are offsets from this base
    to a u16 that is itself a ``name`` operand. No header field holding
    the base has been located, so it is calibrated against two
    independent constraints:

    * the module's type-reference table sits a fixed distance before the
      base and is self-describing (its header repeats its own byte
      length), which rejects almost every wrong candidate outright;
    * every declaration operand must land on a resolvable identifier,
      scored so that procedure names and typed parameters -- the
      strongest signals -- outweigh incidental hits.

    Scoring rather than first-fit matters: a base shifted by one
    parameter stride still resolves *some* operands, and a procedure
    whose name collides with the runtime operand table resolves through
    the secondary space, so a strict all-or-nothing rule picks wrong.
    """
    from pcode_names import NAME_OPERAND_BASE

    dis = disassemble_module_stream(module_stream, is_64bit=is_64bit)
    operands = [(a, v) for line in dis.lines for i in line.instructions
                for a, v in i.operands if a in ("func_", "var_", "rec_")]
    if not operands:
        return None
    decls = [v for a, v in operands if a in ("func_", "rec_")]
    variables = [v for a, v in operands if a == "var_"]

    def name_at(base: int, offset: int) -> int | None:
        p = base + offset
        if p < 0 or p + 2 > len(module_stream):
            return None
        return int.from_bytes(module_stream[p:p + 2], "little") & DECL_NAME_MASK

    best_base, best_score = None, 0
    for base in range(min(len(module_stream), 4096)):
        if not find_type_table(module_stream, base):
            continue
        score = 0
        ok = True
        for v in decls:                     # procedure / Type declarations
            u = name_at(base, v)
            if u is None or not resolve_name(u, table):
                ok = False
                break
            score += 4 if u >= NAME_OPERAND_BASE else 3
        if not ok:
            continue
        for v in variables:
            u = name_at(base, v)
            if u is not None and resolve_name(u, table):
                score += 3
        for v in (v for a, v in operands if a == "func_"):
            for k in range(16):             # typed parameters
                off = v + PROC_PARAM_OFFSET + k * PROC_PARAM_STRIDE
                u = name_at(base, off)
                if u is None or not resolve_name(u, table):
                    break
                if resolve_type(module_stream, off, base) is None:
                    break
                score += 2
        if score > best_score:
            best_base, best_score = base, score
    return best_base


# The record's name field carries a flag in bit 0 (set on object /
# class-typed declarations), so mask it before resolving. Name operands
# are always even, which is what makes the bit free.
DECL_NAME_MASK = ~0x01

def resolve_decl(module_stream: bytes, operand: int, base: int | None,
                 table: list[Identifier]) -> str | None:
    """Resolve a ``func_`` / ``var_`` / ``rec_`` declaration operand."""
    if base is None:
        return None
    p = base + operand
    if p + 2 > len(module_stream):
        return None
    stored = int.from_bytes(module_stream[p:p+2], "little")
    return resolve_name(stored & DECL_NAME_MASK, table)

def annotate(module_stream: bytes, vba_project_stream: bytes,
             *, is_64bit: bool = True) -> str:
    """Disassembly listing with every name operand resolved."""
    ids = parse_identifiers(vba_project_stream)
    dis = disassemble_module_stream(module_stream, is_64bit=is_64bit)
    dbase = find_decl_base(module_stream, ids, is_64bit=is_64bit)
    out=[f"; p-code: {dis.num_lines} lines, {len(ids)} identifiers, decl_base={dbase}"]
    for line in dis.lines:
        if not line.instructions:
            out.append(f"; line {line.line_no:3d}: (no p-code)")
            continue
        parts=[]
        for ins in line.instructions:
            txt=ins.mnemonic
            for a,v in ins.operands:
                if a=="name":
                    nm=resolve_name(v,ids)
                    txt+=f" {nm}" if nm else f" name{v:04X}"
                elif a=="0x":
                    txt+=f" 0x{v:X}"
                elif a in ("func_","var_"):
                    dn=resolve_decl(module_stream,v,dbase,ids)
                    txt+=f" {dn}" if dn else f" {a}{v:08X}"
                else:
                    txt+=f" {a}{v:08X}"
            if ins.payload is not None:
                try:
                    txt+=' "'+ins.payload.decode("latin-1")+'"'
                except Exception:
                    txt+=f" <{ins.payload.hex()}>"
            parts.append(txt)
        out.append(f"; line {line.line_no:3d}: "+" | ".join(parts))
    return "\n".join(out)

def reconstruct(module_stream: bytes, vba_project_stream: bytes,
                *, is_64bit: bool = True) -> str:
    """Best-effort VBA source reconstruction from p-code.

    Covers the statement forms the corpus exercises: procedure
    declarations, Dim, literal assignment, and call statements. Lines
    whose opcodes are not yet mapped are emitted as a commented
    disassembly so nothing is silently dropped.
    """
    ids=parse_identifiers(vba_project_stream)
    dis=disassemble_module_stream(module_stream,is_64bit=is_64bit)
    dbase=find_decl_base(module_stream,ids,is_64bit=is_64bit)
    def declname(i):
        for a,v in i.operands:
            if a in ("func_","var_"):
                return resolve_decl(module_stream,v,dbase,ids)
        return None
    src=[]
    for line in dis.lines:
        ins=list(line.instructions)
        if not ins:
            continue
        mn=[i.mnemonic for i in ins]
        def nm(i):
            for a,v in i.operands:
                if a=="name":
                    return resolve_name(v,ids) or f"name{v:04X}"
            return "?"
        # Sub/Function declaration
        if mn[0] in ("FuncDefn","FuncDefnSave"):
            src.append(f"Sub {declname(ins[0]) or '<proc>'}()")
            continue
        if mn[0] in ("EndSub",):
            src.append("End Sub")
            continue
        if mn[0]=="EndFunction":
            src.append("End Function")
            continue
        # Dim
        if mn[0]=="Dim":
            decls=[]
            for i in ins:
                if not i.mnemonic.startswith("VarDefn"):
                    continue
                vn=declname(i) or "<var>"
                vt=None
                for a,v in i.operands:
                    if a=="var_":
                        vt=resolve_type(module_stream,v,dbase)
                decls.append(f"{vn} As {vt}" if vt else vn)
            src.append("    Dim "+(", ".join(decls) if decls else "<var>"))
            continue
        # literal assignment:  Lit* ... St <name>
        if mn[-1] in ("St","SetStmt") and len(ins)>=2:
            lit=ins[0]
            val=None
            if lit.payload is not None:
                val='"'+lit.payload.decode("latin-1")+'"'
            elif lit.operands:
                val=str(lit.operands[0][1])
            src.append(f"    {nm(ins[-1])} = {val}")
            continue
        # call statement: [args...] ArgsCall <name> <argc>
        if any(m in _CALLISH for m in mn):
            call=next(i for i in ins if i.mnemonic in _CALLISH)
            args=[]
            for i in ins:
                if i is call:
                    break
                if i.payload is not None:
                    args.append('"'+i.payload.decode("latin-1")+'"')
                elif i.mnemonic.startswith("Lit") and i.operands:
                    args.append(str(i.operands[0][1]))
                elif i.mnemonic=="Ld":
                    args.append(nm(i))
            src.append(f"    {nm(call)}"+(" "+", ".join(args) if args else ""))
            continue
        src.append("    ' [unmapped] "+" | ".join(mn))
    return "\n".join(src)


def read_param_mode(module_stream: bytes, operand: int,
                    base: int | None) -> tuple[str, bool, bool]:
    """Return ``(prefix, optional, has_default)`` for one parameter."""
    if base is None:
        return ("", False, False)
    p = base + operand + PARAM_MODE_OFFSET
    if p + 2 > len(module_stream):
        return ("", False, False)
    mode, flags = module_stream[p], module_stream[p + 1]
    if mode & PARAM_BYVAL:
        prefix = "ByVal "
    elif mode & PARAM_BYREF_KEYWORD:
        prefix = "ByRef "
    else:
        prefix = ""
    return (prefix, bool(flags & PARAM_OPTIONAL),
            bool(flags & PARAM_HAS_DEFAULT))


def read_signature(module_stream: bytes, func_operand: int,
                   base: int | None, table, *, is_function: bool,
                   local_offsets: frozenset[int] | None = None,
                   type_table: list[TypeRefEntry] | None = None,
                   defaults: list[str] | None = None):
    """Return ``(name, [(param, type), ...], return_type)``.

    The parameter list is a linked list starting at the procedure
    record's ``PROC_PARAM_HEAD``, so locals sharing the slot region are
    never mistaken for parameters. ``defaults`` carries the literals a
    ``ConstFuncExpr`` pushed before the ``FuncDefn``, consumed in order
    by the parameters whose record says they have one.
    """
    del local_offsets, is_function          # superseded by the record links
    if base is None:
        return (None, [], None)
    resolver = (lambda op: resolve_name(op, table)) if table is not None else None
    name = resolve_decl(module_stream, func_operand, base, table)

    def link(offset: int, field: int) -> int:
        p = base + offset + field
        if p + 4 > len(module_stream):
            return NO_LINK
        return int.from_bytes(module_stream[p:p + 4], "little")

    pending = list(defaults or [])
    params: list[tuple[str, str | None]] = []
    seen: set[int] = set()
    off = link(func_operand, PROC_PARAM_HEAD)
    while off != NO_LINK and off not in seen and len(params) < 64:
        seen.add(off)
        pname = resolve_decl(module_stream, off, base, table)
        ptype = resolve_type(module_stream, off, base, type_table, resolver)
        if not pname:
            break
        prefix, optional, has_default = read_param_mode(module_stream, off, base)
        if optional:
            prefix = "Optional " + prefix
        suffix = ""
        if has_default and pending:
            suffix = f" = {pending.pop(0)}"
        params.append((prefix + pname, (ptype + suffix) if ptype else None))
        off = link(off, PARAM_NEXT_OFFSET)

    ret = resolve_type(module_stream, func_operand + PROC_RETURN_FIELD,
                       base, type_table, resolver)
    return (name, params, ret)
