"""P-code decompiler: compiled VBA7 bytecode -> readable listing/source.

Host-agnostic: consumes the module-stream p-code region (CAFE) plus the
project identifier table, both of which are identical in layout across
Excel, Word, PowerPoint, and Access.
"""
from __future__ import annotations
import sys
sys.path.insert(0,"F:/GitHub/pyOpenVBA/src")
from pyopenvba.vba_pcode import disassemble_module_stream
from pcode_names import parse_identifiers, resolve_name, Identifier

# Procedure layout, relative to DECL_BASE + func_operand:
#   +0x00              the procedure's own name entry
#   +0x58 + k*0x20     parameter k (name entry; VARTYPE at +14 as usual)
# The return-type entry repeats the procedure name with a type set; its
# offset is not fixed, so it is located by scanning.
PROC_PARAM_OFFSET = 0x58
PROC_PARAM_STRIDE = 0x20

# Declared-type byte lives at DECL_BASE + var_operand + TYPE_FIELD_OFFSET
# and holds a standard OLE Automation VARTYPE code.
TYPE_FIELD_OFFSET = 14
VARTYPE_NAMES: dict[int, str] = {
    0: "Empty", 1: "Null", 2: "Integer", 3: "Long", 4: "Single",
    5: "Double", 6: "Currency", 7: "Date", 8: "String", 9: "Object",
    10: "Error", 11: "Boolean", 12: "Variant", 13: "Unknown",
    14: "Decimal", 17: "Byte", 20: "LongLong",
}

# Flag bits carried alongside the VARTYPE in the declared-type byte.
VARTYPE_MASK = 0x3F
VARTYPE_FLAG_CONST = 0x40

def resolve_type(module_stream: bytes, operand: int, base: int | None) -> str | None:
    """Resolve a ``var_`` operand to its declared VBA type name.

    The byte carries flags above the VARTYPE: ``0x40`` marks a constant
    (``0x43`` is Const + Long, ``0x48`` Const + String). Arrays encode
    differently -- both ``As Long`` and ``As String`` arrays leave
    ``0x10`` in the low bits -- so their element type is not read here
    (it lives in the ``type_`` indirect table).
    """
    if base is None:
        return None
    p = base + operand + TYPE_FIELD_OFFSET
    if p >= len(module_stream):
        return None
    return VARTYPE_NAMES.get(module_stream[p] & VARTYPE_MASK)


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

    ``func_`` / ``var_`` operands are offsets from this base to a u16
    that is itself a ``name`` operand. The base is not (yet) known to
    live in a header field, so it is calibrated: the unique offset at
    which every such operand resolves to a real identifier.
    """
    dis = disassemble_module_stream(module_stream, is_64bit=is_64bit)
    ops = [v for l in dis.lines for i in l.instructions for a, v in i.operands
           if a in ("func_", "var_")]
    # A FuncDefn's operand points directly at its own name entry, which
    # is a far stronger constraint than "some offset resolves": prefer a
    # base where every procedure name lands. Calibrating on var_ alone
    # can settle on a shifted base that still resolves (off by 0x2C in
    # observed single-procedure modules).
    fops = [v for l in dis.lines for i in l.instructions for a, v in i.operands
            if a == "func_"]
    if fops:
        # Score candidate bases instead of taking the first that fits.
        # A procedure name may legitimately resolve through the built-in
        # space (a Sub called B collides with built-in 'b'), so a strict
        # project-table-only rule rejects the true base; conversely a
        # shifted base can satisfy the names alone by landing on the
        # parameter slots. Counting typed parameters as well
        # disambiguates: only the true base makes both line up.
        from pcode_names import NAME_OPERAND_BASE
        best_base, best_score = None, -1
        for base in range(0, min(len(module_stream), 4096)):
            score = 0
            ok = True
            for v in fops:
                q = base + v
                if q + 2 > len(module_stream):
                    ok = False; break
                u = int.from_bytes(module_stream[q:q+2], "little")
                if not resolve_name(u, table):
                    ok = False; break
                score += 4 if u >= NAME_OPERAND_BASE else 3
            if not ok:
                continue
            for v in ops:                       # every declared variable
                q = base + v
                if q + 2 <= len(module_stream) and resolve_name(
                        int.from_bytes(module_stream[q:q+2], "little"), table):
                    score += 3
            for v in fops:                      # typed parameters
                for k in range(16):
                    off = v + PROC_PARAM_OFFSET + k * PROC_PARAM_STRIDE
                    q = base + off
                    if q + 2 > len(module_stream):
                        break
                    if not resolve_name(
                            int.from_bytes(module_stream[q:q+2], "little"), table):
                        break
                    if resolve_type(module_stream, off, base) is None:
                        break
                    score += 2
            if score > best_score:
                best_base, best_score = base, score
        if best_base is not None:
            return best_base
    if not ops:
        return None
    # Maximise resolutions rather than demanding all of them: an operand
    # may reference the secondary identifier region (see pcode_reference
    # section 5.1), which would otherwise veto an otherwise-correct base.
    best_base, best_hits = None, 0
    for base in range(0, min(len(module_stream), 4096)):
        hits = 0
        for v in ops:
            p = base + v
            if p + 2 > len(module_stream):
                continue
            if resolve_name(int.from_bytes(module_stream[p:p+2], "little"), table):
                hits += 1
        if hits > best_hits:
            best_base, best_hits = base, hits
            if hits == len(ops):
                break
    return best_base

def resolve_decl(module_stream: bytes, operand: int, base: int | None,
                 table: list[Identifier]) -> str | None:
    """Resolve a ``func_`` / ``var_`` / ``rec_`` declaration operand."""
    if base is None:
        return None
    p = base + operand
    if p + 2 > len(module_stream):
        return None
    return resolve_name(int.from_bytes(module_stream[p:p+2], "little"), table)

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
                try: txt+=' "'+ins.payload.decode("latin-1")+'"'
                except Exception: txt+=f" <{ins.payload.hex()}>"
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
            src.append(f"Sub {declname(ins[0]) or '<proc>'}()"); continue
        if mn[0] in ("EndSub",):
            src.append("End Sub"); continue
        if mn[0]=="EndFunction": src.append("End Function"); continue
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
            src.append("    Dim "+(", ".join(decls) if decls else "<var>")); continue
        # literal assignment:  Lit* ... St <name>
        if mn[-1] in ("St","SetStmt") and len(ins)>=2:
            lit=ins[0]
            val=None
            if lit.payload is not None:
                val='"'+lit.payload.decode("latin-1")+'"'
            elif lit.operands:
                val=str(lit.operands[0][1])
            src.append(f"    {nm(ins[-1])} = {val}"); continue
        # call statement: [args...] ArgsCall <name> <argc>
        if any(m in _CALLISH for m in mn):
            call=next(i for i in ins if i.mnemonic in _CALLISH)
            args=[]
            for i in ins:
                if i is call: break
                if i.payload is not None: args.append('"'+i.payload.decode("latin-1")+'"')
                elif i.mnemonic.startswith("Lit") and i.operands: args.append(str(i.operands[0][1]))
                elif i.mnemonic=="Ld": args.append(nm(i))
            src.append(f"    {nm(call)}"+(" "+", ".join(args) if args else "")); continue
        src.append("    ' [unmapped] "+" | ".join(mn))
    return "\n".join(src)


def read_signature(module_stream: bytes, func_operand: int,
                   base: int | None, table, *, is_function: bool,
                   local_offsets: frozenset[int] | None = None):
    """Return ``(name, [(param, type), ...], return_type)`` for a procedure.

    The name sits at ``base + func_operand``; parameters follow at
    ``+0x58`` in ``0x20`` strides. The return type is a second entry
    repeating the procedure name with a VARTYPE set, located by scan.

    Parameters and a procedure's *locals* share that slot region, so a
    naive scan reads locals as extra parameters. ``local_offsets`` (the
    ``var_`` operands of ``VarDefn``, i.e. everything the body declares
    with ``Dim``) marks the slots to stop at.
    """
    locals_ = local_offsets or frozenset()
    if base is None:
        return (None, [], None)
    name = resolve_decl(module_stream, func_operand, base, table)
    params: list[tuple[str, str | None]] = []
    for k in range(64):
        off = func_operand + PROC_PARAM_OFFSET + k * PROC_PARAM_STRIDE
        p = base + off
        if p + 2 > len(module_stream):
            break
        if off in locals_:
            break                       # a Dim-declared local, not a parameter
        pname = resolve_decl(module_stream, off, base, table)
        ptype = resolve_type(module_stream, off, base)
        # Every parameter carries a declared type (Variant when implicit);
        # a typeless entry means the parameter list has ended.
        if not pname or pname == name or ptype is None:
            break
        prefix = "ByVal " if is_byval(module_stream, off, base) else ""
        params.append((prefix + pname, ptype))
    ret = None
    if is_function and name:
        for off in range(func_operand, min(func_operand + 0x400, len(module_stream) - base), 2):
            if resolve_decl(module_stream, off, base, table) == name:
                t = resolve_type(module_stream, off, base)
                if t:
                    ret = t
                    break
    return (name, params, ret)
