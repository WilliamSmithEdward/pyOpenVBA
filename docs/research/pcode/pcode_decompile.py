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
    if not ops:
        return None
    for base in range(0, min(len(module_stream), 4096)):
        good = True
        for v in ops:
            p = base + v
            if p + 2 > len(module_stream):
                good = False; break
            if not resolve_name(int.from_bytes(module_stream[p:p+2], "little"), table):
                good = False; break
        if good:
            return base
    return None

def resolve_decl(module_stream: bytes, operand: int, base: int | None,
                 table: list[Identifier]) -> str | None:
    """Resolve a ``func_``/``var_`` declaration operand to its name."""
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
            names=[declname(i) or "<var>" for i in ins if i.mnemonic.startswith("VarDefn")]
            src.append("    Dim "+(", ".join(names) if names else "<var>")); continue
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
