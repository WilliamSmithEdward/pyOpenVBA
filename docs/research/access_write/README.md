# Writing executable VBA into Access, from pure Python

Research code behind the 2026-08 Access write breakthrough. Everything
here is dev-only and lives outside `src/`; `AccessReader` remains
read-only.

The claim this directory supports: **an Access `.accdb` module's logic can
be recompiled and rewritten entirely in pure Python, and real Access will
compile and execute the result.** Verified by running the macro in desktop
Access and reading the value it returns -- not by reading the file back,
which proves nothing (see `docs/msaccess_lessons_learned.md`).

## Why the earlier attempts failed

Phase 5 (2026-05) could rewrite the OVBA source cache but every non-trivial
edit was reverted or rejected. Two beliefs were wrong.

**"Same-length edits only."** Growing an LVAL row needs more than page
surgery. Each VBA long-value is described by a `MSysAccessStorage` catalog
row holding `<u16 length> 00 40 <slot><page>`, and Access trusts that
length. A resized row whose catalog length still reads the old value makes
Access fault while loading the project. Section 5 of the lessons document
listed this binding as "never located" -- it is that length field, and
updating it is what makes resizing work.

**"A p-code assembler is a multi-month effort."** The instruction encoding
was already solved. What made code generation tractable is that
**VBA p-code control flow is structured, not jump-based**: `IfBlock` /
`ElseBlock` / `EndIfBlock`, `DoWhile` / `Loop`, and `For` / `NextVar` are
plain markers carrying no branch offsets, so there are no jump targets to
fix up -- normally the hardest part of a code generator.

## What is proven to execute

| Result | Evidence |
|---|---|
| Access runs the canonical `0xCAFE` p-code, not the `rU@` execodes | Editing only the CAFE region changed `5+3` to `5+9`; Access returned 14 while every `rU@` row stayed stale |
| Source alone is display-only | Source edited to `5 + 3 + 100` with p-code untouched still returned **8** |
| A grown module row loads and runs | `5+3` -> `5+3+10` (+12 bytes) returned 18, `compile_project` OK |
| Our p-code equals Microsoft's | 33 statements across 4 modules re-emitted byte-for-byte identically |
| Rewritten logic executes | `acc = 9 * 9`, `idx = acc + 19`, `Calc = acc * 2 + idx` returned **262** |

The source/p-code test is the important negative: Access has **no
load-time recompile-from-source trigger**. Excel treats a version mismatch
as "discard p-code and rebuild from source", which is what pyOpenVBA's
Excel writer relies on. Access does not, so writing source text alone
produces a module that displays one thing and does another.

## Files

- `accdb_write.py` -- storage layer: LVAL page reflow, the
  `MSysAccessStorage` length field, a literal-only MS-OVBA compressor, and
  `Perf`, which parses and rebuilds a module's `0xCAFE` region.
- `vba_compile.py` -- VBA to p-code compiler: full operator precedence,
  assignment, `If`/`ElseIf`/`Else`, `Do While`/`Loop`, `For`/`Next`.
- `verify_compiler.py` -- gate: recompiles an Access-built module and
  requires byte-identical output. Non-zero exit on any difference.
- `rewrite_module.py` -- rewrites a module's statements end to end.

```bash
python docs/research/access_write/verify_compiler.py sample.accdb
```

```bash
python docs/research/access_write/rewrite_module.py in.accdb out.accdb "acc = 9 * 9" "idx = acc + 19" "Calc = acc * 2 + idx"
```

## Format notes

Identifier operands index the project identifier table as
`operand = 524 + 2*index`.

Most records are positional, but a few names bind to a pre-existing
low-numbered slot and use a variant record that carries it explicitly:

```
00 00 <u16 slot> <u8 len> 80 <6B descriptor> <name>
```

It has no trailing id / `10 00` pair, and its operand is `2*slot + 2`
rather than `524 + 2*index`. Such a record takes no position, so counting
it would misname every identifier after it. Observed for `b` (slot 11,
operand 24) and `f` (slot 81, operand 164); the other 24 single letters
are ordinary positional records. `AccessReader.identifiers()` exposes
these with `slot` set and `index == -1`.

Beware that VBA is case-insensitive when checking a name resolved: in a
module `M` with function `G`, the variables `g` and `m` do not create
identifiers at all -- they bind to the existing `G` and `M`.

Implicit (undeclared) variables need **no** declaration record -- `zz = 7`
is just `LitDI2 7 | St(name)`. Only `Dim`-ed variables emit
`Dim | VarDefn(var_=...)`. That is why arbitrary logic can be generated
without touching the declaration tables.

A 12-byte line record is `<flags> <0x80|0x81> <0x08|0x09> <indent>`, then a
u16 p-code length at `+4`, a u16 frame-size hint at `+6`, and a u32 p-code
offset at `+8`. `indent` is the source line's leading-space count. Line
offsets are 8-byte aligned and the p-code region ends with an 8-byte
trailer, so
`total = align8(last_offset + last_length) + 8`.

## What still does not work

**Changing the statement count.** The pre-`0xCAFE` header carries a table
of 24-byte slot records -- named locals plus anonymous compiler
temporaries (`4004feff`) -- sized by Access when it compiles. A module with
8 more statements had 64 more header bytes, entirely in that table, plus
several count and pointer fields (a statement count at two offsets, and
u32 pointers) that move with it. Generating those requires modelling VBA's
temporary allocation and the frame-size hint at record `+6`, which depends
on static type inference (an implicit Variant assignment scores 20 where
the same statement on a `Long` scores 18). Until that is reproduced,
statement counts must be preserved.

**New identifiers.** Adding a name means extending the `_VBA_PROJECT`
identifier table. The record format and its hash are solved
(`docs/research/pcode/pcode_hash.py`, `LHashValOfNameSysA`), but the row
also holds Access-generated symbol buckets that change wholesale between
builds and are not yet understood. So generated code can only use names
already present in the project.

**New procedures.** `FuncDefn` carries a `func_` offset into the
declaration tables, which has the same unsolved shape as the slot table.

A read-path gap turned up in passing and is now fixed: the explicitly
slotted records described above were absorbed into the next entry's
`prefix` instead of being parsed, so `b` and `f` never appeared in
`AccessReader.identifiers()`. The positional numbering was always
correct -- excluding those records is what keeps `524 + 2*index` valid --
so the fix surfaces them without renumbering anything. Every p-code
`Ld`/`St` operand across the sample databases now resolves to a name.
