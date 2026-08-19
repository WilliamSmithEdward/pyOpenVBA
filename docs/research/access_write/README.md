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
| Our p-code equals Microsoft's | 171 statements across 9 modules re-emitted byte-for-byte identically, including a 120-statement module |
| Rewritten logic executes | `acc = 9 * 9`, `idx = acc + 19`, `Calc = acc * 2 + idx` returned **262** |
| Statement count can change | A 3-statement template regrown to 13 statements (nested `Do While` + `If`/`ElseIf`/`Else`) returned **160**; shrunk to 2 returned **42** |
| New identifiers can be added | A 12-statement program using three variables Access never created (`cnt`, `tot`, `best`) returned **80** |
| One module of several can be targeted | Rewriting only `ModB` of a two-module project returned **142**, with `ModA` unchanged at **6** |
| Modules larger than a page can be rewritten | A 122-statement module chained across three LVAL pages, rewritten to 11116 bytes, returned **21660** |
| Any procedure, not just the first | Rewriting `F3` of four to `F3 = n * 10` returned **50** for `F3(5)`, with F1, F2 and F5 unchanged |
| Generated code can carry comments | A commented loop summing odd numbers below 10 returned **25**; comment encoding matches Access on 2503 real comment lines |
| A procedure body can be emptied | Clearing `F2` to a comment made it return **0** while F1, F3 and F5 kept their original behaviour |
| A database can be made from nothing | `AccessReader.create_new()` writes an embedded template, and filling its `Main` with a generated loop returned **56** -- no COM anywhere in that path |

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
- `verify_compiler.py` -- gate: recompiles every module in an
  Access-built project and requires byte-identical output.
- `verify_identity.py` -- gate: rebuilds every module *unchanged* and
  requires the database back byte for byte. Every storage rule here was
  found by this failing; run it before trusting a new one.
- `rewrite_module.py` -- replaces a procedure's body end to end, with a
  free statement count (`--file program.vba` to read the body from a file).

```bash
python docs/research/access_write/verify_compiler.py sample.accdb
```

```bash
python docs/research/access_write/rewrite_module.py in.accdb out.accdb "acc = 9 * 9" "idx = acc + 19" "Calc = acc * 2 + idx"
```

```bash
python docs/research/access_write/rewrite_module.py in.accdb out.accdb --module ModB --file program.vba
```

Access routinely stores several modules on one LVAL page, so a module is
identified by the `Attribute VB_Name` its row decompresses to, never by
its page. Getting that wrong is not a loud failure -- it silently returns
a neighbouring module's p-code, which is what `AccessReader` used to do
(fixed alongside this work).

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

VBA is case-insensitive, so a name may resolve to something that already
exists rather than creating an identifier: in a module `M` with function
`G`, the variables `g` and `m` produce no new records at all -- they bind
to `G` and `M`.

Implicit (undeclared) variables need **no** declaration record -- `zz = 7`
is just `LitDI2 7 | St(name)`. Only `Dim`-ed variables emit
`Dim | VarDefn(var_=...)`. That is why arbitrary logic can be generated
without touching the declaration tables.

A long value is described in `MSysAccessStorage` by an 8-byte descriptor,
`<u32 length | flags><u8 slot><u24 page>`. Flag `0x40000000` means the row
*is* the payload; flags `0` mean it heads a chain whose rows each begin
with `<u8 next_slot><u24 next_page>`, terminated by `(0, 0)`.

Chains fill greedily -- each chunk to capacity, the last running short
(measured 4072 / 4072 / 583) -- and Access keeps a **4-byte gap** between
a page's slot table and its lowest row, so a full chain page reads
`free=4` with the row starting at offset 20. Spreading a payload evenly
instead, or consuming that gap, produces a chain Access refuses to load
even when the bytes round-trip exactly. With both rules applied, rewriting
a chain with its own bytes reproduces the original file byte for byte.

Every procedure owns a pair of u16 line counters at `base + func_`, where
`func_` is its `FuncDefn` operand. The base is 516 for an ordinary
standard module, but it is not universal -- a class module was measured
at 612, with its first procedure's `func_` starting at 56 rather than 0 --
so it is derived per module by finding the one offset at which every
procedure's stored pair already matches what the layout implies. Each
pair holds

```
min(its EndFunc line, line count - 2) - the previous EndFunc line
```

the lines it spans since the procedure before it, with a procedure that
ends the module stopping one line short. Computing these outright rather
than shifting them by the module's line delta is what lets a procedure
other than the first be rewritten. Rebuilding an unchanged module
reproduces its pre-`0xCAFE` header byte for byte across 11 procedures in
10 modules, which is the check that pins the rule down.

A **comment line** is stored as text in the same region as the p-code and
pointed at by a line record of kind `0x09` (code lines are `0x08`):

```
E3 00 <u16 indent> <u16 text length> <text>        padded to even length
```

The leading apostrophe is dropped, the indent is the source line's
leading-space count, and the line record carries neither indent nor
frame-size hint. This encoding reproduces all 2503 comment lines in the
sample databases byte for byte.

The p-code region's size is recorded twice: a **u16 in the 10-byte gap**,
and the **u32 at offset 29** holding the region's end. The u16 overflows --
a 65544-byte region stores as 8 -- so the u32 is the one to trust. That
overflow is why two large modules first looked unmodelled.

The leading `Attribute` block is **not one line**: a standard module
carries only `VB_Name`, a class module carries five or more. Line-table
index *i* corresponds to source line *i* only after skipping that block,
which `Perf.source_lines()` does.

A 12-byte line record is `<flags> <0x80|0x81> <0x08|0x09> <indent>`, then a
u16 p-code length at `+4`, a u16 frame-size hint at `+6`, and a u32 p-code
offset at `+8`. `indent` is the source line's leading-space count. Line
offsets are 8-byte aligned and the p-code region ends with an 8-byte
trailer, so
`total = align8(last_offset + last_length) + 8`.

## The statement-count barrier, and why it was not the slot table

An earlier reading of the pre-`0xCAFE` header concluded that the statement
count was fixed by a table of 24-byte slot records (named locals plus
anonymous compiler temporaries, `4004feff`), because a module with 8 more
statements carried 64 more header bytes. That was a confounded comparison:
the two modules also differed in variables and control flow.

Holding those constant and varying only the statement count settles it.
Five modules with one variable and 1..5 statements have an **identical**
header size and an identical 2-record slot table; only 13 header bytes
differ, of which the build timestamp, the `*\R...` per-compile cookie and
the checksum account for most. Slot records track variables and
control-flow blocks, not statements -- `x = 1` and
`x = 1 + 2 * 3 + 4 * 5 + 6 * 7` both allocate two.

What actually blocked it was two u16 fields at `+516` and `+518` holding
the procedure's line count. They are patched by the line delta, and with
that alone a module regrown from 3 to 13 statements compiles and runs.
The frame-size hint at record `+6` turned out to be **advisory**: the
13-statement module ran correctly with deliberately wrong values.

## Adding identifiers, and the symbol buckets that were not buckets

Adding a name looked blocked by "Access-generated symbol buckets" in the
`_VBA_PROJECT` row that change wholesale between builds. They do change --
but building the *same source twice* changes them too, along with 62 other
bytes, so they are per-build scratch state and carry nothing that has to be
reproduced. An earlier attempt to fit a `hash % N` bucket model to them was
fitting noise: it matched two samples and failed on every independent one.

What Access does validate is a pair of u16 counters sitting immediately
*before* the identifier table -- the record count at `table_start - 10` and
the slot count at `table_start - 12`. Appending a record without bumping
both makes Access hang on load. With them corrected, a name appended
before the `02 FF FF 01 01` sentinel is picked up normally, and its p-code
operand follows the usual `524 + 2*index`.

So adding an identifier is: append
`<u8 len><u8 type=0><name><u16 hash><10 00>`, where the hash is the OLE
`LHashValOfNameSysA` value already solved in
`docs/research/pcode/pcode_hash.py` (verified against Access: `zz` ->
`0x6031`, `_B_var_zz` -> `0xf4d9`); bump both counters; then update the
row's `MSysAccessStorage` length like any other resize. `rewrite_module.py`
does this automatically for every name a program introduces.

## What still does not work

**Modules the layout model does not cover.** `rewrite_module.py` rebuilds
the module unchanged first and refuses to write unless the result is
byte-identical, so an unmodelled layout is a refusal rather than a
corrupted database. One known case: a large comment-heavy module stores
its comment text in a plaintext region *after* the p-code, with those line
records (byte 2 = `0x09` rather than `0x08`) pointing into it instead of
into the p-code. The p-code region then does not end where the model
expects, and the 1 MB fixture's `Module1` and `Class1` are refused for
exactly that reason.

**Creating a procedure from nothing.** Adding one grows the pre-`0xCAFE`
header by 360 bytes spread over several structures at once: an 88-byte
declaration record at the new procedure's `func_` offset (the second
procedure's is 88, matching the stride seen in a four-procedure module),
two more 24-byte slot records for its locals, and roughly 220 further
bytes of counts and pointers that shift with them. Reading, rewriting,
and emptying the procedures a module already has all work; synthesising a
new one needs that structure modelled.

Filling in a procedure a template already declares -- including an empty
one -- covers much of the same ground today, since a VBA project has to
exist before any of this applies. `AccessReader.create_new()` supplies
that starting point from bytes embedded in the library, the same way
`ExcelFile.create_new()` does, so no Access install is needed to obtain
one. `bake_access_template.py` regenerates the embedded copy; an
`.accdb` is mostly empty pages and compresses to about 4% of its size.

**Allocating pages.** A module already stored as a chain can be rewritten
up to that chain's capacity (12216 bytes for a three-page chain), and a
single-row module up to its page's free space -- around 23 statements for
a small template. Growing past either needs a *new* page, which means
reproducing Access's page allocator and its usage maps. Both limits raise
`ValueError` rather than corrupting.

Shortening a chain below the rows it occupies is refused, and that
refusal is deliberate rather than unimplemented. Leaving the surplus
chunks carrying only their 4-byte link makes Access reject the project.
Releasing them -- terminating the chain early and tombstoning the freed
slots -- works for some modules and not others: shrinking a two-row chain
in one standard module loaded and ran correctly, while the same operation
on a class module, and on a four-row chain, did not. Something further
tracks those rows, most likely the page usage maps. A write path that
sometimes produces a database Access will not open is worse than one that
declines, so it declines.

**A procedure's body is bounded by its `FuncDefn` and `EndFunc` lines**,
not by "the statements we can recompile" -- that way a procedure holding
no executable statements, empty or entirely comments, still has a findable
body.

**Do not assume where the catalog lives.** `MSysAccessStorage` sits on
page 48 in a freshly created database, but the 1 MB fixture in this repo
keeps its long-value descriptors on page 168. Descriptors are found by
searching for their exact 8 bytes instead.

**Beware: running a database mutates it.** Opening an `.accdb` read-write
in Access rewrites parts of the file -- in one case relocating the
`_VBA_PROJECT` row so it could no longer be found. Keep templates
pristine and always work on copies.

A read-path gap turned up in passing and is now fixed: the explicitly
slotted records described above were absorbed into the next entry's
`prefix` instead of being parsed, so `b` and `f` never appeared in
`AccessReader.identifiers()`. The positional numbering was always
correct -- excluding those records is what keeps `524 + 2*index` valid --
so the fix surfaces them without renumbering anything. Every p-code
`Ld`/`St` operand across the sample databases now resolves to a name.
