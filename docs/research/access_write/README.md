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
| Access *sometimes* runs the canonical `0xCAFE` p-code | Editing only the CAFE region changed `5+3` to `5+9`; Access returned 14 while every `rU@` row stayed stale. **This does not generalise** -- see "Correction" below, where the execodes shadow the p-code instead |
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
| Comment-heavy multi-procedure modules rewrite | With 2, 10 and 60 comments per procedure, rewriting one returned **42** while its neighbour kept returning **2** |
| Generated code can call procedures | `a = Twice(7)` / `b2 = Twice(a)` / `Main = a + b2` returned **42**; `MsgBox`, `DoCmd.Beep` and a user call all re-emit byte-identically |

The source/p-code test is the important negative: Access has **no
load-time recompile-from-source trigger**. Excel treats a version mismatch
as "discard p-code and rebuild from source", which is what pyOpenVBA's
Excel writer relies on. Access does not, so writing source text alone
produces a module that displays one thing and does another.

## Files

- `accdb_write.py` -- storage layer: LVAL page reflow, the
  `MSysAccessStorage` length field, a literal-only MS-OVBA compressor, and
  `Perf`, which parses and rebuilds a module's `0xCAFE` region.
- `vba_compile.py` -- VBA to p-code compiler covering a subset:
  expressions with full operator precedence, assignment,
  `If`/`ElseIf`/`Else`, `Do While`/`Loop`, `For`/`Next`, `Exit Do`/`For`,
  comments, and calls -- `Foo(a, b)` in an expression, `Foo a, b` as a
  statement, and `obj.Member` in either. Anything outside that --
  `Select Case`, `Dim`, `Set`, `With`, arrays, `On Error` -- is refused
  rather than mis-emitted.
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

## How p-code reaches outside the module

Four mechanisms, all visible in one module:

| Source | p-code |
|---|---|
| `Len(s)`, `Abs(-3)` | `FnLen`, `FnAbs` -- a dedicated opcode per intrinsic, no operand |
| `Left(s, 2)` | args, then `ArgsLd(name=220, argc)` -- a **built-in** name at a low slot |
| `Helper(7)`, `MsgBox "hi"` | args, then `ArgsLd` / `ArgsCall` -- names living in the **project** table |
| `DoCmd.Beep` | `Ld(DoCmd)`, args, then `ArgsMemCall(Beep, argc)` |
| `Debug.Print s` | fully special-cased: `Debug`, `PrintObj`, value, `PrintItemNL` |

The operand space is one flat table of slots addressed as `2*slot + 2`.
Slots 0..260 are **pre-populated built-ins** -- `Left` is slot 109, and
the `b` and `f` oddity above is simply slots 11 and 81. Project
identifiers start at slot 261, which is why they read as
`524 + 2*index`. So an external reference is either an intrinsic opcode,
a built-in slot, or an ordinary project identifier record; a library name
like `MsgBox` or `DoCmd` is just appended to the project table like any
other name, which is why calling one needs no new machinery.

A statement call -- one whose result is discarded -- carries **op_type 16**
in the high bits of the instruction word: `MsgBox "hi"` emits `0x4041`
where an expression call emits `0x0041`.

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

Compression matters more than it looks. An early round of this work
concluded Access rejected the repository's `compress()` and fell back to
literal-only chunks; that test was confounded by a stale catalog length,
and with the length correct Access accepts it. On one module the real
compressor produced 148 bytes where literal-only produced 366 -- matching
Access's own output exactly -- and on another 784 against 4913. The
difference decides whether a rewritten module still fits its page.

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

Shortening a chain works by releasing the surplus rows: the chain
terminates early and each freed slot is tombstoned. Leaving those chunks
in place carrying only their 4-byte link does *not* work -- Access
rejects the project -- so they have to go. Verified on a standard module,
a class module, and a three-row chain shrinking to two.

Access itself does something different when a chained module shrinks: it
abandons the chain entirely and rewrites the value into freshly allocated
storage as a single row. Reusing the rows already held avoids needing an
allocator at all.

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

## Correction: Access runs execodes, not the p-code we write

An earlier note in this file said Access executes the canonical `0xCAFE`
p-code, and this session repeated that a rewritten procedure "returns 42
from real Access". **Neither generalises.** On every database tried here,
a rewritten procedure kept running its *old* code. Measured 2026-08:

| Level | Content | Role |
|-------|---------|------|
| `rU@` rows | a second compiled form, in its own `MSysAccessStorage` rows | **what Access actually executes** |
| OVBA source | the compressed text at `MODULEOFFSET` | what a decompile rebuilds from |
| `0xCAFE` p-code | the canonical stream this compiler targets | what the VBE and our tools read |

Two probes are needed, because one alone cannot separate the levels.

**Execodes beat the canonical p-code.** Rewrite a procedure so source and
p-code both say 777 while the untouched execodes still encode 5. The
file provably contains no other copy of the old code -- a byte search
finds `LitDI2(777) St(Probe)` once and `LitDI2(5)` nowhere -- and Access
still returns **5**. The VBE meanwhile displays `Probe = 777`.

**Source beats the canonical p-code, once the execodes are gone.** Edit
*only* the source text, leaving every p-code byte as Access compiled it,
so the file says `Probe = 123` and executes `LitDI2(5) St(Probe)`:

```
before /decompile   Eval("Probe()") -> 5      (execodes, matching the p-code)
after  /decompile   Eval("Probe()") -> 123    (rebuilt from source)
```

So a decompile rebuilds from **source**, not from the canonical p-code.
Rewriting the p-code changes what the VBE displays and what every reader
sees, and so far has never changed what runs.

What this does *not* establish is a universal rule. The earlier `5+3` ->
`5+9` result above was a real observation, and under the model here it
should not have happened -- so Access must sometimes treat the execodes
as invalid and fall back. The condition that decides this is not yet
identified, and finding it is the open question.

`DoCmd.RunCommand(126)` ("Compile And Save All Modules") does not help --
Access reports success while continuing to return the stale value, so it
evidently considers the project already compiled.
The only lever found so far is launching `MSACCESS.EXE <db> /decompile`,
which discards the execodes; Access then rebuilds from source and the
written code runs. 156 of the 163 databases in this repo carry execode
rows, so this is the normal case, not an edge case.

Two consequences worth being blunt about:

- **For execution, the source text is what matters**, plus something that
  invalidates the execodes. The p-code compiler is not on that path.
- **The p-code compiler is still worth having**, but for different
  reasons than assumed: byte-exact agreement with Microsoft's compiler is
  a strong correctness oracle, it keeps a rewritten module internally
  consistent, and reading it is how the disassembler and decompiler work.

What is still unknown: whether the execode rows can be invalidated from
pure Python, which is what would make an Access-free write path execute.

## Establishing p-code coverage

"Does the compiler handle VBA?" is not answerable by inspection, and the
opcode table does not answer it either -- that table is closed by
construction (264 entries, 0..263, the last literally named `Illegal`),
so a disassembler cannot meet an unknown opcode. Generation is the open
problem.

The instrument is `construct_matrix.bas`: one module exercising the
constructs on purpose, compiled by Access itself via `build_matrix.ps1`,
then diffed statement by statement with `verify_compiler.py`. Every
statement either matches Microsoft byte for byte or is reported.

It earned its keep immediately -- six defects on first run, three of them
**silent miscompiles**, which are far worse than refusals because they
emit valid p-code for the wrong program:

| Source | Was emitted | Should be |
|--------|-------------|-----------|
| `r = a ^ b` | `Ld(a) St(r)` -- operand dropped | `Ld(a) Ld(b) Pwr St(r)` |
| `arr(1) = 10` | `ArgsCall(arr)` -- an assignment compiled as a call | `LitDI2 LitDI2 ArgsSt(arr)[1]` |
| `d.Add "k", 1` | object pushed first | arguments first, object last |
| `x = d.Count` | `ArgsMemLd(Count)[0]` | `MemLd(Count)` |
| `For i = 10 To 1 Step -1` | `Step` silently ignored | step expression then `ForStep` |
| `v = 3.5` | crash | `LitR8` with the raw double |

The member-call order bug had survived because every earlier probe called
a **zero-argument** member (`DoCmd.Beep`), and with no arguments the two
orders are byte-identical. A probe that cannot distinguish two hypotheses
is not evidence for either.

The grammar now refuses any statement it cannot fully consume, so an
unparsed tail raises instead of quietly shrinking the program. After the
fixes: **75 statements byte-identical to Access, 0 differing**, and the
matrix module rebuilds byte-for-byte through `verify_identity.py`.

Uncovered and refused, not approximated: `With`, `Dim`/`Const`/`ReDim`/
`Erase`, `On Error`, line labels, single-line `If ... Then <statement>`,
date literals, and built-ins living in the pre-populated slots below 261.

## References added through the References menu

They need no new p-code machinery, because **early and late binding
compile identically**. `dEarly.Add "k", 1` on a `Scripting.Dictionary`
and `dLate.Add "k", 1` on an `Object` differ by exactly one byte -- the
operand naming the variable:

```
b90001006b00ac00010020003202424040020200   Dim dEarly As Scripting.Dictionary
b90001006b00ac00010020003602424040020200   Dim dLate  As Object
                          ^^^^
```

No DISPID, no vtable slot, no type-library token in the instruction
stream: the member is an identifier slot and binding happens at run time.
A referenced library's names -- `Scripting`, `Dictionary`, `Add`,
`CompareMode`, the enum constant `TextCompare`, and `kernel32` from a
`Declare` -- are ordinary entries in the project identifier table,
indistinguishable from names the user wrote.

The library shows up in exactly three other places:

1. A `REFERENCED` record in the dir stream, which `AccessReader` already
   parses: `*\G{420B2830-...}#1.0#0#C:\Windows\System32\scrrun.dll#Microsoft Scripting Runtime`.
2. The declared type on a `Dim`, through the `type_` descriptor's typeref.
3. `New Scripting.Dictionary`, which emits `New` with an **import index**
   (`imp_=0`, `imp_=8` -- an 8-byte stride) rather than a name operand.
   This is the only construct that reaches the type library from the
   instruction stream.

A `Declare PtrSafe Function ... Lib "kernel32"` is not special either: it
is a `FuncDefn` like any other, and its call site is an ordinary
`ArgsLd`, indistinguishable from a call to a local `Sub`.
