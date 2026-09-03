# Writing executable VBA into Access, from pure Python

> **Status: research, reopened 2026-09-03. Not a supported feature and
> not on any roadmap.** Nothing here is imported by `src/` and nothing
> here ships. This directory is a research record: the measurements are
> reproducible and the dead ends are written down so nobody has to
> rediscover them.
>
> Reopened because the storage engine in `pyopenvba.access` arrived after
> it was parked. Module **rename and delete now work**, verified by
> running the result in Access. **Create** writes a module Access loads,
> the VBE enumerates and reads, and the project will take further
> components; editing an existing module's code in such a project is what
> still fails.

Everything here is dev-only, needs Windows with desktop Access, and was
verified by running the macro in Access and reading the value it returns
-- never by reading the file back, which proves nothing (see
`docs/msaccess_lessons_learned.md`).

## What was established

An `.accdb` module's **procedure bodies** can be recompiled and rewritten
in pure Python, and real Access executes the result:

* A VBA-to-p-code compiler whose output is **byte-identical to
  Microsoft's** across the probe corpus, covering expressions, all the
  control-flow forms, calls, literals, `With`, `On Error`, and `Dim`.
* Arbitrary statement counts, including growing a body from empty.
* `Dim` declarations added, removed and retyped, byte-exact against
  Access in both directions across 18 transitions.
* The reason writes appeared to do nothing: Access executes an `__SRP_*`
  compiled cache, and dropping it is what makes a rewrite take effect.

Module **rename** and **delete** work end to end, in pure Python, with no
COM in the write path -- the eight places a name lives, and every
structure a module occupies. See "Module rename, delete and create".

## What does not work, and is not close

* **Creating a procedure** in a module, which is what stops a created
  module from being useful and stops `Dim`'s neighbours from working.
  Module rename and delete now work; create writes a module Access opens
  and lists but will not extend. See "Module rename, delete and create".
* **`Const`, arrays, `Static`, fixed-length strings.** Each reshapes the
  module header its own way; all measured, none implemented, all refused.
* **`Set x = New <Class>`**, which needs the import table.
* **Page allocation** for the module rewriter, so a module can only grow
  into space it already has. The storage engine allocates pages properly,
  so this is a matter of routing the rewriter through it rather than an
  open question.

## If this is picked up again

Read "Auditing the guards" first. Six guards in this code were wrong in
the same way -- narrow tests that passed by not looking -- and three of
the fixes replaced a table this code maintained with a value read from
the file being edited. That is the pattern worth carrying forward.

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
| Access runs the canonical `0xCAFE` p-code, once its `__SRP_*` cache is dropped | Editing only the CAFE region changed `5+3` to `5+9` and Access returned 14. With the cache present it returns the stale value instead -- see the `__SRP_` section below |
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

Module CRUD, all of it built on the storage engine rather than on byte
patching:

- `module_rename.py` -- rename, all eight places, `__SRP_` drop included.
- `module_delete.py` -- delete, every structure a module occupies.
- `module_create.py` -- create, cloning an existing module's compiled
  shape (`donor=` to take it from another database, `skip=` to bisect the
  pieces). Access loads the result, the VBE enumerates and extends it;
  editing an existing module's code in such a project still fails.
- `module_stream.py`, `project_streams.py`, `vba_project_table.py`,
  `vba_module_table.py`, `dir_records.py` -- one structure each.
- `attribute_pages.py` -- says which table owns each page two databases
  differ on and diffs its rows. This is the instrument that made the
  rename map readable; a raw page diff of the same edit is sixty pages of
  recompiled project.
- `drive_access.py` -- drives Access through `pyvbaharness`, so a VBA
  error comes back as data rather than as a modal dialog nobody can see.
  Access refuses several `DoCmd` verbs over a bare COM boundary that it
  accepts from inside its own VBA, which is why the operations are driven
  this way rather than with `module_ops.ps1`.
- `peek_project.ps1` -- opens a database read-only and reports what its
  VBA project looks like, for checking a write without letting the
  harness inject anything.
- `exercise_module.ps1` -- puts a procedure into a named module and runs
  it, reporting each step, so a created module can be exercised without
  the harness injecting one of its own. It can hang on a modal dialog;
  give it a timeout.

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

## The __SRP_ cache is what Access executes, and dropping it is the lever

A rewritten procedure kept running its **old** code, on every database
tried. The cause, and the fix, both turned out to be something the
library already handles for Office.

Access keeps a second compiled form of every module in storage rows named
`__SRP_0`, `__SRP_1`, ... -- the same performance cache [MS-OVBA]
describes as `__SRP_*` streams, and that `_host.py` already drops when
writing Office files:

```python
cfb.drop_streams_in_storage("VBA", lambda n: n.startswith("__SRP_"))
```

Access *executes* that cache. Rewrite a procedure so source and p-code
both say 777 while the untouched cache still holds 5, and Access returns
**5** while the VBE displays `Probe = 777`. A byte search finds the new
p-code once and the old value nowhere. `DoCmd.RunCommand(126)` reports
success and changes nothing.

Drop the `__SRP_*` rows and Access returns **777**. 156 of the 163
databases here carry them, so this is the normal case.

The rows are catalogued in `MSysAccessStorage`, and **the catalog row is
what must go**. Deleting only the long-value rows leaves the catalog
pointing at nothing, and Access rejects the entire project with "can't
find the function" -- which is what made the first attempts look like the
cache was load-bearing rather than stale. `drop_srp_cache()` marks the
slot-table entry deleted, the way Access retires a row.

`MSACCESS.EXE <db> /decompile` does the same job from outside, and is
what confirmed the diagnosis before the pure-Python route was found: it
strips the `__SRP_` rows and 383 bytes of compiled state from the module
header, leaving the p-code and source intact.

**So the write path executes, from pure Python, with no COM.** A rewrite
that keeps the module's line count is verified end to end: a program of
twelve generated statements is refused for the reason below, but a
single-statement rewrite returns its new value from real Access.

### And the p-code is what executes, so the compiler is not optional

With the cache gone, does Access run our p-code or recompile from our
source? Make the two disagree and ask. Both directions agree:

| source | p-code | `__SRP_` | Access returns |
|--------|--------|----------|----------------|
| `Probe = 123` | `LitDI2(5)` | dropped | **5** |
| `Probe = 5` | `LitDI2(777)` | dropped | **777** |

The p-code wins every time. Writing source alone would change nothing,
so the compiler is load-bearing rather than a nicety, and byte-exact
agreement with Microsoft's own output is the property that matters.

Source is for display and for structural validation. That is also what
the line-count bug below is: Access parses the source against the
header's line counters, and errors when they disagree, even though
execution itself comes from the p-code.

Only `/decompile` promotes source over p-code, by regenerating the p-code
from it. Dropping `__SRP_` is not a recompile.

### Changing the line count: the per-procedure counters

Growing or shrinking a module used to produce "Compile error: Expected
End Function" the moment Access recompiled. The pre-0xCAFE header carries
a pair of u16 counters per procedure at ``base + func_``, and the rule
they hold was wrong.

A controlled series settles it: the same module compiled by Access with
one to six body lines, then two- and three-procedure variants and a class
module. Exactly two header offsets move with the line count, and they
move by one per line:

```
N body lines:  1  2  3  4  5  6
@516 and @518: 3  4  5  6  7  8
```

The value is **the number of source lines the procedure spans**: from its
``FuncDefn`` up to the next procedure's, or to the end of the module for
the last one, so a blank separator counts toward the procedure above it.

| module | procedures (FuncDefn line) | counters |
|--------|---------------------------|----------|
| one procedure, 6 body lines | 1, module ends at 9 | 8 |
| two procedures | 1, 5, module ends at 9 | 4, 4 |
| three procedures | 1, 5, 10, module ends at 15 | 4, 5, 5 |

``EndFunc`` plays no part, which is what the old rule got wrong: a
``Declare`` emits a ``FuncDefn`` with no matching ``EndFunc``, so pairing
them drops its counter and shifts every later one. The base is 516 for a
standard module and 612 for a class module.

The new rule reproduces Access exactly in **97 of the 103 modules** in
this repo that contain a procedure, up from 4. The six exceptions are
fixtures whose p-code was deliberately left inconsistent with their
source.

With this and the ``__SRP_`` drop, arbitrary statement counts work:
twelve generated statements -- ``Set``, ``CreateObject``, member calls
with arguments, ``^`` and ``Select Case`` -- written into an empty
procedure body return 145 from real Access.

### Gate what runs, not just what parses

Every other gate here is static, and both passed on databases Access
refused to run. ``verify_execution.py`` closes that: it rewrites a
procedure, drops the cache and asks Access for the answer, over a
same-count body, a grown one and one built from nothing.

A warning about probes, since this cost real time twice: bare
``Application.Eval`` over COM **hangs** behind a modal VBA compile-error
dialog. Use ``pyvbaharness``, which reports ``modal-blocked`` instead.

## Module rename, delete and create

**Rename and delete work. Create loads but is not finished.** What
unblocked them was not new p-code work: it was the storage engine in
`pyopenvba.access`, which arrived after this directory was parked. Every
place a module's name lives is an ordinary table row, so each edit is an
`update_row` that resizes the row properly. The note below that said
"resizing a storage-catalog row corrupts it" was a limit of patching
bytes in place, not of the format.

`attribute_pages.py` is what made the map: it names the table that owns
each page two databases differ on and diffs the rows, so an Access
operation reads as "MSysObjects row 26 Name changed" rather than as
sixty pages of noise.

### Where a module's name lives

Eight places, and Access needs them to agree:

| where | what changes |
|-------|--------------|
| its `MSysObjects` row | `Name` |
| its `MSysNavPaneObjectIDs` row | `Name` |
| `MSysAccessStorage` `\x03DirData` | `04 <len> <name UTF-16>`, len being the name's bytes plus four |
| `MSysAccessStorage` `PROJECTwm` | `<name MBCS> 00 <name UTF-16> 00 00` |
| `MSysAccessStorage` `PROJECT` | the `Module=<name>` line and the `[Workspace]` line |
| the `dir` stream | MODULENAME and MODULENAMEUNICODE |
| the module's own stream | `Attribute VB_Name` |
| `_VBA_PROJECT` | an appended identifier, and the module's UTF-16 record |

The earlier attempt updated six of them and Access still showed the old
name. The two it missed were `MSysNavPaneObjectIDs` and `_VBA_PROJECT`.

Three findings from asking Access rather than reading the file:

**MSysObjects is not what Access displays.** Renaming the catalog row and
the navigation-pane row alone leaves both the VBE and
`CurrentProject.AllModules` showing the old name, and the code still
runs. DirData, PROJECTwm and PROJECT decide `AllModules`; the VBE's name
comes from `_VBA_PROJECT`.

**Access repairs what it disagrees with.** A rename that reaches the dir
stream and `Attribute VB_Name` but not `_VBA_PROJECT` is *reverted* the
moment Access opens the database -- the file on disk says `Gamma` before
the open and `Alpha` after it. Reading a file back after Access has
touched it measures Access's repair, not the write.

**A rename appends an identifier, it does not rewrite one.** VBA is
case-insensitive and one identifier record serves every use of a name, so
renaming the record in place would rename a variable that happens to
share it. Access appends `<u8 len> <u8 4> <name> <u16 hash> 10 00` before
the `02 ff ff 01 01` sentinel, bumps a slot counter at `start - 14` and a
record counter at `start - 12`, and points the module's UTF-16 record at
the new slot. The operand is **`2 * slot + 2`**, taking the slot counter's
value from before the bump; the familiar `524 + 2*index` is the same rule
for slot `261 + index`. The hash is the OLE `LHashValOfNameSysA` value
`pcode_hash.py` already computes. Both edits come out byte-identical to
Access's own rename.

The table itself begins after a `00 00 00 00 02 00` anchor, which also
occurs in unrelated data, so a candidate only counts when the record
behind it reads as a named one.

### What rename reproduces

`module_rename.py` renames all eight places and drops the `__SRP_` cache.
Checked against real Access on a project of three modules, one of which
calls `Alpha.AlphaGo`:

| check | result |
|---|---|
| the VBE and `AllModules` both show the new name | yes |
| the renamed module's code still runs | `AlphaGo` returns 42 |
| a caller naming the old module still runs | `CallIt` returns 42 |
| the VBE can add a procedure to it and Access compiles and runs it | returns 999 |
| a shorter name (`Zed`) and a much longer one | both work |

`Application.Run "Module.Procedure"` fails on a renamed module -- and on
an untouched one, and on a module Access itself renamed, so it is an
Access limitation rather than a defect here. Controls are what settled
it.

### Delete

`module_delete.py`, measured against `DoCmd.DeleteObject`:

* the dir stream loses the module's whole block, MODULENAME through
  MODULEEND, and PROJECTMODULES drops by one
* DirData loses its entry **and the four bytes that follow it** -- leaving
  those behind produced a file Access opened with an empty `AllModules`,
  which is the shape of a list Access could not walk
* PROJECTwm and PROJECT lose their entries
* `_VBA_PROJECT` loses the module's entry from the table that follows the
  project cookie, and that table's count drops by one, while the
  identifier table keeps the name and its counters do not move
* the module's stream row, its `MSysObjects` row and its
  `MSysNavPaneObjectIDs` row all go

DirData, PROJECTwm and the module table come out byte-identical to
Access's own delete, `AllModules` matches it exactly, and the project
still compiles and runs afterwards.

One rule the storage engine learned here: deleting a **catalog** row
leaves the page it was alone on alive and owned, where a filtered DELETE
retires it (`Table.delete_row(rid, retire_empty=False)`).

### Create, as far as it goes

`module_create.py` writes everything a new module needs. Access opens the
result, the VBE enumerates the new module with its type and line count,
its source reads back, `CurrentProject.AllModules` lists it, the
project's other modules still execute, and the VBE **will add further
components to it**. What still fails is editing an existing module's code
in a project built this way.

What it writes:

* three `MSysAccessStorage` rows -- a numbered storage folder under
  `Modules`, a 13-byte `PropData` under it, and the module's stream under
  `VBA` with a row name of 28 random capitals
* an entry in `Modules/PropData`, the list of those folders:
  `05 09 02 <folder name, one UTF-16 character> "CB0"`, eleven bytes each.
  The folder name is a single character and Access takes the one after
  the highest in use -- `0`, then `4`, then `5`.
* a dir block of eleven records and PROJECTMODULES up by one
* entries in DirData, PROJECTwm and PROJECT
* in `_VBA_PROJECT`: an identifier for the name, a flag in the per-module
  list, a 32-byte per-module record, a module entry appended after the
  last, the project's reserve advanced, and the module index grown by two
  slots and rehashed
* an `MSysObjects` row of type -32761 and a navigation-pane row

**Ids come from the table's own AutoNumber, not from max + 1.** Every
database Access wrote has `MSysAccessStorage`'s AutoNumber counter equal
to its highest id. Inserting rows with explicit ids leaves the counter
behind, and Access's own next insert then collides with a row that
already exists -- which is what made a created project refuse a new
component while reading and enumerating perfectly. Letting the counter
assign the ids is the fix, and it is the single change that turned "will
not take a component" into "will".

Three structures found the same way, by asking why Access called the
project corrupt:

**The 32-byte per-module record.** Immediately before the identifier
table's counters sits one record per module past the first,
`ff ff ff ff 01 00 00 00 ff ff ff ff <u32 reserve> <16-byte GUID>`, then
`80 00 00 00 00 00` and the counters. Growing one project from one module
to five gave records 0x228, 0x278, 0x298, 0x2b8 -- each new one carrying
the reserve of the module that was last before it.

**The reserve rule.** A new module's reserve word is the value the
project's trailer offers, and the trailer then advances by 0x20: 0x208,
then 0x278, 0x298, 0x2b8, 0x2d8 as the project grew. Every module also
carries its own 20-byte cookie, `<two characters><the project's own
eight>`; two sharing one is not something Access writes.

**The flag list.** Ahead of the module table sits `<u16 module count>
<count u16 flags, each 1> <u16 n> <n four-byte records>`; adding a module
bumps the count and inserts a flag.

### The module index

The table past the identifier table's `02 ff ff 01 01` sentinel is a u32
size and then a hash of six-byte slots, `<u16 key> <u16 value> <u16
chain>`, an empty slot being all `ff`. Every module has an entry -- the
first included -- whose key is its **name operand plus one** and whose
value is its index in the module table, alongside six fixed entries a
project always carries (0x0006, 0x0058, 0x020e, 0x021c, 0x021e, 0x0221).
Placement is

    slot = (key >> 1) % slots

with linear probing, which fits all 44 entries of five projects of one to
five modules with no exceptions. Since an operand is `2 * slot + 2`, that
hash is the identifier's own slot number plus one. Access grows the table
by two slots per module and rehashes it.

A module's entry chains to the slot of the module before it and the first
module's chains to itself; the entries that are not modules carry `ffff`.
Rebuilding the table that way reproduces Access's own index **byte for
byte** -- placement, keys, values and chains -- on the three-module
project it made by adding to a two-module one.

### What still fails, and what is known about it

Editing an existing module's code in a created project fails, where the
same edit succeeds on every project Access built. `module_create.py`
takes a `skip` set so the pieces can be bisected, and the bisect says the
`_VBA_PROJECT` edit is responsible: write everything else and leave that
row alone, and the edit succeeds (the module then has no name in the VBE,
since the VBE reads its list from the dir stream). Appending only the
identifier is also fine -- that is what `module_rename.py` does, and
renamed projects take edits happily. Every smaller subset of the
`_VBA_PROJECT` work leaves the project inconsistent in some other way, so
the bisect cannot narrow it further without a structure that has not been
modelled yet.

Two facts found on the way, both worth keeping:

**A cloned module brings its names with it.** Cloning a module whose code
calls `ZetaGo` leaves the p-code naming an identifier the target project
does not have. Access's own new module appends two identifiers, the
module's name and its procedure's; a donor with no procedures needs only
the one.

**`Module1` in the shipped template does not round-trip.** Rebuilding it
unchanged through `Perf.build` gives a row one byte shorter, differing at
the source-length field (`3fb0` against `3eb0`), so the first attempts at
create were cloning a malformed stream. A module Access has just made
rebuilds byte for byte, which is the check to run on any donor first.

**Three traps when checking this by hand.** Opening a database in Access
rewrites it, so a second measurement is measuring the repair -- make a
fresh copy per operation. `VBComponents.Add` over a bare COM boundary
fails with "Device I/O error" on a database Access itself built, so a
failure there says nothing. And splicing a structure out of another
project rather than generating it makes Access discard the whole project
and rebuild it empty, which reads as success from the outside: check with
`peek_project.ps1` before believing it.

### The per-module entry in `_VBA_PROJECT`

After the project cookie comes a module count and one entry each:

```
<u16 stream-name bytes> <stream name UTF-16>
<u16 20> <10-character cookie UTF-16>
ff ff <u16 name operand> <u16 name bytes> <name UTF-16>
ff ff <u16 module cookie> 00*6 <u16> 00 00 00 <u32 module offset>
```

Entries are separated by `ff ff`; the first follows the count directly.
The module cookie is the dir stream's MODULEEND2 for that module and the
offset its MODULEOFFSET, so the two structures carry the same facts
twice.

## Auditing the guards

Every real defect in this directory lately has been a guard or table that
looked complete, so the guards themselves were audited. Three were wrong.

**`find_counter_base` could pick a base that does not exist.** It scanned
every even offset for one where the predicted counters already matched,
and returned it if exactly one did. On six modules whose p-code is
inconsistent with their source the rule predicts nonsense -- 1005 lines
for a four-line procedure -- and the scan found a single offset, 596,
where the nonsense happened to match. It would then have written counters
into arbitrary header bytes with full confidence. It now tries only the
two measured bases, 516 and 612, and refuses otherwise; a predicted count
larger than the module's line count is rejected outright.

**`verify_identity` was checking the header and calling it the row.** It
compared `rebuilt[:cafe]` and then did a "no-op rewrite" that wrote back
the *original* row rather than the rebuilt one -- so the file comparison
was trivially true and tested only `set_lval_payload`. Both halves are
fixed, and the gate now compares the whole row.

That exposed the third: **a no-op rebuild was not byte-identical on a
single module in the repository.** Access does not zero the padding
between p-code lines, and `build` did:

```
original  d01c ffff780000009604...      rebuilt  d01c 000000000000 9604...
```

Harmless at run time -- every module written this way executes -- but it
meant a gate reporting "modules rebuild identically" had never compared
the p-code region at all. Preserving the original padding takes the
corpus from **0 to 127 of 142 modules rebuilding byte for byte**, and
`_require_reproducible` now checks the whole row too.

**`RESERVED_SLOT` was load-bearing and only 18 of 261 slots deep.** A
name VBA pre-interns resolves to a slot, not a project identifier.
Interning one anyway turns out to be harmless for a *variable* -- a
module using `Text` as a local runs either way -- but fatal when the name
is the **procedure's**, because that is the binding the function result
uses:

```
Go interned as a project identifier   ->  Access refuses to compile
Go resolved to slot 92                ->  42
```

Completing a 261-entry table by guessing candidate names is not a
strategy. Instead the operand is now **read out of the module**: a
`Function` assigns to its own name, and that `St` already carries
whatever operand Access chose. Removing `go` from the table entirely and
re-running gives 42, so correctness no longer depends on the table being
complete. The harvest was extended anyway (CDec, RGB, StrComp, and the
call-form names) because it still improves fidelity.

**`_plan_declarations` matched on names and ignored types.** Rewriting
`Dim aa As Long` to `Dim aa As Double` kept the Long record, so
`aa = 3.7` came back as **4** while the source plainly said Double --
source and behaviour disagreeing with nothing to warn you. The prefix now
compares name *and* type, and a changed type becomes a release followed
by an append, which already worked.

**The chained-row guard did not exist.** `load_module` set an
`info["chained"]` flag and a comment said `write_module` would refuse on
it; nothing ever read it. Writing a chained module turns out to work when
the result fits one row -- the spanning-pages fixture rewrites and runs
-- so the comment was wrong rather than the code. What does fail is a
chain whose MODULEOFFSET does not address the assembled bytes, which used
to surface as an unhandled decompression error; it is now a plain
refusal.

The pattern is worth naming: each of these passed by not looking, and
each was found by asking the guard to prove what its name claimed. Three
of the six fixes replaced a table this code maintains with a value read
from the file being edited, which is the shape that generalises.

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

## Calling VBA's built-ins

`Left`, `Len` and `Abs` used to be refused as unknown identifiers. A probe
module calling eighty built-ins (`builtins_probe.bas`) shows there are
four mechanisms, and which one applies is not guessable from the name:

| how it is reached | count | examples |
|---|---|---|
| an ordinary project identifier | 56 | `Trim`, `UCase`, `Chr`, `Now`, `Replace`, `IIf`, `Nz`, `DLookup` |
| a dedicated opcode | 7 | `Len`→`FnLen`, `Abs`→`FnAbs`, `InStr`→`FnInStr`, `Int`, `Fix`, `Sgn`, `StrComp` |
| a pre-populated slot, via `ArgsLd` | 6 | `Left`=109, `Mid`=124, `String`=173, `Format`=85, `CurDir`=37, `FreeFile`=87 |
| its own opcode plus a slot | 1 | `Array(...)` -> `ArgsArray` naming slot 8 |

The largest group needs nothing: those names already flow through
`_add_missing_identifiers` exactly like a user-written one, which is why
`MsgBox` worked long before any of this.

Three further shapes fall out of the same probe. The conversion functions
share opcode `Coerce` and differ only in op_type -- `CVar` 0, `CInt` 2,
`CLng` 3, `CDbl` 5, `CDate` 7, `CStr` 8, `CBool` 11. `UBound`/`LBound`
carry the dimension as an *operand*, not a stack argument. And `Date` is
a value rather than a call: Access rewrites `Date()` to `Date` and emits
`Ld` of slot 44.

### A nested call was silently miscounting arguments

The probe also caught a defect that had nothing to do with built-ins.
`arg_list()` recorded its argument count on the parser, so an inner call
overwrote the outer one:

```
DateAdd("d", 1, Now())    Access: ArgsLd(DateAdd)[3]   ours: [1]
Join(Array(1, 2), ",")    Access: ArgsLd(Join)[2]      ours: [4]
```

Any `f(g())` was affected. The count is now returned rather than stashed.
Another silent miscompile that only a differential probe would find.

## Where the compiler stands

Measured across the three probe modules: **195 statements byte-identical
to Access, 0 differing.** What is still refused, and why:

| refused | count | blocked on |
|---------|-------|-----------|
| `Dim`, `Const`, `ReDim`, user-defined `Type` | 28 | growing the pre-CAFE header (below) |
| procedure headers and footers, `Option` | 26 | creating a procedure, not a body statement |
| `Declare` | 2 | same |
| `Set x = New <Class>` | 2 | the import table `New` indexes into |

So every remaining refusal in a *procedure body* is a declaration. The
statement grammar itself is done for practical purposes.

Two notes from getting there. Explicit parentheses are not free: Access
records a `Paren` marker after the grouped expression, so `(a + b) * a`
differs from `a + b * a` by more than precedence. And the corpus gate now
separates lines whose **source is ahead of their p-code** from real
differences -- pyOpenVBA's source-only write path leaves databases in
that state, and counting them as code-generation defects made the total
move every time the compiler learned a new statement.

## `Dim`: the local-name hash table, solved

`Dim x As Long` is eight bytes of p-code -- `Dim | VarDefn(var_)`, the
same eight whatever the type -- plus a 24-byte record in the pre-0xCAFE
header carrying everything else. Two experiments ruled out every cheap
route before the real one was found:

- **Source-only** (compile the `Dim` away, let the variable be an
  implicit Variant) runs, but silently miscompiles:
  `Dim x As Long : x = 3.7 : Probe = x` returns **4** in real Access and
  **3.7** source-only.
- **Real p-code, no header record** crashes Access outright: `var_` then
  points at unrelated bytes.

### The table

Records live at `464 + var_`, `var_` starting at 88 and striding by 24,
and form a linked list; the type is a plain VARTYPE code in the first
field. Past the records sits a **16-bucket hash table**, 4 bytes per
bucket, at `464 + 88 + 24*ndecl + 48`. And the index is not a new
mystery:

```
bucket = identifier_hash(name) % 16
```

the *same* OLE `LHashValOfNameSysA` the project identifier table uses,
already solved in `pcode_hash.py`. Verified on 9 modules across two name
series, D5 included.

The table holds **procedures as well as variables** -- in a module whose
function is `Go`, bucket 7 is `Go`'s record offset, and it shifts by 24
each time a declaration is inserted ahead of it, while variable entries
(`var_` values) stay put.

### Collisions

A new name takes the bucket, and **the record immediately before it takes
custody of whatever was displaced** -- the previous variable's own
offset, the procedure's offset when a procedure lost the bucket, or the
null marker when the bucket was empty. `collision_probe.bas` forces four
names into one bucket to pin this down.

That single rule explains every case, including the one that looked
anomalous for hours: a module where a new variable collided with the
procedure name, putting the procedure's offset into a variable's record.

### The arena

Records are carved out of an arena whose remaining space is tracked by an
optional `ffffffff <free>` pair after a marker, with a flag 36 bytes past
the records saying whether an arena exists at all. Measured on modules of
one to nine declarations:

```
ndecl    1    2    3    4    5       6     7     8     9
free    88   64   40   16   (none)  480   456   432   408
flag     0    0    0    0   248       0     0     0     0
```

`free` drops by 24 per declaration; when it can no longer cover one the
pair is removed and the flag goes to 248; the next declaration allocates
a fresh 480-byte arena and clears the flag again. So the header grows by
24 normally, by 16 when the arena is exhausted, and by 32 when one is
allocated -- which is exactly the "resize" that looked like an
unmodelled reorganization.

### Procedures are in the same chain

The linked list does not begin at the first variable: a procedure has a
record of its own immediately before it, at `464 + 64`, in the same list
and the same shape. Adding a module's *first* declaration therefore links
the procedure's record to it, and inherits the owner from there -- which
is why frame offsets run `-32` for the procedure and then `-40`, `-48`
for the variables.

Two consequences worth stating. There is always a previous record to
convert, so no special case is needed for the first declaration. And that
record must be **patched field by field, not rewritten** from the
variable template: a procedure record uses fields a variable does not,
and overwriting them corrupts it.

### What reproduces



Appending is **byte-identical to Access on 17 of 18 measured
transitions**, zero through nine declarations across four independent
name series, including four-way hash collisions, a collision with the
procedure name, arena exhaustion and arena reallocation. There is no
declaration ceiling. The one exception is a single u32 in the per-build
scratch region whose value is arbitrary across modules (0, 1, 0xffff...),
the same class as the cookie at offset 41 -- and the module runs.

`Dim` is wired through `rewrite_module.py`, so it works from the command
line like any other statement:

```
Dim aa As Long / Dim bb As Long / Dim cc As Double
aa = 20 / bb = 22 / cc = 0.5 / Go = (aa + bb) * cc     -> 21
```

### The other declaration forms, and a corruption path closed

`Dim x As Long` is the only form modelled. The others were probed
(`array_probe.bas`), and each reshapes the header its own way:

| form | p-code | header |
|------|--------|--------|
| `Const K As Long = 10` | `Dim~1`, the value, `VarDefn~2` | record type gains bit `0x40` |
| `Static s As Long` | `Dim~32`, `VarDefn~1` | region shifts by 8 |
| `Dim a(1 To 5) As Long` | bounds pushed, then `VarDefn~1` | +40 bytes of array descriptor |
| `Dim s As String * 8` | length pushed, then `VarDefn~1` | +8 bytes |
| `ReDim a(1 To 3)` | `Redim(name)` with two operands | none, it is a statement |

Measuring them mattered for a reason beyond completeness. `Dim arr(1 To
5) As Long` does not match the scalar pattern, so `is_declaration`
returned None and the count guard compared 0 against 0 and passed. A
rewrite whose new body simply omitted that line then appended a record
**on top of the array's descriptor**, and the tool reported success:

```
K2_clob.accdb: body 3 -> 3 lines, ... 4 __SRP_ cache row(s) dropped
Access: the server threw an exception
```

Silent corruption, produced by the guard being narrower than the thing it
guarded. `declares_storage` now matches every form that reserves storage,
modelled or not, and a module containing one is refused outright.

### Releasing a record

`remove_declaration` is the inverse, and the same probe corpus tests it
by running every transition backwards. It restores the displaced entry
from the previous record's custody field, closes that record again, and
returns 24 bytes to the arena -- re-creating the free pair when the
arena had been exhausted.

**Both directions reproduce Access byte for byte on all 18 pairs.**

One trap: the record patch has to happen *after* the pointer fixups. One
of the absolute fixups lands at offset 540, which is inside the procedure
record, so patching first and fixing up second turned a zero into -24 and
crashed Access. The same ordering applies to both operations -- the third
time this session a "fixed offset" turned out to be a field with meaning.

The new body and the old must agree on a prefix: whatever the old body
has beyond it is released, newest first, and whatever the new body has
beyond it is appended. Reordering, or removing from the middle, would
renumber every later `var_` and the p-code referring to them, so it is
refused.

The written declaration **runs, with its type honoured**: a synthesized
`Dim bb As Long` makes `bb = 3.7` return 4, and the same code with
VARTYPE Double returns 3.7 -- matching Access exactly, and passing the
test that killed the source-only shortcut.

Two bugs worth recording, both found only by diffing:

- All-ones is a **null sentinel**, not a number. Incrementing it during
  pointer fixup turned a null into offset 23 and Access crashed.
- One "pointer fixup" at a fixed offset was really **hash bucket 13**.
  Bumping it corrupted a live entry whenever a name happened to land
  there.
- Another was really the **arena's free counter**. Once the arena was
  modelled properly the fixed-offset version double-counted it, which is
  why removing it from the list fixed thirteen transitions at once.

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

## Calling VBA's built-ins

`Left`, `Len` and `Abs` used to be refused as unknown identifiers. A probe
module calling eighty built-ins (`builtins_probe.bas`) shows there are
four mechanisms, and which one applies is not guessable from the name:

| how it is reached | count | examples |
|---|---|---|
| an ordinary project identifier | 56 | `Trim`, `UCase`, `Chr`, `Now`, `Replace`, `IIf`, `Nz`, `DLookup` |
| a dedicated opcode | 7 | `Len`→`FnLen`, `Abs`→`FnAbs`, `InStr`→`FnInStr`, `Int`, `Fix`, `Sgn`, `StrComp` |
| a pre-populated slot, via `ArgsLd` | 6 | `Left`=109, `Mid`=124, `String`=173, `Format`=85, `CurDir`=37, `FreeFile`=87 |
| its own opcode plus a slot | 1 | `Array(...)` -> `ArgsArray` naming slot 8 |

The largest group needs nothing: those names already flow through
`_add_missing_identifiers` exactly like a user-written one, which is why
`MsgBox` worked long before any of this.

Three further shapes fall out of the same probe. The conversion functions
share opcode `Coerce` and differ only in op_type -- `CVar` 0, `CInt` 2,
`CLng` 3, `CDbl` 5, `CDate` 7, `CStr` 8, `CBool` 11. `UBound`/`LBound`
carry the dimension as an *operand*, not a stack argument. And `Date` is
a value rather than a call: Access rewrites `Date()` to `Date` and emits
`Ld` of slot 44.

### A nested call was silently miscounting arguments

The probe also caught a defect that had nothing to do with built-ins.
`arg_list()` recorded its argument count on the parser, so an inner call
overwrote the outer one:

```
DateAdd("d", 1, Now())    Access: ArgsLd(DateAdd)[3]   ours: [1]
Join(Array(1, 2), ",")    Access: ArgsLd(Join)[2]      ours: [4]
```

Any `f(g())` was affected. The count is now returned rather than stashed.
Another silent miscompile that only a differential probe would find.

## Where the compiler stands

Measured across the three probe modules: **195 statements byte-identical
to Access, 0 differing.** What is still refused, and why:

| refused | count | blocked on |
|---------|-------|-----------|
| `Dim`, `Const`, `ReDim`, user-defined `Type` | 28 | growing the pre-CAFE header (below) |
| procedure headers and footers, `Option` | 26 | creating a procedure, not a body statement |
| `Declare` | 2 | same |
| `Set x = New <Class>` | 2 | the import table `New` indexes into |

So every remaining refusal in a *procedure body* is a declaration. The
statement grammar itself is done for practical purposes.

Two notes from getting there. Explicit parentheses are not free: Access
records a `Paren` marker after the grouped expression, so `(a + b) * a`
differs from `a + b * a` by more than precedence. And the corpus gate now
separates lines whose **source is ahead of their p-code** from real
differences -- pyOpenVBA's source-only write path leaves databases in
that state, and counting them as code-generation defects made the total
move every time the compiler learned a new statement.

## Why `Dim` stays refused: the shortcut is a silent miscompile

Two experiments settle whether `Dim` can be shipped cheaply, and both say
no.

**Source-only, variable implicit.** Compile `Dim x As Long` to an empty
p-code line -- source keeps it for display, and every use of `x` is an
implicit Variant. It runs: `Dim x As Long / x = 7 / Probe = x + 1`
returns 8. But the declared type changes coercion, and dropping it is
wrong, not merely lossy:

```
Dim x As Long : x = 3.7 : Probe = x
   real Access (As Long):  4     (rounds to the declared type)
   source-only (Variant):  3.7
```

That is precisely the silent-miscompile class every gate in this
directory exists to stop, so the shortcut is rejected rather than shipped.

**Real p-code, header left alone.** Emit the genuine
`Dim | VarDefn(var_=88)` bytes but do not grow the header. Access
crashes (`server threw an exception`): the `VarDefn` points at offset
`464 + 88`, which in an ungrown header is unrelated data. The declaration
record is not optional.

So `Dim` needs the real record in a grown header, and stays refused until
that exists -- refusing is the correct outcome, not a gap.

## What `Dim` needs: the declaration record, mapped

`Dim` is still refused, but the structure behind it is now measured
rather than guessed. `dim_probe.bas` builds the same module with zero to
five declarations and one per declared type.

The p-code half is trivial, and **identical for every type**:

```
Dim aa As Long      5d00 f504 5800 0000     Dim | VarDefn~1(var_=88)
```

`Long`, `String`, `Variant`, `Double`, `Boolean` and `Object` all emit
those exact eight bytes. Everything that distinguishes them is in the
header.

### The 24-byte record

Records live at `464 + var_`, with `var_` starting at 88 and striding by
24. They form a linked list, and take one of two shapes depending on
whether another declaration follows:

```
        +0    +2    +4    +6    +8    +10   +12   +14   +16   +18   +20  +22
not last  TT  ffff  0000  0000  8460  next  ffff  ffff  frame ffff  ffff ffff
last      TT  ffff  0000  0000  ffff  ffff  0000  0000  8302  owner ffff ffff
```

* `TT` is the declared type as a plain VARTYPE code -- Long 3, Double 5,
  String 8, Object 9, Boolean 11, Variant 12. The same numbering the
  `Coerce` op_types use, so the format has one type table, not two.
* `next` is the following declaration's name operand; `owner` on the last
  record is the procedure's.
* `frame` is the variable's frame offset, `-40` for the first and eight
  less for each after it -- so every local occupies eight bytes whatever
  its type. The last record has no frame field: that slot carries the
  owner link instead.

Adding a declaration therefore rewrites the previously-last record into
the non-last shape, which is exactly what Access does -- `aa` gains its
`-40` frame offset only once `bb` exists.

### The header fields that move

Growing the declaration region moves a handful of fields. Measured across
zero to five declarations:

| offset | size | per declaration |
|--------|------|-----------------|
| 488 | u32 | **+24, always** -- the declaration region's own size |
| 492 | u32 | -8 -- frame size, signed and negative |
| 516, 518 | u16 | +1 -- the line counters, because the module gained a line |
| 9, 25, 444, 540 | u32 | +24 for the first four, then +16 |
| 29 | u32 | +44, then +36 -- end of p-code, so it also absorbs the new line record |
| 41 | u32 | changes every build; **not validated**, see above |

That last row of +16 is the open question. The declaration region itself
grows by exactly 24 every time -- field 488 says so without exception --
so something else in the header lost eight bytes on the fifth variable.
The likely culprit is the project identifier region, whose compact
records pack differently from build to build; the same per-build
variability that made the `_VBA_PROJECT` "symbol buckets" a dead end.

So the remaining work is not the record, which is understood, but
inserting it: region C past the records holds a **table of `var_`
offsets** (the u32s 88, 112, 136 appear there verbatim), so adding a
declaration shifts every later one and every pointer into them must be
fixed up. On top of that, a local-name **hash table** in region C resized
on the fifth declaration -- the 8-byte shrink is `ffffffff10000000`, one
empty bucket -- so the header's size is not a fixed multiple of the
record count. Whether that table is validated was the next thing to test
with the deadbeef method that cleared the offset-41 cookie -- and unlike
the cookie, **it is**. Overwriting the empty-bucket markers (`ffffffff`)
of a working three-declaration module with `0xde`, dropping the cache and
running, crashes Access every time. So the buckets are read, not
rebuilt on load, and faithful `Dim` needs the local-name hash table
constructed correctly: the new name's bucket inserted, the table resized
when load demands it, the `var_` offset table extended, and every
internal pointer fixed up. That is the "symbol buckets" structure that
was a dead end earlier, now confirmed load-bearing rather than scratch.
This is the real frontier for declarations, not a quick win.

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
