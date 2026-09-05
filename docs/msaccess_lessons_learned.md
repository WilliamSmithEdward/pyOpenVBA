# MS Access VBA: Lessons Learned

**Historical.** This chronicle ends, in August 2026, with a read-only
`AccessReader` and the conclusion that a pure-Python writer was out of
reach. That conclusion no longer stands: since September 2026
`AccessDatabase` writes Access databases, VBA included, through a
pure-Python implementation of the storage engine, and the module
operations this document could not reach (create, rename, delete) are
done and checked in live Access. See [access_engine.md](access_engine.md)
for the engine and the top-level README for the API. The account below
is kept for what it found along the way.

This document chronicles pyOpenVBA's attempt to extend its
read-and-write VBA support from Excel, Word, and PowerPoint to
Access `.accdb` databases, and the reasons it shipped, at the time, only
a read-only `AccessReader` API.

The work spanned multiple reverse-engineering phases (referred to as
phases 5a through 5h in the development history). Every phase
expanded our understanding of the Access on-disk format; none of them
produced a write surface that survived a full Access GUI roundtrip on
a non-trivial mutation.

The end state is intentional: `AccessReader` exposes a rich pure-Python
read API and nothing else. If you want to edit VBA in an Access
database from Python, drive Access COM automation directly.

---

## 1. Goal

Feature parity with `ExcelFile`, `WordFile`, and `PowerPointFile`:

- `AccessReader(path)` opens a database in memory.
- `read_vba_module(name)`, `vba_modules()`, `iter_vba_modules()`,
  `pull_modules()` all work.
- A symmetric write surface (`set_module`, `replace_module`,
  `rename_module`, `delete_module`, `import_module`, `push_modules`)
  produces a `.accdb` that reopens cleanly in the Access GUI with the
  edits visible in the VBA editor and the modules runnable at the
  database level.

The read side is straightforward and shipped. The write side is what
this document is about.

---

## 2. Reverse-engineering phases

Each phase peeled back a layer of the Access VBA storage stack.

| Phase | Layer reached                                  | Outcome                                |
|-------|------------------------------------------------|----------------------------------------|
| 5a    | LVAL catalog row discovery and OVBA decompress | Read works; cache row identified.      |
| 5b    | OVBA cache rewrite (same-length swap)          | Python reader sees the new bytes.      |
| 5c    | MSysObjects row rename / delete / add          | Python reader stays consistent.        |
| 5e    | dir-stream `MODULENAME` rewrite                | Catalog stays internally consistent.   |
| 5f    | Rename catalog (header + per-record)           | All record IDs stay valid.             |
| 5g    | OVBA compressor parity with MS-OVBA            | Roundtrip is byte-exact.               |
| 5h    | PROJECT INI rewrite + p-code invalidation      | Cache rewrites do not survive Access.  |

By phase 5h the pure-Python side of the story was complete: any
sequence of mutations we performed in memory was perfectly mirrored
back through `AccessReader`'s own reader on disk reload. The Access
GUI disagreed.

---

## 3. Empirical results matrix

Every mutation was driven by a diagnostic script (`diag_*`) against
the canonical corpus sample `040__sub_msgbox_hello.accdb` and the
hand-edited `New Microsoft Access Database.accdb` fixture. "Python
roundtrip" means `AccessReader` reads back the mutated state correctly.
"Access GUI" means: copy the mutated file to a Windows machine, open
it in Microsoft Access, navigate to the VBA editor, and observe the
result.

| Experiment                                  | Python roundtrip | Access GUI               |
|---------------------------------------------|------------------|--------------------------|
| Same-length body swap (`replace_text`)      | PASS             | PASS                     |
| In-place body grow (`replace_text_resize`)  | PASS             | SOFT ERROR on save       |
| Rename module (catalog + cache + MSysObj)   | PASS             | "cannot read VBA project"|
| Delete module (catalog + MSysObj tombstone) | PASS             | HARD CRASH               |
| Add module (catalog + MSysObj insert)       | PASS             | HARD CRASH               |
| Decompose-and-rename (rebuild from scratch) | PASS             | HARD CRASH               |
| Tombstone `rU@` p-code marker only          | PASS             | No effect (cache ignored)|
| Tombstone `CAFE` p-code rows only           | PASS             | No effect (cache ignored)|
| Tombstone both + 100-line cache rewrite     | PASS             | No effect (cache ignored)|

The pattern is consistent: as soon as Access's own state machine has
to reconcile the cache with its compiled module table, our edits get
discarded or the database is rejected.

### 3.1 Follow-up (2026-08, p-code tooling + automated execution harness)

The original matrix judged edits by opening the database in the Access
GUI and looking. With the VBA7 p-code decoder/encoder now available and
an automated harness -- build a module via COM (`RunCommand 126`,
acCmdCompileAndSaveAllModules), edit bytes in Python, reopen and *run*
the macro via `pyvbaharness.AccessSession`, and read a sentinel file the
macro writes -- edits can be judged by what actually **executes**, not
just what displays. Three results sharpen the picture:

| Experiment                                   | Executes as edited? |
|----------------------------------------------|---------------------|
| Same-length interned-string (`B9`) edit      | YES                 |
| Same-length p-code literal edit (`LitDI2`)   | YES                 |
| `_VBA_PROJECT` version-cookie bump            | NO -- runner-error  |

The first two are the important additions. Access executes p-code that
references interned string-literal rows by index; a same-length edit to
either the p-code stream or a `B9` literal row changes the value the
compiled code produces. `5 + 3` re-encoded to `5 + 9` returns 14 from
real Access; `"ORIGINALSTR"` overwritten with `"CHANGEDXSTR"` prints the
new string. So a bounded but genuinely **executable** write already
works, and this is the first validation of the p-code assembler (built
and previously checked only against Excel) against Access's own runtime.

The version-cookie result is the useful negative. Excel treats a
`_VBA_PROJECT` version mismatch as "discard p-code, recompile from
source"; Access does not. Bumping the cookie yields a runner-error, not
a recompile. There is no load-time recompile-from-source trigger to
exploit -- confirming, with a cleaner test than the p-code tombstones
above, that arbitrary writing must produce correct p-code rather than
lean on Access to regenerate it.

The remaining barrier is unchanged and is about **size, not
correctness**: every edit above is length-preserving. Growing a row --
which any real source change needs -- still requires reproducing
Access's LVAL page allocator and `MSysObjects` chunk-pointer mutation
(section 5).

### 3.2 Breakthrough (2026-08): executable pure-Python writes

The barrier in 3.1 was stated as "size, not correctness". That was right
about the symptom and wrong about the cause. Growing a row does not need
Access's page allocator at all; it needs one field.

Every VBA long-value is described by an `MSysAccessStorage` catalog row
holding `<u16 length> 00 40 <slot><page>`, and Access trusts that length
over the page slot table. Resizing a row while that length still reads the
old value makes Access fault while loading the project -- which is exactly
the "soft error on save" the phase-5 resize experiment hit. Updating it
makes resizing work. **This field is the "internal index binding a module
name to its compiled rows" that section 5 lists as never located.**

Three further results follow from it:

| Experiment | Result |
|---|---|
| Edit only the canonical `0xCAFE` p-code, leave every `rU@` row stale | Access returns the **new** value -- the CAFE region is what executes |
| Edit only the source text, leave p-code alone | Access returns the **old** value -- source is display-only |
| Grow a module row by 12 bytes and fix the catalog length | Runs correctly; `compile_project` succeeds |

The second row settles a question 3.1 left open. There is no load-time
recompile-from-source trigger, so writing source text alone yields a
module that displays one thing and does another. Any Access writer must
generate correct p-code.

Generating it turned out to be tractable because **VBA p-code control flow
is structured rather than jump-based**: `IfBlock`/`ElseBlock`/`EndIfBlock`,
`DoWhile`/`Loop` and `For`/`NextVar` carry no branch offsets, so there are
no jump targets to fix up. Implicit (undeclared) variables also need no
declaration record. A compiler covering full operator precedence,
assignment, `If`/`ElseIf`/`Else`, `Do While`/`Loop` and `For`/`Next` now
re-emits **33 of 33 statements byte-for-byte identically to Microsoft's
own compiler**, and a module whose logic was rewritten entirely in Python
computes the right answer in real Access.

The statement count can change too. That looked blocked by a table of
24-byte slot records in the pre-`0xCAFE` header, but the comparison behind
that reading was confounded: holding variables and control flow constant
while varying only the statement count leaves the header size and the slot
table **identical**. Slot records track variables and control-flow blocks,
not statements. The real constraint was two u16 fields at `+516`/`+518`
holding the procedure's line count; patched by the line delta, a module
regrown from 3 statements to 13 compiles and runs (the per-line frame-size
hint turned out to be advisory).

New **identifiers** can be added too. That looked blocked by
Access-generated "symbol buckets" in the `_VBA_PROJECT` row, but building
the same source twice changes those bytes as well, so they are per-build
scratch state with nothing to reproduce. What Access validates is a pair
of u16 counters just before the identifier table (record count at
`table_start - 10`, slot count at `-12`); appending a record without
bumping both makes Access hang on load. With them corrected, a name
appended before the table sentinel resolves normally at `524 + 2*index`,
using the already-solved `LHashValOfNameSysA` hash.

What remains is larger than that summary once suggested, and the work is
**parked as of 2026-08-19**: creating a procedure needs the `FuncDefn`
declaration tables; module create, rename and delete need coordinated
edits across six structures and do not work; `Const`, arrays, `Static`
and fixed-length strings each reshape the module header differently and
are refused; and growing a module past its 4 KB page needs the LVAL chain
allocator. What does work is rewriting a procedure *body* -- arbitrary
statements, arbitrary control flow, and `Dim` declarations added, removed
or retyped. Code and the full status in `docs/research/access_write/`.

---

## 4. Why Access is different

Excel, Word, and PowerPoint all use the **MS-OVBA** CFB container as
their authoritative VBA store. The OVBA streams are the source of
truth. When the host application opens the file it recompiles from
those streams on demand. pyOpenVBA's writes to those streams are
therefore the only thing the host has to consume, and the host always
agrees with what was written.

Access is different. An `.accdb` file stores VBA in **two** places:

1. **OVBA cache rows in LVAL pages.** These are OVBA-compressed copies
   of the `dir`, `_VBA_PROJECT`, `PROJECT`, and per-module source
   streams. This is what pyOpenVBA reads. It is byte-compatible with
   the MS-OVBA spec.
2. **Compiled p-code, stored in independent LVAL rows.** Each module
   has an `rU@ ...` active-pcode marker row and one or more `CAFE`
   rows that hold the VBA7-compiled bytecode. This is what the Access
   VBA editor actually executes and displays.

The OVBA cache is **passive**. We could find no recompile trigger
that forces Access to discard the compiled p-code and rebuild it from
the cache rows. Tombstoning the `rU@` marker, the `CAFE` rows, or
both did not work: Access either ignored our changes or refused to
load the project.

The compiled p-code rows are **authoritative**. Editing them
correctly requires a full VBA7 p-code assembler: instruction stream
generation, identifier table re-indexing, jump fixup tables,
versioning negotiation against the database's stored runtime
version. This is a multi-month effort with no public specification
and no off-the-shelf reference implementation.

There is also a third actor: the **`MSysObjects`** system table holds
the row that names each module and ties it to its physical storage.
Its schema is private and partially undocumented (Type 5 entries with
`Lv` long-value columns whose payload is the p-code chunk pointer).
Successfully renaming or adding a module requires updating
`MSysObjects` in a way Access trusts.

---

## 5. What was needed but not built

To deliver a production-quality writer we would have needed all of
the following, in addition to the OVBA-cache surgery we already had:

- A complete **VBA7 p-code assembler** that emits `CAFE` rows
  bit-compatible with Access's native compiler output.
- An exact reproduction of Access's **MSysObjects** mutation path,
  including the long-value chunk allocator and chunk pointer rewrites.
- A reproduction of Access's **page-allocator** behavior so that
  newly-allocated LVAL/DATA pages land in the same regions Access
  would have chosen.
- A way to invalidate or refresh whatever **internal index** Access
  uses to bind a module name to its compiled `rU@`/`CAFE` rows.
  This index was never located.

Without all of these, any write that crosses a module-identity
boundary (rename, delete, add, anything that grows the source by more
than the cache row can absorb) will either be silently reverted or
will corrupt the database.

**Update (2026-08).** Two of these four are resolved, and one was
misdiagnosed. The p-code assembler exists and matches Microsoft's output
byte-for-byte on every statement tested (section 3.2). The "internal
index" is the `MSysAccessStorage` length field, and updating it is what
lets a row grow -- no page allocator is involved, because a resized row
still fits its existing page; the allocator only matters once a row
outgrows one. What remains is the compiled header's temporary-slot table,
which fixes the statement count, and the `_VBA_PROJECT` symbol buckets,
which fix the identifier set.

---

## 6. Final decision (superseded, see the note at the top)

`AccessReader` ships as a **read-only** class. The library cleanly
extracts every VBA module, attribute block, identifier list, and
p-code stream we can decode, and writes nothing back. The
`pull_access` top-level helper mirrors the `pull` / `pull_word` /
`pull_ppt` API so you can dump every module to a directory of
`.bas` / `.cls` files and version them in git.

If you need to programmatically modify VBA inside a `.accdb`, drive
Access through COM (`win32com.client.Dispatch("Access.Application")`).
That is the supported Microsoft path and it interacts with the
compiled p-code through the same code paths the VBA editor uses.

The pure-Python write path remains an interesting reverse-engineering
problem. We would happily revisit it if a public spec or reference
implementation for the VBA7 p-code format ever emerges.
