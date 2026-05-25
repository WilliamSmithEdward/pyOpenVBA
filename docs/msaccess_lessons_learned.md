# MS Access VBA: Lessons Learned

This document chronicles pyOpenVBA's attempt to extend its
read-and-write VBA support from Excel, Word, and PowerPoint to
Access `.accdb` databases, and the reasons we ultimately shipped only
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

---

## 6. Final decision

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
