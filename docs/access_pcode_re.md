# Access VBA p-code reverse-engineering

This document accretes findings from the multi-session effort to fully
understand the `.accdb` VBA storage format so that pyOpenVBA can perform
structural writes (add/rename/remove modules, change source) on Access
databases without invoking Office COM.

## Goal

Pure-Python, dependency-free, bidirectional read/write parity for VBA
inside `.accdb` files, matching what `pyopenvba.excel`, `pyopenvba.word`
and `pyopenvba.powerpoint` already deliver for OOXML containers.

## What is already shipped (read + same-length plaintext write)

* `AccessFile` reads OVBA-compressed module source via the LVAL-page
  chain walker (`_lval_segments`, `_read_lval_chain`,
  `_find_ovba_signature_offsets`). Parity with the COM oracle is
  verified on the canonical fixture.
* `AccessFile.replace_text` performs same-length, in-place plaintext
  patches against the authoritative `E3 00 00 00 <u16 len> <ascii>`
  comment-row table and `B9 00 <u16 len> <ascii> <12B trailer>`
  literal-row table. Verified end-to-end against the live VBA editor.
* `AccessFile.save` writes the in-memory mutated bytes back to disk.

## What is NOT yet supported (the open problem)

Structural changes:

* add module
* remove module
* rename module
* arbitrary-length source change (length-mismatched source replacement)

Each of those requires Access's symbol/name table, p-code table, OVBA
catalog and (for length changes) the per-page B-tree of the LVAL store
to be rewritten consistently. None of those tables are documented.

We confirmed earlier in the project that the per-module OVBA blob is a
**passive cache** -- Access reads its compiled p-code from a separate,
authoritative byte region. Force-recompile attempts (zero OVBA + zero
per-module compile-cache 4-tuple) did NOT cause Access to recompile
from OVBA. The p-code tables are therefore the real target.

## Storage layout (current understanding)

Each `.accdb` that contains a VBA project stores the full project state
in a single long-value (LVAL) record. The record's payload is structured
as a concatenation of:

1. **Plaintext PROJECT file** -- the same line-based format Excel/Word/
   PowerPoint ship in their CFB `PROJECT` stream:
   ```
   ID="{<guid>}"
   Module=<head module name>
   Name="<project name>"
   HelpContextID="0"
   VersionCompatible32="393222000"
   CMG="..."
   DPB="..."
   GC="..."

   [Host Extender Info]
   &H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000

   [Workspace]
   <ModuleName>=38, 38, 4512, 1443, Z
   ```
2. **Binary catalog / symbol-table region.** Undocumented. This is the
   primary target of the RE corpus diff work.
3. **One MS-OVBA-compressed stream per module**, beginning with the
   canonical
   `Attribute VB_Name = "<name>"\r\nOption Compare Database\r\n<body>`.
4. **Plaintext source-row tables** -- `B9 00 <u16 len> <ascii> <12B
   trailer>` for string literals, `E3 00 00 00 <u16 len> <ascii>` for
   comments. These are the bytes `AccessFile.replace_text` already
   mutates.

The LVAL record is stored across one or more LVAL pages (page type
`0x01`, tag `LVAL` at +4). Each LVAL page may hold chunks of multiple
distinct long-value records, indexed by a per-row slot table that
begins at page offset 14. The continuation pointer for a given chunk
is a `(page, slot)` tuple stored at the tail of the row, NOT a single
page-level next-pointer in the page header. **This row format is not
yet decoded** -- it is the first blocker for any multi-page chain
reassembly.

## Phased plan

### Phase 1 -- corpus + tooling (DONE this session)

* `tests/live_access_test/_corpus_generate.ps1` -- COM-driven corpus
  generator. Produces `baseline_empty.accdb`, `baseline_empty_proj.accdb`
  and 25 minimal samples (IDs 010..051) covering: empty StdModule with
  varying names, empty ClassModule, single empty Sub, single statements
  (MsgBox/Dim/Let/comment), basic If/For. Output lives in
  `tests/live_access_test/re_corpus/` (gitignored).
* `scripts/_diff_accdb.py` -- page-aware byte diff with page-type
  tagging.
* `scripts/access_re_chain.py` -- locates the project-VBA LVAL head
  page by the `ID="{` plaintext fingerprint, extracts the head-page
  payload slice, splits it into known section types and dumps it.

### Phase 2 -- LVAL row format

Reverse-engineer the per-row layout on an LVAL page so that the full
multi-page chain can be walked deterministically:

* row offset/length encoding in the slot table at `page[14:]`,
* row prefix (length, flags, owner-table backreference),
* row trailing continuation `(page, slot)` pointer,
* relationship between an LVAL row and the data-page row in
  `MSysObjects`/system catalog that points to it.

Deliverable: `access_re_chain.py` extended to reassemble the full
multi-page payload for the project VBA record.

### Phase 3 -- symbol/catalog table

Diff the binary section between the end of the PROJECT plaintext and
the first OVBA stream across corpus samples that differ only in:

* module name (010..016): isolates module-name table entries
* module kind (Std vs Class) (010 vs 020): isolates kind flag
* procedure name (030..033): isolates proc-name table entries
* source body (040..051): isolates body-only deltas

Deliverable: byte-precise documentation of the catalog structure.

### Phase 4 -- p-code opcode field guide

Decompress each per-module OVBA stream via
`pyopenvba.vba.decompress(..., stream_name="VBA")` and diff the
resulting bytes across the body-varying samples (040..051) to learn the
opcode encoding.

Deliverable: `docs/access_pcode_re.md` opcode table section, expanded
turn by turn.

### Phase 5 -- structural write primitives

Implement and test, in dependency order:

1. `set_module_source(name, src)` for length-matched source (already
   works via `replace_text`; promote to a public API once Phase 4 lets
   us regenerate the matching p-code).
2. `set_module_source(name, src)` for length-mismatched source --
   requires Phase 2 (B-tree rebalance) + Phase 4 (p-code regeneration).
3. `add_module(name, kind, src)` -- catalog insert + new p-code emit.
4. `rename_module(old, new)` -- catalog patch only.
5. `remove_module(name)` -- catalog delete + LVAL free.

Each capability ships with COM-oracle regression tests in
`tests/test_access.py` and a freshly baked fixture.

## Reference assets

* Canonical fixture: `tests/live_access_test/New Microsoft Access Database.accdb`.
* RE corpus baselines & samples: `tests/live_access_test/re_corpus/`
  (regenerable via `_corpus_generate.ps1`).
* COM oracle: `tests/live_access_test/_oracle.ps1`.
* COM mutator (for ground-truth before/after pairs):
  `tests/live_access_test/_com_mutate.ps1`.
* Chain locator: `scripts/access_re_chain.py`.
* Page-aware byte diff: `scripts/_diff_accdb.py`.

## Repo-memory notes

Working notes that should NOT be repeated in this doc live in
`/memories/repo/access-vba-storage.md` (general format) and will live
in `/memories/repo/access-vba-pcode-re.md` (per-opcode findings) once
Phase 4 starts.
