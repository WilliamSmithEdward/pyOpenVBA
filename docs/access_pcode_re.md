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

### Phase 2 -- LVAL row format (DONE 2026-05)

Reverse-engineered the per-row layout so the multi-page chain can be
walked deterministically and module discovery works on arbitrary
.accdb files (verified on all 25 corpus samples + canonical fixture).

LVAL page layout (4 KiB pages):

* `[0]`        `page_type = 0x01`
* `[4:8]`      `'LVAL'` tag
* `[12:14]`    u16 LE slot count `N`
* `[14:14+2N]` u16 LE slot table. Top nibble `0xD` = tombstone; else
  the low 12 bits are the row's byte offset within the page.
* Rows grow downward from `PAGE_SIZE`. A row ends at the smallest
  higher non-tombstone slot offset, or `PAGE_SIZE` for the top row.

Long-value chunk continuation prefix (present only on chained rows):

* `row[0]`   u8 next_slot
* `row[1:4]` u24 LE next_page
* `row[4:]`  chunk payload
* Terminator: `(next_slot, next_page) == (0, 0)`.

A long-value that fits in one chunk is stored standalone -- the row IS
the payload, with NO continuation prefix. Standalone vs chain-head is
not encoded in the row itself; the working heuristic is "treat
`row[1:4]` as a page number; if it's in range AND points to another
LVAL page, walk it as a chain, otherwise treat as standalone."

Implementation: `pyopenvba.access.AccessFile.iter_vba_modules` walks
every non-tombstone LVAL row, scans for MS-OVBA stream signatures
(`0x01` followed by a u16 LE chunk header with sig bits `0b011`),
and accepts any candidate that decompresses to a stream beginning
with `Attribute VB_Name = "..."`.

### Phase 3 -- symbol/catalog table (DONE 2026-05)

**Finding: there is no Access-specific symbol-table format. The
"binary catalog" is just an MS-OVBA `dir` stream (MS-OVBA section
2.3.4.2) OVBA-compressed in a single LVAL row.**

In every corpus sample examined, exactly one LVAL row OVBA-decompresses
to bytes starting with the PROJECTSYSKIND record header
`01 00 04 00 00 00`. That row contains the full standard dir-stream
TLV record sequence:

* PROJECTSYSKIND / PROJECTLCID / PROJECTLCIDINVOKE / PROJECTCODEPAGE
* PROJECTNAME (`baseline_empty_proj` in the empty-project corpus)
* PROJECTVERSION (special-cased: no real size field; 10-byte payload)
* PROJECTREFERENCES -- `stdole` plus `DAO` (ACEDAO.DLL) are standard
* PROJECTMODULES blocks: MODULENAME / MODULENAMEUNICODE /
  MODULESTREAMNAME (a mangled obfuscated identifier; Access does not
  use it as a real stream name since there is no CFB) /
  MODULEDOCSTRING / MODULEOFFSET / MODULETYPE
  (0x0021 = procedural, 0x0022 = class) / MODULEREADONLY /
  MODULEPRIVATE / module terminator 0x002B
* Dir-stream terminator 0x0010

The catalog row's slot index is **not** stable across files (010..015
put it at (68, 1); class-module sample 020 places it at (68, 2)). We
locate it by content -- decompress every LVAL row and accept the one
whose decompressed prefix is `01 00 04 00 00 00`.

Production wiring (`AccessFile.read_project_info`) reuses
`pyopenvba.vba.parse_dir_stream` directly (same parser that drives the
Excel/Word/PowerPoint paths). `vba_module_names()` now prefers the
catalog's authoritative ordered list when available, falling back to
the OVBA-scan path only if the catalog row cannot be located.

Deliverable: DONE. Byte-precise documentation lives in MS-OVBA
section 2.3.4.2; the parser in `pyopenvba/vba.py:parse_dir_stream`
handles all currently observed records. Reference-record decoding
(beyond the standard `registered` kind) and write-back (catalog
mutation when modules are added/removed/renamed) are still pending and
fall under Phase 5.

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
