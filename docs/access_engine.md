# The Access engine

**Status: in progress, started 2026-09-01.** Read layer shipping first;
each later phase lands only when live Access agrees with it.

## Context

pyOpenVBA read Access VBA through `AccessReader`, a page scanner that finds
MS-OVBA blobs by signature and decodes `MSysObjects` with a hand-written,
17-column row decoder. Writing was parked in August 2026: rewriting a
procedure body worked and Access executed the result, but creating,
renaming or deleting a module never did. The record in
`docs/research/access_write/README.md` shows why: the catalog side --
`MSysObjects`, `MSysAccessStorage`, `MSysNavPaneObjectIDs` and the B-trees
over them -- was byte-patched, not written through anything that
understood rows, usage maps or indexes.

On 2026-09-01 the direction changed to replicating the whole engine.

## Decision

Build a Jet 4 / ACE storage engine in pure Python as the `pyopenvba.access`
subpackage, one private module per layer of the file:

| module | owns |
|---|---|
| `_pages` | the page array, the masked definition page, row slots, usage maps |
| `_tdef` | table definitions: columns, real and logical indexes, usage-map refs |
| `_rows` | row splitting and the codec for every column type |
| `_lval` | Memo/OLE long values: inline, single row, chained |
| `_index` (next) | B-tree pages, entry masks, sort-key encoding |
| `_catalog` (next) | `MSysObjects` and the other system tables as typed objects |
| `database` | `AccessDatabase` and `Table`, the public facade |

Order of work: read everything; write rows; write indexes; write schema;
write the VBA project through the same writer (un-parking the 2026-08
research); queries, relationships and properties; forms, reports and
macros; a SQL executor over the lot.

`AccessReader` keeps its public API and will be rebuilt on the engine once
the engine finds the same rows it does today.

## Ground truth, in order of authority

1. **Live Access**, through `pyvbaharness` (`RUN_LIVE_ACCESS=1`), test-time
   only -- pyOpenVBA never uses COM. Access builds the data and the engine
   reads it back; DAO's view is the reference
   (`tests/test_live_access_engine_gate.py`). For writes the direction
   reverses: the engine writes, Access reads back; and Access performs
   the same edit so the bytes can be diffed. This is the method that
   found every UserForm rule the spec left out, and it is the only way to
   learn a rule Access enforces but does not write down.
2. **Access-authored fixtures**: the five databases listed in
   `tests/test_access_engine.py`. Files produced by the earlier
   experimental writer (`tests/live_access_test/_write_*.accdb`,
   `tests/output/**`, `docs/research/**`) are not ground truth; the engine
   flags eleven of them as inconsistent (an LVAL row shorter or longer
   than its 12-byte definition says), which Access tolerated at the time.
3. **Format references** -- the mdbtools `HACKING` notes and Jackcess --
   for a rule not visible in a file. Consulted for layout knowledge; no
   code is copied and every rule taken from them is re-measured here.

## Format facts established, and how each was checked

* **Page 0 mask.** Bytes 0x18..0x96 (126) are XORed with the RC4 keystream
  of key `C7 DA 39 6B`. Checked: code page 1252 at 0x3C, LCID 0x409 at
  0x6E, creation date at 0x72 decoding to the fixture's creation day.
* **Table definition layout** (see the docstring of `_tdef.py`): 12-byte
  real-index headers, 25-byte column headers, names, 52-byte real index
  definitions, 28-byte logical index definitions, names, then a 10-byte
  usage-map pair per long-value column ending in `FF FF`. Checked: the
  bytes consumed equal the length the page declares, for every definition
  in every authored fixture; the parser refuses otherwise.
* **Row layout.** u16 column count, fixed data at each column's fixed
  offset, variable data, then `(var count + 1)` u16 offsets in reverse
  order, the var count, and a null mask whose set bit means "has value".
  Boolean columns exist only in that mask. Checked: every table in every
  authored fixture counts to its definition's row count, and the catalog
  names match the shipped reader's independent decoder.
* **Slot flags.** `0x4000` alone is a live overflow pointer (row byte,
  three-byte page) to a row moved to another page; the moved row's own
  slot carries `0x8000`, which therefore means "deleted" only on a page
  reached directly; `0xC000` is a dead slot, often with a stale offset
  shared with its neighbour. Found on the 1 MB fixture, whose `Table2` row
  the old reader garbled.
* **Usage maps.** Kind 0: u32 start page, then a 64-byte bitmap. Kind 1:
  u32 page numbers of type-5 pages whose bytes from offset 4 are bitmap
  chunks. The global map (page 1, row 0) marks *free* pages and counts
  pages past the end of the file as free.
* **Long values.** 12-byte definition: u24 length, kind, row, u24 page,
  four unused bytes. Kind `0x80` inline after the definition, `0x40` one
  LVAL row of exactly that length, `0x00` a chain whose rows each start
  with a 4-byte next pointer, `(0, 0)` last.
* **Index leaf pages.** Owner at 4, prev/next/tail at 12/16/20, a 453-byte
  entry mask at 27 over entry data from 480; a leaf entry is the encoded
  key columns followed by a three-byte big-endian page and a one-byte
  row. Longs are stored big-endian with the sign bit flipped behind a
  `0x7F` flag byte; text uses a collation table (observed: a = A = 0x4A,
  c 0x4D, D 0x4F, e 0x51, g 0x55, P 0x66, s 0x6B, t 0x6D) that will be
  generated from Access rather than transcribed.

## Alternatives considered

* **Drive Access through COM at runtime.** Breaks the library's one hard
  rule and needs Windows with Office.
* **Wrap mdbtools or Jackcess.** Neither is pure Python, neither writes
  the ACE catalog the way Access does, neither knows VBA.
* **Keep byte-patching.** The parked write path is the measurement that
  this does not scale past a procedure body.

## Consequences

* The package grows a database engine. It stays zero-dependency, and an
  unedited database must save back byte for byte before any write feature
  ships, as the UserForm writer did.
* Jet 4 (`.mdb`, version 1) and ACE (versions 2, 3, 5) share the page
  format and are all in scope. Jet 3 (Access 97, 2 KiB pages) is refused.
* The engine is strict: an LVAL row that does not match its definition,
  or a definition that does not reconcile, raises rather than guessing.
  That makes it a corruption detector as well as a reader.

## Phases

| # | phase | status |
|---|---|---|
| 1 | pages, header, usage maps, definitions, rows, long values, catalog | reading. Every table in every authored fixture decodes; the live gate matches ACE field for field on all 16 column types, null rows, chained 5000-character memos and a 151-column definition spanning two pages |
| 2 | indexes: walk B-trees, decode entries, sort keys for every type | next |
| 3 | write rows: insert/update/delete, free-space and owned-page maps, LVAL allocation, counters | |
| 4 | write indexes: key encoding from an Access-generated collation table, B-tree insert and split | |
| 5 | write schema: create/drop table, catalog rows, permissions, navigation pane rows | |
| 6 | VBA project through the writer: module create/rename/delete | |
| 7 | queries (`MSysQueries` to SQL and back), relationships, properties | |
| 8 | forms, reports, macros: the binary object formats nobody has published | |
| 9 | SQL executor over the engine | |
