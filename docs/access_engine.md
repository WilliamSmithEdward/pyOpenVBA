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
| `_index` | B-tree pages, entry masks, the key codec both ways |
| `_collation` | text sort keys; `_collation_general_legacy` is its generated table |
| `_datapage` | in-place editing of a data page's slots and rows |
| `_alloc` | page allocation from the global usage map, usage-map bit writes |
| `_btree` | index insert and delete, page compression and splits |
| `_schema` | table definitions written from specs or re-serialized, map pages, index roots |
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
* **Index pages** (type 3 node, type 4 leaf). Free bytes at 2, owner at
  4, prev/next/tail at 12/16/20, a u16 prefix length at 0x18, a 453-byte
  bit mask at 27 over the entry area from 480. A set mask bit marks the
  END of an entry, the first starting at 0; every entry after the first
  is stored without its first `prefix length` bytes, which equal the
  first entry's. An entry is the encoded key, then the row's home slot
  (three-byte big-endian page, one-byte row), then on node pages the
  big-endian child page; a node entry carries the last key of its child,
  and the tail pointer names the child holding everything greater. Index
  entries point at a row's home slot, which survives the row being moved
  to an overflow page. The 12-byte index header's u32 at +4 is the
  engine's distinct-key count, null counting as one key. Checked: every
  index of an ACE-written 1500-row table -- one per column type, one
  descending, one two-column, one unique ignore-nulls -- decodes to the
  row values it points at, in order, with every node entry equal to its
  child's last entry.
* **Key encoding.** Per column a flag byte: `0x7F` value ascending,
  `0x80` value descending, `0x00` null ascending, `0xFF` null descending.
  A descending value is the ascending bytes inverted. Ascending, all
  big-endian: Boolean `0x00` True / `0xFF` False (True sorts first, as
  -1 does); Byte as is; Integer, Long, BigInt and Currency (scaled by
  10 000) with the sign bit flipped; Single, Double and DateTime as IEEE
  bits with the sign bit flipped when positive and every bit inverted
  when negative; Decimal `0xFF` + 16-byte magnitude when positive, `0x00`
  + inverted magnitude when negative; Binary in eight-byte chunks each
  followed by `0x09` while more follow and by the count of real bytes in
  the last; GUID as its 16 bytes in textual order through the binary
  scheme; Text as collation bytes, `0x01`, up to four `0x01`-separated
  extra-weight sections with trailing empty ones omitted, `0x00`.
* **Text collation** (sort order 1033 version 0, the "General" order of
  Jet 4 and Access 2007 files; DAO-created databases of both formats
  carry it). `_collation.py` reproduces it and
  `_collation_general_legacy.py` is generated by
  `scripts/generate_access_collation.py` from one indexed row per BMP
  code point plus 146 composition samples; the generator re-encodes all
  63 632 strings and gets every one back byte for byte. The rules: case
  is not stored (a unique index treats `a` and `A` as duplicates); each
  character yields zero or more *elements* of one or two bytes (19 585
  code points yield none, `ß` yields `ss`, `ﬃ` yields `ffi`); trailing
  spaces are dropped and other spaces weigh `0x07`; a combining mark
  folds into its base when a precomposed letter exists, a first mark
  without one takes the weight it gives any precomposed letter, and a
  further mark adds the weight it has standing alone. After the
  primaries and `0x01` come up to four `0x01`-separated sections with
  trailing empty ones omitted: one diacritic weight per element (`0x02`
  placeholder, trailing placeholders trimmed); nothing ever seen; kana
  as a bit stream (`10` marker then two bits per kana, three kana per
  byte, `11` full-size and `10` small, cut after the last small one)
  followed by `ff 02 80 ff 80` and one more `ff` when section 4 follows;
  ignorable-but-recorded characters (hyphen, apostrophe, controls) as
  `80 <7 + 4 * elements before it> 06 <code>`. The engine stores at most
  510 key bytes and cuts longer keys without a clean end, so the encoder
  refuses those instead of guessing.

* **Writing rows** (measured by having the engine perform single
  inserts, updates and deletes on a small table and diffing every page).
  Rows lie contiguously below the slot table in slot order; free space at
  2 is exactly `4096 - 14 - 2 * slots - bytes below the lowest row`.
  Deleting shifts the rows below the hole up, leaves the slot flagged
  `0xC000` at the boundary it now sits on, and does not clear the freed
  bytes; replacing a row shifts the rows below it by the size change;
  inserting appends a slot below the lowest row. The definition's row
  count at 0x10 tracks live rows; 0x14 holds the last AutoNumber handed
  out; an index header's u32 at +4 grows only when a new distinct key
  arrives and never shrinks. A null fixed-length field is left holding
  whatever the engine's buffer had (one attachment row carries stale
  text); pyOpenVBA writes zeros, which the null mask makes equivalent.
  When no page in the free-space map can take a row the engine drops
  that page from the map and takes the lowest free page of the global map,
  growing the file to reach it, then registers the page with both of the
  table's maps. Text is compressed only when the column's byte 16 says so
  (Access sets it, SQL DDL does not unless `WITH COMPRESSION`), except in
  the engine's own catalog tables where it always is. A fixed-size Binary
  column stores its full width, zero-padded. Page 0 carries a counter at
  0xE02 the engine bumps per SQL statement; it is left alone.
* **Growing indexes.** An entry is inserted in full-byte order (key then
  row pointer) into the leaf found by descending the node separators. A
  page keeps its prefix length until an entry no longer fits; then it is
  compressed with the full common prefix; then it splits. A leaf that
  fills while entries are appended stays full and the next leaf starts
  with the new entry (602 then 298 for 900 sequential Longs); a middle
  insert splits in half (457 and 443 for 900 random ones). The root page
  number is fixed: a splitting root becomes a node with one separator, a
  tail child and level 1 at 0x1A, both children fresh pages. Deleting an
  entry rewrites the leaf compactly over its old bytes. A single insert
  and a single delete written by pyOpenVBA are byte-identical to the
  engine's own on every page but page 0.

* **Long values** (Memo and OLE), measured by inserting values of every
  size through the engine. Up to 64 bytes of value live inline behind the
  12-byte definition (kind `0x80`); up to 3816 bytes go as one row on an
  LVAL page shared by the column's values (kind `0x40`, the page chosen
  from the column's free-space map, a fresh page otherwise); anything
  longer is a chain (kind `0x00`) of 4072-byte payloads, each behind a
  4-byte next pointer, one chunk per fresh page. Memo text is compressed
  inline when that is shorter (a one-character memo is not) and stored
  uncompressed outside the row, whatever the column's compression flag.
  A chained value's definition ends in a 4-byte stamp that also sits at
  offset 8 of its first page; the engine uses one stamp per session and
  the file reads fine with any stamp as long as the two match. Clearing
  or deleting a chained value releases its pages to the global map and
  drops them from the column's owned map, content left in place; a
  single-page value is tombstoned on its LVAL page.
* **Overflow rows.** When an updated row no longer fits its page, the
  engine writes it as a row on a page from the table's free-space map
  with slot flag `0x8000`, and replaces the row at home with a 4-byte
  pointer (row byte, three-byte page) under flag `0x4000`; index entries
  keep pointing at the home slot. When it shrinks enough it comes home and
  the copy is tombstoned; deleting it tombstones both slots.

* **Definitions longer than one page**, measured with tables of 111 to
  151 columns whose names were tuned to land the definition length on
  4086 to 4100 bytes, then 8111, 10547 and 13244. The engine counts 4088
  bytes (4096 minus an 8-byte reserve) per page, so a definition takes
  `ceil(length / 4088)` pages: 4088 bytes fit one page and 4089 already
  take two, the second holding nothing but its header. Physically the
  first page holds the first 4096 bytes and each continuation the next
  4088 after an 8-byte header (`02 01`, free word, next page, 0); the
  free word is `4088 * pages - length` on the last page and 0 on every
  other. Continuation pages are allocated last when a table is created
  (after the index roots), in ascending order, and chained in reverse:
  the first allocated page ends the chain. Every rewrite of a definition
  (CREATE INDEX) allocates a fresh chain the same way and only then
  releases the old continuation pages, bytes intact, even when the page
  count does not change. DROP TABLE marks only the first page 0x08;
  continuation pages just return to the free map. The catalog row's
  DateUpdate is stamped when the definition is complete, so on a
  150-column table it runs a couple of milliseconds after DateCreate;
  `create_table(created=, updated=)` takes both.
* **Stamps carry more than a millisecond**, seen in 14 of 112 catalog
  timestamps the engine wrote: their doubles sit one bit away from any
  millisecond value, and no arithmetic tried (nearest, ceiling, floor,
  twenty operation orders) reproduces them from a datetime. A datetime
  is stored as the nearest double; the stored serial is exposed
  (`CatalogEntry.date_create_serial`) and accepted wherever a DateTime
  goes, and an update keeps the bytes of every column it does not touch
  rather than re-encoding decoded values, which is what made two catalog
  rows differ by a bit.
* **Released pages wait for the next session**, measured by dropping a
  table and creating another in one DAO session (the new table took pages
  past the end of the file while the dropped table's stayed free), then
  creating one more in a fresh session (it took the dropped table's
  pages, lowest first). Pages a session releases, whether by DROP TABLE,
  a definition rewrite or a freed long value, are not handed out again
  until the database is reopened. An `AccessDatabase` instance is a
  session: `PageStore.released` holds what it has released, the
  allocator skips those, and a new instance starts clean.
* **Creating and dropping tables**, measured by diffing `CREATE TABLE`,
  `CREATE INDEX` and `DROP TABLE` page by page. A new table takes its
  definition page and one data-shaped page (owner 0) holding its usage
  maps as 69-byte inline rows: owned pages, free-space pages, one per
  index, two per Memo/OLE column. Each index gets an empty leaf as root.
  In the definition, real indexes keep creation order while the logical
  list and its names are stored sorted by name (each logical entry
  naming its own index number); the per-table tag at 0x0C and in every column header is the
  database's (0x659 everywhere seen); a real index definition begins
  `83 07 00 00`, a logical one carries `04 04` before its kind byte; a
  Boolean column is "fixed" of length 1 but takes no row space, a GUID is
  variable-length, a fixed Binary keeps its declared width. The catalog
  gets an MSysObjects row (Id = definition page, parent the Tables
  container, Type 1, uncompressed name) written in two steps -- inserted
  without an owner, then updated with one -- and three MSysACEs rows. The
  pages are taken in the order definition, map page, whatever the catalog
  rows need, index roots. `CREATE INDEX` appends the real index, re-sorts
  the logical list, appends a map row, allocates a root and re-stamps the
  table's DateUpdate. Dropping releases the owned pages and kills the
  owned-map row first, then clears every index and long-value map, then
  kills the remaining map rows in order, marks the definition page type 8
  and releases every page; that order is what decides which stale bytes
  the dead map rows keep.
* **Two rules found on the way.** An update stays in place when its
  growth fits the page's free space and otherwise moves the row behind a
  pointer (a two-byte growth stays with three bytes free and moves with
  one); a page a row moved off leaves the free-space map. An insert needs
  room for the row and its slot entry, else the next page -- a catalog
  row whose home is a pointer from the moment it is created got that way
  from the two-step write above, not from an insert rule (a first reading
  of the bytes said otherwise; the growth comparison below corrected it).
* **Growing past 512 pages.** An inline usage map covers 8 pages per
  bitmap byte from its start page. When a page beyond its reach is
  added, a map that holds pages grows its bitmap in 8-byte steps to the
  least size covering the page (573 pages: the global map's row goes from
  69 to 77 bytes; 1708 pages: 221), and an empty map is re-based to the
  page's 8-aligned start instead (a table whose only data page is 542
  gets start 536). The global map is extended one step at a time when it
  lists no free page, the 64 new pages counting as free. The reference
  form of a map has not yet been seen written by the engine below 1708
  pages. Checked: 450 memo rows carrying a database from 121 to 573 pages
  leave every page but page 0 identical to the engine's own.

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
| 2 | indexes: walk B-trees, decode entries, sort keys for every type | done for reading. Every index on every fixture and on the live 1500-row table checks out, and `encode_key` rebuilds all 25 500 of its entries from the row values, text included |
| 3 | write rows: insert/update/delete, free-space and owned-page maps, LVAL allocation, counters | done: every column type including Memo/OLE of every storage kind, overflow rows, unique-index enforcement, page allocation and all counters; the engine reads the result, keeps working on it and compacts it; single edits and memo inserts byte-identical to the engine's |
| 4 | write indexes: key encoding from the engine-generated collation table, B-tree insert and split | done: entries inserted and removed, pages compressed when full and split, root pinned; single edits byte-identical to the engine |
| 3b | large files: usage maps growing past 512 pages | done for inline maps (growth and re-base as the engine does); the reference form is read but not yet written |
| 5 | write schema: create/drop table, create index, catalog rows | done: `create_table`, `create_index`, `drop_table`; byte-identical to the engine's CREATE TABLE, CREATE INDEX and DROP TABLE on every page but page 0; the engine inserts into, reads and compacts a table pyOpenVBA created; definitions over one page (up to the 255-column limit) are chained and rewritten as the engine does, byte-identical. Not yet: a second map page, navigation-pane rows (the Access layer adds those itself) |
| 6 | VBA project through the writer: module create/rename/delete | |
| 7 | queries (`MSysQueries` to SQL and back), relationships, properties | |
| 8 | forms, reports, macros: the binary object formats nobody has published | |
| 9 | SQL executor over the engine | |
