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
* **Relationships**, measured by diffing `ALTER TABLE ... ADD CONSTRAINT
  ... FOREIGN KEY ... REFERENCES` twice against one parent. The
  referencing table gets a non-unique index named after the constraint
  (flags 0x80, an empty root, a map row, entries for existing rows) and a
  logical entry of kind 2 whose byte 12 is 2, bytes 13..16 the parent's
  new logical index number, bytes 17..20 the parent's definition page,
  bytes 21 and 22 the cascade-update and cascade-delete flags (normal
  indexes carry `04 04` there). The referenced table gets a logical entry
  named `.r` plus the letter at that index number (`.rB` for its second
  logical index, `.rC` for the third) sharing the unique index the key
  refers to, byte 12 = 1, pointing back at the child's page and logical
  number. MSysRelationships gets one row per column pair (`grbit` is
  DAO's RelationAttributeEnum: 0x100 cascade updates, 0x1000 cascade
  deletes, 2 not enforced). The relationship is also a catalog object:
  an MSysObjects row of type 8 under the Relationships container, id one
  past the highest negative id, owner as usual, DateCreate = DateUpdate,
  with three MSysACEs rows whose ACMs are 0xF00FE, 0xFFFFF, 0xFFFFF on
  the three default SIDs in order. Both tables' DateUpdate is stamped,
  at slightly different instants.
* **A second usage-map page**, measured with a table of 32 indexes and 12
  Memo columns (58 map rows). The map page holds 57 rows of 69 bytes;
  the 58th went to row 0 of a fresh map page (a data page with owner 0,
  like the first) allocated the moment the current one was full, in the
  place CREATE INDEX takes its map row: before the index's root page and
  the definition's rewrite. `_new_map_rows` places later rows on the
  first of the table's map pages with room.
* **ALTER COLUMN**, measured by resizing a Text(40) to Text(80) and
  retyping a Long to a Double on a table with ten rows. The engine adds a
  new column under the next number with the old column's name, place in
  the definition and position field (a fixed one placed past the highest
  fixed column, its variable-index field holding the variable count as
  with ADD COLUMN), re-encodes every row with the value copied or
  converted into it while the old column's bytes stay in the row as a
  phantom (its null-mask bit still set, its fixed slot or variable slot
  still occupied), then drops the old header; the definition, same
  length as before, is written once, the rows in place. How a row is
  re-encoded depends on whether its variable-column count still matches
  the definition's: when it does, the fixed block is copied byte for byte
  from the old row (a never-written fixed slot can inherit old text
  bytes) and the variable slots are carried by index, phantoms included;
  when variable columns have come or gone since the row was written, the
  engine keeps the old body verbatim, its variable table and count
  included, writes any fixed value over its slot in that body, and
  appends a fresh variable table with every slot empty behind it.
  `encode_row(template=...)` reproduces both. Since a replacement column
  keeps its place but not its number, `Table.columns` follows the
  definition's order rather than column-number order.
* **Renaming a column** (a Field's Name set through DAO) changes the name
  in the definition and nothing else in it (the header bytes are
  identical), rewrites the catalog row's property blob with the column's
  block renamed, changes every MSysRelationships row naming the column,
  and stamps the catalog row; indexes refer to columns by number and are
  untouched.
* **Renaming a table** (a TableDef's Name set through DAO) changes the
  catalog row's Name and DateUpdate and every MSysRelationships row that
  names the table as its object or referenced object; the definition,
  which does not carry the name, is untouched. The rename also showed
  that an index's distinct-key count grows on inserts only: an update
  that moves a row to a new key, distinct or not, never raises it, and
  the row counter below is what lowers it.
* **ALTER TABLE**, measured with ADD COLUMN (Long, then Text(30), then a
  Long again after a drop) and DROP COLUMN (the Long, then the original
  Text) on a table holding ten rows. An added column takes the next
  column number (the definition's maximum column count, which only ever
  grows), a fixed column the offset just past the highest fixed column
  present (a dropped column's slot is reused), a variable column the next
  variable index; its header and name are appended and the definition
  rewritten. A dropped column's header and name leave; the other columns
  keep their numbers, offsets and variable indexes, the variable count
  and the maximum column count stay. Rows are never rewritten: a row
  written before the change keeps its old column count and reads back
  with the new column null. Only the definition page and the catalog
  row's DateUpdate change. A Memo or OLE column added this way also gets
  two 69-byte usage-map rows on the table's map page and a map pair in
  the definition; dropped, its map bits are cleared, its pages released
  untouched and its two map rows killed, the pair leaving the definition,
  while the rows keep their stale value references.
* **Saved queries**, measured with DAO's `CreateQueryDef` on a plain
  select, a joined DISTINCT TOP GROUP BY HAVING ORDER BY DESC query, a
  parameter query and a DELETE. A query is a catalog object of type 5
  under the Tables container (Flags 0x20 for an action query), owner as
  usual, three table-style permission rows, and a property blob holding
  DAO's ODBCTimeout (Integer 60) and MaxRecords (Long 0), appended one at
  a time so the first blob sits inline in the row before the second
  moves it to a long-value page; an action query's flag comes with that
  last write, and so does the final DateUpdate. The row is stamped three
  times in all: at the insert, when the owner is set (that version's
  bytes outlive the final one) and with the last write, which is why
  `create_query` takes `created`, `owner_updated` and `updated`. Its definition is a set of MSysQueries rows (`ObjectId,
  Attribute, Order, Name1, Name2, Expression, Flag`), `Order` a four-byte
  big-endian sequence per attribute, inserted in this order: attribute 0
  (Flag 0), 255 (the end marker), 1 (the type: 5 delete, 4 update, 3
  append, 2 make-table; absent for a select), 2 per PARAMETERS entry
  (Name1 the name, Flag the DAO type), 6 per output column (Expression,
  Name1 the alias, Flag 0), 7 per join (Name1/Name2 the tables,
  Expression the condition, Flag 1 inner 2 left 3 right), 5 per source
  table (Name1, Name2 the alias), 8 WHERE, 9 per GROUP BY expression
  (Flag 0), 10 HAVING, 11 per ORDER BY expression (Name1 `d` for DESC),
  and last 3 when the select flags are not 0 (0x01 `*`, 0x02 DISTINCT,
  0x04 DISTINCTROW, 0x10 TOP with the count in Name1). Expressions are
  stored as written. Action queries measured too: UPDATE puts the type
  row (Flag 4) and the table before one attribute-6 row per SET item
  (Name2 the column, Expression the value); INSERT INTO ... SELECT puts
  the type row (Flag 3, Name1 the target table) first, then attribute-6
  rows with the target column in Name2 and the source expression, then
  the source tables; SELECT ... INTO puts the type row (Flag 2, Name1 the
  new table) after the columns; a UNION stores each member SELECT
  verbatim (the first keeping its trailing space) as attribute-5 rows
  named `X7YZ_____1`, `X7YZ_____2`, ... around a type row of 9, with a
  flags row of 3. The catalog Flags is DAO's QueryDefTypeEnum: 32 delete,
  48 update, 64 append, 80 make-table, 128 union. `QueryDefs.Delete`
  removes the rows, the catalog object (freeing its blob) and the three
  permission rows, deleting the rows in the order of the (ObjectId,
  Attribute, Order) index; that order showed in the page's compaction
  residue, and the dead slots it left, all recording one boundary,
  showed that a dead slot sitting exactly at a later compaction boundary
  moves with the block below it, even when that block is empty: dead
  slots parked at the lowest row's start follow the data start when
  that row goes too (a SQL DELETE over a page holding three overflow
  copies showed the case). `_queries.py` turns that subset of Jet SQL
  into rows and back.
* **Properties** live in the catalog row's `LvProp` long value as an
  `MR2` blob, measured on a blob DAO wrote (table Description, field
  Caption and Description) and confirmed on all 17 Access-authored blobs
  in the fixtures, which serialize back byte for byte. After the
  signature come blocks of `u32 length, u16 kind, body`: kind 0x80 is
  the name table (`u16 byte length` + UTF-16 per name), kind 0x00 the
  object's own properties, kind 0x01 one column's. A property block
  starts with `u16 name-part length` (6 when unnamed), `u16 0`, `u16
  name byte length` and the UTF-16 column name, then records of `u16
  length, u8 flags, u8 DAO type, u16 name index, u16 value length,
  value`. Values follow the DAO type (text as uncompressed UTF-16; Access
  writes ColumnWidth and ColumnOrder with four bytes under type 3, so
  integers decode by width). Each `Properties.Append` rewrites the whole
  blob: the new value is stored first and the old one freed, and the
  row's stamps are not touched. Names are indexed in first-use order and
  blocks keep their order across rewrites.
* **Dropping a relationship** (`DROP CONSTRAINT`) clears the foreign-key
  index's map bits, kills its map row and releases its pages untouched,
  removes the logical entry on each side (the remaining entries keep
  their numbers: `.rC` stays `.rC`), deletes the MSysRelationships rows,
  the catalog object and its three permission rows, and stamps both
  tables' DateUpdate. A rewritten definition is written up to its new
  length plus the eight reserved bytes the free word counts, zeroed; when
  it shrinks, the dropped entries' bytes beyond that stay readable. A new
  table's definition page is filled fresh.
* **Catalog rows are written twice**: CREATE TABLE inserts the row with
  DateUpdate equal to DateCreate and then updates it with the owner and
  the final DateUpdate; the first version's bytes stay below the slot
  table when the row moves, which is how the order shows.
* **Where a single-row long value lands**, measured with DAO on Memo
  columns in one session and across sessions. The engine first tries the
  LVAL page it last wrote a value to in this session; if the value does
  not fit there it takes the first page in the column's free-space map
  that has room; failing that it allocates a fresh page. (Page A held
  1080 bytes free and the last write had gone to page B: a 900-byte value
  went to B, and the next one, B full, to A.) A page is listed in the
  free-space map while more than 256 bytes are free: left with 256 it is
  unlisted, with 258 listed. An update stores the new value first and
  frees the old one afterwards, so the new value never lands in the
  hole it is about to open; a delete that leaves the page above the
  threshold lists it again. `PageStore.lval_cursor` holds the per-column
  cursor for the session.
* **Emptied pages are retired, truncation releases them untouched**,
  measured with DAO deletes on tables of 24 500-byte rows over four
  pages with three single-row Memo values. A filtered `DELETE` that
  takes the last row off a data page or an LVAL page retires it: type
  byte 0x09, every slot 0xD000 (dead, at the page end), free word
  `4096 - 14 - 2 * slots`, rows and owner left in place, the page
  released to the global map and dropped from the table's (or the
  column's) owned and free-space maps. The table's first data page is
  never retired, even emptied; an LVAL page always is, a column's only
  one included. A page that lost rows but keeps some rejoins the
  free-space map. `DELETE FROM t` with no filter takes another path:
  every data, long-value and non-root index page is released with its
  bytes untouched, the maps are emptied to all-zero rows, each index
  root becomes an empty leaf with a distinct count of 0, and the
  AutoNumber counter stays; `Table.truncate()` does the same. (A
  filtered delete that happens to remove every row keeps the distinct
  count: deletes never lower it.) Two things a delete leaves alone: the
  page that held a moved row's copy is only written back (emptied, it
  stays type 0x01, owned and unreleased) -- though when the row comes
  home through an update, the emptied copy page is retired -- and a
  home page that held only
  the 4-byte pointer is not re-listed, while a home page that lost a
  15-byte row is. That is also what DROP TABLE's deletion of the
  table's catalog rows shows, so there is no separate catalog path.
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
  allocator skips those, and a new instance starts clean. A transaction
  changes nothing here (DROP and CREATE inside one transaction still
  took fresh pages, and so did seven tables dropped and recreated one
  after another). Pages retired by a filtered delete and pages an
  unfiltered DELETE releases are quarantined the same way (the rows
  inserted next in the session took fresh pages). Two other kinds of
  release behave differently. The pages of a freed long-value chain come
  back at once when the chain existed before the database was opened
  (deleted and rewritten, a 10 KB value took its three pages back in
  order, and a new table took them for its definition and maps); a chain
  created within the session waits. The continuation page a definition
  rewrite replaces waits too, and the waiting pages come back into use
  together once five have piled up: over 31 CREATE INDEX statements the
  engine reused its released continuation pages for later roots, a map
  page and a continuation in one batch, lowest first, and kept growing
  the file before that. `PageStore.released`, `pending` and `allocated`
  carry the three distinctions.
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
  added, a map that holds pages grows its bitmap to cover two pages past
  the one being added, rounded up to four bytes -- the two spare pages
  being what a page taken from the end of the file leaves ahead of it.
  The rule shows where it differs from the obvious one: taking page 542
  grew a map to 72 bytes where covering page 542 alone wanted 68, and a
  map taking page 30 016 went to exactly 3756. Thirty growths across
  five scenarios, from 68 bytes to 3756, agree. An empty map is re-based
  to the page's 8-aligned start instead (a table whose only data page is
  542 gets start 536). The global map is extended one 8-byte step at a
  time when it lists no free page, the 64 new pages counting as free.
  Checked: 450 memo rows carrying a database from 121 to 573 pages leave
  every page but page 0 identical to the engine's own.
* **The reference form of a usage map**, which the engine writes once a
  map's row can no longer grow inside its page. The row becomes 69
  bytes: a kind byte of 1 and seventeen four-byte chunk pointers, each
  naming a page whose bytes from offset 4 are one 32 736-page bitmap
  (type 5, second byte 1). Seventeen chunks reach further than a
  database may grow, so the row never grows again. The conversion is
  exactly what the fit demands: a map row reached 3761 bytes with 3796
  of room, and the step that would have made it 3797 converted it
  instead. The global free map converts the same way, and each chunk it
  gains marks every page of that chunk the file has not reached as free
  -- so a 33 000-page database lists the whole of the second chunk,
  through page 65 471, as free. A bitmap page is taken from the end of
  the file like any other page. Checked byte for byte on a 130 MB
  database: 130 megabyte-sized long values, one column map converted
  with two chunks and the global map with two of its own, every page
  identical to the engine's but for the per-session stamps.
* **Jet DDL type words**, one CREATE TABLE per word, read back from the
  engine's own definition. `INTEGER`, `INT` and `INTEGER4` are the
  four-byte Long here, the two-byte one being `SHORT`, `SMALLINT` or
  `INTEGER2` -- the reverse of what the names suggest. `CHAR` and
  `CHARACTER` make a *fixed-width* Text column, which is the only Text
  column the engine marks fixed. `NOT NULL` changes nothing in the
  column header: a NOT NULL column's twenty-five bytes are a nullable
  column's. `WITH COMPRESSION` is refused by the Jet parser, so a column
  made through DDL never has Unicode compression, where a column made in
  the Access window does. `DECIMAL` and `NUMERIC` are refused too: a
  Decimal column reaches a file only through another provider. An
  unnamed `PRIMARY KEY` or `UNIQUE` gets a random `Index_...` name, so
  only a named `CONSTRAINT` can be reproduced.
* **A BigInt column brings three properties.** A table holding one gets
  `FCMinReadVer`, `FCMinWriteVer` and `FCMinDesignVer`, all Text
  `16.0.7124.1000`, in its catalog row's `LvProp`; no other type that
  can be created does this. They arrive one at a time, each rewriting
  the whole blob, so a table with fourteen columns and one BigInt leaves
  two dead long-value rows behind the live one.
* **Crosstab queries** (`TRANSFORM ... PIVOT`) are catalog Flags 16 with
  a type row carrying 6. Their column rows say what each column is for:
  flag 0 the value the TRANSFORM aggregates (its alias in Name1), flag 2
  each row heading from the SELECT list, flag 1 the pivot. The GROUP BY
  rows carry flag 2 and the engine adds the pivot to them with flag 1,
  whether or not the SQL named it. DAO writes them in this order: the
  parameters, the type row, the value column, the row headings, the
  joins, the tables, WHERE, GROUP BY, ORDER BY, then last of all the
  pivot -- its group row holding the expression alone and its column row
  the whole clause, `IN` list included. A TOP still writes the flags row
  at the very end. ACE refuses HAVING on a crosstab ("Syntax error in
  TRANSFORM statement"), so `_queries.py` refuses it too. A pass-through
  query is Flags 112 with a single type row: flag 8, Name1 the connect
  string, Expression the SQL sent to the server.  DAO reaches one by
  making an empty QueryDef, setting its Connect -- which writes that
  type row with no SQL in it -- and then its SQL, which deletes that row
  and writes another, so the page keeps a dead row between the two.
  `create_query(name, sql, connect=...)` walks the same path and lands
  on the same bytes.  Text passed to a pass-through is kept exactly as
  given: it is the server's SQL, and Jet does not parse it (a select
  query's SQL, by contrast, comes back reformatted with CRLF between
  clauses and a trailing semicolon).
* **A column header's bytes 7-8** count the variable columns declared
  before that column: for a variable column that is its own index, and
  for a fixed one it is how many came first (a Currency column after two
  Text columns carries 2). ALTER TABLE's rule is the same one, since a
  column added at the end follows every variable column the table has.
* **DROP INDEX** gives back everything the index held and nothing else:
  every page in its usage map is released with its bytes untouched (no
  retirement mark, no zeroing), its map row is deleted from the table's
  map page, the definition is rewritten without its three records (the
  12-byte header, the 52-byte index and the 28-byte logical entry plus
  its name, 102 bytes in this measurement), and the catalog row's
  DateUpdate is stamped. The real indexes after the hole move up, so
  every logical index pointing past it has its B-tree position lowered
  by one, while its own index number stays: relationships elsewhere name
  that number. A map row is never reused -- the next index appends a
  fresh row after the dead slot -- and a page freed by a drop comes back
  only in a later session, the object rule again (dropping and creating
  in one DAO session took a page freed by an earlier drop, not the one
  just released).
* **Building an index over rows.** The engine sorts the keys first, so a
  leaf fills at its end rather than in the middle, and it fills with
  *uncompressed* entries. When the next entry will not fit, the page is
  compressed: the bytes its entries share become the page prefix at 0x18
  and each entry after the first is stored without them, which on a
  400-row text index freed 1650 of 3616 bytes. The compressed page then
  keeps taking entries, but only ones carrying that prefix: the first
  key without it closes the page and opens the next. So the leaves break
  where the shared prefix changes, not where the page fills -- 111, 111,
  111 and 67 entries for keys `value number 1..400`, the three closed
  leaves each holding one leading digit with 1504 bytes still free, and
  the last leaf left uncompressed (prefix 0) because it never
  overflowed. A page's prefix is fixed once set. The bytes the
  uncompressed fill left beyond the compressed entries stay on the page,
  and when the root splits they travel with it: the left half is written
  over the root's own page image, and the new node keeps that image too.
* **Index row counters.** Each real index's 12-byte header carries two
  counts: the distinct keys at +4 (above) and, at +0, the rows the
  index holds. The second is written only when the index is built over
  existing rows -- CREATE INDEX or ADD CONSTRAINT on a populated table
  records the rows that got an entry (nulls left out by a unique
  ignore-nulls index) -- and an index that predates its rows keeps 0
  for good: inserts never touch it, however many rows they add. What
  lowers it is a row leaving the index: one off per deleted row, and
  one off per row an UPDATE writes through it, meaning every index
  whose columns appear in the SET list, even where the value does not
  change (`SET M = M` dropped the counter all the same) and however
  many statements it takes. A row whose key is null in an
  ignore-nulls index costs nothing, the count stops at zero rather
  than going negative, and the distinct count is capped at what is
  left after each step; an unfiltered DELETE zeroes both. Measured
  against DAO on a 30-row table with 17 indexes (an UPDATE of one
  indexed column and a filtered DELETE left the three indexes over
  that column at 0 and 0, the rest at 9 and 9) and on two six-row
  tables, byte for byte.
* **SQL over the engine.** `_sql.py` tokenizes and parses Jet SQL
  expressions (precedence climbing; three-valued logic; LIKE with `*`,
  `?`, `#` and character lists; `[parameters]`) and runs SELECT (joins
  by nested loops, GROUP BY over evaluated keys, aggregates, HAVING,
  ORDER BY with Null first, DISTINCT, TOP) and DML through the table
  writers, coercing each written value to its column's type. Checked by
  running the same statements through DAO on the same database: twelve
  SELECT shapes give identical names and values (including Currency
  arithmetic and Avg staying Currency, Boolean grouping True before
  False, `Expr1000` names), and two UPDATEs (one in place, one growing
  rows) plus a filtered DELETE leave every page identical; a SQL DELETE
  deletes row by row like the recordset path, and `DELETE FROM t` alone
  is the truncation path above. One Jet rule the operators do not share:
  `&` reads Null as an empty string and gives Null only when both sides
  are Null, while `+` propagates it (`Null & 'x'` is `'x'`, `Null + 'x'`
  is Null), which is why an engine UPDATE appending text to a null Memo
  writes the appended text where a Null-propagating reading writes
  nothing.

* **A column's rules are properties, not header bits.** Required,
  AllowZeroLength, DefaultValue, ValidationRule and ValidationText live
  in the definition's property blob, on the column's block; a table's own
  ValidationRule and ValidationText sit on the table's. `CREATE TABLE`
  with `NOT NULL` leaves the column header byte for byte the header of a
  nullable column and writes a Required property instead, one blob write
  per column in column order, each stamping the catalog row's DateUpdate
  (measured: a two-column NOT NULL table wrote the blob twice). These
  five take the engine's own DAO type and a flags byte of 1 -- Required
  and AllowZeroLength Boolean, DefaultValue and ValidationRule Memo,
  ValidationText Text -- where a property a client appends (Caption,
  Description) takes the client's type and no flags, and appending one
  leaves the stamps alone. `NOT NULL` stores Required as `01` where
  DAO's `Field.Required = True` stores `FF`; both read as True.
  `Field.ValidationRule` leaves a trailing NUL in a column's rule, which
  the table's rule does not have.
* **What the engine does with those rules.** A DefaultValue fills any
  column an INSERT does not name, through SQL as well as AddNew; its text
  is a Jet expression with an optional leading `=`, and a name it cannot
  resolve is its own text (`hello` defaults to the string, `a & b` to
  `ab`, `1+1` to 2, `=Date()` to today). Required refuses a null,
  omitted or explicit, on insert and on update: "You must enter a value
  in the 'T.B' field." A ValidationRule is checked on both; a rule that
  opens with an operator is about its own column, so `>0` reads as
  `[A]>0`; it refuses a row only when it comes out False, so a null
  passes; the message is the ValidationText when there is one, else
  "One or more values are prohibited by the validation rule '>0' set for
  'T.A'. Enter a value that the expression for this field can accept."
  AllowZeroLength is stored and read but not enforced: an empty string
  went into a text column with it off.
* **Which LVAL page a value goes on.** The free-space map lists a page
  while more than 256 bytes are free, so every listed page has room for
  a value of 256 bytes or fewer and the engine takes the first of them.
  A larger value needs the page checked and the engine checks one: the
  last page the map lists. When that page cannot take the value it
  starts another rather than looking further back. Measured on one file
  with the first page holding 3827 free and the last 1676: values of 254
  and 256 bytes shared the first, 258 and 1546 went to the last, and
  1706 -- too big for the last -- went to a new page, never to the first.
  Filling the first page until it dropped out of the map sent the next
  small value to the last page, so the pages are not segregated by size.

* **Statement shapes, checked the same way.** `TOP n PERCENT` keeps that
  share of the rows rounded up (one percent of four rows is one); `ORDER
  BY 2` names the second output column; a comparison against `ALL`, `ANY`
  or `SOME` of a subquery reads every row or any one of them, and over no
  rows `ALL` holds where `ANY` does not; and when two sources hold the
  same column name the engine names every one of them for its table,
  `a.Id` and `b.Id`, where a single one keeps the plain name. One shape
  the engine refuses that this executor takes: an output alias in `ORDER
  BY`, which Jet reads as a parameter it has not been given.

* **Jet has no Boolean of its own in an expression.** Every computed
  truth value comes back as -1 or 0 -- a comparison, `Not`, `And`, `Or`,
  `Is Null`, the literal `True`, `IsNumeric`, `CBool` -- and so does
  every aggregate, the Boolean column included: over True, False, True,
  `Max` gave 0, `Min` and `First` -1 and `Sum` -2. A Boolean column
  selected on its own is the exception: it keeps its type, and it sorts
  True before False, which is -1 before 0.
* **The expression functions**, each measured against DAO over four rows
  holding a null, an empty string, a negative number, a leap day,
  midnight and a second before it: the text, maths, conversion and date
  functions listed in `_sql.py`, plus `Format` and `Partition` in
  `_format.py`. Three of them are not what a reading of the
  documentation would give: `DateDiff` counts boundaries crossed rather
  than whole units (December to January is one month, 23:00 to 01:00 one
  day), `Str` drops the leading zero of a value under one (` .125`), and
  `Format` of Null is the empty string rather than Null. `Format`
  covers the named formats (General/Long/Medium/Short Date and Time,
  General Number, Currency, Fixed, Standard, Percent, Scientific,
  Yes/No, True/False, On/Off), custom date patterns down to `q`, `y`,
  `w` and `ww`, custom number patterns with `0`, `#`, thousands, `%` and
  up to four `;` sections, and the `@`, `&`, `>` and `<` text patterns.

* **A linked table is a catalog row and nothing else.** Type 6 under the
  Tables container, with the file in `Database`, the table's name over
  there in `ForeignName`, `Connect` holding whatever prefix is left once
  the leading `;` is gone (empty for a link to another Access file,
  `Text;` for a folder of text files), and the linked flag 0x200000 --
  0xA00000 when the source is not another Access file. There is no
  definition page, no data and no index: the id is the next one up from
  the lowest in use, the ids the engine gives objects that have no
  definition page of their own. It gets the same three permission rows a
  table gets, and the values that say where the rows are arrive on the
  second write, with the owner, so the catalog row takes its page before
  the long values take theirs. Checked against `TableDefs.Append` for a
  Jet link and a text link, and `TableDefs.Delete` for the removal, byte
  for byte. That delete showed one thing a filtered DELETE does not: the
  page the catalog row was alone on stays alive and owned, where a
  filtered delete would have retired it.

* **Jet 4 `.mdb` files** take the writers as `.accdb` files do: a table
  created in a database DAO made with `CreateDatabase(..., dbVersion40)`
  lands on the same pages, with the same definition, rows and index, and
  the engine reads it back. The one difference is the permission rows.
  A new object inherits from its container, and a database whose
  workgroup differs from the shipped template's has different SIDs and
  access masks there: a DAO-made `.mdb` gave its new table two rows
  where the template's container gives three. Reproducing that means
  the workgroup security model, which stays out of scope. Jet 3 could
  not be measured at all: this ACE runtime refuses to create one
  ("could not find installable ISAM").

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
| 3b | large files: usage maps growing past 512 pages | done. Inline maps grow and re-base as the engine does, and a map whose row outgrows its page becomes the reference form, the global free map included; a 130 MB database with 130 long values is byte-identical to the engine's (live gate) |
| 5 | write schema: create/drop table, create index, catalog rows | done: `create_table`, `create_index`, `drop_index`, `drop_table`; byte-identical to the engine's CREATE TABLE, CREATE INDEX and DROP TABLE on every page but page 0; the engine inserts into, reads and compacts a table pyOpenVBA created; definitions over one page (up to the 255-column limit) are chained and rewritten as the engine does, byte-identical; `add_column` / `drop_column` match ALTER TABLE ADD COLUMN / DROP COLUMN byte for byte (live gate). `rename_table`, `rename_column` and `alter_column` match DAO renames and ALTER COLUMN; a table's map rows spill onto a second map page as the engine's do. `drop_index` matches DROP INDEX byte for byte, as does an index built over four hundred rows whose B-tree spans four leaves. A column's Required, DefaultValue, ValidationRule and ValidationText are written as the engine writes them, `CREATE TABLE ... NOT NULL` included (live gate), and the writers apply and enforce them on every row. Linked tables are read with `db.links()` and written with `db.link_table(...)` / `db.drop_link(...)`, byte-identical to DAO's TableDefs.Append and Delete for a Jet link and a text link (live gate). Not yet: navigation-pane rows (the Access layer adds those itself) |
| 6 | VBA project through the writer: module create/rename/delete | done, as research in `docs/research/access_write` rather than shipped code. Create, rename, delete and replacing a module's source all run in Access (live gate, `RUN_LIVE_ACCESS_VBA=1`): a created standard or class module is listed by `AllModules`, reads back, compiles and executes, a second create works with Access never opening the file in between, and the project still takes edits. The lever is that `_VBA_PROJECT` is a performance cache -- write a `Version` VBA does not recognise and it recompiles the project from the module streams, so a new module needs no p-code and none of the compiled tables have to be generated. One of them could not have been: the 32-slot table before the module table is runtime state, and adding the same module to the same file twice gives two different tables. Writing source this way also reaches every form the p-code compiler refuses (a new procedure, `Const`, arrays, `Static`, fixed-length strings, `Set x = New Class`), at the cost of a recompile on the next open. Two allocations are computed, not chosen: a module's storage folder is named `chr(0x30 + rows under Modules - 1)` and `MSysObjects` ids step by 4 |
| 7 | queries (`MSysQueries` to SQL and back), relationships, properties | done. Relationships: `create_relationship` / `drop_relationship` / `relationships()`, byte-identical to the engine's ADD CONSTRAINT ... FOREIGN KEY for a first and a second relationship on one parent and to DROP CONSTRAINT (live gate). Properties done: `table.properties()`, `column_properties()`, `set_properties()`, `db.database_properties()`; DAO's three property appends reproduced byte for byte (live gate). Queries done for SELECT, PARAMETERS, DELETE, UPDATE, INSERT INTO ... SELECT, SELECT ... INTO, UNION and crosstabs (`TRANSFORM ... PIVOT`, with an `IN` list, TOP, a join or a parameter): `db.queries()`, `db.query()`, `db.create_query(name, sql)`, `db.drop_query(name)`; thirteen CreateQueryDef calls and a QueryDefs.Delete reproduced byte for byte (live gate). Pass-through queries are written too, by the create-then-convert route DAO takes. Subqueries save too: in a WHERE, as a value in the select list, and as a table of their own, where the engine puts the bracketed SELECT in the row's expression and only the alias in Name2 |
| 8 | forms, reports, macros: the binary object formats nobody has published | |
| 9 | SQL executor over the engine | done: `db.execute(sql)` runs Jet DDL (CREATE TABLE with named keys and inline or table constraints, CREATE [UNIQUE] INDEX, DROP TABLE, DROP INDEX, ALTER TABLE ADD / ALTER / DROP COLUMN and ADD / DROP CONSTRAINT), byte-identical to DAO's Execute on the same statements (live gate), and SELECT (column list or `*`, INNER / LEFT / RIGHT JOIN, WHERE, GROUP BY with Count / Sum / Avg / Min / Max, HAVING, ORDER BY, DISTINCT, TOP; comparison, logical, arithmetic and `&` operators, LIKE, IN, BETWEEN, IS NULL, `[parameters]`, the common string, numeric and date functions), INSERT ... VALUES, INSERT ... SELECT, UPDATE and DELETE through the row writers, coercing values to the column type. Eleven SELECT shapes answer exactly as DAO does on the same database and an UPDATE plus a DELETE write the same bytes DAO's Execute writes (live gate). Subqueries run too: `IN`, `NOT IN`, `EXISTS`, a scalar subquery in any expression, all of them correlated when they name the outer query, plus a bracketed SELECT or a saved query as a FROM source, and `UNION`/`UNION ALL` folded left to right. A crosstab runs as well as saves: `TRANSFORM ... PIVOT` groups by its row headings, makes one column per pivot value (`<>` for Null, an `IN` list fixing the columns and their order) and hands the rows back sorted by their headings, which is how the engine hands them back. The operators now include Jet's `Mod`, `\` and `^`, in VBA's order, and the function list reaches every one a Jet expression can name (text, maths, conversion, dates, `Format`, `Partition`, `Switch`, `Choose`, the `Is` tests) plus the `First`, `Last`, `StDev`, `StDevP`, `Var` and `VarP` aggregates, each answering what DAO answers on the same rows (live gate). `with db.transaction():` rolls back to the exact bytes on an exception, and a committed one leaves what DAO's BeginTrans/CommitTrans leaves (live gate) |
