# Changelog

All notable changes to pyOpenVBA are documented here. This project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **Every control type pyOpenVBA can read, it can now write.**
  `add_control` wrote a Label or a TextBox; it now writes all eighteen --
  CommandButton, ToggleButton, OptionButton, CheckBox, OptionGroup,
  ListBox, ComboBox, Rectangle, Line, Image, PageBreak,
  BoundObjectFrame, ObjectFrame, Subform, Tab and Page as well. Every
  slot's id, code, value type and width was read back from a control
  Access itself made, and each type gets only the slots it has: a page
  break carries a top and nothing else, a tab control no left or top at
  all, an image no overlap flags, a combo box its GUID ahead of its
  name. A button is written with the padding Access gives it, a list or
  combo box with `Table/Query` as its row source type, and anything that
  takes the focus with the next tab index.

- **A control can hold controls.** `add_control(..., parent="Tabs")` puts
  a page on a tab control, which is written as a group of its own right
  after it. Reading a design is now a tree walk rather than a flat scan,
  since a section's count is of its own controls and not of everything
  beneath them. A page must have a parent tab and nothing else may have
  one, both of which are refused rather than written.

  The live gate puts one of every writable type on a single form and has
  Access name them all back, and builds a tab control with two pages
  between a text box and a button -- Access reports the tab as each
  page's parent and the form as the parent of the two beside it.

- **Jet 3 (Access 97) databases read.** `AccessDatabase` opens a 2 KiB
  page `.mdb` and reads its catalog, tables, rows and long values the
  same way it reads an `.accdb`. Everything that moved between the two
  versions lives in one `Layout` record, so there is a single parser, a
  single row splitter and a single set of value decoders rather than a
  second implementation: the page halves, a row counts its columns and
  its variable-column offsets in bytes rather than words, text is stored
  in the code page page 0 names rather than UTF-16, and the definition,
  column, index and name headers all shrink and move.

  Every offset was measured against files the Jet engine wrote, and the
  parser checks that what it consumed equals the length the page
  declares. The live gate has DAO 3.6 -- which still creates Access 97
  files though Access dropped the format in 2013 -- build a database with
  every Jet 3 column type, a 4000-character memo, code page text, a
  deleted row and four hundred rows over many pages; pyOpenVBA reads the
  same file with no COM involved and the two agree cell for cell.

  Writing a Jet 3 file is refused rather than attempted, at the page
  store itself, so nothing can put a Jet 4 shape into one.

- **Compaction.** `db.compact()` gives back the free pages at the end of
  the file and says how many went. That is what a dropped table or a
  large delete leaves behind, and it is the part of compaction that can
  be done without moving a page: nothing moves and no object is
  rewritten, so nothing can be lost. Free pages in the middle keep their
  place, and the first two pages are never dropped.

  A 954 KB database with a dropped 3000-row table came back at 300 KB
  with its other table intact. The live gate has the engine itself read
  a compacted file row for row and then run its own Compact and Repair
  over it, which rebuilds every page it kept.

  This is not Access's Compact and Repair, which also renumbers pages,
  resets AutoNumber counters and drops deleted rows from the middle of a
  table.

- **The domain functions.** `DLookup`, `DCount`, `DSum`, `DAvg`, `DMin`,
  `DMax`, `DFirst`, `DLast`, `DStDev`, `DStDevP`, `DVar` and `DVarP` run
  in `db.execute(...)`. Each is a query over another table, so each runs
  as one: `DLookup("Total", "Orders", "Id = 1")` is
  `SELECT Total FROM Orders WHERE Id = 1`. A criteria can name a column
  of the row it is evaluated in, which is what makes them worth having.

  These are Access's own rather than the database engine's, so DAO cannot
  answer for them: the gate compares every one against Access's `Eval` on
  the same database, and all seventeen agree.

- **The password guard reaches Access.** `db.vba_is_protected()` reads
  the `PROJECT` stream's `DPB` record, and `db.save()` refuses to write a
  VBA change into a protected project unless it is told
  `allow_protected=True` -- which is what the other hosts already do on
  their own `save`. A change that is not to the VBA project saves as
  before.

- **The project's references.** `db.references()` reads the libraries a
  VBA project points at, `db.add_reference(name, guid, major, minor,
  path=, description=)` adds one and `db.drop_reference(name)` removes
  it. A live gate adds the Scripting Runtime, writes a module that uses
  `Scripting.Dictionary`, and has Access compile and run it.

  Access keeps them only in the dir stream, three records each --
  `REFERENCEORIGINAL`, its Unicode twin, and `REFERENCEREGISTERED`
  holding `*\G{GUID}#major.minor#lcid#path#description`. `PROJECT`
  carries no `Reference=` line, and the two libraries every project has,
  VBA itself and Access, are not in the file at all. The version is
  written in **hex**: DAO 12.0 is stored as `c.0`.

- **Forms and reports.** `db.forms()`, `db.reports()`, `db.form(name)`,
  `db.report(name)`, `db.create_form(name)`, `db.create_report(name)`,
  `db.delete_form(name)` and `db.delete_report(name)`. A live gate opens
  what this writes in Access's own designer, and a created form also
  opens in form view.

  A design is a stream of property records, `<u32 id><u16 code><u32
  type><u32 width><u32 length><value>`, with the ids ascending inside one
  object. Three ids are not properties but markers that open the next
  object: `0xFE` a section, `0xFD` the next object at the same level, and
  `0xFF` a control, which carries a second `u16` naming its type. Every
  design measured -- an empty form, a form with a label and a text box,
  and a report with its three sections -- rebuilds byte for byte.

  `db.add_control(design, type, name, ...)` puts a Label or a TextBox on
  one, and Access reads back every measurement it was given. Thirty-three
  property codes are named, worked out by exporting a design with
  `SaveAsText` -- which writes the same records with their names, in the
  same order -- and walking the two together.

  A control belongs to a section and is written immediately after it, and
  its marker depends on **how many controls that section holds**: one is
  a single child, `0xFE`; two or more open a group, `0xFF` then `0xFD`.
  Access writes both in one report -- a page header holding one control
  and a detail band holding two -- and refuses each in the other's place,
  with "saved in an invalid format" for one of them.

  `db.set_design_code(name, code)` puts code behind one, creating the
  module when the design has none, and Access runs it. A document module
  belongs to its design rather than to `Modules`: no storage folder, no
  catalog row, and a `DocClass=` line in `PROJECT` where a class module
  gets `Class=`. Without that line Access loads the module and the form
  still does not answer to it. The design's `TypeInfo` and the module's
  `VB_Base` share a CLSID, and a byte in the design folder's `PropData`
  records that it has a module at all.

  Creating cuts from a captured empty design with a GUID of its own
  patched in, since the catalog row repeats the one the design carries.

- **Macros.** `db.macros()`, `db.macro(name)`, `db.create_macro(name,
  actions)` and `db.delete_macro(name)`, with `MacroAction(name,
  arguments)` for each step. A live gate creates a macro, has Access run
  it with `DoCmd.RunMacro`, and reads back the value it set.

  Access stores a macro as a binary blob, not as the XML its designer
  shows: a 32-byte header, a length-prefixed `"33"`, then one record per
  action carrying the action id, the row number, fourteen `u16` slots
  holding byte offsets into a string area, and the strings themselves.
  Arguments occupy slots from 4 upward and an empty one takes no slot,
  so a gap in the middle reads back as an empty string. Every blob in
  the fixture rebuilds byte for byte.

  Twenty-four action ids, measured by loading one macro each through
  `LoadFromText` and pairing storage folders with `MSysObjects` rows in
  id order. A macro's object id steps by **one** where a module's steps
  by four, and a macro gets no navigation-pane group row where a module
  does -- so the step is what an object reserves for itself rather than
  a global stride.

- **Attachments and multi-valued columns.** `db.complex_columns()` finds
  them, `table.attachments(column, key)` and
  `table.multi_values(column, key)` read them, `set_attachments` /
  `set_multi_values` write them, and
  `table.add_complex_column(name, kind)` creates one. An inserted row is
  given its complex id automatically. A live gate builds a database, a
  table and both kinds of column from nothing and has the ACE engine read
  the bytes back.

  Creating one costs four things beyond the column itself, and three were
  invisible until the engine refused the result. The flat table keeps its
  two Long bookkeeping columns **among the variable columns**, where a
  Long would normally sit in the fixed block, with no collation and no
  fixed bit -- `ColumnSpec(..., variable=True)` now says so. Its catalog
  row carries `Flags` `0x800A0000`. And the table that *has* the column
  carries `0x40000`, which no other table does; without it DAO opens the
  child recordset and finds no fields in it.

  A `Complex` column keeps only a Long in the row -- an id shared by
  every complex column in that row, handed out from a counter at 0x1C of
  the table definition and never reused. The values live one per row in
  `f_<GUID>_<Column>`, joined on that Long, and `MSysComplexColumns`
  names the pairing. `FileData` is a container of its own: a flag and an
  inflated size, then either a zlib stream or the bytes as they are, and
  inside that a header carrying the file's extension.

  Access decides whether to compress **by file type**, measured across 45
  extensions: it leaves `docx`, `gif`, `jpeg`, `jpg`, `png`, `pptx`,
  `xlsx` and `zip` alone and compresses everything else, including `7z`
  and `mp4`. Those eight are written byte-identically; a compressed one
  is not, because Access's deflate is not zlib's and no combination of
  level, memLevel, strategy or window size reproduces a stream it wrote.

  Two corrections to the table definition came out of this: the field at
  0x18 is a constant (1 in ACE, -1 in Jet 4) and not a counter, and the
  complex-id counter is the u32 at 0x1C. A complex column is flagged
  AutoNumber like any other, so the row writer had been handing it a
  value from the ordinary AutoNumber counter, which gave two columns in
  one row two different ids.

- **Access VBA is writable.** `AccessDatabase` gained `modules()`,
  `module(name)`, `create_module(name, code, kind="module"|"class")`,
  `set_module_source(name, code)`, `rename_module(old, new)` and
  `delete_module(name)`. Standard and class modules both, with whatever
  source you give them.

  The route is not a p-code writer. `_VBA_PROJECT` is [MS-OVBA]'s
  PerformanceCache, and its `Version` field names the build of VBA that
  compiled it; writing a version the host does not recognise makes VBA
  discard the cache and compile the project from the module streams, the
  same thing Access's `/decompile` does. A module's stream is therefore
  the compressed source alone with MODULEOFFSET at zero, and none of the
  compiled tables have to be generated -- one of which could not have
  been, since the 32-slot table ahead of the module table is runtime
  state and adding the same module to the same database twice gives two
  different tables. The cost is a recompile on the next open: the code
  has to compile, and the cache stops matching what Access wrote until
  Access rewrites it.

  Three of the rules are invisible from the file and were caught only by
  asking Access: a module's storage folder is named from the rows its
  container already holds and Access will not look under any other name,
  `MSysObjects` ids step by four rather than one, and a delete has to
  free the folder or the next create picks a name Access rejects. A live
  gate (`RUN_LIVE_ACCESS_VBA=1`) runs the result in Access and compares
  the value the code returns, class instantiation included.

- **More of the statement.** `TOP n PERCENT`, `ORDER BY <position>`, and
  a comparison against `ALL`, `ANY` or `SOME` of a subquery. A column
  name two sources share is now qualified in the output the way the
  engine qualifies it, and a crosstab pivoting on a comparison names its
  columns -1 and 0.

- **The rest of the Jet expression functions.** Replace, Space, String,
  StrComp, StrReverse, Asc, Chr, Sgn, Sqr, Exp, Log, Fix, Val, Str, Hex,
  Oct, the CBool/CByte/CCur/CSng/CDate family, DateAdd, DateDiff,
  DatePart, DateSerial, TimeSerial, Weekday, WeekdayName, MonthName,
  DateValue, TimeValue, Time, IsNull, IsNumeric, IsDate, Switch, Choose,
  Format and Partition, with the First, Last, StDev, StDevP, Var and
  VarP aggregates. A gate runs 131 expressions through DAO and through
  the executor over the same four rows and compares them cell by cell.

- **Truth values as Jet writes them.** A computed comparison, logical
  operator or yes-or-no function now answers -1 or 0, and so does an
  aggregate over a Boolean column, which is what the engine answers. A
  Boolean column selected on its own still reads as a Boolean.

- **Linked tables.** `db.links()` and `db.link(name)` read the tables a
  database only points at; `db.link_table(name, database, source,
  connect=...)` writes one and `db.drop_link(name)` forgets it, byte for
  byte as DAO's `TableDefs.Append` and `Delete` do, for a link to another
  Access file and to a folder of text files. Following a link is left to
  the caller: the path comes out of the database, so opening it is not
  something the library does on its own.

- **A column's own rules.** Required, DefaultValue, ValidationRule and
  ValidationText are properties on the column, not bits in its header,
  and the writers now put them there the way the engine does: one blob
  write per column, the engine's DAO type and flags, and the catalog
  stamp that goes with it. `CREATE TABLE ... NOT NULL` and
  `DEFAULT <expr>` set them, and `ColumnSpec` carries them.

- **The engine's rules applied to every row.** A column an INSERT does
  not name takes its DefaultValue, evaluated as Jet evaluates it; a null
  in a Required column and a value against a ValidationRule, the
  column's or the table's, are refused with the engine's own message, on
  insert and on update. A live gate has the engine reject the same four
  statements.

- **Which LVAL page a long value lands on**, measured: a value of 256
  bytes or fewer goes on the first page the free-space map lists, a
  larger one on the last, and when that page cannot take it, on a new
  page rather than an earlier one.

- **A gate for a database built without the engine at all**: every
  column type, keyed and unique indexes, a foreign key, long values,
  four saved queries and table and column properties, all written by
  pyOpenVBA, then read back field for field by DAO and compacted by it,
  which fails on any structure the engine cannot follow.

- **Subqueries in saved queries.** A subquery in a WHERE, one used as a
  value, and a bracketed SELECT standing where a table does all save as
  the engine saves them; a derived table's text goes in its row's
  expression with only the alias naming it.

- **Pass-through queries.** `db.create_query(name, sql, connect=...)`
  saves one, byte for byte as DAO does, dead row and all; the SQL is
  kept exactly as given because the server parses it, not Jet.
  `query.connect` reads the connect string back.

- **Usage maps past their row.** A map whose bitmap can no longer grow
  inside its page becomes the engine's reference form: a row of chunk
  pointers, each naming a page holding one 32 736-page bitmap. The
  global free map converts the same way and marks each new chunk's
  unreached pages free. A 130 MB database of long values is now
  byte-identical to the engine's, both maps converted.

- **`with db.transaction():`** groups writes so they all land or none
  do; an exception puts the pages and the session's state back exactly
  as they were. The engine writes the same bytes either way, and a live
  gate checks that against DAO running the same statements inside
  BeginTrans/CommitTrans.

- **Crosstabs run, not just save.** `db.execute("TRANSFORM ... PIVOT
  ...")` returns the pivoted rows: one column per pivot value, `<>` for
  a Null heading, an `IN` list fixing the columns and their order, an
  aggregate allowed among the row headings, and the rows sorted by those
  headings as the engine sorts them. Two of these answer exactly as DAO
  does (live gate).
- **Jet's `Mod`, `\` and `^` operators**, in VBA's order of
  precedence, with both sides rounded half to even and the division
  truncating toward zero.

- **Subqueries, unions and saved queries in `db.execute`.** `IN`,
  `NOT IN`, `EXISTS`, `NOT EXISTS` and a scalar subquery work in any
  expression, correlated when they name the outer query; a bracketed
  SELECT or a saved query's name can stand where a table does in FROM;
  and `UNION` / `UNION ALL` fold left to right with a trailing ORDER BY
  over the result. Eight of these shapes answer exactly as DAO does on
  the same database (live gate).

- **DDL through `db.execute`.** CREATE TABLE (every Jet type word, named
  primary keys, inline and table constraints, foreign keys), CREATE
  [UNIQUE] INDEX with ASC/DESC columns and WITH IGNORE NULL, DROP TABLE,
  DROP INDEX, and ALTER TABLE ADD / ALTER / DROP COLUMN and ADD / DROP
  CONSTRAINT. Fourteen statements leave the same bytes DAO's Execute
  leaves for the same SQL (live gate). What the Jet parser refuses --
  `CHAR`, `DECIMAL`, `NUMERIC`, `WITH COMPRESSION` -- is refused here
  with the reason.
- **A table with a BigInt column carries the engine's version
  properties** (`FCMinReadVer`, `FCMinWriteVer`, `FCMinDesignVer`),
  written one at a time as the engine writes them.

- **Crosstab saved queries.** `db.create_query(name, "TRANSFORM ... PIVOT
  ...")` writes the rows DAO writes, byte for byte, with an `IN` list, a
  TOP, a join or a parameter; `db.query(name).sql` gives the statement
  back. HAVING is refused there, as the engine refuses it.

- **`db.drop_index(table, name)`**, byte-identical to the engine's DROP
  INDEX: the index's pages released with their bytes untouched, its
  usage-map row deleted, its records taken out of the definition with
  the indexes after it moved up, and the catalog row stamped. The
  primary key and an index a relationship rests on are refused, as the
  engine refuses them.

- **SQL executor.** `AccessDatabase.execute(sql, parameters)` runs Jet
  SQL against the engine in pure Python: SELECT with a column list or
  `*`, INNER / LEFT / RIGHT JOINs, WHERE, GROUP BY with Count, Sum, Avg,
  Min and Max, HAVING, ORDER BY, DISTINCT and TOP; the comparison,
  logical, arithmetic and `&` operators, LIKE with the engine's
  wildcards, IN, BETWEEN, IS NULL, `[parameters]` and the common string,
  numeric and date functions; INSERT ... VALUES, INSERT ... SELECT,
  UPDATE and DELETE through the row writers, with values coerced to the
  column type as the engine coerces them. Three-valued logic follows the
  engine. Eleven SELECT shapes answer exactly as DAO does on the same
  database, name for name and value for value, and an UPDATE plus a
  DELETE write the same bytes DAO's Execute writes (live gate).
- **Index row counters.** An index built over existing rows records how
  many rows it holds next to its distinct-key count. Every row that
  leaves the index takes one off, whether deleted or written by an
  UPDATE that names one of the index's columns, and the distinct count
  is capped at what is left; a null key in an ignore-nulls index costs
  nothing, the count stops at zero, and an unfiltered DELETE zeroes
  both. Inserts leave the row counter alone, as the engine does.

### Fixed

- **A section holding three or more controls lost all but two.** The word
  an `0xFF` marker carries is the number of objects in the group it
  opens, and it was written as a constant 2 -- right by coincidence for
  the two-control case that was measured, wrong for every other. Access
  does not refuse a design that gets it wrong: it opens the form and
  shows only as many controls as the number claims, which is why nothing
  caught it. A form Access built with eleven controls carries `0xFF 11`
  twice, once over its prototypes and once over its controls, which is
  what settled it.
- An inline usage map now grows to cover two pages past the one being
  added, rounded to four bytes, which is what the engine writes.
  Rounding to eight was right only by accident on small files and gave
  the wrong bitmap on every map above a few hundred bytes.
- A BigInt column is no longer marked fixed: the engine keeps it among
  the variable columns, like a GUID, and a table created with one
  differed from the engine's by that flag.
- An index key over a Currency column now takes a float, rounding it to
  four places exactly as the row writer does. Inserting a float into a
  uniquely indexed Currency column raised instead.
- A fixed column declared after a variable one now carries the count of
  variable columns before it in its header, as the engine writes it. A
  table whose Currency column followed two Text columns differed from
  the engine's by that byte.
- The SQL rebuilt from a saved query with a join now names the joined
  table, not just its alias.
- An index built over existing rows now matches the engine's B-tree.
  Entries are added in key order, a leaf keeps the prefix it was
  compressed with instead of shrinking it (so a key without that prefix
  starts the next leaf), and a splitting root hands its page image to
  the left half. A four-leaf index over four hundred rows differed from
  the engine's on every leaf before; it is byte-identical now.
- `&` in a SQL expression now reads Null as an empty string, as Jet
  does, giving Null only when both sides are Null; `+` still propagates
  it. Appending text to a null Memo through UPDATE wrote nothing
  before, where the engine writes the appended text.
- Dead slots parked at a page's lowest row now follow the data start
  when that row is deleted too, as the engine's do; before, they kept
  the old offset and a page that lost rows through an overflow copy
  differed from the engine's by those slot entries.

- **A Jet 4 / ACE storage engine, read layer** (`pyopenvba.access`,
  in progress; not yet exported from the package root).
  `AccessDatabase(path)` opens an `.accdb` or Jet 4 `.mdb`, lists the
  catalog and reads any table's rows as Python values: every column type
  including Currency, Decimal, GUID, BigInt, Memo and OLE long values,
  Unicode-compressed text, overflow rows and definitions that span pages.
  Checked against the ACE engine itself (`RUN_LIVE_ACCESS=1`, DAO driven
  from PowerShell as a test-time oracle) field for field, and against
  every Access-authored fixture table for table. The plan and the format
  facts established so far are in `docs/access_engine.md`.
- **Indexes read.** `table.indexes`, `table.index(name)` and
  `table.primary_key` expose each B-tree; `index.entries()` walks the
  leaves in key order and decodes every key type (`index.rows()` yields
  the rows in that order). Verified against seventeen ACE-written
  indexes covering every indexable column type, descending order,
  two-column keys and unique ignore-nulls, entry for entry.
- **Text collation reproduced.** The sort keys the engine writes for
  text -- case-blind primaries, diacritic weights, kana marks, recorded
  ignorables -- are generated from the engine's own output for every
  BMP code point (`scripts/generate_access_collation.py`) and
  re-encoded exactly, 63 632 of 63 632 strings. With that,
  `encode_key` produces the stored bytes for any index key from Python
  values, the inverse of the decoder, checked against every entry of
  seventeen live indexes.
- **Rows can be written.** `table.insert_row(values)`,
  `table.update_row(row_id, changes)` and `table.delete_row(row_id)`
  edit a table the way the engine does: rows laid down and compacted on
  the data page, AutoNumber and row counters maintained, every index
  updated with pages compressed when full and split when needed, new
  pages taken from the global usage map and registered with the table's
  maps. `AccessDatabase.save()` writes the file. Verified live: the ACE
  engine reads back every row pyOpenVBA wrote across all scalar column
  types, keeps working with the table, and compacts the database; a
  single insert and delete are byte-identical to the engine's own.
- **Memo and OLE values are written** in the storage kind the engine
  would choose -- inline, one row on a shared long-value page, or a
  chain of pages -- and freed when replaced or deleted; a row that
  outgrows its page moves behind an overflow pointer and comes back
  when it fits. A duplicate key in a unique index is refused. All of it
  read back by the engine live; single-page and chained memo inserts
  byte-identical to its own.
- **Tables can be created and dropped.** `db.create_table(name,
  columns, indexes)` with `ColumnSpec` / `IndexSpec` (every column type,
  AutoNumber, primary key, unique, descending and multi-column indexes),
  `db.create_index(table, spec)` and `db.drop_table(name)` write the
  definition page, the usage-map page, the index roots and the catalog
  rows exactly as the engine does: pyOpenVBA's CREATE TABLE, CREATE
  INDEX and DROP TABLE leave every page but page 0 identical to the
  engine's own, and the engine inserts into, reads and compacts a table
  pyOpenVBA created. `pyopenvba.access` exports `AccessDatabase`,
  `Table`, `Index`, `RowId`, `ColumnSpec`, `IndexSpec`.
- **Table definitions longer than one page** are written the way the
  engine writes them: `ceil(length / 4088)` pages, continuation pages
  allocated after the index roots and chained in reverse, the free word
  `4088 * pages - length` on the last page; CREATE INDEX rewrites onto a
  fresh chain and releases the old pages; DROP TABLE marks only the
  first page. Tables of up to 255 columns with long names now round-trip
  byte for byte against the engine (live gate).
- **Relationships.** `db.create_relationship(name, table, columns,
  referenced_table, referenced_columns, cascade_updates=, cascade_deletes=)`
  writes a foreign key the way `ALTER TABLE ... ADD CONSTRAINT ... FOREIGN
  KEY` does: the index on the referencing columns, the paired logical
  entries on both tables (`.rB`, `.rC`, ... on the referenced side), the
  MSysRelationships rows, the type-8 catalog object with its permission
  rows, and both tables' stamps; `db.relationships()` reads them back and
  `db.drop_relationship(name)` removes one as `DROP CONSTRAINT` does,
  leaving the shrunken definitions' old bytes in place as the engine
  does. Byte-identical to the engine for two relationships on one parent
  and the drop of the first (live gate). CREATE TABLE's catalog row now carries DateCreate as its first
  DateUpdate, as the engine's does, which the stale bytes under the slot
  table revealed.
- **A second usage-map page.** When a table's map page is full (57
  rows), the next index or Memo column gets its map row on a fresh map
  page, as the engine does; a table with 32 indexes and 12 Memo columns
  is byte-identical to the engine's (live gate). The same run pinned down
  when the engine reuses pages it released earlier in a session:
  dropped, truncated or retired pages never, a freed value chain's pages
  at once when they predate the session, and a definition rewrite's old
  continuation pages in a batch once five are waiting.
- `db.rename_table(name, new_name)` renames a table as DAO does: the
  catalog row and every relationship row naming it (live gate).
  `table.rename_column(name, new_name)` renames a column the same way:
  the definition's name, the column's property block, the relationship
  rows naming it, the catalog stamp (live gate).
  `table.alter_column(name, ColumnSpec(...))` retypes or resizes a column
  as ALTER COLUMN does: a replacement column, every row re-encoded with
  the value converted, the old header dropped (live gate).
- **Columns can be added and dropped.** `table.add_column(ColumnSpec(...))`
  and `table.drop_column(name)` do what `ALTER TABLE ... ADD COLUMN` and
  `DROP COLUMN` do, byte for byte in the live gate: the definition
  rewritten, rows left as they are (old rows read the new column as
  null), the catalog row stamped. Memo and OLE columns bring or give back
  their usage maps and pages; indexed columns are refused. Found on the
  way: a definition rewrite
  after inserts in the same session wrote back the indexes' distinct-key
  counts as they were when the definition was read; the live counts now
  go in.
- **Saved queries.** `db.queries()` and `db.query(name)` read
  MSysQueries into `SavedQuery` objects whose `.sql` spells the Jet SQL
  back in DAO's own layout; `db.create_query(name, sql)` saves a SELECT
  (DISTINCT, DISTINCTROW, TOP, aliases, INNER/LEFT/RIGHT JOIN, WHERE,
  GROUP BY, HAVING, ORDER BY), PARAMETERS, DELETE, UPDATE, INSERT INTO
  ... SELECT, SELECT ... INTO or UNION query as `CreateQueryDef` does:
  the rows, the type-5 catalog object with its properties, permissions
  and DAO's query type in Flags; `db.drop_query(name)` removes one as
  `QueryDefs.Delete` does. Eight query shapes and a deletion
  byte-identical to DAO in the live gate.
- **Properties.** `table.properties()`, `table.column_properties(name)`
  and `db.database_properties()` decode the `MR2` blob in a catalog row's
  LvProp (Description, Caption, Format, ColumnWidth, the database's own
  settings, ...); `table.set_properties({...}, column=...)` rewrites it
  the way DAO's `Properties.Append` does, one call per append, byte for
  byte in the live gate. All 17 Access-authored blobs in the fixtures
  serialize back identically.
- A freed long-value chain's pages are reusable at once within the
  session when the chain predates the session, unlike pages a DROP TABLE
  releases or pages the session itself allocated (measured with DAO).
- Single-row long values are placed as the engine places them: the page
  last written this session when it has room, else the first listed page
  with room, else a fresh page; LVAL pages stay listed while more than
  256 bytes are free (was 64); an update stores the new value before
  freeing the old; a delete re-lists its page. Byte-exact live gate.
- Deleting the last row of a data or LVAL page retires the page as the
  engine does (type 0x09, slots 0xD000, released, out of the maps; a
  table's first data page stays), pages that lost rows rejoin the
  free-space map, and `Table.truncate()` mirrors an unfiltered `DELETE
  FROM`: pages released untouched, maps emptied, indexes reset. All
  byte-exact against DAO in the live gate.
- Pages released in a session (dropped tables, rewritten definitions,
  freed long values) are not reallocated until the database is reopened,
  as the engine does; an `AccessDatabase` instance is the session.
- `CatalogEntry.date_create_serial` / `date_update_serial` carry the
  stored stamps as doubles, and DateTime columns, `create_table` and
  `create_index` accept such serials, so a stamp copied from another
  database lands bit for bit (a datetime cannot carry the last bit).
  `update_row` keeps the stored bytes of every column it does not touch,
  and index keys are built from the stored serial rather than a decoded
  datetime.
- `AccessDatabase.create_new(path)` writes a blank database from the
  embedded Access-authored template; `AccessDatabase`, `ColumnSpec` and
  `IndexSpec` are exported from the package root, and the README has a
  section on writing Access tables without Office.
- **Files grow past 512 pages** the way the engine grows them: inline
  usage maps enlarge their bitmaps in 8-byte steps, an empty map is
  re-based to its first page, and the global map is extended a step at
  a time. Growing a database from 121 to 573 pages leaves every page but
  page 0 identical to the engine's own.

## [3.5.1] - 2026-08-31

### Changed

- The README now introduces the form designer where a reader starts.
  3.5.0 documented it in full, but the "Why use this?" pitch and the
  "good fit for" list still described module operations only, and the
  architecture map predated `forms.py`, `_oforms_records.py`,
  `_oforms_pages.py` and `_ppt_container.py`, the `forms` CLI command,
  and the `.xlam` and `.accdb` templates.  PyPI renders a project page
  from the README in the released sdist, so correcting it there takes a
  release.

The library itself is unchanged: 3.5.0 and 3.5.1 are the same code.

## [3.5.0] - 2026-08-31

### Added

- **UserForm designs are now read and written, not just preserved**
  (issue #15).  A form's *code* was always a module like any other; its
  *design* -- which controls exist, how they nest, and what their
  properties are -- lived in streams the library carried verbatim.  It is
  now a first-class surface, with no Office installed:

  ```python
  with pyopenvba.ExcelFile("book.xlsm") as wb:
      form = wb.add_form("Wizard", caption="Setup", width=300, height=200)
      form.add_control("Frame", "Shipping", left=12, top=40, width=200, height=80)
      form.add_control("OptionButton", "Ground", container="Shipping")
      form.add_control("MultiPage", "Tabs", left=12, top=140, width=280)
      form.add_page("Tabs", name="Review")
      form.control("Ground").set_property("Caption", "Ground shipping")
      wb.save()
  ```

  `host.forms()` reads the tree; `host.add_form()` composes one from
  nothing; `form.add_control()` / `remove_control()` / `add_page()` /
  `remove_page()` and `control.set_property()` edit it.  Containers
  recurse -- a `Frame`'s children and a `MultiPage`'s pages live in
  storages of their own -- and each is created and deleted with its
  storage.  Geometry is in points.  `python -m pyopenvba forms <file>`
  prints the tree; `--mask` gives the raw property bits instead.

- **Only what the developer set.**  MSForms stores a property just when it
  differs from that control's default, so `control.properties()` is the
  set the author chose -- which a live COM read cannot distinguish from
  inherited and default values.  That is the reason this belongs in a
  file-level library.

- **Writing is lossless.**  An unedited form saves back byte for byte:
  alignment padding, raw string bytes, pictures and any tail the property
  tables do not model are all replayed as read.  Bytes inside a record
  that the tables cannot explain are refused rather than dropped, and a
  form whose streams do not reconcile raises `FormParseError` rather than
  returning a partly guessed control list.

- **Verified against live Excel and live PowerPoint**, which is where four
  defects surfaced that no structural check could catch: an added control
  colliding with the last one's id (`NextAvailableID` is the highest
  handed out, not the next free), a MorphData record omitting reserved
  mask bit 31 ([MS-OFORMS] 2.2.5.2), a container written with a leaf's
  site, and a designer edit leaving the `_VBA_PROJECT` cache stale.

- **Path-addressed CFB navigation and editing**: `CFB.list_storages_at`,
  `list_streams_at`, `get_stream_at`, `write_stream_at`, `add_stream_at`,
  `add_substorage_at` (which can set a storage's CLSID), and
  `remove_storage_at` (recursive).  Nested designer storages repeat
  names -- every container owns an `f` -- so a name-based lookup finds
  whichever comes first in directory order.

- `VBAForm`, `FormControl`, `Size` and `FormParseError` are exported from
  the package root.

### Fixed

- **A UserForm edit left the VBA performance cache stale.**  Only module
  changes counted as mutating, so a designer-only save kept a
  `_VBA_PROJECT` cache describing the form's old members and Office
  refused to load the form.  A designer edit now invalidates it too.
- **`.ppt` was advertised but could not be read** (issue #17).
  `PowerPointFile` listed `.ppt` and failed on every real one with
  "No 'dir' stream found", which reads like file corruption and is not.
  Unlike `.doc` and `.xls`, a binary presentation's CFB root carries no
  VBA storage: the project is a whole CFB, zlib-deflated, inside an
  `ExOleObjStg` record of the `PowerPoint Document` stream, reached
  through the persist chain.  Both directions now work; the write path
  splices the record back in and shifts every absolute offset past it.
  Verified against live PowerPoint, each check run first against an
  untouched control: a rewritten presentation opens with its slides,
  titles and body text intact, and an edited macro returns the new value.

## [3.4.0] - 2026-08-03

### Fixed

- **Non-Latin module names were corrupted in the PROJECT stream**
  (issue #11).  The PROJECT stream is code-page ANSI per [MS-OVBA]
  2.3.1, but four sites hardcoded cp1252, so any rewrite of it -- add,
  rename, or delete -- re-encoded module names with `errors="replace"`.
  A cp1251 project containing `МодульТест` came out as
  `Module=??????????` while the dir stream kept the real name; Excel
  cross-checks those declarations, so the project was left internally
  inconsistent.  `serialize_project_stream`, `parse_project_stream`, and
  `parse_projectwm` now take the project's `code_page` (defaulting to
  1252 for standalone callers) and the save path passes it.  Verified in
  live Excel: a cp1251 workbook whose module is *named* `МодульТест`
  now compiles and returns `Привет, мир` from a Cyrillic-named function.
- **Vietnamese text was destroyed on encode** (issue #13).  Python's
  charmap codecs do no composition, so `'Tiếng Việt'.encode('cp1258')`
  lost every stacked-diacritic character -- and NFD does not help, since
  cp1258 stores `ệ` as precomposed `ê` plus a combining dot-below rather
  than its canonical decomposition.  The new
  `pyopenvba.vba.encode_mbcs` decomposes unmappable characters and folds
  each combining mark back into the base until the codec accepts the
  result, emitting the remaining marks as combining bytes.  Text the
  codec already encodes directly is returned byte-for-byte unchanged.
- **Code pages resolved differently on Windows than on Linux/macOS.**
  CPython falls through to the operating system's code-page registry on
  Windows, so `cp10000`, `cp20866`, `cp21866`, `cp28592`, and `cp28595`
  resolved there while raising `LookupError` elsewhere -- text in those
  pages decoded correctly on one platform and became latin-1 mojibake
  on another.  `_CODEPAGE_ALIASES` now maps 30 Windows code-page
  identifiers (Macintosh, KOI8, the ISO-8859 family, ISO-2022, EUC, GB,
  UTF-7, GB18030) to portable Python codec names and is consulted
  first, so every platform resolves identically.  Found by the new
  cross-OS CI job on its first run; two tests now assert portability
  against the pure-Python codec registry so a regression fails on every
  platform rather than only the affected one.
- **Unresolvable code pages failed silently** (issue #12).  Falling back
  to latin-1 now emits a `UserWarning` instead of quietly producing
  mojibake that survives round-trip checks.
- **ANSI and Unicode dir records are reconciled** (issue #12).  When a
  module's name, stream name, or doc string disagrees between its ANSI
  record and its UTF-16 partner, the Unicode record -- lossless by
  construction -- is now authoritative.

### Added

- **20-language code-page test matrix** (issue #13, ported from
  `xlide_vscode`): one native-language sample per supported code page,
  each asserting zero substitution bytes on encode, an NFC-normalized
  round trip, and a full write -> read -> list -> validate cycle on a
  workbook whose PROJECTCODEPAGE is that page, plus native-language
  module names for cp1251 / cp932 / cp936.  The zero-substitution
  assertion is the load-bearing one: with `errors="replace"` a wall of
  `?` round-trips happily.  Fixtures are generated by patching one
  template's dir record, so no per-language binaries are committed.
- **Dedicated cross-OS `languages` CI job** running that matrix on
  ubuntu and windows, mirroring the equivalent job in the port, so a
  code-page regression names its own OS.
- Live Excel gate case for a Cyrillic-named module (opt-in via
  `RUN_LIVE_EXCEL=1`).

## [3.3.0] - 2026-08-01

### Added

- **Excel fixture CI on real Office** (#4, contributed by
  @DecimalTurn): a Windows workflow that builds fixture workbooks with
  the checked-out pyOpenVBA (no Office needed for the build), installs
  Excel on the runner via the SHA-pinned `DecimalTurn/setup-vba`
  action, runs each fixture's macro over COM, verifies its sentinel
  output, and uploads a desktop screenshot on failure.  Path-filtered
  to fixture and harness changes.  Complements the local
  `RUN_LIVE_EXCEL` gate with per-PR live-Office coverage -- the
  `with_class` fixture is a genuine VBE-export-form class module, so
  the issue #1 bug class is now regression-tested on real Excel in CI.
- **`ExcelFile.create_new` supports `.xlam`** (Excel add-in), joining
  `.xlsm` and `.xlsb`.  The baked-in template is captured from a
  freshly Excel-authored add-in (`ThisWorkbook`, `Sheet1`, bare
  `Module1`) via the new `scripts/bake_xlam_template.py`, following
  the existing bake pattern.

## [3.2.0] - 2026-08-01

### Changed

- **Decompression is 1.76x faster, byte-for-byte identical** (issue #5).
  `decompress` now emits output with slice operations wherever the spec
  allows -- non-overlapping copy tokens move as one slice, runs of
  literal tokens within a flag byte extend once -- and recomputes the
  copy-token masks only when the chunk-local output size crosses a
  power of two.  Overlapping copies keep the spec's byte-at-a-time
  semantics.  Measured 12.4 -> 21.8 MB/s across the 31 module and dir
  streams in the live fixtures; new oracle-equivalence tests pin the
  optimized decoder against the original per-byte implementation,
  including identical error messages and offsets on malformed input.
- **Module source loads lazily** (issue #5).  Decompressing module
  source is 88-96% of the cost of opening a project, so
  `parse_vba_project` now decompresses only the first chunk of each
  module stream (enough for the `Attribute VB_*` header; for
  single-chunk modules it already is the whole source) and defers the
  rest until the first `VBAModule.source` access.  Stream lookup and
  MODULEOFFSET bounds checks stay eager.  Opening the large-module
  fixture for `module_names()` drops from 1.47 ms to 0.79 ms.  Two
  visible consequences: a corrupt chunk past the first one raises
  `VBAProjectError` at first access instead of at parse time, and
  `VBAModule` is now a regular class rather than a dataclass -- the
  constructor signature is unchanged, a new `source_loaded` property
  reports materialization, but dataclass-generated field equality and
  repr are gone (equality is identity).

### Added

- `decompress(..., max_bytes=N)` stops at the first chunk boundary at
  or beyond N output bytes and returns the chunk-aligned prefix.  Copy
  tokens never cross chunk boundaries (the decoder enforces it), so
  the prefix is byte-identical to the same range of a full
  decompression.

## [3.1.0] - 2026-07-22

### Fixed

- **Class modules built from VBE-exported `.cls` sources now compile in
  the host** (GitHub issue #1). `add_module(kind=VBAModuleKind.other)`,
  `set_module`, and `push_modules` normalize class sources from
  file-export form to stream form via the new
  `pyopenvba.vba.normalize_class_source()`: a leading
  `VERSION 1.0 CLASS` / `BEGIN` / `END` preamble is stripped, and
  `Attribute VB_Base` is inserted after `VB_Name` when missing.  On
  replacement of an existing module the prior header's `VB_Base` line is
  preserved, so document-module host CLSIDs are never overwritten.
  Previously a supplied header was written into the stream verbatim: a
  missing `VB_Base` made Excel raise "Invalid procedure call or
  argument" at the first `New` site, and a VERSION preamble in the
  stream raised "Compile error: Expected: end of statement" (both
  verified against live Excel, as is the fix).  Supersedes the 2.0.1
  guidance that callers must supply the `VB_Base` line themselves.
- `pyopenvba.__version__` reported 2.0.0 while PyPI shipped 3.0.x.  A
  new test pins it to the installed package metadata so the two sources
  cannot drift again.
- CFB `get_stream_in_storage` / `write_stream_in_storage` /
  `list_streams_in_storage` now operate on the named storage's own
  child subtree instead of linear-scanning the whole directory.  The
  old scan could read or overwrite a same-named stream in a different
  storage (two UserForms both carry `o` / `f` streams) and reported
  root-level streams as members of every storage.  The host facades now
  address `PROJECTwm` at the project root, where [MS-OVBA] 2.2.1 puts
  it.  Byte output for well-formed files is unchanged (verified by
  hashing a 25-case save matrix across all live fixtures).
- `python -m pyopenvba pull / push / ls` now route Word and PowerPoint
  files by extension instead of assuming Excel; legacy `.xls` / `.doc`
  / `.ppt` are accepted everywhere the modern extensions are.  `disasm`
  no longer advertises `.xltm` / `.ppam`, which no facade accepts.
- `python -m pyopenvba access-pull` delegates to
  `AccessReader.pull_modules`, so Access class modules export as
  `.cls` (previously everything was written as `.bas`).
- README support section named the wrong project; roadmap.md's link to
  the feature-gate matrix pointed outside `docs/`.

### Changed

- The MS-OVBA compressor's LZ encoder uses a 3-gram position index
  instead of re-scanning the whole window at every position: about 60x
  faster on the 17 KB large-module fixture (0.44 s to 0.007 s) and
  0.4 s on a 1 MB input.  Output is byte-for-byte unchanged -- Access
  validates OVBA cache blobs against exact compressor output -- pinned
  by new naive-oracle equivalence tests across random, repetitive, and
  boundary inputs.
- `AccessReader.pull_modules` walks the database's LVAL rows once
  instead of four times per call.
- `save()` emits pending module additions and deletions in sorted
  order, making multi-add saves byte-deterministic across processes
  (Python randomizes set iteration per process via string hashing).
- `ExcelFile`, `WordFile`, and `PowerPointFile` are now thin subclasses
  of a single shared implementation
  (`pyopenvba._host.VBAHostFile`), removing three hand-synchronized
  copies of the read/edit/pull/push/save pipeline (~900 duplicated
  lines).  The public API is unchanged and the refactor was verified
  byte-identical against the previous implementation on every live
  fixture and save operation.

### Added

- **Live Excel compile-and-run gate** (`tests/test_live_excel_gate.py`
  plus `tools/live_excel/`): builds a workbook with an export-form
  class module, runs its macro in desktop Excel under a popup-aware
  bounded harness (VBE modals are dismissed, captured, and reported
  instead of deadlocking the run), and requires a clean run plus the
  macro's sentinel output.  Opt-in via `RUN_LIVE_EXCEL=1` on Windows;
  skipped in CI.  Issue #1 shipped because "opens without a repair
  prompt" was the strongest live verification; this gate closes that
  gap.
- CI matrix now tests Python 3.14 (the classifiers already claimed it).

## [3.0.0] - 2026-05-24

### Added

- **`AccessReader`** (EXPERIMENTAL) -- pure-Python **read-only** support for
  Microsoft Access `.accdb` / `.mdb` (ACE / Jet 4) databases:
  - `AccessReader(path)` parses the 4 KiB page-layout file header and
    validates the ACE / Jet signature.
  - `iter_vba_modules()` yields every embedded VBA module (`VBAModule`
    dataclass with `name`, `start_offset`, `attributes_text`, `source`).
    Modules are discovered by scanning for MS-OVBA stream signatures and
    walking the LVAL page chains they live on -- no Access COM, no
    MSysObjects parser required.
  - `vba_module_names()` deduplicates shadow / undo copies and returns
    the live module name list.
  - `read_vba_module(name)` returns the user-visible source string with
    `\r\n` line endings preserved; matches Access COM
    `CodeModule.Lines()` output byte-for-byte (verified on a 1000-line
    Module + 1000-line Class + 500-line Module live fixture against an
    Access COM oracle).
  - Re-exported from `pyopenvba` as `AccessReader`.
  - Write path (re-compress + re-allocate LVAL pages) is not implemented;
    Access support is read-only by design.

### Changed

- **BREAKING**: Renamed `pyopenvba.access` module to `pyopenvba.access_read`
  and renamed the `AccessFile` class to `AccessReader` to make the
  read-only access posture explicit.
- Adopted strict static analysis: pyright `typeCheckingMode = "strict"`
  and a curated ruff lint configuration (`E, F, W, B, UP, SIM, I, RUF,
  PIE, C4, PERF, N, TC, RET, TRY`) now run clean across `src/` and
  `tests/` with 0 errors.

### Removed

- Pruned ~1800 lines of dead Access write-path / probe code and the
  associated tests that exercised never-public APIs.

## [2.0.1] - 2026-05-24

### Added

- **`synthesize_class_header(name)`** -- new public helper (importable from
  `pyopenvba`) that returns the standard eight-line attribute header for a
  plain VBA class module, including the universal `VB_Base` CLSID. It is
  now also emitted automatically by `add_module(kind=VBAModuleKind.other)`
  when a bare body is supplied, matching the existing behaviour for standard
  modules. Callers no longer need to construct or hard-code the CLSID
  constant themselves.

### Fixed

- **README relative links were broken on PyPI.** The links to `LICENSE.md`,
  `docs/roadmap.md`, `docs/architecture.md`, and
  `docs/ms-ovba-implementation-guide_v2.md` were relative paths that
  resolved correctly on GitHub but 404'd on the PyPI project page. All
  five occurrences are now absolute `github.com/blob/main/...` URLs.

### Changed

- Demo scripts (`create_new_excel_with_class_demo.py`,
  `create_new_with_class_demo.py`, `create_new_word_with_class_demo.py`,
  `inject_xlsb_with_class_demo.py`) updated to use the body-only
  `add_module` call, removing the manual `_CLASS_VB_BASE` constant and
  `DATAMODEL_HEADER` block.
- README Architecture section updated to include `synthesize_class_header`
  in the `__init__.py` public API listing.

## [2.0.0] - 2026-05-24

### Added

- **`WordFile`** -- full read/write support for Word macro-enabled files:
  `.docm`, `.dotm` (OOXML/ZIP), and legacy `.doc` (raw CFB/BIFF8).
  Exposes the same API as `ExcelFile`: `module_names()`, `get_module()`,
  `set_module()`, `vba_project()`, `save()`, `pull_modules()`,
  `push_modules()`.
- **`PowerPointFile`** -- full read/write support for PowerPoint
  macro-enabled files: `.pptm`, `.potm` (OOXML/ZIP), and legacy `.ppt`
  (raw CFB). Same API surface as `ExcelFile` and `WordFile`.
- **`WordFile.create_new(path)`** -- create a brand-new `.docm` from
  scratch without launching Word. Ships with `ThisDocument` and an empty
  `Module1`; opens cleanly with no repair prompt.
- **`PowerPointFile.create_new(path)`** -- create a brand-new `.pptm`
  from scratch without launching PowerPoint. Ships with an empty
  `Module1`; opens cleanly with no repair prompt.
- **`ExcelFile.create_new()` now supports `.xlsb`** in addition to
  `.xlsm`. The extension in the path controls which baked-in template is
  used.
- **`pull_word(document, dest_dir)`** / **`push_word(src_dir, document)`**
  -- disk-based pull/push helpers for Word, mirroring the Excel `pull()`
  / `push()` API.
- **`pull_ppt(presentation, dest_dir)`** / **`push_ppt(src_dir, presentation)`**
  -- disk-based pull/push helpers for PowerPoint.
- **`scripts/bake_xlsb_template.py`** -- bakes the empty `.xlsb` template
  blob into `_templates/__init__.py` using the same splice pattern as the
  docm/pptm bake scripts.
- Class module creation is now fully supported across all three hosts.
  When adding a class module via `add_module(kind=other)`, callers must
  supply the full attribute header including
  `Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"` (the
  universal VBA class CLSID); without it Office raises "Invalid procedure
  call or argument" on instantiation.

### Changed

- `pyproject.toml` description updated to reflect all three supported
  Office hosts; `word`, `powerpoint`, `docm`, and `pptm` added to
  keywords.
- README fully updated: tagline, supported formats tables, 30-second tour,
  `create_new` section, and pull/push workflow section now cover Excel,
  Word, and PowerPoint.

## [1.1.1] - 2026-05-22

### Fixed

- **Editing a document module's source via `set_module()` (e.g. `ThisWorkbook`,
  `Sheet1`) silently broke the workbook in Excel.** The leading
  `Attribute VB_Name = "ThisWorkbook"` / `Attribute VB_Base = "..."` /
  `Attribute VB_PredeclaredId = True` header lines that bind a document
  module to its host object were being stripped on a source replacement.
  Excel then re-compiled the module without those bindings and either
  silently dropped the code or showed an empty module in the VBE.

### Added

- **VBE-style body-only source edits.** `ExcelFile.set_module(name, text)`
  now accepts either a full source replacement (text beginning with
  `Attribute VB_*` or `VERSION ... CLASS`) or a bare body. When a bare
  body is supplied, the module's existing attribute header is
  automatically re-prepended, matching the VBE UX where the user only
  types the executable code.
- **`VBAModule.body`** property: read or write a module's executable body
  without touching its attribute header.
- **`VBAModule.attribute_header`** field: the contiguous leading
  `VERSION ... CLASS` block + `Attribute VB_*` lines + separator,
  captured at parse time.
- **`split_attribute_header(source) -> (header, body)`** public helper.
- **`add_module(name, body, kind=standard)` now synthesizes a minimal
  `Attribute VB_Name = "<name>"` header** when the caller doesn't supply
  one. Caller-supplied headers are passed through unchanged.
- **`add_module(kind=other)` requires an explicit attribute header.**
  pyOpenVBA refuses to invent class or document module headers since
  their host-binding metadata can't be safely guessed.
- **`rename_module()` re-keys the in-source `Attribute VB_Name = "..."`
  line** to the new logical name so the source matches the dir-stream
  binding.
- New `TestAttributeHeaderPreservation` test class covering:
  header splitting (standard, document, class, headerless),
  `set_module` body-only preservation on a document module,
  `set_module` full-source replacement,
  `add_module` header synthesis vs. caller-supplied,
  `add_module(kind=other)` rejection without a header,
  and the `VBAModule.body` property round-trip.

## [1.1.0] - 2026-05-22

### Added

- **`ExcelFile.create_new(path)`** -- create a brand-new macro-enabled
  workbook from scratch in pure Python, without ever launching Excel.
  The new file ships with a fresh VBA project containing `ThisWorkbook`,
  `Sheet1`, and an empty `Module1`, opens cleanly in Excel with no
  "found a problem with some content" repair prompt, and is ready for
  immediate edits via the normal `vba_project()` / `save()` flow.
- New `TestExcelFileCreateNew` test class covering write-out, expected
  modules, empty `Module1`, round-trip with user code, overwrite of an
  existing file, and creation of missing parent directories.

### Internal

- New `src/pyopenvba/_templates/__init__.py` module embedding a
  byte-for-byte clone of a freshly Excel-authored empty `.xlsm` as a
  zlib-compressed base85 constant. No binary fixtures are shipped in the
  wheel; the template is regenerated by `scripts/bake_empty_template.py`
  from `tests/live_excel_testing/freshly_touched.xlsm`.

## [1.0.1] - 2026-05-22

### Fixed

- **Excel rejected modules whose source spanned more than one 4 KB chunk**
  with *"An error occurred while loading <Module>"*. The MS-OVBA compressor
  was emitting raw (CompressedChunkFlag = 0) chunks for full 4096-byte
  blocks. Although spec-legal, Office itself never writes raw chunks for
  module source streams -- empirically confirmed against an Excel-authored
  workbook containing a 16,881-byte module (all five of its chunks were
  token-compressed). The compressor now always emits token-compressed
  (flag = 1) chunks for module source; raw chunks remain only as a fallback
  for adversarial 4096-byte high-entropy input that overflows LZ encoding.
- **Re-running an add-module workflow after a delete produced duplicate
  `PROJECT` entries**, which Excel treats as corruption. Calling
  `add_module(name, ...)` after `delete_module(name)` in the same save now
  cancels the pending delete and treats the operation as a source rewrite,
  matching Excel's own behaviour. `serialize_project_stream` additionally
  scrubs duplicate `Module=` and workspace declarations on every structural
  save, healing files that were corrupted by earlier versions.

### Added

- `demo/` folder containing a runnable end-to-end demo
  (`push_demo_module.py` + `test_macro_workbook.xlsm` + `demo.md`).
- New regression tests:
  - `TestCompress.test_full_chunk_emitted_as_token_compressed_not_raw` and
    `TestCompress.test_long_module_round_trip_through_excel_save` verify
    that no raw chunks are produced for realistic VBA source.
  - `TestLargeModuleFixture` uses an Excel-authored 16 KB module as an
    empirical anchor and round-trips it through pyOpenVBA's saver.
  - `test_delete_then_readd_same_name_does_not_duplicate_project_decl` and
    `test_save_heals_preexisting_duplicate_project_declarations` cover the
    PROJECT-stream fix.
- `tests/live_excel_testing/large_vba_module.xlsm` fixture (Excel-authored
  reference for multi-chunk module compression).

## [1.0.0] - 2026

Initial public release. Pure-Python read/write support for VBA projects
inside `.xlsm`, `.xlsb`, and `.xls` containers, covering CFB parsing,
MS-OVBA compression, module add/edit/rename/delete, `PROJECT`/`PROJECTwm`
serialization, `_VBA_PROJECT` cache invalidation, and round-trip
preservation including password-protected projects.

[2.0.0]: https://github.com/WilliamSmithEdward/pyOpenVBA/compare/v1.1.1...v2.0.0
[1.1.1]: https://github.com/WilliamSmithEdward/pyOpenVBA/compare/v1.1.0...v1.1.1
[1.1.0]: https://github.com/WilliamSmithEdward/pyOpenVBA/compare/v1.0.1...v1.1.0
[1.0.1]: https://github.com/WilliamSmithEdward/pyOpenVBA/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/WilliamSmithEdward/pyOpenVBA/releases/tag/v1.0.0
