# Complex columns: the fixtures behind the shipped support

Attachment and multi-valued columns are **shipped** -- see
`pyopenvba.access._complex`, `AccessDatabase.complex_columns()` and the
`attachments` / `multi_values` pair on `Table`. What is here is the DAO
that built the fixtures those were measured against, kept so the
measurements can be repeated.

| script | what it builds |
|---|---|
| `build_fixture.ps1` | `Things`, an attachment column and a multi-valued text column, three rows, one of them empty. Committed as `tests/live_access_test/complex_columns.accdb`. |
| `build_type_fixture.ps1` | A zip, a png and a 40 KB text file, so the stored-raw and the chained cases are both covered. |
| `attachment_types.ps1` | One identical compressible payload under 51 extensions, which is how the compression rule was found. |

Run any of them with `-Target <path>` on Windows with the Access database
engine installed. They use DAO and are dev-only; nothing in `src/` uses
COM.

## What they established

**Access compresses attachments by file type, not by outcome.** It stores
`docx`, `gif`, `jpeg`, `jpg`, `png`, `pptx`, `xlsx` and `zip` as they are
and compresses the other 37 probed, `7z`, `rar`, `gz`, `mp3` and `mp4`
included -- a 72-byte PNG of nearly constant bytes goes in uncompressed
while a text file of the same size does not.

**It refuses some extensions outright**: `bas`, `accdb`, `cab`, `exe`,
`iso` and `msi`. That is Access's own security policy, and not something
the library enforces.

**Its deflate is not zlib's.** No combination of level, memLevel,
strategy or window size reproduces a stream Access wrote, so a compressed
attachment written by this library inflates to the same file and packs
differently. The eight stored-raw types are byte-identical.

**The complex id lives at 0x1C of the table definition**, not at 0x18 as
the engine's model had it -- 0x18 is 1 on every table in every database
measured. The counter moves once per inserted row, is shared by every
complex column in that row, and is not reused after a delete: deleting
the middle of five rows and adding one gave 6.

**A complex column is flagged AutoNumber** like any other, which is what
made the row writer hand two complex columns in one row two different
ids until it learned to skip them.

## Creating one, and the three rules that were invisible

Shipped as `table.add_complex_column(name, kind)`. The shape was easy to
read off a database Access made; what took the measuring was why the
engine refused the result even once the shape matched.

* the parent column is type `0x12`, four bytes, flagged AutoNumber, and
  its header's collation slot carries the **ComplexID** (`Files` had 1,
  `Tags` 2, matching `MSysComplexColumns`)
* it needs a unique index named `<Column>_<32 uppercase hex>`
* the flat table is `f_<32 uppercase hex>_<Column>`, a *different* GUID
  from the index's
* its columns are `_<Column>`, the type's own columns, and
  `<Parent>_<Column>`, with misc flags 8, 16 (or 0 for a scalar) and 4
* three indexes: `_<Column>` on the key, `IdxFKPrimaryScalar` unique on
  `(<key>, FileName)` or `(<key>, Value)`, and `MSysComplexPKIndex`
  unique and primary on `<Parent>_<Column>`

Three more, each found by giving the result to DAO and watching it fail:

**The two Long bookkeeping columns sit among the variable columns.** A
Long is normally fixed, and `create_table` put them there; Access does
not. They carry no fixed bit and no collation -- flags 2 and 6 with sort
order 0, against 3 and 7 with 1033 for an ordinary Long.

**The flat table's catalog `Flags` are `0x800A0000`**, not the
`0x80090000` a first reading of the same row suggested.

**The table that has the column carries `0x40000`**, and no other table
in a database does -- checked against every table of two of them, where
only `Things` and `MSysResources` have it. Without it DAO opens the child
recordset and finds no fields in it, which is a failure that says nothing
about its cause.
