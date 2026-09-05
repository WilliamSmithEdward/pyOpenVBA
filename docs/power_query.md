# Power Query in an Excel workbook

How Excel stores Get and Transform queries, and what pyOpenVBA writes.

Every statement below was measured, not read off a specification. Two
oracles were used, and each fact says which one settled it:

* **Live Excel**, through `pyvbaharness`: build a workbook, open it, ask
  the object model what it sees, and have the mashup engine evaluate a
  query. This is what says a written package works.
* **Microsoft's own packaging assemblies**, which ship beside Excel in
  `ADDINS\Microsoft Power Query for Excel Integrated\bin`. Loading
  `Microsoft.Mashup.Client.Packaging.dll` in PowerShell and calling
  `PackageMetadataSerializer`, `QueriesMetadataSerializer` and
  `SerializedMetadataEntry` gives the exact bytes Excel's own code would
  write for a given input. This is what settled the encodings, and the
  golden values in `tests/test_powerquery_metadata.py` came out of it.

---

## Where it lives

The whole project sits in one custom XML part, `customXml/itemN.xml`,
written in UTF-16 with a byte-order mark:

```xml
<?xml version="1.0" encoding="utf-16"?>
<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">BASE64</DataMashup>
```

The element sometimes carries an `sqmid` attribute, so the wrapper is
kept as it was found.

A workbook that has never held a query needs four more things before the
part means anything, and pyOpenVBA writes all of them:

| Part | What it carries |
| --- | --- |
| `customXml/itemPropsN.xml` | `ds:datastoreItem` with a schema reference to the DataMashup namespace |
| `customXml/_rels/itemN.xml.rels` | the item's pointer at its properties |
| `[Content_Types].xml` | an Override for the properties part; the item itself rides the `xml` Default |
| `xl/_rels/workbook.xml.rels` | a `customXml` relationship to the item |

Nothing else is needed for queries that only hold a connection (measured
against workbooks Excel wrote, and against a workbook built here from
one that had no Power Query at all).

---

## The blob

Five pieces, each after the first behind a length:

```
u32 version (0)
u32 length + the OPC package
u32 length + the permission list (XML)
u32 length + the metadata section
u32 length + the permission bindings
```

Measured with live Excel, by emptying one piece at a time:

* **The permission list is required.** Empty it and Excel opens the
  workbook with an error instead of its queries.
* **The metadata's content package is required.** Cut the metadata short
  of it and the same thing happens.
* **The bindings are not required.** They hold a signature protected by
  the Windows data-protection API, tied to the machine that wrote them.
  Excel opens and refreshes a workbook whose bindings are empty, so
  pyOpenVBA drops them when it changes the package: a signature that no
  longer covers what it signs is worse than none.

**Excel does not rewrite the blob when it saves a workbook it did not
edit.** A blob written here comes back byte for byte after Excel opens
and saves the file. So the writer aims at Excel's own bytes, not
merely at bytes Excel tolerates.

### The package

An OPC ZIP with three parts, in this order: `Config/Package.xml`,
`[Content_Types].xml`, `Formulas/Section1.m`.

Excel writes it with settings of its own, and pyOpenVBA reproduces them:

* raw deflate at **level 6** with the default memory level. Level 6 is
  the only level that explains every entry measured: a 1075-byte section
  rules out 7 and above, and a 73 KB one rules out 5 and below.
* general-purpose flag `0x0002`, version made by 45, version needed 20.
* a 28-byte Open Packaging growth hint (`0xA220`) on every local header
  and none in the central directory.

Rebuilding each fixture's package from its parts gives back Excel's
bytes exactly (`tests/test_powerquery_package.py`).

### Formulas/Section1.m

```
section Section1;<CRLF><CRLF>
[ Description = "what it does" ]<CRLF>     (only when there is one)
shared Name = <expression>;<CRLF><CRLF>
...
shared Last = <expression>;                (no newline at the end)
```

CRLF throughout. A name that is not a plain identifier is written
`#"Like This"`.

### The metadata section

```
u32 version (0)
u32 length + the XML (UTF-8, with a byte-order mark)
u32 length + a content package (an empty ZIP, 22 bytes)
```

The XML is a `LocalPackageMetadataFile` listing one `Item` per formula:

```xml
<Item>
  <ItemLocation><ItemType>Formula</ItemType>
    <ItemPath>Section1/Order%20Lines</ItemPath></ItemLocation>
  <StableEntries><Entry Type="IsPrivate" Value="l0" /></StableEntries>
</Item>
```

* The first item is `AllFormulas`, with an empty path. It carries what
  belongs to the whole document: the query groups, and the switches for
  relationship and type detection.
* **A query is an item with at least one entry.** An item whose
  `StableEntries` is empty is a step. This was measured: a `shared`
  member added with an entry-less item left Excel showing three queries,
  and any single entry was enough to make the fourth appear.
* **A member with no item is worse than useless.** Excel fails to open
  the queries at all ("the index is out of bounds"), so the two halves
  are always written together.
* Steps come from the **top-level** `let` bindings, one item each,
  immediately after the query's own item. A nested `let` contributes
  nothing, an expression that is not a `let` gets no steps, and a
  function gets none either.
* Item paths are percent-encoded by .NET's `Uri.EscapeDataString` rules:
  `A-Za-z0-9-._~!'()*` survive, everything else becomes `%XX` over UTF-8.

### Entry values

The first character is the type, from Microsoft's `SerializedMetadataEntry`:

| Prefix | Type | Example |
| --- | --- | --- |
| `l` | 64-bit integer | `l0`, `l-7` |
| `f` | double, .NET's `G15` | `f1.5`, `f0.333333333333333`, `f1E+20` |
| `s` | text | `sConnectionOnly` |
| `c` | GUID | `c1111...` |
| `d` | timestamp | `d2026-09-05T13:45:12.2500000Z` |

Excel writes `QueryID` as a text entry, not a GUID one, and pyOpenVBA
does the same.

What a new query carries, exactly as Excel writes it: `IsPrivate` 0,
`FillEnabled` 0, `FillObjectType` `ConnectionOnly`,
`FillToDataModelEnabled` 0, `QueryID`, and `ResultType` `Function` when
the expression is a function.

`ResultType` is Excel's record of what a query *evaluated* to, not of how
it is written: Excel marks `each _ + 1`, `(x) => x` and
`let F = (x) => x in F` alike. The last cannot be told from the text
without running the query, so pyOpenVBA answers the syntactic question
and lets Excel correct the entry on its next refresh.

### Query groups

The folders in the Queries pane live in one entry on the `AllFormulas`
item, and the value is **not JSON**. It is Microsoft's own binary
serialization, base64-encoded:

```
u32 count
per group:
  u32 0
  16 bytes  Guid Id (.NET layout)
  string    Name          (7-bit-encoded length, UTF-8)
  string    Description   (never null; Excel writes "")
  u8        has parent, then 16 bytes Guid ParentId when set
  i32       Order
```

A query joins a group through its own `QueryGroupID` entry.

pyOpenVBA's writer is byte-identical to
`QueriesMetadataSerializer.SerializeQueryGroups` over 60 randomized
cases. Excel parses the value with the same reader: given a deliberately
corrupt one, it threw out of
`QueryGroupMetadataSet.Deserialize(BinarySerializationReader)` and could
not open the queries, while the valid one written here survived Excel's
own edit of the workbook untouched.

---

## Loading a query onto a sheet

The metadata says where a query loads, and **saying it changes nothing**:
a connection-only query marked `FillEnabled = 1`, `FillObjectType = Table`
opened in Excel, refreshed, and produced no table at all. Excel loads a
query only when the workbook also carries the objects that do it:

* a connection in `xl/connections.xml` through
  `Provider=Microsoft.Mashup.OleDb.1` with `Location=<query>` and
  `command="SELECT * FROM [<query>]"`,
* `xl/queryTables/queryTableN.xml` bound to that connection,
* `xl/tables/tableN.xml` with `tableType="queryTable"`,
* the sheet's `tableParts` entry and relationship, a header row, and a
  hidden `ExternalData_N` defined name.

`load_to_sheet()` writes all of it. The column names have to be supplied,
because working them out means evaluating the query and only the mashup
engine can do that; Excel reconciles them with the real result and fills
the rows on the first refresh.

`unload()` takes every piece back out. One rule there is easy to get
wrong: **a connections part holding no connections is one Excel refuses**,
so removing the last connection removes the part, its content type and
its relationship as well.

---

## What Excel refuses

* A query name containing a dot. `Queries.Add` rejects it with
  `0x80070057`, so a workbook should not carry one either.

## What is not covered

* Loading to the data model. The connection for it is a different shape,
  and the model itself is a separate store.
* Evaluating M. Nothing here runs a query; that is the engine's job.
* Credentials and privacy levels beyond the permission list, which is
  written as Excel writes it for a workbook whose queries hold none.
