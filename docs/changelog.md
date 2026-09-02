# Changelog

All notable changes to pyOpenVBA are documented here. This project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

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
