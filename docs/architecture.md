# pyOpenVBA Architecture

This document describes the internal layering, module responsibilities,
data flow, and conventions of pyOpenVBA. It is the canonical reference
for contributors. End-users should read the [top-level README](../README.md)
instead.

For the language-agnostic spec walkthrough (what to build if you are
porting this to another language), see
[ms-ovba-implementation-guide_v2.md](ms-ovba-implementation-guide_v2.md).

---

## 1. Layering

pyOpenVBA is organized in three strictly-layered modules. Each layer
knows about the layer below it but not the layer above.

```
+--------------------------------------------------------+
| excel.py / word.py / powerpoint.py   host facades      |
|   - thin subclasses of _host.VBAHostFile               |
|   - per-host constants: container extensions,          |
|     vbaProject.bin entry path, message wording         |
|   - create_new() from a baked-in template              |
+--------------------------------------------------------+
| _host.py        VBAHostFile shared implementation      |
|   - ZIP / raw-CFB dispatch by extension                |
|   - save() pipeline (safety gates, structural rewrites)|
|   - pull / push disk workflow                          |
|   - class-source normalization at set_module / push    |
+--------------------------------------------------------+
| access_read.py  AccessReader facade  (READ-ONLY)         |
|   - ACE / Jet 4 page reader (.accdb, .mdb)             |
|   - MS-OVBA blob discovery via signature scan          |
|   - LVAL page-chain walker                             |
|   - read_vba_module(name) -> str (pure Python)         |
|   - get_module / vba_modules / iter_vba_modules        |
|   - pull_modules / export_modules / export_module      |
|   - read_vba_module_with_attributes (full preamble)    |
|   - read_project_info / identifiers                    |
|   - find_interned_strings / find_module_streams        |
|   - iter_pcode_streams / read_module_pcode_stream      |
|   - disassemble_module (via vba_pcode)                 |
|   - iter_msys_objects / find_msys_module               |
|       (MSysObjects system catalog reader)              |
|   - NO write APIs; see msaccess_lessons_learned.md     |
+--------------------------------------------------------+
| access/         Jet 4 / ACE storage engine (IN PROGRESS)|
|   - _pages: page array, masked page 0, row slots,      |
|       usage maps                                       |
|   - _tdef: table definitions (columns, indexes, maps)  |
|   - _rows: row splitting + every column type's codec   |
|   - _lval: Memo / OLE long values                      |
|   - _index: B-tree pages, key codec both ways          |
|   - _collation: text sort keys (generated table)       |
|   - database: AccessDatabase / Table / Index facade    |
|   - plan and format facts: docs/access_engine.md       |
+--------------------------------------------------------+
| vba.py          VBA project layer                      |
|   - MS-OVBA compression / decompression                |
|   - dir / PROJECT / PROJECTwm / PROJECTlk parsers      |
|   - VBAProject + VBAModule data model                  |
|   - cache invalidation, signature detection            |
+--------------------------------------------------------+
| vba_pcode.py    VBA7 p-code disassembler  (EXPERIMENTAL)|
|   - 264-entry VBA7 opcode table (factual data)         |
|   - per-line p-code streaming parser                   |
|   - source-line correlation (to_annotated_listing)     |
|   - dependency-free (no oletools, no pcodedmp)         |
+--------------------------------------------------------+
| cfb.py          MS-CFB layer                           |
|   - Compound File Binary parser/writer                 |
|   - stream + storage CRUD (case-insensitive lookup)    |
+--------------------------------------------------------+
```

Other files:

- `exceptions.py` — exception hierarchy shared by all layers.
- `__init__.py` — public re-exports (`ExcelFile`, `pull`, `push`,
  `VBAModuleKind`, exceptions).
- `__main__.py` — `python -m pyopenvba
  {pull,push,ls,access-ls,access-pull,access-disasm,disasm}` CLI.
  `disasm` / `access-disasm` accept `--with-source` to interleave
  the original VBA source with the decoded p-code.
- `_templates/__init__.py` — generated module embedding a
  zlib-compressed base85 blob of a freshly Excel-authored empty `.xlsm`,
  consumed by `ExcelFile.create_new()`. Regenerated from
  `tests/live_excel_testing/freshly_touched.xlsm` by
  `scripts/bake_empty_template.py`. No binary fixtures ship in the wheel.

### 1.1 Layer rules

- `cfb.py` is **completely VBA-agnostic** and could be lifted into a
  separate library. It knows nothing about MS-OVBA.
- `vba.py` operates on a `CFB` instance but never opens files on disk.
- `_host.py` is the only module that touches the filesystem, ZIP
  containers, or host file paths; the three facades only supply
  constants and `create_new()` templates.
- Cross-layer leaks are a code smell. Adding a `pathlib.Path` import to
  `vba.py` or `cfb.py` is a red flag.

---

## 2. Public API surface

Everything in `__all__` of [`src/pyopenvba/__init__.py`](../src/pyopenvba/__init__.py)
is supported and version-stable. Everything else is internal and may
change without notice.

| Public name           | Defined in   | Purpose                              |
|-----------------------|--------------|--------------------------------------|
| `ExcelFile`           | `excel.py`   | Excel facade; context manager        |
| `ExcelFile.create_new`| `excel.py`   | Build a brand-new .xlsm from baked template |
| `WordFile`            | `word.py`    | Word facade; context manager         |
| `VBAModuleKind`       | `vba.py`     | enum: standard / class / document / designer |
| `pull`                | `__init__.py`| One-call disk export (Excel)         |
| `push`                | `__init__.py`| One-call disk import + save (Excel)  |
| `pull_word`           | `__init__.py`| One-call disk export (Word)          |
| `push_word`           | `__init__.py`| One-call disk import + save (Word)   |
| `PyOpenVBAError`      | `exceptions.py` | Library root exception            |
| `UnsupportedFormatError` | `exceptions.py` | Bad extension / host           |
| `CFBError`            | `exceptions.py` | Malformed CFB                     |
| `VBAProjectError`     | `exceptions.py` | Malformed VBA / refused mutation  |

Internal-but-useful symbols (importable from `pyopenvba.vba` for
advanced users):

- `compress`, `decompress` — MS-OVBA codec.
- `parse_vba_project`, `parse_project_stream`, `parse_projectwm`,
  `parse_projectlk`.
- `serialize_dir_stream`, `serialize_project_stream`,
  `serialize_projectwm`, `serialize_projectlk`.
- `invalidate_vba_project_cache`, `detect_signature`,
  `compute_v3_content_hash`.
- `VBAProject`, `VBAModule`, `VBAReference`.

These are intentionally not re-exported from the top-level package;
they are stable enough to call but should not be considered the
recommended user surface.

---

## 3. Data flow

### 3.1 Read path

```
ExcelFile(path)
    -> _open()                            (excel.py)
        -> _open_zip() / _open_cfb_direct()
            -> extract xl/vbaProject.bin  (excel.py)
                -> CFB.from_bytes(...)    (cfb.py)
                    -> parse_vba_project(cfb)         (vba.py)
                        -> decompress(dir_stream)     (vba.py)
                        -> _parse_dir_stream(...)     (vba.py)
                        -> parse_project_stream(...)  (vba.py)
                        -> [for each module]
                             cfb.get_stream_in_storage("VBA", name)
                             slice from MODULEOFFSET
                             decompress(...)
                             decode(..., code_page)
```

### 3.2 Write path (the `save()` pipeline)

This is the most important code path in the library. It lives in
`ExcelFile.save()` ([excel.py](../src/pyopenvba/excel.py)) and runs in
a fixed order:

```
1. Detect mutation                      (any rename / add / delete / dirty module)
2. Safety gates                         (protected? signed?)
3. Apply renames                        (CFB stream rename)
4. Apply adds                           (CFB stream create with seeded body)
5. Apply deletes                        (CFB stream remove)
6. Replace dirty module bodies          (compress + write back)
7. Rewrite dir / PROJECT / PROJECTwm    (if structure changed)
8. Invalidate _VBA_PROJECT cache        (zero PerformanceCache body)
9. Drop __SRP_* streams                 (force regeneration)
10. Drop signature streams              (if mutating, with warning)
11. Serialize CFB                       (cfb.to_bytes())
12. Write outer container               (replace single ZIP entry, or raw CFB)
```

If any of steps 3-10 fail, the on-disk file is **never modified** —
all work is done on the in-memory CFB and the final bytes are written
in one atomic-ish operation.

### 3.3 Access read-only model (`AccessReader`)

`.accdb`/`.mdb` files are exposed through a **read-only** facade.
After an extensive reverse-engineering effort (chronicled in
[docs/msaccess_lessons_learned.md](msaccess_lessons_learned.md))
pyOpenVBA does not support writing back to Access databases. Use
Access COM (`win32com.client.Dispatch("Access.Application")`) if you
need to programmatically modify VBA inside a `.accdb`.

The Access on-disk layout differs from Excel/Word/PowerPoint:

* The MS-OVBA blob on the LVAL chain is a **passive cache**. Zero-filling
  the entire blob has no effect on what Access (or the VBA editor)
  displays — verified end-to-end on a live fixture. This is what
  pyOpenVBA reads.
* The **authoritative** sources Access reads from are:
  * **Project metadata** — the MS-OVBA `dir` stream (Section 2.3.4.2)
    OVBA-compressed in a single LVAL row; located by content
    (decompressed prefix = `01 00 04 00 00 00`). Parsed by
    `AccessReader.read_project_info()` -> `AccessVBAProject` (system kind,
    LCID, code page, project name, references, modules with class flag,
    private/read-only flags).
  * **Module bytecode** — the compiled VBA p-code. Lives in LVAL rows
    whose payload starts with magic `72 55 40 00` ('rU@\0'). The
    **module-active** one is identified deterministically by the
    12-byte prefix `72 55 40 00 00 00 00 00 00 00 40 00`
    (byte 10 = 0x40). Exposed via
    `AccessReader.read_module_pcode_stream()` and `iter_pcode_streams()`.
    Compiled p-code is **fully anonymised**: opcodes and slot
    references with no user-authored text. The standard MS-OVBA
    `0xCAFE` module stream coexists alongside `rU@` in a separate
    LVAL row and is what `AccessReader.disassemble_module()` consumes
    via the pure-Python `pyopenvba.vba_pcode` decoder.
  * **Identifier inventory** — every project-level identifier name
    (typelib refs, project name, module/proc/variable names, intrinsic
    call targets such as `MsgBox`) is enumerated in the
    `_VBA_PROJECT`-equivalent stream stored UNCOMPRESSED in the LVAL
    row whose first 2 bytes are the `CC 61` Office magic. Exposed via
    `AccessReader.identifiers() -> tuple[AccessVBAIdentifier, ...]`.
  * **Comment text** — stored verbatim in plaintext rows tagged
    `E3 00 00 00 <u16-LE length> <ASCII payload>` (apostrophe stripped).
  * **String literal text** — stored verbatim in rows tagged
    `B9 00 <u16-LE length> <ASCII payload> <12-byte trailer>`,
    exposed via `AccessReader.find_interned_strings()`.

The **MSysObjects** system catalog is also surfaced read-only:
`AccessReader.iter_msys_objects()`, `msys_objects()`,
`iter_msys_modules()`, `find_msys_object(name, *, type_=None)`,
`find_msys_module(name)`. Each row yields an `AccessSysObject`
dataclass (id_, parent_id, type_, flags, name, page, slot). VBA code
modules carry `type_ == MSYS_TYPE_MODULE` (-32761, 0x8007) and are
parented to the `Modules` container row.

### Excel-symmetric VBA module read API

`AccessReader` exposes the same ergonomic read surface as `ExcelFile` /
`WordFile` / `PowerPointFile`: `get_module(name)`,
`vba_modules() -> dict[str, str]`, `read_vba_module(name)`,
`read_vba_module_with_attributes(name)`, `iter_vba_modules()`,
`pull_modules(dest_dir)`, `export_module(name)`,
`export_modules(dest_dir, *, include_attributes=False)`. The
top-level helper `pyopenvba.pull_access(database, dest_dir)` mirrors
`pull` / `pull_word` / `pull_ppt`. There is intentionally no
`push_access` symbol.

### Why no write path

Access stores compiled VBA p-code (the `CAFE` rows) separately from the
OVBA source cache, and executes an `__SRP_*` compiled cache in
preference to either, so mutations to the source alone do not change
behaviour. Research later found the lever -- dropping those cache rows
makes a rewritten module take effect -- and got as far as rewriting
procedure bodies, including declarations, with output byte-identical to
Microsoft's compiler. It did **not** get to creating, renaming or
deleting a module, which is what a useful writer would need. That work
is parked and unsupported; a production-quality writer would still
require the `FuncDefn` declaration tables and ACE page allocator parity.
See
[docs/msaccess_lessons_learned.md](msaccess_lessons_learned.md) for
the empirical results matrix and the reasoning in full.

---

## 4. Mutation safety gates

| Condition                              | Default behavior                 | Override                                |
|----------------------------------------|----------------------------------|-----------------------------------------|
| Project is password-protected (DPB)    | `VBAProjectError` on mutation    | `save(allow_protected=True)`            |
| Project has any signature stream       | Drop streams + `UserWarning`     | `save(allow_invalidate_signature=True)` |
| Non-mutating save (`save()` no edits)  | Pass through unchanged           | n/a                                     |

Both gates only apply when `save()` detects an actual change. A no-op
save on a protected or signed workbook is always allowed and never
removes signatures.

---

## 5. Outer-container dispatch

`ExcelFile._open()` and `WordFile._open()` each select the container by extension:

**ExcelFile**

| Extension                    | Container       | Implementation       |
|------------------------------|-----------------|----------------------|
| `.xlsm`, `.xlsb`, `.xlam`    | ZIP (OOXML)     | `_open_zip()`        |
| `.xls`                       | raw CFB (BIFF8) | `_open_cfb_direct()` |
| anything else                | n/a             | `UnsupportedFormatError` |

**WordFile**

| Extension        | Container         | VBA entry path           |
|------------------|-------------------|---------------------------|
| `.docm`, `.dotm` | ZIP (OOXML)       | `word/vbaProject.bin`    |
| `.doc`           | raw CFB (Word 97) | (whole file is CFB)      |
| anything else    | n/a               | `UnsupportedFormatError` |

**PowerPointFile**

| Extension        | Container              | VBA entry path         |
|------------------|------------------------|------------------------|
| `.pptm`, `.potm` | ZIP (OOXML)            | `ppt/vbaProject.bin`   |
| `.ppt`           | CFB (PPT 97)           | embedded in `PowerPoint Document` |
| anything else    | n/a                    | `UnsupportedFormatError` |

For the ZIP case, the VBA project is at the fixed path
`xl/vbaProject.bin`. On save, exactly that entry is replaced while
every other ZIP entry is preserved byte-for-byte including its
compression method, external attributes, create system, and timestamp.

For `.xls` and `.doc`, the entire file *is* the CFB; `cfb.to_bytes()`
is written straight to the output path.

`.ppt` is the exception among the legacy containers. Its root holds no
VBA storage at all: the project is a whole CFB, zlib-deflated, inside an
`ExOleObjStg` record of the `PowerPoint Document` stream, found through
the persist chain (`Current User` -> `UserEditAtom` -> `PersistDirectoryAtom`).
`_ppt_container.py` extracts it on open and splices it back on save,
shifting every absolute offset past the resized record. The two hooks
`VBAHostFile._vba_cfb_bytes` / `._container_bytes` are the seam; they are
identities for every other format.

If a `.xlsm` exists but does not contain `xl/vbaProject.bin` (no VBA
project has ever been created), `ExcelFile` raises a structured
`VBAProjectError` with the workbook path — it is **not** a corruption.

---

## 5a. UserForm designer streams

A form's design lives beside the VBA storage, not inside it: a root
storage named for the form, holding `f` (the sites: which controls, in
what order), `o` (each control's own property record), and the
`VBFrame` text. Containers nest into storages of their own, named
for the site id -- a `Frame`'s children in `i02`, a `MultiPage`'s Pages
in `i08` / `i09` under its own `i06`.

`forms.py` reads that tree and writes it back; `_oforms_records.py`
carries one property table per control class. It never guesses: every
structure is counted or length-prefixed, so a misread collapses rather
than yielding a plausible control list, and it raises `FormParseError`
instead. Three checks have to agree -- `CountOfBytes` runs exactly to the
end of `f`, the per-site `ObjectStreamSize` values sum to exactly
`len(o)`, and every child storage is claimed by a site.

Writing is lossless first: alignment padding is captured and replayed
(the spec leaves those bytes undefined), string bytes are kept raw beside
their decoded text, pictures stay opaque runs, and any tail the tables do
not model is preserved. Bytes inside a record's `cb` that the tables
cannot explain are refused rather than dropped. The gate is that an
unedited form serializes to the bytes it was read from.

Four details cost the most and none are obvious from a first reading of
[MS-OFORMS]:

- a site's `cbSite` counts from the **mask**, so the next site begins at
  `start + 4 + cbSite`;
- mask **bit 8 carries no fixed field**; reading two bytes for it puts
  every name two characters late;
- a `MultiPage`'s `f` carries a trailing MultiPage record after the
  FormControl, so the sites do not close the stream exactly there. It is
  version-stamped and length-prefixed, so the reader checks for it rather
  than merely tolerating a remainder;
- bit 8 selects no DataBlock field but does select an `fmPosition` in the
  ExtraDataBlock, between the Name/Tag strings and the rest.

Three things a *written* form needs that reading never reveals, each
found by Excel refusing the result:

- **`NextAvailableID` is the highest id already handed out**, not the next
  free one. A new control takes `NextAvailableID + 1`; using the field
  as-is repeats the last control's id and MSForms refuses the form.
- **MorphData's mask bit 31 is reserved and MUST be 1**
  ([MS-OFORMS] 2.2.5.2). Setting it is the single change that makes a new
  `TextBox` or `OptionButton` load.
- **A designer edit must invalidate the `_VBA_PROJECT` performance cache.**
  Adding or removing a control changes the form class's members, and with
  a stale cache Office loads a member list the form no longer matches.
- **A container's storage is bound by its CLSID and its `\x01CompObj`**,
  which names the kind fm20 should treat it as. Get either wrong and the
  container loads without erroring and simply does not appear, so both are
  reproduced verbatim from an Excel-authored fixture.
- **A container's site is not a leaf's.** It carries `BitFlags` and no
  `ObjectStreamSize`, because its record is the `f` of its own storage
  rather than a slice of the parent's `o`.
- **`NextAvailableID` is per container, and it is the highest id anywhere
  beneath it** -- the fixture's MultiPage carries 11, which is a control
  two levels down on one of its pages. A new id is recorded on every
  ancestor up to the form.

Composing a form from nothing needs one more thing the format does not
announce: the **empty class-table count word**. With `BooleanProperties`
defaulted, `DONTSAVECLASSTABLE` is 0 and fm20 reads a count before
`CountOfSites`; omit the word and the low bytes of `CountOfSites` are read
*as* the count. An empty form survives that by luck -- both are zero -- and
the form's first control makes the misread count 1, so fm20 parses garbage
as class info and refuses the whole form.

A form is also two things at once: a designer storage and a code-behind
module, declared `BaseClass=` (not `Class=`) in the PROJECT stream. A
storage without a module is not a component the host shows; a module
without a storage is an ordinary class.

A page is four structures at once, and `_oforms_pages.py` moves them
together: a site and a storage of its own; an entry in each of the
MultiPage's five parallel TabStrip arrays (`Items`, `TipStrings`,
`TabNames`, `Tags`, `Accelerators`) with one flag word per tab after
them; `TabData` and `TabsAllocated`; and its position in the `x` stream,
which holds one more `PageProperties` record than there are pages (the
first ignored) followed by the page site ids. Page names are scoped to
their MultiPage, not to the form -- Excel gives a second MultiPage its own
`Page1` and `Page2` -- and pages never appear in `Designer.Controls`.

Nesting is resolved by matching a child storage's numeric suffix against
a site id, not by rebuilding the storage name from the id: the file says
which storages exist, and the padding of that name is only observed.

Reading `f` and `o` needs path-addressed CFB navigation
(`list_storages_at` / `get_stream_at`), because names repeat -- every
container owns an `f` -- and a name-based lookup finds whichever comes
first in directory order.

---

## 6. Encoding conventions

- The VBA `PROJECTCODEPAGE` value (typically `1252`) drives every
  MBCS encode/decode in `vba.py`. Never hardcode `cp1252`.
- UTF-16LE strings in the project (the `*UNICODE` records and the
  unicode halves of `PROJECT*` records) are **not BOM-prefixed**.
- Module source is normalized to CRLF (`\r\n`) line endings on write.
  Read returns whatever the file contained.
- Disk pull/push uses UTF-8 by default, configurable via the
  `encoding=` kwarg on both helpers.

---

## 7. Exception hierarchy

```
PyOpenVBAError
    UnsupportedFormatError      bad extension / host
    CFBError                    malformed CFB structure
    VBAProjectError             malformed VBA project, or mutation refused
```

Rules:

- Every parser MUST raise from this hierarchy (or `ValueError` for
  caller-input bugs). No bare `Exception` and no `KeyError` escaping
  to user code.
- Mutation refusals (protected, missing module, name collision) raise
  `VBAProjectError` with a message that names the affected module or
  stream.

---

## 8. Test layout

```
tests/
  test_cfb.py                  CFB primitives + round-trip
  test_vba.py                  MS-OVBA codec + dir/PROJECT parsers
  test_excel.py                ExcelFile facade, end-to-end
  test_word.py                 WordFile facade, end-to-end
  test_powerpoint.py           PowerPointFile facade, end-to-end
  test_pull_push.py            Disk workflow
  test_gates.py                Per-roadmap-gate regression tests
  test_access.py               AccessReader read path: page walk,
                               LVAL chain walk, MS-OVBA blob decode,
                               byte-for-byte oracle parity (EXPERIMENTAL)
  test_access_engine.py        Storage engine read layer against the
                               Access-authored fixtures
  test_live_access_engine_gate.py
                               Opt-in (RUN_LIVE_ACCESS=1): the ACE engine,
                               driven through DAO by
                               live_access_test/dao_oracle.ps1, builds every
                               column type; the engine must read it back
                               field for field
  fuzz_corpus/                 Persistent fuzz seeds (see test_gates.py Gate 23)
  live_excel_testing/          Real fixture workbooks (Excel)
  live_word_testing/           Real fixture documents (Word)
  live_powerpoint_testing/     Real fixture presentations (PowerPoint)
  live_access_test/            Real fixture databases (Access) + COM oracle
```

- **Always** run pytest with `-p no:randomly` to keep ordering
  reproducible.
- **Strict Pyright** (0 errors) must pass on `src/` and `tests/`
  before any merge: `pyright src tests`.
- New behavior changes must land with a test in the same commit.

---

## 9. Conventions and house rules

- **No runtime dependencies.** The stdlib only. Dev tools (`pytest`,
  `pyright`) are not part of `pyproject.toml` `dependencies`.
- **Python 3.10+.** Use `|` union types, `match` where natural,
  `dataclasses` for record types.
- **No silent corruption.** Any code path that could leave a workbook
  inconsistent must either succeed completely or raise. Partial state
  is never written to disk.
- **Preserve what you don't understand.** Bytes the library does not
  interpret (PROJECTlk, signature payloads on no-op saves) are
  round-tripped verbatim -- and so are the parts of a structure it *does*
  interpret but cannot explain: a control record's alignment padding, an
  unmodelled tail, a CompObj blob.
- **Drop perf caches on mutation.** `_VBA_PROJECT` body is zeroed,
  `__SRP_*` streams are removed. Office regenerates both.
- **ASCII only in user-facing strings.** Warnings, error messages,
  CLI output: no emoji, no smart quotes, no en-dashes.

---

## 10. Files to keep up to date

When making the listed structural changes, update **all** of these
files in the same commit:

| Change                                       | Files to update                                   |
|----------------------------------------------|---------------------------------------------------|
| New public API symbol                        | `__init__.py` `__all__`, README, this file        |
| New layer / module                           | This file (sections 1 and 2), README architecture box |
| New roadmap gate or status change            | `docs/roadmap.md`                                 |
| New mutation safety gate                     | This file section 4, README "Safety guards"       |
| New supported file extension                 | Facade `_zip_formats` / `_cfb_formats`, `__main__._HOST_BY_SUFFIX`, README table, this file section 5 |
| New test file or fuzz target                 | This file section 8                               |
| Spec/behavior change worth telling other implementers | `docs/ms-ovba-implementation-guide_v2.md` |

This table is the canonical sync checklist.
