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
| excel.py        ExcelFile facade                       |
|   - ZIP / raw-CFB dispatch by extension                |
|   - save() pipeline (safety gates, structural rewrites)|
|   - pull / push disk workflow                          |
+--------------------------------------------------------+
| word.py         WordFile facade                        |
|   - ZIP / raw-CFB dispatch by extension                |
|   - identical save() pipeline as ExcelFile             |
|   - pull / push disk workflow                          |
+--------------------------------------------------------+
| powerpoint.py   PowerPointFile facade                  |
|   - ZIP / raw-CFB dispatch by extension                |
|   - identical save() pipeline as ExcelFile             |
|   - pull / push disk workflow                          |
+--------------------------------------------------------+
| access.py       AccessFile facade  (EXPERIMENTAL)      |
|   - ACE / Jet 4 page reader (.accdb, .mdb)             |
|   - MS-OVBA blob discovery via signature scan          |
|   - LVAL page-chain walker                             |
|   - read_vba_module(name) -> str (pure Python)         |
|   - replace_text(old, new) same-length plaintext patch |
|   - replace_text_resize(old, new) reflow-aware patch   |
|   - write_lval_row / lval_free_space LVAL primitives   |
|   - rename_module / delete_module / add_module_catalog |
|   - modify_module_cache / replace_module               |
|   - export_modules(dir) / import_module(path|source)   |
|   - iter_msys_objects / find_msys_module               |
|       (MSysObjects system catalog reader)              |
|   - save([path]) -> persists in-memory edits           |
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
- `excel.py` is the only module that touches the filesystem, ZIP
  containers, or workbook paths.
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

### 3.3 Access write model (`AccessFile`)

`.accdb`/`.mdb` does **not** follow the OOXML/CFB write pipeline above
because Access stores VBA source very differently from Excel/Word/PowerPoint:

* The MS-OVBA blob on the LVAL chain is a **passive cache**. Zero-filling
  the entire blob has no effect on what Access (or the VBA editor)
  displays — verified end-to-end on a live fixture.
* The **authoritative** sources Access reads from are:
  * **Project metadata** — the MS-OVBA `dir` stream (Section 2.3.4.2)
    OVBA-compressed in a single LVAL row; located by content
    (decompressed prefix = `01 00 04 00 00 00`). Parsed by
    `AccessFile.read_project_info()` -> `AccessVBAProject` (system kind,
    LCID, code page, project name, references, modules with class flag,
    private/read-only flags).
  * **Module bytecode** — the compiled VBA p-code. Lives in LVAL rows
    whose payload starts with magic `72 55 40 00` ('rU@\0'). Each
    database has 2-3 such rows; the **module-active** one is identified
    deterministically by the 12-byte prefix
    `72 55 40 00 00 00 00 00 00 00 40 00` (byte 10 = 0x40). Exposed via
    `AccessFile.read_module_pcode_stream()` and `iter_pcode_streams()`.
    Compiled p-code is **fully anonymised**: it contains opcodes and
    slot references but no user-authored text. Procedure names,
    module names, comments, and string literals all live in separate
    catalog / symbol / plaintext rows and are referenced from the
    bytecode by slot id only. Opcode field guide in progress; see
    `docs/access_pcode_re.md` Phase 4.
  * **Identifier inventory** — every project-level identifier name
    (typelib refs, project name, module/proc/variable names, intrinsic
    call targets such as `MsgBox`) is enumerated in the
    `_VBA_PROJECT`-equivalent stream stored UNCOMPRESSED in the LVAL
    row whose first 2 bytes are the `CC 61` Office magic. Each entry
    on disk uses the record layout
    `<u8 len><u8 type>[<6B descriptor>]<ASCII name><u16 id><10 00>`.
    Exposed via `AccessFile.identifiers()` ->
    `tuple[AccessVBAIdentifier, ...]`. Note: this is a project-wide
    inventory only; the `name_id` operands in compiled p-code are
    per-procedure reference-table slots (resolution pending).
  * **Comment text** — stored verbatim in plaintext rows tagged
    `E3 00 00 00 <u16-LE length> <ASCII payload>` (apostrophe stripped).
  * **String literal text** — stored verbatim in rows tagged
    `B9 00 <u16-LE length> <ASCII payload> <12-byte trailer>`.

Consequently the only safe pure-Python write primitive today is
`AccessFile.replace_text(old, new)`: a **same-length byte substitution**
inside the plaintext rows. This is sufficient to patch comment bodies
and string literal contents (verified through Access COM and the live
VBA editor). Changing literal/comment lengths or modifying code structure
(renaming subs, adding statements) is out of scope until the p-code
tables are reverse-engineered.

For **length-changing** edits inside a single LVAL row,
`AccessFile.replace_text_resize(old, new)` reflows the affected page's
slot table in place: the row is grown/shrunk, neighbour rows shift
within the page, and other pages are untouched. The underlying LVAL
mutation primitives (`_lval_resize_row`, `_lval_write_row`,
`_lval_tombstone_slot`, `_lval_append_row`) are also exposed publicly
as `write_lval_row` and `lval_free_space` for callers that have
located a specific row via `_iter_lval_rows`. Edits that would (a)
straddle row boundaries, (b) overflow the page free space, or (c)
require allocating a new LVAL page (chain growth) are rejected with
`AccessError`.

Module **rename / delete / add / source-cache modify** are exposed
as catalog-level operations:
`AccessFile.rename_module(old, new)` rewrites MODULENAME +
MODULENAMEUNICODE in the dir-stream and updates the `Attribute
VB_Name` line inside the corresponding OVBA cache row.
`AccessFile.delete_module(name)` excises the module's record block
from the dir-stream, decrements PROJECTMODULES count, and tombstones
every matching OVBA cache row (Access keeps redundant copies).
`AccessFile.add_module_catalog_entry(name, *, is_class_module,
is_private, is_read_only)` appends a new module record block before
DIR-TERM. `AccessFile.modify_module_cache(name, new_source)` rewrites
the OVBA cache row's decompressed payload (Attribute VB_Name line
preserved). `AccessFile.replace_module(name, new_source)` is a
convenience alias. `AccessFile.export_modules(dest_dir)` writes every
module to `<name>.bas` / `<name>.cls` on disk (class modules detected
via the dir-stream MODULETYPE record); `AccessFile.import_module(
path_or_source, *, name=None, is_class_module=None, replace_existing
=False)` adds a new module from a `.bas`/`.cls` file or raw source
string, optionally overwriting an existing module's cache. These
operations make the changes visible to pure-Python readers
(`read_project_info`, `iter_vba_modules`); they do NOT update the
Access engine's MSysObjects catalog or p-code tables, so the live
Access VBA editor may not immediately reflect every change. Those
engine-level writes remain documented as future RE phases in
`memories/repo/access-vba-storage.md`.

`AccessFile.iter_msys_objects()` (and the convenience helpers
`msys_objects()`, `iter_msys_modules()`, `find_msys_object(name,
*, type_=None)`, `find_msys_module(name)`) read the Jet/ACE
**MSysObjects** system catalog: the master index of every persistent
object inside a .accdb (tables, queries, forms, reports, scripts,
modules, and the container hubs that group them). Each row is
surfaced as an `AccessSysObject` dataclass (id_, parent_id, type_,
flags, name, page, slot). VBA code modules carry `Type ==
MSYS_TYPE_MODULE` (-32761, 0x8007) and have `parent_id` equal to
the Id of the row named `Modules` (Type=3 container). This reader
is the foundation for the future write path that will make
`rename_module` / `delete_module` / `add_module_catalog_entry`
visible inside the live Access UI.

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
| `.ppt`           | raw CFB (PPT 97)       | (whole file is CFB)    |
| anything else    | n/a                    | `UnsupportedFormatError` |

For the ZIP case, the VBA project is at the fixed path
`xl/vbaProject.bin`. On save, exactly that entry is replaced while
every other ZIP entry is preserved byte-for-byte including its
compression method, external attributes, create system, and timestamp.

For `.xls`, the entire file *is* the CFB; `cfb.to_bytes()` is written
straight to the output path.

If a `.xlsm` exists but does not contain `xl/vbaProject.bin` (no VBA
project has ever been created), `ExcelFile` raises a structured
`VBAProjectError` with the workbook path — it is **not** a corruption.

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
  test_access.py               AccessFile read path: page walk,
                               LVAL chain walk, MS-OVBA blob decode,
                               byte-for-byte oracle parity (EXPERIMENTAL)
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
- **Preserve what you don't understand.** Bytes that the library does
  not interpret (designer sub-storages, PROJECTlk in v1, signature
  payloads on no-op saves) are round-tripped verbatim.
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
| New supported file extension (Excel)         | `excel.py` dispatch, README table, this file section 5       |
| New supported file extension (Word)          | `word.py` dispatch, README table, this file section 5        |
| New supported file extension (PowerPoint)    | `powerpoint.py` dispatch, README table, this file section 5  |
| New test file or fuzz target                 | This file section 8                               |
| Spec/behavior change worth telling other implementers | `docs/ms-ovba-implementation-guide_v2.md` |

This table is the canonical sync checklist.
