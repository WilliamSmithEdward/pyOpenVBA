# Changelog

All notable changes to pyOpenVBA are documented here. This project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
