# pyOpenVBA Roadmap

This roadmap tracks pyOpenVBA's progress against the 26 hard gates defined in
[`xlsm_feature_completeness_gates.md`](../xlsm_feature_completeness_gates.md).
Per-gate status keys:

- **PASS** — implemented; gate tests in [`tests/test_gates.py`](../tests/test_gates.py) pass.
- **PARTIAL** — minimum bar met; secondary assertions are `xfail`.
- **TODO** — not implemented; gate tests are `xfail(strict=True)`.
- **VERBATIM** — bytes are preserved through round-trip but not interpreted.
- **OUT OF SCOPE** — explicitly excluded from current development; tests are `skip` with a reason.

## Current scope declaration

pyOpenVBA today is best described as:

> **"MS-OVBA round-tripper with module source reader/writer (.xlsm focus)."**

### Supported
- `.xlsm`, `.xlsb`, `.xlam` (OOXML ZIP containers with `xl/vbaProject.bin`).
- `.xls` (legacy BIFF8 — vbaProject.bin is the entire file).
- Read all VBA module sources (standard, class, document, designer code-behind).
- Replace source of any existing module.
- **Disk-based push/pull** (`pyopenvba.pull` / `pyopenvba.push`, and
  `python -m pyopenvba {pull,push,ls}`) — export modules to `.bas` / `.cls`
  files for use with any text editor / version control, then push edits back
  into the workbook.
- No-op round-trip preservation of every non-VBA ZIP entry.
- No-op round-trip preservation of every module's performance-cache prefix.
- Pure Python 3.10+, zero runtime dependencies.

### Unsupported (today)
- UserForm layout / Office Forms editing (form layout bytes survive
  verbatim through the CFB round-trip but the library does not interpret
  them).
- ActiveX license editing (PROJECTlk). ActiveX controls are deprecated; license bytes are round-tripped verbatim.
- Project password / protection editing (parsing-only; save refuses to mutate protected projects unless `allow_protected=True`).
- Digital signature re-signing (out of scope; stale signature streams are dropped on mutating save with a `UserWarning`).
- Office-compatible V3 / agile content-hash recomputation (a stable
  pyOpenVBA-internal digest is available via `compute_v3_content_hash`).
- Non-ASCII module identifiers (Excel's VBA IDE does not permit them; the parser nevertheless round-trips Latin-1 supplement names through cp1252).

## Gate-by-gate status

| Gate | Title | Status | Notes |
|------|-------|--------|-------|
| 0 | Scope Declaration | PASS | `ExcelFile` rejects unsupported hosts; CFB and VBA layers separated; `vba_project_bytes()` exposes raw `vbaProject.bin`. |
| 1 | Host Package | PASS | No-op + single-module-edit save preserves every other ZIP entry on both xlsm and xlsb. `xlsm` workbooks that have never had a VBA project initialised raise a structured `VBAProjectError` at open. "Opens in Excel without repair" is verified manually against the live corpus. |
| 2 | OLE/CFB Container | PASS | Reader + writer round-trip; case-insensitive lookup; `CFB.remove_stream` / `drop_streams_in_storage` / `rename_stream_in_storage` / `add_stream_to_storage` all rebuild the directory subtree; SRP streams are auto-dropped on `ExcelFile.save`. |
| 3 | Binary Parsing Discipline | PASS | Bounds-checked, signature-checked; `decompress()` carries `stream_name` + byte offset in `VBAProjectError` messages. |
| 4 | Compression / Decompression | PASS | Spec-compliant chunk-based codec; randomized round-trips up to 32 KB. |
| 5 | `_VBA_PROJECT` / Performance Cache | PASS | Module performance-cache prefix preserved verbatim across no-op writes. On any mutating save (add / rename / delete / source edit) the `_VBA_PROJECT` stream body is zeroed (5-byte header preserved) so Office regenerates the cache on next open ([MS-OVBA] 2.3.4.1: PerformanceCache MUST be ignored on read). Writers never emit `__SRP_*` streams. |
| 6 | PROJECT Stream | PASS | `parse_project_stream()` decodes the full plain-text grammar (project section, `[Host Extender Info]`, `[Workspace]`). `serialize_project_stream()` rewrites `Module=`/`Class=`/`BaseClass=`/`Document=` and `[Workspace]` keys to follow logical-name renames. |
| 7 | PROJECTwm | PASS | `parse_projectwm()` + `serialize_projectwm()` round-trip the live fixture byte-for-byte; `ExcelFile.save()` rebuilds `PROJECTwm` whenever the module set changes (add / rename / delete). |
| 8 | PROJECTlk | PASS | `parse_projectlk()` + `serialize_projectlk()` round-trip `LicenseInfoRecord`s; `ExcelFile.save()` preserves PROJECTlk bytes verbatim until ActiveX license editing is required. |
| 9 | dir Project Information | PASS | All PROJECTINFORMATION records decoded: code page, name, SysKind, LCID(invoke), DocString, HelpFile, HelpContext, LibFlags, Version, Constants, CompatVersion. |
| 10 | dir References | PASS | REFERENCENAME / REFERENCEREGISTERED / REFERENCEPROJECT / REFERENCECONTROL / REFERENCEORIGINAL records exposed as `VBAReference` entries on `VBAProject.references`. |
| 11 | dir Module Records | PASS | Module name (MBCS + Unicode), stream name (MBCS + Unicode), offset, type, read-only, private, doc-string (MBCS + Unicode), help-context, cookie all decoded. `serialize_dir_modules_section()` re-emits the full block. |
| 12 | Module Stream | PASS | Source decompressed from `MODULEOFFSET`; replacement preserves cache prefix; reparse yields identical source. |
| 13 | Module Mutation | PASS | Replace, add, rename, and delete all persist end-to-end (CFB stream create/rename/remove + dir rewrite + PROJECT rewrite). |
| 14 | Designer / UserForm | PASS | UserForm sub-storage and all four designer child streams (`f`, `o`, `\x01CompObj`, `\x03VBFrame`) survive a no-op save byte-for-byte on the live xlsm fixture (`test_designer_storage_preserved`). Generic sub-storage round-trip is also covered (`test_synthetic_substorage_round_trips_through_cfb`). |
| 15 | Content Hash / Integrity | OUT OF SCOPE | `compute_v3_content_hash()` provides a stable SHA-1 digest over normalized module sources for internal use. The Office-compatible V3 / agile content hash (host-specific tokenization that matches Excel's signature payload) is intentionally out of scope; reaching parity would require Excel-side reference vectors and only matters for re-signing (also out of scope, see Gate 17). |
| 16 | Protection / Encryption / Password | PASS | `ProjectProtection` exposes raw obfuscated CMG/DPB/GC plus `has_password`. A real password-protected fixture (`workbook_with_password_protected_vba_modules.xlsm`) is parsed end-to-end. `ExcelFile.save()` refuses to mutate a protected project unless `allow_protected=True`; with the opt-in, the password material is preserved verbatim. Password decryption / re-encryption is intentionally out of scope. |
| 17 | Digital Signature | OUT OF SCOPE | `detect_signature()` identifies legacy / agile / V3 signature streams. `ExcelFile.save()` drops stale signature streams when the project is mutated and emits a `UserWarning` (silenced with `allow_invalidate_signature=True`). Re-signing modified projects with PKCS#7 / VBA digital signatures is intentionally out of scope. |
| 18 | Encoding | PASS | Latin-1 supplement module names + source round-trip end-to-end on a cp1252 project (`test_latin1_supplement_module_name_round_trip`). Non-cp1252 module identifiers are not exercised because Excel's VBA IDE does not permit them (out of scope). |
| 19 | Cross-Structure Consistency | PASS | `VBAProject.validate(cfb)` reports duplicates and missing streams. |
| 20 | Round-Trip Preservation | PASS | No-op parse-write-reopen preserves every module source, every ZIP entry, and every module-stream cache prefix. Manual Excel verification covers no-op, source-edit, add, rename, and delete on xlsm; no-op and source-edit on xlsb; and no-op + opt-in source-edit on the password-protected fixture (9/9 open in Excel with no repair dialog; see `scripts/build_excel_verification_set.py`). |
| 21 | Mutation Round-Trip | PASS | Replace-source, add, rename, and delete mutations all round-trip through save/reopen (parsed model, CFB streams, dir stream, and PROJECT stream all consistent). UserForm code-behind edits persist while the sibling designer sub-storage stays byte-for-byte identical (`test_replace_userform_code_behind_round_trip`). |
| 22 | Corpus | PASS | In-scope corpus complete: `test_macro_workbook.xlsm` (std + class + document + UserForm), `test_macro_workbook.xlsb` (binary host), `workbook_with_password_protected_vba_modules.xlsm` (password-protected), `xlsm_file_with_no_vba_entered_yet.xlsm` (no-VBA negative case). ActiveX, signed, and non-ASCII workbooks are explicitly out of scope. |
| 23 | Fuzz / Malformed Input | PASS | Truncated, zero-length, and random-byte inputs fail cleanly. Bit-flip fuzz harnesses exercise the CFB, dir, PROJECT, and PROJECTwm parsers (~120 mutated inputs per run, seeded for reproducibility). A persistent on-disk corpus lives at [`tests/fuzz_corpus/`](../tests/fuzz_corpus/) (~50 seeded inputs across `cfb/`, `decompress/`, `dir/`, `project/`, `projectwm/`); every file is replayed as a parametrized test (`test_persistent_fuzz_corpus`) and new regression seeds can be added by dropping files into the appropriate subdirectory. Regenerate / extend with `python scripts/seed_fuzz_corpus.py` (idempotent and additive). |
| 24 | API Contract | PASS | Layered modules: `pyopenvba.cfb`, `pyopenvba.vba`, `pyopenvba.excel`. Mutation surface (`add_module`/`rename_module`/`delete_module`) persists end-to-end through `save()`. |
| 25 | Documentation | PASS | `README.md` carries the scope statement, supported formats, push/pull workflow, and safety-guard summary; `docs/roadmap.md` tracks per-gate status. |

## Near-term roadmap (in priority order)

_No open near-term items: all in-scope gates are PASS. See the "Out of scope" section below for explicitly deferred work._

## Out of scope (no current plans)

- Re-signing modified projects with arbitrary PKCS#7 / VBA digital signatures (Gate 17).
- Office-compatible V3 / agile content-hash recomputation (Gate 15) — only meaningful as a prerequisite for re-signing.
- Breaking, removing, or bypassing project protection passwords.
- Editing UserForm layout (controls, positions, properties). Form layout
  bytes will continue to be round-tripped verbatim.
- Editing ActiveX licenses beyond preserving `PROJECTlk` verbatim.
- BIFF8 record-level editing of legacy `.xls` workbooks (only the embedded
  CFB / VBA project is supported; the workbook stream is treated as opaque).

## Files to keep up to date

When any structural change lands, update:

- `docs/roadmap.md` (this file).
- `tests/test_gates.py` (remove `xfail` markers as gates become PASS).
- `README.md` scope statement.
- The "Current scope declaration" section above.
