# pyOpenVBA Roadmap

This roadmap tracks pyOpenVBA's progress against the 26 hard gates defined in
[`xlsm_feature_completeness_gates.md`](../xlsm_feature_completeness_gates.md).
Per-gate status keys:

- **PASS** — implemented; gate tests in [`tests/test_gates.py`](../tests/test_gates.py) pass.
- **PARTIAL** — minimum bar met; secondary assertions are `xfail`.
- **TODO** — not implemented; gate tests are `xfail(strict=True)`.
- **VERBATIM** — bytes are preserved through round-trip but not interpreted.

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
- ActiveX license editing (PROJECTlk).
- Project password / protection editing (parsing-only).
- Digital signature re-signing (detection-only).
- Office-compatible V3 / agile content-hash recomputation (a stable
  pyOpenVBA-internal digest is available via `compute_v3_content_hash`).
- Non-ASCII module names (untested -- no corpus fixture).
- PROJECTwm name-mapping editing (reader-only).

## Gate-by-gate status

| Gate | Title | Status | Notes |
|------|-------|--------|-------|
| 0 | Scope Declaration | PASS | `ExcelFile` rejects unsupported hosts; CFB and VBA layers separated; `vba_project_bytes()` exposes raw `vbaProject.bin`. |
| 1 | Host Package | PARTIAL | No-op + single-module-edit save preserves every other ZIP entry. "Opens in Excel without repair" is not asserted from Python. |
| 2 | OLE/CFB Container | PASS | Reader + writer round-trip; case-insensitive lookup; `CFB.remove_stream` / `drop_streams_in_storage` / `rename_stream_in_storage` / `add_stream_to_storage` all rebuild the directory subtree; SRP streams are auto-dropped on `ExcelFile.save`. |
| 3 | Binary Parsing Discipline | PASS | Bounds-checked, signature-checked; `decompress()` carries `stream_name` + byte offset in `VBAProjectError` messages. |
| 4 | Compression / Decompression | PASS | Spec-compliant chunk-based codec; randomized round-trips up to 32 KB. |
| 5 | `_VBA_PROJECT` / Performance Cache | PARTIAL | Module performance-cache prefix preserved verbatim across writes. Stale-cache invalidation logic that Office uses (e.g. zeroing the `_VBA_PROJECT` stream contents) is not yet performed. |
| 6 | PROJECT Stream | PASS | `parse_project_stream()` decodes the full plain-text grammar (project section, `[Host Extender Info]`, `[Workspace]`). `serialize_project_stream()` rewrites `Module=`/`Class=`/`BaseClass=`/`Document=` and `[Workspace]` keys to follow logical-name renames. |
| 7 | PROJECTwm | PASS | `parse_projectwm()` decodes (MBCS, Unicode) module-name pairs; writer is pending. |
| 8 | PROJECTlk | PASS | `parse_projectlk()` decodes `LicenseInfoRecord`s; writer is pending. |
| 9 | dir Project Information | PASS | All PROJECTINFORMATION records decoded: code page, name, SysKind, LCID(invoke), DocString, HelpFile, HelpContext, LibFlags, Version, Constants, CompatVersion. |
| 10 | dir References | PASS | REFERENCENAME / REFERENCEREGISTERED / REFERENCEPROJECT / REFERENCECONTROL / REFERENCEORIGINAL records exposed as `VBAReference` entries on `VBAProject.references`. |
| 11 | dir Module Records | PASS | Module name (MBCS + Unicode), stream name (MBCS + Unicode), offset, type, read-only, private, doc-string (MBCS + Unicode), help-context, cookie all decoded. `serialize_dir_modules_section()` re-emits the full block. |
| 12 | Module Stream | PASS | Source decompressed from `MODULEOFFSET`; replacement preserves cache prefix; reparse yields identical source. |
| 13 | Module Mutation | PASS | Replace, add, rename, and delete all persist end-to-end (CFB stream create/rename/remove + dir rewrite + PROJECT rewrite). |
| 14 | Designer / UserForm | PASS | UserForm sub-storage and all four designer child streams (`f`, `o`, `\x01CompObj`, `\x03VBFrame`) survive a no-op save byte-for-byte on the live xlsm fixture (`test_designer_storage_preserved`). Generic sub-storage round-trip is also covered (`test_synthetic_substorage_round_trips_through_cfb`). |
| 15 | Content Hash / Integrity | PARTIAL | `compute_v3_content_hash()` provides a stable SHA-1 digest over normalized module sources. The Office-compatible V3 / agile content hash (host-specific tokenization) is not implemented. |
| 16 | Protection / Encryption / Password | PARTIAL | `ProjectProtection` exposes raw obfuscated CMG/DPB/GC plus a `has_password` heuristic. Password decryption / re-encryption is not implemented. |
| 17 | Digital Signature | PARTIAL | `detect_signature()` identifies legacy / agile / V3 signature streams. Editing a signed project will still leave a stale signature; re-signing is out of scope. |
| 18 | Encoding | PARTIAL | Latin-1 supplement module names + source round-trip end-to-end on a cp1252 project (`test_latin1_supplement_module_name_round_trip`). Non-cp1252 project code pages (e.g. CJK, Cyrillic) still need a fixture. |
| 19 | Cross-Structure Consistency | PASS | `VBAProject.validate(cfb)` reports duplicates and missing streams. |
| 20 | Round-Trip Preservation | PARTIAL | No-op parse-write-reopen preserves every module source, every ZIP entry, and every module-stream cache prefix. "Opens in Excel" requires manual verification. |
| 21 | Mutation Round-Trip | PASS | Replace-source, add, rename, and delete mutations all round-trip through save/reopen (parsed model, CFB streams, dir stream, and PROJECT stream all consistent). UserForm code-behind edits persist while the sibling designer sub-storage stays byte-for-byte identical (`test_replace_userform_code_behind_round_trip`). |
| 22 | Corpus | PARTIAL | One fixture: `tests/live_excel_testing/test_macro_workbook.xlsm` (standard + class + document + UserForm modules with Office Forms 2.0 designer storage). ActiveX, non-cp1252-codepage, password-protected, and signed fixtures pending. |
| 23 | Fuzz / Malformed Input | PARTIAL | Truncated and zero-length inputs fail cleanly. No structured fuzz corpus yet. |
| 24 | API Contract | PASS | Layered modules: `pyopenvba.cfb`, `pyopenvba.vba`, `pyopenvba.excel`. Mutation surface (`add_module`/`rename_module`/`delete_module`) persists end-to-end through `save()`. |
| 25 | Documentation | PARTIAL | This roadmap exists. README needs to be expanded with the scope statement above. |

## Near-term roadmap (in priority order)

1. **PROJECTwm writer** (Gate 7). Required as soon as non-ASCII module names enter the corpus.
2. **Non-cp1252 (CJK / Cyrillic) corpus fixture + Gate 18 hardening** for projects whose code page is not Western European.
3. **Office-compatible V3 / agile content hash** (Gate 15 full). The current SHA-1 digest is stable for internal use but does not match Excel's signature payload.
4. **Additional corpus fixtures** (ActiveX, password-protected, signed, non-cp1252 codepage) to close the Gate 22 catalog.
5. **Protected-project refuse-to-edit gate** (Gate 16 hardening). Save-on-protected-project should fail closed unless the caller opts in.
6. **Signed-project staleness reporting** (Gate 17 hardening). Save-on-signed-project should warn that the signature will be invalidated.

## Out of scope (no current plans)

- Re-signing modified projects with arbitrary PKCS#7 / VBA digital signatures.
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
