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
- Adding, renaming, or deleting modules.
- UserForm layout / Office Forms editing (form layout bytes survive
  verbatim through the CFB round-trip but the library does not interpret
  them).
- ActiveX license editing (PROJECTlk).
- Project password / protection editing.
- Digital signature detection or re-signing.
- Content-hash / agile-content-hash recomputation.
- Non-ASCII module names (untested — no corpus fixture).
- PROJECTwm name-mapping editing.
- Full PROJECTINFORMATION and REFERENCE record decoding (records survive
  verbatim via the unchanged `dir` stream, but are not exposed in the Python
  data model).

## Gate-by-gate status

| Gate | Title | Status | Notes |
|------|-------|--------|-------|
| 0 | Scope Declaration | PASS | `ExcelFile` rejects unsupported hosts; CFB and VBA layers separated; `vba_project_bytes()` exposes raw `vbaProject.bin`. |
| 1 | Host Package | PARTIAL | No-op + single-module-edit save preserves every other ZIP entry. "Opens in Excel without repair" is not asserted from Python. |
| 2 | OLE/CFB Container | PARTIAL | Reader + writer round-trip; case-insensitive lookup. SRP-stream dropping on write is TODO (requires CFB directory-entry removal API). |
| 3 | Binary Parsing Discipline | PARTIAL | Bounds-checked, signature-checked. Error messages do not yet carry stream name + offset metadata. |
| 4 | Compression / Decompression | PASS | Spec-compliant chunk-based codec; randomized round-trips up to 32 KB. |
| 5 | `_VBA_PROJECT` / Performance Cache | PARTIAL | Module performance-cache prefix preserved verbatim across writes. Stale-cache invalidation logic that Office uses (e.g. zeroing the `_VBA_PROJECT` stream contents) is not yet performed. |
| 6 | PROJECT Stream | VERBATIM | Stream survives round-trip unchanged; grammar not parsed. |
| 7 | PROJECTwm | VERBATIM | Stream survives round-trip unchanged when present. |
| 8 | PROJECTlk | VERBATIM | Stream survives round-trip unchanged when present. |
| 9 | dir Project Information | PARTIAL | PROJECTCODEPAGE is decoded; other records survive in the unchanged `dir` bytes. |
| 10 | dir References | VERBATIM | Reference records survive in the unchanged `dir` bytes; not exposed in the model. |
| 11 | dir Module Records | PARTIAL | Module name, stream name, offset, type, read-only, private parsed. MODULEDOCSTRING, MODULEHELPCONTEXT, MODULECOOKIE not decoded. |
| 12 | Module Stream | PASS | Source decompressed from `MODULEOFFSET`; replacement preserves cache prefix; reparse yields identical source. |
| 13 | Module Mutation | PARTIAL | Replace works for standard/class/document modules. Add / rename / delete not implemented. |
| 14 | Designer / UserForm | VERBATIM | Designer storages survive round-trip; no fixture/test yet. |
| 15 | Content Hash / Integrity | TODO | V3 content hash, agile content hash, project normalized data — none implemented. |
| 16 | Protection / Encryption / Password | TODO | Protected projects round-trip verbatim but cannot be safely modified. |
| 17 | Digital Signature | TODO | Signed projects round-trip verbatim. Editing a signed project will leave a stale signature; the library does not yet report this. |
| 18 | Encoding | PARTIAL | cp1252 source round-trips through the project code page. Non-ASCII module names / source untested. |
| 19 | Cross-Structure Consistency | PASS | `VBAProject.validate(cfb)` reports duplicates and missing streams. |
| 20 | Round-Trip Preservation | PARTIAL | No-op parse-write-reopen preserves every module source, every ZIP entry, and every module-stream cache prefix. "Opens in Excel" requires manual verification. |
| 21 | Mutation Round-Trip | PARTIAL | Replace-source mutations round-trip for standard, class, and document modules. UserForm code-behind / add / rename / delete pending. |
| 22 | Corpus | PARTIAL | One fixture: `tests/live_excel_testing/test_macro_workbook.xlsm` (standard + class + document modules). UserForm, ActiveX, non-ASCII, password-protected, and signed fixtures pending. |
| 23 | Fuzz / Malformed Input | PARTIAL | Truncated and zero-length inputs fail cleanly. No structured fuzz corpus yet. |
| 24 | API Contract | PASS | Layered modules: `pyopenvba.cfb`, `pyopenvba.vba`, `pyopenvba.excel`. Mutation surface (`add_module`/`rename_module`/`delete_module`) pending. |
| 25 | Documentation | PARTIAL | This roadmap exists. README needs to be expanded with the scope statement above. |

## Near-term roadmap (in priority order)

1. **PROJECT stream grammar parser** (Gate 6). Required for safe add/rename/delete (Gate 13) because every mutation must update the `Module=...` declarations in `PROJECT` alongside `dir` and the CFB stream name.
2. **`dir` stream writer** (Gate 11 full). Required for any mutation that changes module identity — replace-source today works only because we never rewrite `dir`.
3. **`add_module`, `rename_module`, `delete_module`** (Gate 13). Builds on (1) and (2).
4. **PROJECTwm reader/writer** (Gate 7). Required as soon as non-ASCII module names enter the corpus.
5. **Non-ASCII corpus fixture + Gate 18 hardening**.
6. **Signature detection** (Gate 17). Detect-and-report only is enough to fail closed on signed workbooks; full re-signing is out of scope.
7. **Content-hash recomputation** (Gate 15). Required before signature detection can do better than refuse-to-edit.
8. **UserForm corpus fixture + Gate 14 round-trip assertion**.
9. **Protected-project detection** (Gate 16). Detect-and-refuse is the minimum; password manipulation is explicitly out of scope.
10. **CFB stream-removal API + SRP-drop on write** (Gate 2 closure). Per MS-OVBA 2.3.4.1, SRP streams MUST be ignored on read and should not be emitted by non-host writers. Implementation requires red-black-tree directory surgery in `CFB.to_bytes()`.

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
