# MS-OVBA Feature-Completeness Gates

An implementation is NOT feature-complete unless every HARD gate below passes.

If any gate is intentionally unsupported, the implementation must advertise itself as a
subset, for example:

- "VBA module extractor only"
- "VBA module source replacer only"
- "MS-OVBA reader with limited writer"
- "No UserForm/designer support"
- "No protected project/password/hash support"

Do not call the implementation "feature complete" unless all HARD gates pass.

---

## GATE 0 — Scope Declaration Gate

HARD PASS requires:

- The implementation explicitly declares its supported host file formats.
- The core implementation can operate directly on a raw `vbaProject.bin`.
- For `.xlsm`, it can locate, extract, modify, and reinsert `/xl/vbaProject.bin`.
- For unsupported hosts such as `.docm`, `.pptm`, `.xlsb`, or legacy `.xls`, the implementation fails with a clear "unsupported host container" error.
- The implementation distinguishes between:
  - Office package handling
  - OLE/CFB handling
  - MS-OVBA project handling
  - VBA source text handling

Failure condition:

- The implementation only handles exported `.bas`, `.cls`, or `.frm` files and claims to support MS-OVBA.

---

## GATE 1 — Host Package Gate

For `.xlsm` support, HARD PASS requires:

- Opens the Office Open XML ZIP package without corrupting unrelated workbook parts.
- Locates the VBA project part correctly.
- Preserves workbook XML, relationships, content types, media, custom properties, and unrelated package parts.
- Replaces only the intended VBA project binary unless explicitly asked to modify more.
- Output workbook opens in Excel without a repair dialog.
- Macros remain visible in the VBA editor after save/open.

Failure condition:

- Excel reports repaired records, removed macros, invalid content types, broken relationships, or inaccessible VBA project.

---

## GATE 2 — OLE/CFB Container Gate

HARD PASS requires:

- Reads and writes `vbaProject.bin` as an OLE Compound File Binary container.
- Handles storages and streams case-insensitively where required.
- Correctly locates the Project Root Storage.
- Correctly locates the `VBA` storage.
- Correctly locates required streams:
  - `PROJECT`
  - `VBA/_VBA_PROJECT`
  - `VBA/dir`
  - every module stream named by `MODULESTREAMNAME`
- Correctly handles optional streams/storages:
  - `PROJECTwm`
  - `PROJECTlk`
  - designer storages
  - `VBFrame`
- Ignores SRP streams on read.
- Does not emit SRP streams on write.
- Preserves unknown/unmodified streams unless the spec requires removal.

Failure condition:

- The implementation treats `vbaProject.bin` as flat bytes or assumes module text is directly visible.

---

## GATE 3 — Binary Parsing Discipline Gate

HARD PASS requires:

- All numeric fields are parsed with correct byte order.
- All record IDs, sizes, reserved fields, and terminators are validated.
- Reads are bounds-checked.
- Malformed streams fail cleanly.
- Unknown/reserved values are preserved when possible.
- Parser does not silently resynchronize after corruption unless explicitly operating in recovery mode.
- Every parse error includes stream name, offset, expected field, and actual value.

Failure condition:

- The parser guesses structure after a failure and continues as if valid.

---

## GATE 4 — Compression/Decompression Gate

HARD PASS requires:

- Implements MS-OVBA compressed container decompression.
- Implements MS-OVBA compressed container compression.
- Handles raw chunks.
- Handles compressed chunks.
- Handles chunk headers.
- Handles token sequences.
- Handles literal tokens.
- Handles copy tokens.
- Correctly computes copy-token offset and length.
- Correctly handles 4096-byte decompressed chunk boundaries.
- Correctly handles empty and boundary-sized streams.
- Passes the spec examples:
  - no compression
  - normal compression
  - maximum compression
- `decompress(compress(bytes)) == bytes` for a large randomized corpus.
- `compress(decompress(existingCompressed))` produces a valid stream accepted by Excel, even if bytes are not identical.

Failure condition:

- The implementation can decompress but cannot recompress valid output.

---

## GATE 5 — `_VBA_PROJECT` / Performance Cache Gate

HARD PASS requires:

- Parses `_VBA_PROJECT` enough to determine version-dependent project information.
- Does not rely on performance caches for semantic data.
- Ignores performance caches on read.
- Writes interoperable output that forces Office to ignore stale performance caches.
- Removes or invalidates stale cache streams when required.
- Ensures stale module performance cache bytes do not override updated source.

Failure condition:

- Replaced source appears correct in the binary but Excel/VBE shows old code because a stale cache was honored.

---

## GATE 6 — `PROJECT` Stream Gate

HARD PASS requires:

- Parses and writes the `PROJECT` stream using the project code page.
- Supports the complete project-property grammar:
  - project ID
  - document modules
  - standard modules
  - class modules
  - designer modules
  - packages
  - help file
  - executable name
  - project name
  - help context ID
  - description
  - version compatibility
  - protection state
  - password data
  - visibility state
  - host extenders
  - workspace records
- Maintains consistency between `PROJECT` module declarations and `dir` MODULE records.
- Preserves property ordering where required by the grammar.
- Handles quoted values and escaped quotes correctly.

Failure condition:

- Module type is inferred only from filename extension or stream name.

---

## GATE 7 — `PROJECTwm` Name Mapping Gate

HARD PASS requires:

- Parses `PROJECTwm` when present.
- Writes `PROJECTwm` when required.
- Correctly maps module names between MBCS and UTF-16.
- Correctly handles non-ASCII module names.
- Correctly handles duplicate-looking names that differ by encoding.
- Keeps `PROJECTwm`, `PROJECT`, and `dir` module names consistent after rename/add/delete.

Failure condition:

- Non-ASCII module names are corrupted, dropped, or renamed unexpectedly.

---

## GATE 8 — `PROJECTlk` / ActiveX License Gate

HARD PASS requires:

- Detects `PROJECTlk` when present.
- Parses license information sufficiently to preserve it.
- Does not destroy ActiveX licensing information during round-trip.
- If writing/modifying ActiveX-related project data, updates `PROJECTlk` consistently.
- If ActiveX license editing is unsupported, fails closed instead of producing a partially corrupt project.

Failure condition:

- Workbooks with ActiveX controls open with missing controls, broken references, or license errors after round-trip.

---

## GATE 9 — `dir` Stream Project Information Gate

HARD PASS requires parsing and writing every required project-information record, including:

- `PROJECTSYSKIND`
- `PROJECTCOMPATVERSION`
- `PROJECTLCID`
- `PROJECTLCIDINVOKE`
- `PROJECTCODEPAGE`
- `PROJECTNAME`
- `PROJECTDOCSTRING`
- `PROJECTHELPFILEPATH`
- `PROJECTHELPCONTEXT`
- `PROJECTLIBFLAGS`
- `PROJECTVERSION`
- `PROJECTCONSTANTS`

Additional requirements:

- Correctly uses `PROJECTCODEPAGE` for MBCS text.
- Preserves UTF-16 shadow fields where present.
- Maintains consistency with the `PROJECT` stream.
- Handles empty optional strings.
- Handles long strings up to spec limits.

Failure condition:

- The implementation hardcodes Windows-1252 or UTF-8 for all projects.

---

## GATE 10 — `dir` Stream Reference Gate

HARD PASS requires parsing and writing all reference record types:

- `REFERENCE`
- `REFERENCENAME`
- `REFERENCECONTROL`
- `REFERENCEORIGINAL`
- `REFERENCEREGISTERED`
- `REFERENCEPROJECT`

Additional requirements:

- Preserves registered type-library references.
- Preserves project references.
- Preserves ActiveX control references.
- Handles absolute paths, relative paths, GUIDs, LCIDs, major/minor versions, cookies, and libids.
- Handles missing optional records allowed by the spec.
- Keeps Excel/VBE references dialog semantically unchanged after round-trip.

Failure condition:

- A workbook opens but VBA references become "MISSING" after save.

---

## GATE 11 — `dir` Stream Module Record Gate

HARD PASS requires parsing and writing every module record field:

- `MODULENAME`
- `MODULENAMEUNICODE`
- `MODULESTREAMNAME`
- `MODULEDOCSTRING`
- `MODULEOFFSET`
- `MODULEHELPCONTEXT`
- `MODULECOOKIE`
- `MODULETYPE`
- `MODULEREADONLY`
- `MODULEPRIVATE`
- module terminator records

Additional requirements:

- Supports document modules.
- Supports standard/procedural modules.
- Supports class modules.
- Supports designer modules.
- Correctly distinguishes module name from module stream name.
- Correctly applies `MODULEOFFSET`.
- Keeps module records in sync with actual module streams.
- Keeps module records in sync with the `PROJECT` stream.

Failure condition:

- Source extraction works only when `MODULEOFFSET == 0`.

---

## GATE 12 — Module Stream Gate

HARD PASS requires:

- Reads each module stream named by `MODULESTREAMNAME`.
- Treats bytes before `MODULEOFFSET` as version-dependent performance cache.
- Ignores module performance cache for source extraction.
- Decompresses `CompressedSourceCode` beginning at `MODULEOFFSET`.
- Decodes decompressed source using the project code page.
- Preserves module attributes.
- Preserves line endings or normalizes only by explicit documented policy.
- Replaces source without corrupting attributes, hidden declarations, or module headers.
- Writes valid compressed source back into the module stream.
- Updates `MODULEOFFSET` if cache strategy changes.
- Ensures Excel/VBE displays the new source after opening.

Failure condition:

- The implementation assumes the whole module stream is compressed source.

---

## GATE 13 — Module Mutation Gate

HARD PASS requires valid operations for:

- Read all modules.
- Replace existing standard module source.
- Replace existing class module source.
- Replace existing document module source.
- Replace existing designer module code-behind source without corrupting the form.
- Add standard module.
- Add class module.
- Rename module.
- Delete module.
- Preserve read-only/private flags.
- Preserve document module identity.
- Reject invalid mutations, such as deleting required host document modules, unless explicitly supported.

Each mutation must update all affected structures:

- CFB stream names
- `PROJECT`
- `PROJECTwm`
- `dir`
- module stream
- designer storage references, if applicable

Failure condition:

- Add/rename/delete works by editing one stream only.

---

## GATE 14 — Designer / UserForm Gate

HARD PASS requires:

- Detects designer modules.
- Preserves designer storages.
- Preserves `VBFrame` stream data.
- Preserves Office Forms data associated with UserForms.
- Allows code-behind replacement without changing form layout.
- If full form editing is claimed, implements required designer and Office Forms structures.
- If full form editing is not claimed, form layout bytes must round-trip unchanged.

Failure condition:

- UserForms disappear, lose controls, or lose event-handler linkage after round-trip.

---

## GATE 15 — Content Hash / Integrity Gate

HARD PASS requires:

- Implements all MS-OVBA content-normalization rules that affect integrity verification.
- Implements content hash generation and validation.
- Implements agile content hash generation and validation where applicable.
- Implements V3 content normalized data and V3 content hash where applicable.
- Implements project normalized data where applicable.
- Updates or intentionally removes invalidated integrity metadata according to spec-compatible behavior.
- Never leaves stale hashes that cause Office to distrust or reject the project.

Failure condition:

- A modified project opens with integrity, signature, or trust errors caused by stale hash data.

---

## GATE 16 — Protection / Encryption / Password Gate

HARD PASS requires:

- Parses protection state.
- Parses password data.
- Implements MS-OVBA data encryption/decryption where used by project protection fields.
- Implements password hash data structure.
- Implements null encoding/decoding for password hash data.
- Implements password hash algorithm.
- Implements password hash validation.
- Preserves protected project metadata when not modifying it.
- Fails closed when a protected project cannot be safely modified.
- Does not claim to remove or bypass protection unless explicitly implemented and legally intended.

Failure condition:

- Protected projects are silently corrupted, unlocked unintentionally, or rewritten with invalid protection metadata.

---

## GATE 17 — Digital Signature Gate

HARD PASS requires:

- Detects whether a VBA project is digitally signed.
- Knows which mutations invalidate the signature.
- Does not leave a stale signature that appears valid but covers old content.
- Either:
  - removes invalidated signature metadata cleanly, or
  - re-signs using an explicit signing mechanism.
- Clearly reports signature invalidation to the caller.

Failure condition:

- A signed workbook is modified and then appears signed while containing changed code.

---

## GATE 18 — Encoding Gate

HARD PASS requires:

- Correctly handles the project code page.
- Correctly handles MBCS strings.
- Correctly handles UTF-16 strings.
- Correctly handles module names in both MBCS and Unicode forms.
- Correctly handles non-ASCII source code.
- Correctly handles comments and string literals containing non-ASCII characters.
- Correctly handles null bytes where allowed by structure.
- Does not use UTF-8 unless the spec or host layer explicitly requires it.

Failure condition:

- Unicode module names or source strings are corrupted after round-trip.

---

## GATE 19 — Cross-Structure Consistency Gate

HARD PASS requires validation that all of these agree:

- `PROJECT` module declarations
- `dir` MODULE records
- `PROJECTwm` name mappings
- CFB module stream names
- designer storage names
- document module identities
- module type flags
- source stream offsets

The implementation must expose a `validateProject()` or equivalent operation that reports all inconsistencies.

Failure condition:

- The writer emits a file without first validating cross-structure consistency.

---

## GATE 20 — Round-Trip Preservation Gate

HARD PASS requires a no-op round trip:

Input workbook -> parse -> write -> open in Excel

The output must satisfy:

- Excel opens without repair.
- VBA project opens in VBE.
- All modules are present.
- All references are present.
- All UserForms are present.
- All ActiveX controls remain present.
- Code compiles.
- Exported source text is semantically identical.
- Unmodified binary streams are byte-identical unless the spec requires regeneration.
- Any changed binary streams have documented reasons for change.

Failure condition:

- A no-op parse/write changes unrelated workbook or VBA project behavior.

---

## GATE 21 — Mutation Round-Trip Gate

HARD PASS requires mutation tests for:

- Replace source in a standard module.
- Replace source in a class module.
- Replace source in a document module.
- Replace source behind a UserForm.
- Add standard module.
- Add class module.
- Rename standard module.
- Rename class module.
- Delete standard module.
- Delete class module.
- Preserve non-ASCII module.
- Preserve external references.
- Preserve protected project metadata or fail closed.
- Preserve ActiveX/UserForm metadata.
- Handle signed project by invalidating/removing/re-signing signature explicitly.

Every output file must open in Excel, show expected source, compile, save again, and reopen.

Failure condition:

- Mutation passes binary tests but fails real Excel open/compile tests.

---

## GATE 22 — Corpus Gate

HARD PASS requires a test corpus containing at least:

- Empty VBA project.
- One standard module.
- Multiple standard modules.
- Class module.
- Document modules only.
- Mixed document/standard/class modules.
- UserForm with code-behind.
- UserForm with multiple controls.
- Workbook with ActiveX control.
- Workbook with external registered type-library reference.
- Workbook with project reference.
- Workbook with non-ASCII module names.
- Workbook with non-ASCII source text.
- Workbook using non-default code page.
- Password-protected VBA project.
- Digitally signed VBA project.
- Workbook saved by multiple Office versions if available.
- Corrupt/truncated negative test files.

Failure condition:

- Tests only use one hand-created `.xlsm`.

---

## GATE 23 — Fuzz / Malformed Input Gate

HARD PASS requires:

- Fuzz tests for compressed containers.
- Fuzz tests for CFB stream lookup.
- Fuzz tests for `dir` record lengths.
- Fuzz tests for invalid `MODULEOFFSET`.
- Fuzz tests for invalid chunk headers.
- Fuzz tests for invalid copy tokens.
- Fuzz tests for invalid code pages.
- Fuzz tests for missing required streams.
- Fuzz tests for duplicate module names.
- Fuzz tests for inconsistent `PROJECT` vs `dir`.

The parser must fail safely with diagnostics, not crash, hang, overrun, or emit corrupt output.

Failure condition:

- Malformed input can produce a partially written workbook.

---

## GATE 24 — API Contract Gate

HARD PASS requires the public API to separate these operations:

- `openOfficePackage(path)`
- `extractVbaProject(package)`
- `openVbaProjectBinary(bytes)`
- `parseVbaProject(cfb)`
- `listModules(project)`
- `getModuleSource(project, moduleName)`
- `setModuleSource(project, moduleName, source)`
- `addModule(project, moduleSpec)`
- `renameModule(project, oldName, newName)`
- `deleteModule(project, moduleName)`
- `validateProject(project)`
- `writeVbaProject(project)`
- `reinsertVbaProject(package, vbaProjectBytes)`
- `savePackage(package, path)`

Failure condition:

- API exposes only "replace bytes in zip" or "regex source replacement".

---

## GATE 25 — Documentation Gate

HARD PASS requires documentation stating:

- Supported host formats.
- Unsupported host formats.
- Supported module types.
- Whether UserForms are preserved or editable.
- Whether ActiveX controls are preserved or editable.
- Whether protected projects are supported.
- Whether signed projects are supported.
- Whether hashes are recalculated.
- Whether SRP/performance caches are removed.
- Which Office versions were tested.
- What operations invalidate signatures.
- Exact failure behavior for unsupported features.

Failure condition:

- Users cannot tell whether the library is a full MS-OVBA implementation or a module-only tool.

---

# Final Feature-Complete Rule

The implementation can be called "MS-OVBA feature complete" only if:

1. All hard gates pass.
2. Real Excel opens every generated file without repair.
3. VBE shows the expected project structure.
4. VBA source compiles.
5. No-op round trips preserve semantics.
6. Mutating source updates every dependent structure.
7. Unsupported edge cases fail closed with explicit diagnostics.
8. The test corpus includes modules, classes, document modules, UserForms, references, protection, signatures, non-ASCII names, and malformed files.

Otherwise, describe the implementation by its actual subset.

Examples:

- "MS-OVBA module extractor"
- "MS-OVBA module source reader/writer"
- "Partial MS-OVBA writer without UserForms"
- "VBA project binary round-tripper with source replacement"