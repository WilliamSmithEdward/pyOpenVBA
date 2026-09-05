# MS-OVBA Implementation Guide v2

A practical, language-agnostic guide for building a module-complete,
deterministic reader/writer for VBA modules embedded in Microsoft
Office files. Distilled from the implementation of
[pyOpenVBA](../README.md) against
[`MS-OVBA`](MS-OVBA_official_specification.pdf) v20260519, MS-CFB v3,
and live Excel workbooks (Excel for Microsoft 365, May 2026).

This guide supersedes [`ms-ovba-implementation-guide.md`](ms-ovba-implementation-guide.md).
Read this one. It is shorter, sharper, and reflects what was actually
needed to ship a real round-tripper that survives Excel reopen with no
repair dialog.

---

## 0. What you are building

**Goal.** Read and write VBA module source code embedded in `.xlsm`,
`.xlsb`, `.xlam`, and `.xls` files. Round-trip everything else verbatim.
Never silently corrupt a workbook.

**Non-goals (for v1 of your library).**

- UserForm layout editing (Office Forms binary).
- VBA project password decryption or re-encryption.
- Re-signing digitally signed projects.
- Office-compatible content-hash recomputation (only needed for
  re-signing).
- ActiveX license editing (PROJECTlk semantic edits).

These can be added later as discrete milestones. They are all
**preserve-verbatim** in v1.

---

## 1. Mental model

```
workbook.xlsm                          (host container)
+-- ZIP / Open XML Package
    +-- xl/vbaProject.bin              (the MS-OVBA payload)
        +-- MS-CFB compound file
            +-- /PROJECT               (plain text, MBCS)
            +-- /PROJECTlk  (optional) (ActiveX license info)
            +-- /VBA                   (storage)
            |   +-- _VBA_PROJECT       (performance cache; opaque)
            |   +-- dir                (compressed binary; the spine)
            |   +-- PROJECTwm (optional) (module-name MBCS<->UTF-16 map)
            |   +-- <Module1>          (compressed source w/ cache prefix)
            |   +-- <Module2>
            |   +-- __SRP_*  (optional) (perf cache; NEVER write)
            +-- <UserForm1>            (storage; designer sub-streams)
                +-- f, o, \x01CompObj, \x03VBFrame, ...
```

**The two layers are independent.** Solve them separately:

1. **Outer container layer.** ZIP for `.xlsm`/`.xlsb`/`.xlam`, raw CFB
   for `.xls`. You read/write one file inside the ZIP: `xl/vbaProject.bin`.
2. **Inner CFB / VBA layer.** Everything inside `vbaProject.bin`.

Pseudo-API contract for the host layer:

```text
read_vba_project_bytes(workbook_path) -> bytes
write_vba_project_bytes(workbook_path, new_bytes) -> None
```

For `.xls`, the entire file *is* the CFB. For everything else, replace
the single ZIP entry `xl/vbaProject.bin`, preserving every other entry
**byte-for-byte including compression method and metadata**.

**Critical:** if a `.xlsm` exists but has never had a VBA project
created (user has not entered the VBE), `xl/vbaProject.bin` will be
**absent**. Raise a structured error with the filename, not a generic
KeyError.

---

## 2. Build order (each step is testable in isolation)

| # | Milestone | Verifiable by |
|---|-----------|---------------|
| 1 | Read MS-CFB | Round-trip a known CFB through `parse -> serialize` byte-for-byte |
| 2 | Decompress MS-OVBA | Pass the three spec test vectors in section 3.2 of the PDF |
| 3 | Parse `dir` stream | Enumerate modules with names + offsets matching a known fixture |
| 4 | Read module sources | Diff against `Debug.Print` output from Excel |
| 5 | Write MS-OVBA (compress) | Compress + decompress round-trips every byte of every fixture module |
| 6 | Replace one module's source | Edit, save, reopen in Excel, no repair dialog |
| 7 | Add / rename / delete modules | Same, plus structural assertions on `dir` and `PROJECT` |
| 8 | Protection / signature gates | Refuse mutation without explicit opt-in |
| 9 | Fuzz harness | Bit-flipped inputs never crash unhandled |

Do **not** skip ahead. Step 5 is where most projects fail silently
(see Section 7).

---

## 3. MS-CFB essentials

You can use an existing CFB library or write your own. Either way, your
high-level CFB API needs exactly these operations:

```text
CFB.from_bytes(blob)             -> CFB
CFB.to_bytes()                   -> bytes
CFB.list_streams()               -> list[str]                # root
CFB.get_stream(name)             -> bytes
CFB.write_stream(name, bytes)    -> None
CFB.remove_stream(name)          -> None
CFB.list_streams_in_storage(s)   -> list[str]
CFB.get_stream_in_storage(s, n)  -> bytes
CFB.write_stream_in_storage(s, n, bytes) -> None
CFB.add_stream_to_storage(s, n, bytes)   -> None
CFB.rename_stream_in_storage(s, old, new) -> None
CFB.remove_stream_in_storage(s, n) -> None
CFB.drop_streams_in_storage(s, predicate) -> None
```

Stream / storage lookups MUST be **case-insensitive** but writes MUST
**preserve original casing**.

CFB-side gotchas:

- Sector size is 512 bytes for v3, 4096 for v4. Pick v3 unless you have
  a reason.
- The mini-stream cutoff is 4096 bytes; small streams go in the
  mini-FAT.
- After any stream add/remove/resize you must rebuild the FAT chain,
  the mini-FAT chain, and the directory entry tree. Re-serializing
  end-to-end is simpler than incremental patching and is what most
  implementations do.

---

## 4. MS-OVBA compression

**Used for:** the entire `dir` stream and every module's source body.

### 4.1 Container

```
byte 0:    0x01  (signature)
bytes 1+:  one or more compressed chunks
```

### 4.2 Chunk header (16-bit little-endian)

```
bits 0..11 : ChunkSize - 3       (chunk size INCLUDING the 2-byte header)
bits 12..14: 0b011                (signature)
bit 15     : 0 = raw, 1 = token-compressed
```

Helpers:

```text
chunk_size(h)   = (h & 0x0FFF) + 3
chunk_is_token(h) = (h >> 15) & 1
make_header(size, is_token) = ((size - 3) & 0x0FFF) | 0x3000 | (is_token << 15)
```

### 4.3 Per-chunk rules

- **Decompressed** chunks are at most 4096 bytes.
- Every decompressed chunk except the final one MUST be exactly 4096
  bytes.
- A **raw** chunk carries exactly 4096 data bytes after its header.
  Therefore raw encoding is **forbidden** for a final chunk shorter
  than 4096 bytes (it would append null bytes to the source).

### 4.4 Decompression (reference pseudocode)

```pseudo
function ovba_decompress(input) -> bytes:
    if len(input) == 0 or input[0] != 0x01:
        raise BadCompressedSignature

    pos = 1
    out = bytearray()

    while pos < len(input):
        chunk_start = pos
        header = read_u16le(input, pos)
        size = (header & 0x0FFF) + 3
        signature = (header >> 12) & 0x7
        is_token = (header >> 15) & 1
        if signature != 0b011:
            raise BadChunkSignature

        chunk_end = min(len(input), chunk_start + size)
        pos = chunk_start + 2
        chunk_decompressed_start = len(out)

        if not is_token:
            if pos + 4096 > len(input):
                raise TruncatedRawChunk
            out.extend(input[pos : pos + 4096])
            pos += 4096
            continue

        while pos < chunk_end:
            flag_byte = input[pos]
            pos += 1
            for bit in 0..7:
                if pos >= chunk_end:
                    break
                if (flag_byte >> bit) & 1 == 0:        # literal
                    out.append(input[pos])
                    pos += 1
                else:                                  # copy token
                    if pos + 2 > chunk_end:
                        raise TruncatedCopyToken
                    token = read_u16le(input, pos)
                    pos += 2
                    (offset, length) = unpack_copy_token(
                        token,
                        decompressed_current = len(out),
                        decompressed_chunk_start = chunk_decompressed_start
                    )
                    src = len(out) - offset
                    if src < chunk_decompressed_start:
                        raise BadCopyOffset
                    # Overlap is REQUIRED and intentional.
                    for i in 0..length - 1:
                        out.append(out[src + i])
    return bytes(out)
```

### 4.5 Copy-token packing

Bit allocation between offset and length **changes** as the decoder
advances through the current decompressed chunk:

```pseudo
function copy_token_help(decompressed_current, decompressed_chunk_start):
    difference = decompressed_current - decompressed_chunk_start
    bit_count = max(ceil_log2(difference), 4)         # 4..12
    length_mask = 0xFFFF >> bit_count
    offset_mask = (~length_mask) & 0xFFFF
    max_length  = length_mask + 3
    return (length_mask, offset_mask, bit_count, max_length)

function unpack_copy_token(token, dc, dcs):
    (lm, om, bc, _) = copy_token_help(dc, dcs)
    length = (token & lm) + 3
    offset = ((token & om) >> (16 - bc)) + 1
    return (offset, length)

function pack_copy_token(offset, length, dc, dcs):
    (lm, om, bc, _) = copy_token_help(dc, dcs)
    return ((offset - 1) << (16 - bc)) | ((length - 3) & lm)
```

`ceil_log2(1) = 0` per spec, so `bit_count` is clamped to a minimum of
4. At the start of a chunk the first copy token is impossible (no
backward bytes); the encoder must emit at least one literal first.

### 4.6 Compression

A correct encoder iterates the input in chunks of up to 4096 bytes and
for each chunk searches backward within the **current decompressed
chunk only** (not across chunk boundaries) for the longest match of at
least 3 bytes. Pack literals + copy tokens into groups of 8 prefixed
by a flag byte.

Final-chunk rule that catches every novice:

```pseudo
if chunk_len == 4096:
    emit_raw_chunk(chunk)               # safe and exact
elif chunk_len + ceil(chunk_len / 8) <= 4096:
    emit_literal_only_token_chunk()     # always valid
else:
    emit_full_lz_token_chunk()          # must implement matching
```

Never silently fall back to raw for a short final chunk. Test this with
final lengths `{1, 8, 9, 4093, 4094, 4095, 4096}`.

---

## 5. The `_VBA_PROJECT` stream

```
Reserved1   u16   MUST be 0x61CC          (visible as "cc 61" in little-endian dump)
Version     u16   host-defined cookie
Reserved2   u8    MUST be 0x00
PerformanceCache  variable bytes; "undefined and MUST be ignored on read"
```

**Read:** ignore the body. Validate header bytes if you want.

**Write on a no-op save:** preserve byte-for-byte.

**Write on a mutating save (add/rename/delete/source edit):**

1. Preserve the 5-byte header.
2. Zero the PerformanceCache body in place (same length).

This forces Office to regenerate the cache on next open. Some Office
builds will honor stale cache bytes that still happen to parse, leading
to a repair dialog or wrong p-code on first open. Zeroing the body is
trivial, deterministic, and well-supported by the spec ("MUST be
ignored on read").

Do **not** rewrite `Version` to `0xFFFF` as some older guides suggest;
modern Office writes a real cookie and zeroing the body is sufficient.

---

## 6. The `dir` stream

`dir` is **compressed**. Decompress it first.

### 6.1 Top-level layout

```
PROJECTINFORMATION
PROJECTREFERENCES
PROJECTMODULES
Terminator   u16   0x0010
Reserved     u32   0x00000000
```

Every record is `{ Id: u16, Size: u32, Payload: bytes[Size] }` with a
few exceptions noted in the tables below.

### 6.2 `PROJECTINFORMATION` records (in order)

| Id     | Record               | Payload                                                       |
|--------|----------------------|---------------------------------------------------------------|
| 0x0001 | PROJECTSYSKIND       | u32 (0=Win16, 1=Win32, 2=Mac, 3=Win64)                        |
| 0x004A | PROJECTCOMPATVERSION | u32 (optional)                                                |
| 0x0002 | PROJECTLCID          | u32 (typically 0x00000409)                                    |
| 0x0014 | PROJECTLCIDINVOKE    | u32                                                           |
| 0x0003 | PROJECTCODEPAGE      | u16 (typically 1252)                                          |
| 0x0004 | PROJECTNAME          | MBCS bytes                                                    |
| 0x0005 | PROJECTDOCSTRING     | MBCS, then `0x0040`, then UTF-16LE                            |
| 0x0006 | PROJECTHELPFILEPATH  | MBCS, then `0x003D`, then duplicate MBCS                      |
| 0x0007 | PROJECTHELPCONTEXT   | u32                                                           |
| 0x0008 | PROJECTLIBFLAGS      | u32                                                           |
| 0x0009 | PROJECTVERSION       | reserved u32, VersionMajor u32, VersionMinor u16              |
| 0x000C | PROJECTCONSTANTS     | MBCS, then `0x003C`, then UTF-16LE (optional)                 |

Parse codepage FIRST in your decoder; every MBCS string after it
depends on it.

### 6.3 `PROJECTREFERENCES`

A flat array of `REFERENCE` records, terminated when the next u16 is
`0x000F` (start of `PROJECTMODULES`). Each REFERENCE optionally starts
with a `REFERENCENAME` (0x0016) and then one of:

| Id     | Record              |
|--------|---------------------|
| 0x002F | REFERENCECONTROL    |
| 0x0033 | REFERENCEORIGINAL   |
| 0x000D | REFERENCEREGISTERED |
| 0x000E | REFERENCEPROJECT    |

For v1 you can preserve REFERENCE bytes verbatim; full semantic editing
is rarely required.

### 6.4 `PROJECTMODULES`

```
Id:    u16 = 0x000F
Size:  u32 = 0x00000002
Count: u16
PROJECTCOOKIE { Id: u16 = 0x0013, Size: u32 = 2, Cookie: u16 (write 0xFFFF) }
MODULE * Count
```

### 6.5 `MODULE` record (fixed order)

| Id     | Record            | Payload                                                                 |
|--------|-------------------|-------------------------------------------------------------------------|
| 0x0019 | MODULENAME        | MBCS                                                                    |
| 0x0047 | MODULENAMEUNICODE | UTF-16LE (optional)                                                     |
| 0x001A | MODULESTREAMNAME  | MBCS, then `0x0032`, then UTF-16LE                                      |
| 0x001C | MODULEDOCSTRING   | MBCS, then `0x0048`, then UTF-16LE                                      |
| 0x0031 | MODULEOFFSET      | u32 (offset of source within the module stream)                         |
| 0x001E | MODULEHELPCONTEXT | u32                                                                     |
| 0x002C | MODULECOOKIE      | u16 (write 0xFFFF)                                                      |
| 0x0021 \| 0x0022 | MODULETYPE | (no payload; presence of 0x0021 = procedural, 0x0022 = doc/class/designer) |
| 0x0025 | MODULEREADONLY    | (optional)                                                              |
| 0x0028 | MODULEPRIVATE     | (optional)                                                              |
| 0x002B | terminator        | followed by reserved u32 = 0                                            |

`MODULETYPE` alone cannot tell document from class from designer; for
that you need the `PROJECT` stream (Section 8).

---

## 7. Module streams

Each module's source lives at `/VBA/<MODULESTREAMNAME>`. Layout:

```
[ PerformanceCache prefix ]  MODULEOFFSET bytes (opaque; preserve)
[ CompressedSourceCode    ]  bytes from MODULEOFFSET to EOF
```

**Read:**

```pseudo
stream = cfb.get_stream_in_storage("VBA", module.stream_name)
compressed = stream[module.text_offset:]
raw = ovba_decompress(compressed)
text = decode(raw, code_page=project.code_page)
```

**Write (in-place source replacement):**

```pseudo
old      = cfb.get_stream_in_storage("VBA", module.stream_name)
prefix   = old[:module.text_offset]                          # preserve verbatim
new_raw  = encode(new_text, code_page=project.code_page)
new_body = ovba_compress(new_raw)
cfb.write_stream_in_storage("VBA", module.stream_name, prefix + new_body)
```

**Write (new module):**

```pseudo
seed_body = ovba_compress(encode(initial_source, code_page))
# No prefix needed; set MODULEOFFSET to 0 in the new MODULE record.
cfb.add_stream_to_storage("VBA", new_name, seed_body)
```

Source text MUST be CRLF-terminated (`\r\n`). Normalize on the way in.

Standard module sources usually start with:

```
Attribute VB_Name = "ModuleName"
```

Class modules start with the `VERSION` block + class attributes:

```
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
...
```

**Do not invent these attribute headers.** When adding a new module,
either let the user supply the full text including attributes, or
generate the minimum standard-module header only.

---

## 8. The `PROJECT` stream

Plain MBCS text encoded with the project codepage. The grammar is:

```
ID="{...}"
Document=ThisWorkbook/&H00000000
Document=Sheet1/&H00000000
Module=Module1
Class=Class1
BaseClass=UserForm1
Name="VBAProject"
HelpContextID="0"
VersionCompatible32="393222000"
CMG="..."          ' obfuscated; preserve verbatim
DPB="..."          ' obfuscated; preserve verbatim
GC="..."           ' obfuscated; preserve verbatim

[Host Extender Info]
&H00000001={...};VBE;&H00000000

[Workspace]
ThisWorkbook=0, 0, 0, 0, C
Sheet1=0, 0, 0, 0, C
Module1=0, 0, 0, 0, C
```

Module declaration lines disambiguate `MODULETYPE = 0x0022`:

| Line prefix  | Module kind |
|--------------|-------------|
| `Module=`    | standard procedural |
| `Class=`     | class module |
| `Document=`  | host-owned document (ThisWorkbook, Sheet1, ...) |
| `BaseClass=` | designer module (UserForm, requires sub-storage) |

On mutation rewrite the declaration lines and the `[Workspace]`
entries together. Do not create `Document=` lines (host-owned). Do
not create `BaseClass=` lines unless you also emit a designer
sub-storage.

---

## 9. The `PROJECTwm` stream

A flat array of `{ MBCS\0, UTF-16LE\0\0 }` pairs, terminated by
`0x0000`. Maintain order to match `PROJECTMODULES`. Rewrite whenever
the module set changes (add/rename/delete).

```pseudo
function serialize_projectwm(pairs, code_page):
    out = bytearray()
    for (mbcs_name, unicode_name) in pairs:
        out.extend(encode(mbcs_name, code_page))
        out.append(0x00)
        out.extend(encode_utf16le(unicode_name))
        out.extend(b"\x00\x00")
    out.extend(b"\x00\x00")
    return bytes(out)
```

---

## 10. The `PROJECTlk` stream (ActiveX licenses)

Verbatim-preserve in v1. Format is an array of `LicenseInfoRecord`s
keyed by control CLSID. Round-tripping the raw bytes is sufficient
unless you are editing ActiveX licenses, which is rare and tied to a
deprecated control surface.

---

## 11. UserForms / designer storages

For each `BaseClass=` module there is a **sub-storage** at the CFB root
with the same name as the module. It contains designer streams:
typically `f`, `o`, `\x01CompObj`, `\x03VBFrame`, and sometimes more.

Editing the **code-behind** of a UserForm is identical to editing any
other module's source: the form's *code* lives in `/VBA/<FormName>` and
is round-tripped through the same compressed-source pipeline.

Editing the **layout** (controls, properties, positions) requires the
Office Forms binary format and is out of scope for v1. Preserve the
designer sub-storage byte-for-byte.

---

## 12. Encoding

All MBCS strings inside the VBA project use the codepage declared in
`PROJECTCODEPAGE`. In practice this is overwhelmingly `1252` for
Western installs, but never hard-code it.

`cp1252` covers most of Latin-1, so Latin-1-supplement names like
`Mödüle1` round-trip cleanly. Names that lie outside cp1252 (e.g.
Greek, Cyrillic, CJK) **are rejected by Excel's VBA IDE itself**, so
your library does not need to support them; document this as a
constraint, not a bug.

UTF-16LE strings (the `*UNICODE` records and the UTF-16 halves of
`PROJECT*` records) are always little-endian and **not BOM-prefixed**.

---

## 13. Performance cache streams (`__SRP_*`)

These are Office's compiled p-code cache. They live as `__SRP_0`,
`__SRP_1`, etc. inside `/VBA`.

**Rule:** writers MUST NOT emit `__SRP_*` streams. Drop them
unconditionally on save. Office will regenerate them on next open.

---

## 14. Protection (`CMG` / `DPB` / `GC`)

These three keys in the `PROJECT` stream carry the obfuscated
project-password material:

- `CMG`: Crypted Message Group, encrypted password hash + flags.
- `DPB`: Data Protection Block, encrypted password.
- `GC`: Group Code (visibility flag).

You can **detect** a protected project by parsing `DPB`: if its
decoded payload's first protection-state byte is non-zero, the project
is password-protected.

**You cannot edit the password without the password.** Implementing
RC4-based password decryption/re-encryption is technically possible but
ethically and legally questionable; treat it as out of scope.

**Save safety gate:** if the project is password-protected AND the save
would emit any change, **refuse the save** unless the caller passes an
explicit opt-in flag (e.g. `allow_protected=True`). With the opt-in,
preserve the password material verbatim; the workbook will still
require the original password to unlock the project, and the protection
state is preserved.

---

## 15. Digital signatures

Three signature stream variants may exist inside `/VBA`:

- `_VBA_PROJECT_SIGNATURE` (legacy)
- `_VBA_PROJECT_SIGNATURE_AGILE` (agile)
- `_VBA_PROJECT_SIGNATURE_V3` (V3 / current)

Any change to module source or topology **invalidates** every one of
these signatures. Re-signing requires the original signer's private
key and is therefore out of scope.

**Save behavior on mutation:**

1. Detect any signature streams.
2. Drop them all (remove from both the `/VBA` storage and root).
3. Emit a warning so the caller knows trust has been removed.
4. Allow silencing via an explicit opt-in flag (e.g.
   `allow_invalidate_signature=True`).

---

## 16. The complete save pipeline

```pseudo
function save(cfb, project, *, allow_protected=False, allow_invalidate_signature=False):
    mutating = (project.has_renames
                or project.has_adds
                or project.has_deletes
                or any(module.dirty for module in project.modules))

    # --- Safety gates ---
    if mutating and project.is_password_protected and not allow_protected:
        raise ProtectedProjectError

    if mutating and detect_signature(cfb).present:
        drop_signature_streams(cfb)
        if not allow_invalidate_signature:
            warn("dropped stale signatures")

    # --- 1. Renames: rename existing CFB streams ---
    for (old, new) in project.renames:
        cfb.rename_stream_in_storage("VBA", old, new)

    # --- 2. Adds: create new module streams ---
    for module in project.added_modules:
        seed = compress(encode(module.source, code_page))
        cfb.add_stream_to_storage("VBA", module.stream_name, seed)

    # --- 3. Deletes: remove module streams ---
    for name in project.deleted_streams:
        cfb.remove_stream_in_storage("VBA", name)

    # --- 4. Source edits: replace bodies of pre-existing modules ---
    for module in project.modules:
        if module.dirty:
            replace_module_source(cfb, module, code_page)

    # --- 5. dir / PROJECT / PROJECTwm rewrites if structure changed ---
    if project.structure_changed:
        cfb.write_stream_in_storage("VBA", "dir", compress(serialize_dir(project)))
        cfb.write_stream("PROJECT", serialize_project_stream(...))
        if cfb.has_stream_in_storage("VBA", "PROJECTwm"):
            cfb.write_stream_in_storage("VBA", "PROJECTwm",
                                        serialize_projectwm(project.modules, code_page))

    # --- 6. Cache invalidation ---
    if mutating:
        invalidate_vba_project_cache(cfb)   # zero PerformanceCache body
    cfb.drop_streams_in_storage("VBA", name -> name.startswith("__SRP_"))

    # --- 7. Outer container write ---
    new_bytes = cfb.to_bytes()
    write_vba_project_bytes(workbook_path, new_bytes)   # preserves all other ZIP entries
```

`invalidate_vba_project_cache` (Section 5) is the single most
under-documented step in older guides; without it, Office sometimes
re-prompts for repair after legitimate mutations.

---

## 17. Mutation operations checklist

For each operation, every step must succeed or you must abort.

### 17.1 Replace source

| Step | Touches |
|------|---------|
| Re-encode + recompress source | module stream body (after MODULEOFFSET) |
| Invalidate `_VBA_PROJECT` cache | `_VBA_PROJECT` |
| Drop `__SRP_*` | `/VBA` |
| Drop signatures (warn) | `/VBA`, root |

### 17.2 Add module

| Step | Touches |
|------|---------|
| Create CFB stream w/ seeded source | `/VBA/<NewName>` |
| Append MODULE record | `dir` |
| Increment `PROJECTMODULES.Count` | `dir` |
| Add `Module=` / `Class=` line | `PROJECT` |
| Add `[Workspace]` entry | `PROJECT` |
| Append NAMEMAP pair | `PROJECTwm` |
| Invalidate + drop SRP + signatures | as above |

### 17.3 Rename module

| Step | Touches |
|------|---------|
| Rename CFB stream | `/VBA/<OldName> -> /VBA/<NewName>` |
| Update MODULENAME / MODULENAMEUNICODE / MODULESTREAMNAME | `dir` |
| Update `Module=` / `Class=` line | `PROJECT` |
| Update `[Workspace]` key | `PROJECT` |
| Update NAMEMAP | `PROJECTwm` |
| Update `Attribute VB_Name = "..."` inside the source | module body |
| Invalidate + drop SRP + signatures | as above |

### 17.4 Delete module

| Step | Touches |
|------|---------|
| Remove CFB stream | `/VBA` |
| Drop MODULE record | `dir` |
| Decrement `PROJECTMODULES.Count` | `dir` |
| Drop declaration line | `PROJECT` |
| Drop `[Workspace]` entry | `PROJECT` |
| Drop NAMEMAP pair | `PROJECTwm` |
| Invalidate + drop SRP + signatures | as above |

Forbidden: deleting `Document=` modules (host-owned), deleting
designer modules without also removing their sub-storage.

---

## 18. Outer-container details

### 18.1 OOXML (`.xlsm`, `.xlsb`, `.xlam`)

Replace exactly the `xl/vbaProject.bin` entry. Preserve every other
entry's bytes, compression method, external attributes, create system,
and timestamps:

```pseudo
with ZipFile(input_path) as zin, ZipFile(out_buf, "w") as zout:
    for info in zin.infolist():
        if info.filename == "xl/vbaProject.bin":
            zinfo = ZipInfo(info.filename, info.date_time)
            zinfo.compress_type = info.compress_type
            zinfo.external_attr = info.external_attr
            zinfo.create_system = info.create_system
            zout.writestr(zinfo, new_vba_bin)
        else:
            data = zin.read(info.filename)
            zinfo = clone_zinfo(info)
            zout.writestr(zinfo, data)
write_atomically(out_buf, out_path)
```

`.xlsb` is byte-compatible with this pipeline; the VBA project lives
in the same `xl/vbaProject.bin` entry.

### 18.2 Legacy `.xls` (BIFF8)

The entire file **is** the CFB. Just write `cfb.to_bytes()` to the
output path. No ZIP layer is involved.

### 18.3 Format dispatch

```pseudo
EXT = path.suffix.lower()
if EXT in {".xlsm", ".xlsb", ".xlam"}:
    use_zip_pipeline()
elif EXT == ".xls":
    use_raw_cfb_pipeline()
else:
    raise UnsupportedFormat(EXT)
```

### 18.4 Empty-VBA xlsm

A `.xlsm` that has never had a VBA project initialized **does not
contain** `xl/vbaProject.bin`. Raise a structured error:

```text
"<filename> contains no xl/vbaProject.bin. Make sure the workbook has
a VBA project (save as .xlsm in Excel with at least one macro)."
```

Not having `vbaProject.bin` is not a corruption signal; it's a
first-class user-facing case.

---

## 19. Error model

A minimal exception hierarchy:

```
PyOpenVBAError                  # library root
    UnsupportedFormatError      # bad extension / host
    CFBError                    # malformed CFB
    VBAProjectError             # malformed VBA project, bad codepage, etc.
        ProtectedProjectError   # mutation refused without opt-in
```

Every parser MUST attach `(stream_name, byte_offset)` context to error
messages. This pays for itself a hundred times during fuzz triage.

---

## 20. Testing strategy

### 20.1 Unit / integration

- **Compression conformance**: pass the three test vectors in the PDF
  section 3.2 exactly.
- **CFB round-trip**: parse + reserialize every fixture; binary diff
  must be empty (or have a documented diff set).
- **No-op save round-trip**: open + save without edits; every ZIP entry
  and every CFB stream must be byte-identical.
- **Mutation round-trip**: edit / add / rename / delete; reopen; verify
  the structural model matches.
- **Codepage tests**: at minimum cp1252 + one non-Latin codepage like
  932 (Shift-JIS).

### 20.2 Manual Excel verification

The "opens in Excel without repair" assertion **cannot** be made from
code alone. Build a one-shot script that produces a battery of
mutated workbooks (no-op, source-edit, add, rename, delete; xlsm +
xlsb; protected fixture with opt-in) and open each in Excel. Document
the matrix and version-stamp your `PASS` claim.

### 20.3 Fuzz

Persistent on-disk corpus, one subdirectory per parser target:

```
fuzz_corpus/
  cfb/         (seeds for CFB.from_bytes)
  decompress/  (seeds for ovba_decompress)
  dir/         (seeds for parse_dir_stream)
  project/     (seeds for parse_project_stream)
  projectwm/   (seeds for parse_projectwm)
```

Parametrize a test over every file. Each parser must either succeed
or raise one of its documented exception types. Any other escape is a
fuzz regression. When you find a failing input in the wild, drop it
into the appropriate directory and commit.

Seed the initial corpus with: empty, one byte, truncated header,
random 512B / 2KB blobs, and 4-6 bit-flipped variants of a real
fixture (deterministic RNG seed for reproducibility).

---

## 21. Key gotchas (the things that bite)

1. **`dir` is compressed.** First-time implementers parse it as raw
   binary and get nowhere. Decompress, then parse.
2. **Module stream has a cache prefix.** Source does **not** start at
   byte 0. Always slice from `MODULEOFFSET`.
3. **Source bytes are MBCS.** Decode with `PROJECTCODEPAGE`, not UTF-8.
4. **Source must be CRLF.** Normalize on the way in.
5. **Final compressed chunks shorter than 4096 bytes must be
   token-encoded.** Raw encoding will silently append null bytes.
6. **Copy-token bit allocation is dynamic.** `copy_token_help` depends
   on the current decompressed position within the chunk.
7. **`MODULETYPE` alone does not distinguish class from document from
   designer.** Use the `PROJECT` stream's `Module=` / `Class=` /
   `Document=` / `BaseClass=` lines.
8. **`__SRP_*` streams must never be written.** Drop them.
9. **`_VBA_PROJECT` cache must be invalidated on mutation.** Zero the
   body after the 5-byte header.
10. **Signatures are invalidated by any source change.** Drop them on
    mutation and warn.
11. **Protected projects must not be silently re-saved.** Require an
    explicit opt-in.
12. **CFB stream names are case-insensitive but case-preserving.**
13. **A `.xlsm` with no VBA project has no `vbaProject.bin`.** Surface
    a structured error.
14. **The `Attribute VB_Name = "..."` line inside the source must
    match `MODULENAME`.** Rename touches both.
15. **Renaming a module renames the CFB stream too** (in
    Excel-produced files the logical name equals the stream name).
16. **UserForm code-behind edits are simple module edits**; the
    sibling designer sub-storage must be preserved byte-for-byte.

---

## 22. Suggested public API surface

```text
# Open / save
ExcelFile(path).open() / .save(dest=None, *, allow_protected=False,
                                allow_invalidate_signature=False)

# Read
ExcelFile.module_names() -> list[str]
ExcelFile.get_module(name) -> str
ExcelFile.vba_project() -> VBAProject

# Write (source)
ExcelFile.set_module(name, new_source) -> None

# Write (topology) on the VBAProject
VBAProject.add_module(name, source, *, kind=standard, stream_name=None) -> VBAModule
VBAProject.rename_module(old, new) -> VBAModule
VBAProject.delete_module(name) -> None

# Disk-based workflow
pull(workbook_path, out_dir) -> list[Path]
push(workbook_path, src_dir, *, strict=True, encoding="utf-8") -> list[str]

# Detection
detect_signature(cfb) -> SignatureInfo
invalidate_vba_project_cache(cfb) -> bool

# Errors
PyOpenVBAError, UnsupportedFormatError, CFBError, VBAProjectError
```

---

## 23. Minimum viable read-only extractor (full pseudocode)

If you implement only this much you already have something useful:

```pseudo
function extract_modules(xlsm_path):
    zip = open_zip(xlsm_path)
    if "xl/vbaProject.bin" not in zip:
        raise VBAProjectError("no vbaProject.bin")
    blob = zip.read("xl/vbaProject.bin")

    cfb = CFB.from_bytes(blob)
    dir_raw = ovba_decompress(cfb.get_stream_in_storage("VBA", "dir"))
    project = parse_dir(dir_raw)

    out = []
    for module in project.modules:
        stream = cfb.get_stream_in_storage("VBA", module.stream_name)
        compressed = stream[module.text_offset:]
        source_bytes = ovba_decompress(compressed)
        source_text = decode(source_bytes, code_page=project.code_page)
        out.append((module.name, source_text))
    return out
```

About 25 lines plus the CFB + decompression primitives. That's the
whole MVP.

---

## 24. Closing principles

- **Two layers, two solvers.** Outer container (ZIP/CFB dispatch) is
  not the same problem as inner CFB+VBA. Don't mix them.
- **Preserve verbatim what you don't understand.** Bytes that
  round-trip cleanly are bytes that never break in the wild.
- **Refuse rather than corrupt.** Every mutation that could leave the
  workbook inconsistent (protected, signed, malformed) must fail loudly
  unless explicitly overridden.
- **Test in Excel, not just in code.** "Parsed and re-serialized" is
  necessary but not sufficient; the only real success criterion is
  "Office reopens without repair."
- **Seed your fuzz corpus from day one.** Every weird input is a gift;
  check it in.
