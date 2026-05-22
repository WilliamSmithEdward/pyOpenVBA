# MS-OVBA Implementation Guide for LLM Coding Agents

Source analyzed: `[MS-OVBA] - v20260519`, `Office VBA File Format Structure`, release May 19, 2026. This is a distilled implementation guide, not a replacement for the full Microsoft spec. It is written to help another coding LLM implement direct reading and writing of VBA modules embedded in Office files.

Primary implementation target: extract, replace, add, and remove VBA code modules in macro-enabled Office documents, especially Excel `.xlsm` files. Full support for forms, signatures, project protection, content hashes, and host-specific package creation should be treated as later milestones.

## 1. Mental model

A VBA project is not stored as loose `.bas`, `.cls`, or `.frm` text. It is stored as a structured binary project.

For an Excel `.xlsm`, the usual stack is:

```text
workbook.xlsm
  ZIP / Open XML package
    xl/vbaProject.bin
      OLE Compound File Binary container, MS-CFB
        PROJECT
        PROJECTwm
        PROJECTlk, optional
        VBA/
          _VBA_PROJECT
          dir
          <module stream 1>
          <module stream 2>
          __SRP_*, optional performance cache streams
        <designer storages>, optional
```

The `vbaProject.bin` file is the actual MS-OVBA object. MS-OVBA itself assumes an OLE compound file storage. For `.xlsm`, locating `vbaProject.bin` in the ZIP package is outside the core MS-OVBA parsing problem.

## 2. Scope levels

Implement in layers. Do not try to support every MS-OVBA feature on the first pass.

### Level 1: read-only module extractor

Goal: list modules and export source text.

Required pieces:

1. Open host file.
2. Locate `vbaProject.bin`.
3. Open it as MS-CFB structured storage.
4. Read `/VBA/dir`.
5. Decompress `dir` using MS-OVBA compression.
6. Parse enough `dir` records to get:
   - project code page
   - module count
   - module logical name
   - module stream name
   - module type
   - module source offset
7. For each module stream, slice from `MODULEOFFSET` to EOF.
8. Decompress that slice.
9. Decode bytes using the project code page.
10. Return source text.

### Level 2: source replacement writer

Goal: replace source in existing module streams without changing module names or project topology.

Best low-risk strategy:

1. Preserve the existing CFB project structure.
2. Parse `dir` to find each module's `MODULESTREAMNAME` and `MODULEOFFSET`.
3. Encode replacement source using the project code page.
4. Compress source bytes using MS-OVBA compression.
5. Replace the bytes from `MODULEOFFSET` to EOF in that module stream.
6. Preserve the prefix before `MODULEOFFSET` or set `MODULEOFFSET` to zero and rebuild the `dir` stream.
7. Set `_VBA_PROJECT.Version` to `0xFFFF` so performance caches are ignored.
8. Prefer removing `__SRP_*` streams when doing a strict rewrite.
9. Write a new output workbook. Do not modify the original in place.

### Level 3: topology writer

Goal: add, remove, or rename modules.

This requires rewriting:

- `/VBA/dir`
- `/PROJECT`
- `/PROJECTwm`
- module streams under `/VBA`
- possibly designer storages for UserForms

Standard modules are the safest addition target. Document modules are host-owned. Designer/UserForm modules require companion designer storage and Office Forms binary structures.

## 3. Normative source map

Use these parts of the PDF when implementing:

| Spec area | MS-OVBA sections | PDF pages |
|---|---:|---:|
| Required CFB storage hierarchy | 2.2 | 19-21 |
| `PROJECT` stream text grammar | 2.3.1 | 21-28 |
| `PROJECTwm` module name map | 2.3.3 | 29 |
| `_VBA_PROJECT` stream | 2.3.4.1 | 29-30 |
| compressed `dir` stream | 2.3.4.2 | 30-50 |
| module records | 2.3.4.2.3 | 44-50 |
| module stream source layout | 2.3.4.3 | 50 |
| compression/decompression | 2.4.1 | 54-72 |
| examples and compression test vectors | 3 | 89-111 |

Important: sections 1.7 and 2 are normative. Examples are informative but useful as test vectors.

## 4. Core data model

Use a model like this internally:

```text
VbaProject
  code_page: int
  sys_kind: int
  project_name: string
  references: list[VbaReference]
  modules: list[VbaModule]
  raw_project_stream: bytes or parsed text
  raw_projectwm_stream: bytes or parsed records
  cfb_preserved_items: map[path, bytes]

VbaModule
  name: string
  name_unicode: string optional
  stream_name: string
  stream_name_unicode: string optional
  type: enum {standard, class_or_document_or_designer}
  project_stream_type: enum {Document, Module, Class, BaseClass}
  text_offset: uint32
  help_context: uint32
  cookie: uint16
  is_read_only: bool
  is_private: bool
  compressed_source: bytes
  source_bytes: bytes
  source_text: string
  raw_dir_record: bytes optional
```

Keep both parsed values and raw bytes. Raw preservation lets you support unknown records and host quirks without fully understanding every feature.

## 5. Binary conventions

All record integer fields are little-endian. The PDF packet diagrams use big-endian visual bit numbering, but actual records and enumerations are little-endian.

Required binary helpers:

```text
read_u16le()
read_u32le()
write_u16le()
write_u32le()
read_bytes(n)
read_mbcs_string(byte_count, code_page)
read_utf16le_string(byte_count)
write_mbcs_string(text, code_page)
write_utf16le_string(text)
peek_u16le()
expect_u16le(value, label)
expect_u32le(value, label)
```

Treat invalid size fields as hard parse errors unless you are in a forensic recovery mode.

## 6. Storage hierarchy

The CFB root storage MUST contain:

- storage named `VBA`, case-insensitive
- stream named `PROJECT`, case-insensitive

The `VBA` storage MUST contain:

- stream `_VBA_PROJECT`, case-insensitive
- stream `dir`, case-insensitive
- one module stream for each module record

The `VBA` storage MAY contain `__SRP_*` streams. These are performance caches. Ignore them on read. For strict writing, omit them.

The root storage MAY contain:

- `PROJECTwm`
- `PROJECTlk`
- designer storages

Designer storages are required for designer modules. Do not create UserForms unless you also implement the Office Forms binary structures.

## 7. `_VBA_PROJECT` stream

Layout:

| Field | Size | Write behavior |
|---|---:|---|
| `Reserved1` | 2 | write `0x61CC` |
| `Version` | 2 | write `0xFFFF` |
| `Reserved2` | 1 | write `0x00` |
| `Reserved3` | 2 | ignore or preserve |
| `PerformanceCache` | variable | ignore on read; strict writer omits |

The key implementation rule is `Version = 0xFFFF` on write. That forces interoperable loading and avoids stale version-specific performance caches being used.

## 8. `dir` stream overview

The `dir` stream is compressed using the MS-OVBA compression algorithm. After decompression, its top-level layout is:

```text
PROJECTINFORMATION
PROJECTREFERENCES
PROJECTMODULES
Terminator: u16 = 0x0010
Reserved:   u32 = 0x00000000
```

Parsing order matters because `PROJECTCODEPAGE` appears in `PROJECTINFORMATION` and controls decoding of MBCS strings throughout the project.

For a robust parser, first parse fields structurally as bytes. Decode strings after the code page is known.

## 9. `PROJECTINFORMATION` records

Records occur in fixed order, with a few optional records.

| Record | Id | Payload |
|---|---:|---|
| `PROJECTSYSKIND` | `0x0001` | `Size=4`, `SysKind:u32` |
| `PROJECTCOMPATVERSION` optional | `0x004A` | `Size=4`, `CompatVersion:u32` |
| `PROJECTLCID` | `0x0002` | `Size=4`, `Lcid:u32`, normally `0x00000409` |
| `PROJECTLCIDINVOKE` | `0x0014` | `Size=4`, `LcidInvoke:u32`, normally `0x00000409` |
| `PROJECTCODEPAGE` | `0x0003` | `Size=2`, `CodePage:u16` |
| `PROJECTNAME` | `0x0004` | `SizeOfProjectName:u32`, MBCS bytes |
| `PROJECTDOCSTRING` | `0x0005` | MBCS doc string, reserved `0x0040`, UTF-16 doc string |
| `PROJECTHELPFILEPATH` | `0x0006` | MBCS help path, reserved `0x003D`, duplicate MBCS path |
| `PROJECTHELPCONTEXT` | `0x0007` | `Size=4`, `HelpContext:u32` |
| `PROJECTLIBFLAGS` | `0x0008` | `Size=4`, `ProjectLibFlags:u32`, usually zero |
| `PROJECTVERSION` | `0x0009` | `Reserved=4`, `VersionMajor:u32`, `VersionMinor:u16` |
| `PROJECTCONSTANTS` optional | `0x000C` | MBCS constants, reserved `0x003C`, UTF-16 constants |

`SysKind` values:

| Value | Meaning |
|---:|---|
| `0x00000000` | 16-bit Windows |
| `0x00000001` | 32-bit Windows |
| `0x00000002` | Macintosh |
| `0x00000003` | 64-bit Windows |

For most modern Excel files you will see Windows-oriented values and code page `1252`, but do not hard-code that.

## 10. `PROJECTREFERENCES` records

`PROJECTREFERENCES` is an array of `REFERENCE` records. The array ends when the next `u16` is `0x000F`, which starts `PROJECTMODULES`.

Each `REFERENCE` has an optional `REFERENCENAME` followed by one reference record.

Reference record ids:

| Id | Record |
|---:|---|
| `0x0016` | `REFERENCENAME`, optional prefix |
| `0x002F` | `REFERENCECONTROL` |
| `0x0033` | `REFERENCEORIGINAL` |
| `0x000D` | `REFERENCEREGISTERED` |
| `0x000E` | `REFERENCEPROJECT` |

For module extraction and simple replacement, you can parse references shallowly and preserve their raw bytes. Full semantic support is only needed if you are rewriting references or computing content hashes.

## 11. `PROJECTMODULES` and `MODULE` records

`PROJECTMODULES` begins with:

| Field | Value |
|---|---|
| `Id:u16` | `0x000F` |
| `Size:u32` | `0x00000002` |
| `Count:u16` | number of module records |
| `PROJECTCOOKIE` | id `0x0013`, size `2`, cookie usually `0xFFFF` |

Then parse `Count` `MODULE` records.

A `MODULE` record has this required order:

```text
MODULENAME
MODULENAMEUNICODE, optional
MODULESTREAMNAME
MODULEDOCSTRING
MODULEOFFSET
MODULEHELPCONTEXT
MODULECOOKIE
MODULETYPE
MODULEREADONLY, optional
MODULEPRIVATE, optional
Terminator: u16 = 0x002B
Reserved:   u32 = 0x00000000
```

Module record ids:

| Record | Id | Notes |
|---|---:|---|
| `MODULENAME` | `0x0019` | MBCS module name |
| `MODULENAMEUNICODE` | `0x0047` | optional UTF-16 module name |
| `MODULESTREAMNAME` | `0x001A` | MBCS stream name, reserved `0x0032`, UTF-16 stream name |
| `MODULEDOCSTRING` | `0x001C` | MBCS doc string, reserved `0x0048`, UTF-16 doc string |
| `MODULEOFFSET` | `0x0031` | `Size=4`, `TextOffset:u32` |
| `MODULEHELPCONTEXT` | `0x001E` | `Size=4`, help context |
| `MODULECOOKIE` | `0x002C` | `Size=2`, cookie ignored, write `0xFFFF` |
| `MODULETYPE` | `0x0021` or `0x0022` | `0x0021` standard module; `0x0022` document/class/designer |
| `MODULEREADONLY` | `0x0025` | optional, reserved `0` |
| `MODULEPRIVATE` | `0x0028` | optional, reserved `0` |
| terminator | `0x002B` | followed by reserved `u32=0` |

Important: `MODULETYPE` alone cannot distinguish document, class, and designer modules. Use the `PROJECT` stream to map the module name to one of these text records:

```text
Document=<ModuleIdentifier>/<DocTlibVer>
Module=<ModuleIdentifier>
Class=<ModuleIdentifier>
BaseClass=<ModuleIdentifier>
```

`BaseClass` indicates a designer module. A designer module also needs designer storage in the CFB root.

## 12. Module streams

Each module's source stream lives in `/VBA/<MODULESTREAMNAME>`.

Layout:

```text
PerformanceCache:      MODULEOFFSET bytes
CompressedSourceCode:  bytes from MODULEOFFSET through EOF
```

Reading algorithm:

```pseudo
function read_module_source(vba_storage, module, code_page):
    stream = read_stream(vba_storage, module.stream_name)
    if module.text_offset > len(stream): error
    compressed = stream[module.text_offset:]
    source_bytes = ovba_decompress(compressed)
    return decode_mbc source_bytes using code_page
```

Writing algorithm, minimal source replacement:

```pseudo
function replace_module_source(vba_storage, module, new_text, code_page):
    old_stream = read_stream(vba_storage, module.stream_name)
    prefix = old_stream[0:module.text_offset]
    new_bytes = encode_mbc new_text using code_page
    new_compressed = ovba_compress(new_bytes)
    write_stream(vba_storage, module.stream_name, prefix + new_compressed)
    set_vba_project_version_to_ffff()
```

Canonical strict writer alternative:

```pseudo
module.text_offset = 0
module_stream = ovba_compress(encoded_source)
rebuild_dir_with_updated_MODULEOFFSET(module, 0)
write_stream(vba_storage, module.stream_name, module_stream)
set_vba_project_version_to_ffff()
remove_srp_streams()
```

The minimal strategy avoids rewriting `dir`, but preserves a stale module performance cache prefix. Setting `_VBA_PROJECT.Version` to `0xFFFF` is therefore important.

## 13. `PROJECT` stream

The `PROJECT` stream is MBCS text encoded with the project code page. It contains project properties and the module type declarations used to disambiguate module kinds.

Common module lines:

```text
Document=ThisWorkbook/&H00000000
Document=Sheet1/&H00000000
Module=Module1
Class=Class1
BaseClass=UserForm1
```

For source replacement only, preserve `PROJECT` unchanged.

For add/remove/rename:

- Add or remove the matching `Module=`, `Class=`, or `BaseClass=` line.
- Do not invent document modules. They are tied to the host document.
- Do not add a `BaseClass` designer module unless you also create matching designer storage and form data.

## 14. `PROJECTwm` stream

`PROJECTwm` maps MBCS module names to UTF-16 module names.

Layout:

```text
NAMEMAP[]
Terminator: u16 = 0x0000
```

Each `NAMEMAP` contains:

```text
ModuleName:        null-terminated MBCS string
ModuleNameUnicode: null-terminated UTF-16 string
```

The map order must match the module order in `PROJECTMODULES`.

For source replacement only, preserve `PROJECTwm` unchanged.

For add/remove/rename, update it consistently with `PROJECTMODULES` and `PROJECT`.

## 15. Compression: high-level rules

MS-OVBA compression is required for:

- the entire `dir` stream
- every module's `CompressedSourceCode`

A compressed container starts with signature byte `0x01`, followed by one or more compressed chunks.

Each compressed chunk has:

```text
CompressedChunkHeader: u16 little-endian
CompressedChunkData:   variable
```

Header bits:

```text
bits 0-11:  CompressedChunkSize = chunk_size_in_bytes - 3
bits 12-14: CompressedChunkSignature = 0b011
bit 15:     CompressedChunkFlag, 1 = token-compressed, 0 = raw/uncompressed
```

Helpers:

```pseudo
chunk_size(header) = (header & 0x0FFF) + 3
chunk_flag(header) = (header >> 15) & 1
chunk_signature(header) = (header >> 12) & 0x7
make_header(size, flag) = ((size - 3) & 0x0FFF) | 0x3000 | (flag << 15)
```

A decompressed chunk is at most 4096 bytes. Every decompressed chunk except the final one must be exactly 4096 bytes.

## 16. Decompression pseudocode

This pseudocode is implementation-ready. It uses zero-based arrays and absolute offsets into the output buffer.

```pseudo
function ovba_decompress(input: bytes) -> bytes:
    if len(input) == 0 or input[0] != 0x01:
        error "bad compressed container signature"

    pos = 1
    out = byte_array()

    while pos < len(input):
        chunk_start = pos
        header = read_u16le(input, pos)
        size = (header & 0x0FFF) + 3
        signature = (header >> 12) & 0x7
        flag = (header >> 15) & 0x1

        if signature != 0b011:
            error "bad compressed chunk signature"

        chunk_end = min(len(input), chunk_start + size)
        pos = chunk_start + 2
        decompressed_chunk_start = len(out)

        if flag == 0:
            # Raw chunks contain exactly 4096 data bytes.
            if pos + 4096 > len(input): error
            out.extend(input[pos : pos + 4096])
            pos += 4096
        else:
            while pos < chunk_end:
                flag_byte = input[pos]
                pos += 1

                for bit_index in 0..7:
                    if pos >= chunk_end:
                        break

                    is_copy = (flag_byte >> bit_index) & 1

                    if is_copy == 0:
                        out.append(input[pos])
                        pos += 1
                    else:
                        if pos + 2 > chunk_end: error
                        token = read_u16le(input, pos)
                        pos += 2

                        (offset, length) = unpack_copy_token(
                            token,
                            decompressed_current = len(out),
                            decompressed_chunk_start = decompressed_chunk_start
                        )

                        copy_source = len(out) - offset
                        if copy_source < decompressed_chunk_start: error

                        # Overlap is allowed and required.
                        for i in 0..length-1:
                            out.append(out[copy_source + i])

    return bytes(out)
```

## 17. Copy token decoding

Copy tokens are 16-bit little-endian integers. They encode an offset and length. The bit allocation changes depending on how far into the current decompressed chunk the decoder is.

```pseudo
function copy_token_help(decompressed_current, decompressed_chunk_start):
    difference = decompressed_current - decompressed_chunk_start
    bit_count = ceil_log2(difference)
    bit_count = max(bit_count, 4)

    length_mask = 0xFFFF >> bit_count
    offset_mask = (~length_mask) & 0xFFFF
    maximum_length = length_mask + 3

    return (length_mask, offset_mask, bit_count, maximum_length)

function unpack_copy_token(token, decompressed_current, decompressed_chunk_start):
    (length_mask, offset_mask, bit_count, _) = copy_token_help(
        decompressed_current,
        decompressed_chunk_start
    )

    length = (token & length_mask) + 3
    offset = ((token & offset_mask) >> (16 - bit_count)) + 1
    return (offset, length)
```

`ceil_log2` should return the smallest integer `n` such that `2^n >= difference`. Difference should be at least `1` when a copy token appears.

## 18. Compression implementation notes

A correct writer needs an encoder, not only a decoder. The spec's compression algorithm searches backwards within each 4096-byte decompressed chunk and emits either literal tokens or copy tokens.

Minimum correct encoder strategy:

1. Write container signature `0x01`.
2. Split source into decompressed chunks of up to 4096 bytes.
3. For each full 4096-byte chunk, raw chunk encoding is acceptable and exact.
4. For the final chunk, avoid raw encoding unless it is exactly 4096 bytes, because raw chunks are padded to 4096 bytes and would append extra null bytes to the decompressed source.
5. For non-full final chunks, use token-compressed encoding.
6. Token-compressed chunks group up to eight tokens after one flag byte.
7. A flag bit of `0` means the token is one literal byte.
8. A flag bit of `1` means the token is a two-byte copy token.
9. Use overlapping byte-copy semantics when matching and decoding.

Literal-only compression works only while the token-compressed data fits within the 4096-byte chunk-data limit. Literal-only data size is `n + ceil(n / 8)`. Therefore it cannot encode a 4096-byte chunk and cannot encode a large final chunk near 4096 bytes. Do not silently fall back to raw encoding for a non-full final chunk, because that changes the decompressed source by adding null bytes.

Safer MVP writer policy:

```pseudo
if chunk_len == 4096:
    write_raw_chunk_exactly()
elif chunk_len + ceil(chunk_len / 8) <= 4096:
    write_literal_only_token_chunk()
else:
    use_full_lz_copy_token_encoder_or_fail_loudly()
```

A production writer should implement the full matching algorithm so any final chunk length can be encoded exactly.

## 19. Compression encoder pseudocode outline

```pseudo
function ovba_compress(input: bytes) -> bytes:
    out = byte_array([0x01])
    cursor = 0

    while cursor < len(input):
        chunk = input[cursor : min(cursor + 4096, len(input))]
        cursor += len(chunk)

        if len(chunk) == 4096:
            # Exact and simple. Header size is 4098 total bytes.
            header = make_header(4098, flag=0)
            out.write_u16le(header)
            out.extend(chunk)
            continue

        encoded = encode_token_chunk(chunk)
        if len(encoded) > 4096:
            error "token chunk too large; implement full LZ matching"

        header = make_header(2 + len(encoded), flag=1)
        out.write_u16le(header)
        out.extend(encoded)

    return bytes(out)
```

`encode_token_chunk` should use the current position inside `chunk` as `DecompressedCurrent` and search backward to find the longest match of at least 3 bytes. Then pack copy tokens using the same `copy_token_help` logic as decompression.

Literal-only token chunk:

```pseudo
function encode_literal_only_token_chunk(chunk):
    encoded = byte_array()
    i = 0
    while i < len(chunk):
        group = chunk[i : min(i + 8, len(chunk))]
        encoded.append(0x00)  # all literal tokens
        encoded.extend(group)
        i += len(group)
    return bytes(encoded)
```

Full matching outline:

```pseudo
function find_match(chunk, current):
    best_len = 0
    best_candidate = 0
    candidate = current - 1

    while candidate >= 0:
        c = candidate
        d = current
        length = 0

        while d < len(chunk) and chunk[c] == chunk[d]:
            length += 1
            c += 1
            d += 1
            if c >= current + length: # allow overlap carefully in actual implementation
                pass

        if length > best_len:
            best_len = length
            best_candidate = candidate

        candidate -= 1

    if best_len >= 3:
        (_, _, _, max_len) = copy_token_help(current, 0)
        return (current - best_candidate, min(best_len, max_len))
    else:
        return (0, 0)
```

The production version must carefully allow overlap according to the spec's byte-copy behavior and must avoid reading beyond the chunk.

## 20. Read pipeline

```pseudo
function extract_modules_from_xlsm(xlsm_path):
    zip = open_zip(xlsm_path)
    vba_bin = zip.read("xl/vbaProject.bin")
    return extract_modules_from_vba_project_bin(vba_bin)

function extract_modules_from_vba_project_bin(vba_bin):
    cfb = open_cfb(vba_bin)
    vba = cfb.storage("VBA")

    dir_compressed = vba.stream("dir")
    dir_bytes = ovba_decompress(dir_compressed)
    dir_model = parse_dir(dir_bytes)

    modules = []
    for module_record in dir_model.modules:
        stream = vba.stream(module_record.stream_name)
        compressed_source = stream[module_record.text_offset:]
        source_bytes = ovba_decompress(compressed_source)
        source_text = decode(source_bytes, code_page=dir_model.code_page)
        modules.append(module_record.with_source(source_text))

    return VbaProject(code_page=dir_model.code_page, modules=modules)
```

## 21. Write pipeline for source replacement

```pseudo
function replace_modules_in_xlsm(input_path, output_path, replacements):
    zip_in = open_zip(input_path)
    vba_bin = zip_in.read("xl/vbaProject.bin")

    cfb = open_cfb_mutable(vba_bin)
    vba = cfb.storage("VBA")

    dir_bytes = ovba_decompress(vba.stream("dir"))
    dir_model = parse_dir(dir_bytes)

    for (module_name, new_source_text) in replacements:
        module = dir_model.find_module_by_name(module_name)
        encoded_source = encode(new_source_text, code_page=dir_model.code_page)
        compressed_source = ovba_compress(encoded_source)

        old_stream = vba.stream(module.stream_name)
        prefix = old_stream[0 : module.text_offset]
        vba.write_stream(module.stream_name, prefix + compressed_source)

    patch_vba_project_version(cfb, 0xFFFF)

    new_vba_bin = cfb.serialize()
    zip_out = copy_zip_replacing(zip_in, "xl/vbaProject.bin", new_vba_bin)
    zip_out.write(output_path)
```

Do not update source in place. Create a new workbook. Preserve ZIP entries, compression methods, and metadata where practical.

## 22. Adding a standard module

Adding a standard module is the safest topology change.

Required updates:

1. Add a module stream under `/VBA/<stream_name>`.
2. Add a `MODULE` record to `PROJECTMODULES`.
3. Increment `PROJECTMODULES.Count`.
4. Add `Module=<module_name>` line to `PROJECT` stream.
5. Add a `NAMEMAP` entry to `PROJECTwm` if present.
6. Recompress and write the `dir` stream.
7. Set `_VBA_PROJECT.Version = 0xFFFF`.
8. Remove `__SRP_*` streams for strict output.

Standard module source should start with the correct attribute line:

```text
Attribute VB_Name = "ModuleName"
```

Then include the body text.

## 23. Renaming modules

Renaming touches multiple places:

- `MODULENAME`
- `MODULENAMEUNICODE`
- `MODULESTREAMNAME`, if stream name is changed
- actual CFB stream name, if changed
- `PROJECT` stream module line
- `PROJECTwm` name mapping
- `Attribute VB_Name = "..."` inside module source

Do not implement rename by changing only the source attribute. That creates an inconsistent project.

## 24. Removing modules

Removing a module requires:

- removing its `MODULE` record
- decrementing `PROJECTMODULES.Count`
- removing its module stream
- removing the corresponding `PROJECT` line
- removing the corresponding `PROJECTwm` mapping
- preserving or updating references as needed

Do not remove host document modules such as `ThisWorkbook` or worksheet modules in Excel unless you also understand the host application's object model storage.

## 25. UserForms and designer modules

A designer module has `BaseClass=<ModuleIdentifier>` in the `PROJECT` stream and has related designer storage in the CFB root. For Office Forms, that storage contains Office Forms binary data outside the simple module stream.

Safe behavior:

- You may replace the source code portion of an existing UserForm module.
- You should preserve the designer storage unchanged.
- Do not create new UserForms unless you implement the Office Forms binary format.
- Do not delete UserForms unless you remove both module metadata and designer storage.

## 26. Digital signatures and project protection

Direct modification of `vbaProject.bin` will generally invalidate VBA project signatures. If the host package contains signature-related parts, do not claim signatures are preserved unless you explicitly verify and re-sign.

Project protection fields such as `CMG`, `DPB`, and `GC` in the `PROJECT` stream are obfuscated/encrypted values. For module source replacement, preserve them unchanged. Do not implement password removal as part of a normal reader/writer.

## 27. Error handling

Recommended errors:

| Error | Trigger |
|---|---|
| `MissingVbaProject` | no `vbaProject.bin` found |
| `InvalidCfb` | `vbaProject.bin` is not a readable CFB |
| `MissingVbaStorage` | no `VBA` storage |
| `MissingDirStream` | no `/VBA/dir` stream |
| `BadCompressedSignature` | compressed container does not start with `0x01` |
| `BadChunkSignature` | chunk header bits 12-14 are not `0b011` |
| `UnsupportedCodePage` | project code page cannot be decoded/encoded |
| `MalformedDirRecord` | bad id, bad size, impossible offset, or early EOF |
| `MissingModuleStream` | module stream not found |
| `ModuleOffsetOutOfRange` | `MODULEOFFSET > stream length` |
| `UnsupportedTopologyChange` | requested add/remove/rename cannot be safely done |
| `WouldInvalidateSignature` | caller required signature preservation but source changed |

Default behavior should be strict. Add a separate permissive mode only for forensic extraction.

## 28. Test plan

Use the PDF examples as baseline test vectors, especially the compression examples in section 3.2.

Required tests:

1. Decompress the no-compression example.
2. Decompress the normal-compression example.
3. Decompress the maximum-compression example.
4. Parse a real `.xlsm` with one standard module.
5. Parse a real `.xlsm` with standard, class, document, and UserForm modules.
6. Extract source and compare with VBE export.
7. Replace source in a standard module and open in Excel.
8. Replace source in a class module and open in Excel.
9. Replace source in an existing UserForm code module while preserving form layout.
10. Test non-ASCII source under project code page 1252.
11. Test a non-Latin code page if possible, such as 932, because MBCS can use lead bytes.
12. Test compression with source longer than 4096 bytes.
13. Test final chunk lengths: 1, 8, 9, 100, 3640, 3641, 4095, 4096.
14. Test overlapping copy tokens.
15. Verify that `_VBA_PROJECT.Version` is written as `0xFFFF`.
16. Verify that `__SRP_*` streams are ignored on read.
17. Verify that unsupported signatures are detected and not silently preserved.

## 29. Suggested public API

```text
list_modules(path) -> list[ModuleInfo]
extract_modules(path, output_folder) -> list[ExportedModule]
read_vba_project(path) -> VbaProject
replace_module(path, module_name, source_text, output_path) -> WriteResult
replace_modules(path, replacements, output_path) -> WriteResult
add_standard_module(path, module_name, source_text, output_path) -> WriteResult
remove_standard_module(path, module_name, output_path) -> WriteResult
```

`ModuleInfo` should include:

```text
name
stream_name
kind
text_offset
source_length_decompressed
code_page
is_read_only
is_private
has_designer_storage
```

## 30. Implementation prompt for another LLM

Use this prompt to hand off implementation:

```text
Implement a library for reading and writing VBA modules using MS-OVBA. Use the attached implementation guide as the primary design. Start with read-only extraction from Excel `.xlsm` files. The input file contains `xl/vbaProject.bin`, which is an OLE Compound File Binary object. Implement or use a CFB reader, then parse `/VBA/dir` after MS-OVBA decompression. Extract `PROJECTCODEPAGE`, `PROJECTMODULES`, each `MODULESTREAMNAME`, and each `MODULEOFFSET`. For each module stream, slice from `MODULEOFFSET`, decompress using MS-OVBA compression, decode with the project code page, and return source text. Add tests for the MS-OVBA compression examples and real `.xlsm` samples. After read-only extraction passes, implement source replacement by encoding replacement source, compressing it, replacing the module stream suffix, and setting `_VBA_PROJECT.Version` to `0xFFFF`. Do not support UserForm creation, password removal, or signature preservation in the first implementation.
```

## 31. Key traps

- The module stream name is not always the same as the visible module name. Use `MODULESTREAMNAME`.
- The source does not necessarily start at byte zero. Use `MODULEOFFSET`.
- The source bytes are compressed. Decompress first.
- The decompressed source bytes are MBCS, not necessarily UTF-8.
- `MODULETYPE` id `0x0022` covers document, class, and designer modules. Use `PROJECT` stream lines to disambiguate.
- `dir` itself is compressed.
- CFB stream names are case-insensitive in the spec language, but preserve original casing when rewriting.
- Raw chunks can introduce null padding for non-full final chunks. Do not use raw chunk encoding for a final chunk shorter than 4096 bytes unless you explicitly accept changed output.
- Digital signatures are usually invalidated by any source change.
- UserForms are code plus designer storage. The code stream alone is not the whole form.

## 32. Practical build order

1. Build `ovba_decompress` with spec test vectors.
2. Build a CFB stream lister.
3. Open `vbaProject.bin` and list root + `/VBA` streams.
4. Decompress `/VBA/dir`.
5. Parse `PROJECTCODEPAGE` and `PROJECTMODULES`.
6. Extract module source text.
7. Add export to `.bas`, `.cls`, `.frm` naming conventions.
8. Build `ovba_compress` with round-trip tests.
9. Replace existing module source only.
10. Add standard-module creation.
11. Add rename/remove only after project stream and `PROJECTwm` rewriting are correct.

That sequence keeps failure domains isolated. The critical boundary is this: reading modules only requires understanding enough metadata to find and decode source; writing project topology requires maintaining every cross-reference that makes the VBA project coherent.
