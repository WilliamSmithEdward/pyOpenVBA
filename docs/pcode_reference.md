# VBA7 P-code Reference

Field guide to the compiled VBA7 p-code format, as reverse-engineered
for pyOpenVBA. Covers the on-disk layout, the instruction encoding, and
-- the part no public reference documents -- how compiled name operands
resolve back to real identifiers.

Everything here was verified empirically against real compiled modules
produced by Microsoft Office itself. Where something is unresolved it is
labelled **OPEN**, not glossed over.

---

## 1. Scope: one format, every host

VBA7 p-code is a **single format shared by every VBA host**. Excel,
Word, PowerPoint, and Access all embed the same structures; only the
*container* differs.

| Host | Where the module stream lives |
|------|------------------------------|
| Excel `.xlsm` / `.xlsb` / `.xlam` | CFB stream in `VBA/` inside `xl/vbaProject.bin` |
| Word `.docm` / `.dotm` | same, inside `word/vbaProject.bin` |
| PowerPoint `.pptm` / `.potm` | same, inside `ppt/vbaProject.bin` |
| Access `.accdb` | LVAL rows in the ACE page store (no CFB) |
| Legacy `.xls` / `.doc` / `.ppt` | the whole file is the CFB |

The project-wide identifier table (section 4) likewise uses the same
`CC 61` layout in all hosts. Consequently a p-code reader/writer is
written **once** and works everywhere; only stream location and the
32/64-bit opcode remap vary.

### 1.1 Why p-code matters per host

This is the decisive architectural asymmetry:

- **Excel / Word / PowerPoint treat the MS-OVBA *source* as
  authoritative** and recompile from it. pyOpenVBA deliberately zeroes
  the `_VBA_PROJECT` performance cache on a mutating save
  (`invalidate_vba_project_cache`), and the host rebuilds p-code on next
  open. Editing their p-code is therefore *moot*.
- **Access treats the compiled p-code as authoritative.** Its OVBA
  source cache is passive -- zero-filling it leaves the displayed source
  intact -- and no recompile-from-source trigger was found (including
  the documented `MSACCESS.EXE /decompile`, which does not rescue
  externally-mutated files).

So p-code writing is the enabling capability for **Access** structural
edits, and a p-code *decompiler* is valuable everywhere: it recovers
what the engine will actually execute, independent of the source cache.

---

## 2. Module stream layout

```
+--------------------------------------------------+
| binary header + declaration / indirect / object  |   <- pre-CAFE
| tables, identifier hash buckets                  |
+--------------------------------------------------+
| 0xCAFE magic ..................................  |   <- p-code region
| per-line record table, then per-line instructions |
+--------------------------------------------------+
| MS-OVBA compressed source (MODULEOFFSET onwards) |
+--------------------------------------------------+
```

In pyOpenVBA terms the pre-CAFE + p-code region is exactly
`VBAModule.prefix_bytes` (bytes `0 .. MODULEOFFSET`), which is why the
compiled bytecode is reachable without touching the compressed source.

---

## 3. The CAFE p-code region

Fully specified and **verified byte-exact** by regeneration (section 7).

```
FE CA                      0xCAFE magic, little-endian on the wire
<2 bytes>                  reserved / version (preserve verbatim)
u16 num_lines              count of source-line records
num_lines x 12-byte record:
    bytes [0:4]   reserved / flags   (preserve verbatim)
    bytes [4:6]   u16 line_length    (byte length of this line's p-code)
    bytes [6:8]   reserved           (preserve verbatim)
    bytes [8:12]  u32 line_offset    (offset from pcode_start;
                                      0xFFFFFFFF = source-only line)
<10 bytes>                 gap (preserve verbatim)
pcode_start:               per-line instruction bodies
```

A line whose `line_offset` is `0xFFFFFFFF` (or whose length is 0) carries
no p-code -- `Attribute` lines, `Option` lines, and blank lines. They
still occupy a record, which is why `num_lines` tracks source lines
rather than executable statements.

### 3.1 Instruction encoding

```
u16 header       opcode  = header & 0x03FF
                 op_type = header >> 10        (6 bits of flags)
operands         per the opcode table, in order:
                   "name" / "0x" / "imp_"          -> u16
                   "func_" / "var_" / "rec_" /
                   "type_" / "context_"            -> u32
varg payload     (only for varg opcodes, e.g. LitStr, Rem, QuoteRem):
                   u16 length, then <length> bytes,
                   padded to even length
```

The opcode table has 264 entries (`pyopenvba.vba_pcode.OPCODES_VBA7`).
On 64-bit hosts the raw opcode equals the canonical index; on 32-bit
VBA6/VBA7 hosts opcodes above 173 are shifted by small offsets
(`_translate_opcode`).

---

## 4. Identifier table (`_VBA_PROJECT`)

Project-wide, **not** per module. Magic `CC 61`. Identical layout in
Access (`AccessReader` reads it from a `CC 61` LVAL row) and in the
CFB-based hosts (the `VBA/_VBA_PROJECT` stream).

Record layout:

```
u8  name_len
u8  type_byte
[6-byte descriptor]        present for some type bytes (e.g. 0x80, 0xAC)
ASCII name                 name_len bytes
u16 id_value               per-record cookie (NOT the p-code operand)
0x10 0x00                  fixed trailer
```

Observed `type_byte` values: `0x00` intrinsic (e.g. `MsgBox`), `0x04`
type / reference name, `0x08` library, `0x0C` module or project, `0x80`
special intrinsic (`_Evaluate`), `0xAC` procedure with a body.

Because descriptor size varies, records parse most reliably when
**anchored on the `10 00` trailer** rather than chained forward from a
guessed start.

Two traps, both hit in practice:

- The 6-byte-descriptor variant can fabricate a record a few bytes
  *before* the real table, producing an equal-length chain that starts
  mid-record (symptom: a leading `cel` or `xcel` instead of `Excel`).
- The fix is to validate the type byte: every observed `type_byte` is a
  **multiple of 4 and at most `0xAC`**. That rejects ASCII bytes posing
  as types (e.g. `0x45` `E`). With this check the first record is the
  host name -- `Excel`, `Word`, `PowerPoint` -- on 10/10 fixtures.

The table is ordered, and that ordinal position is what p-code
references.

---

### 4.1 The identifier hash

Each record carries a `u16` alongside the name. It is the low word of the
**OLE Automation name hash** -- `LHashValOfNameSysA` from `OLEAUT32.dll`,
the same hash type libraries use for name lookup. VBE7's identifier
intern routine calls it and keeps the low 16 bits; this was found by
tracing VBE7.DLL from the predefined-identifier setup (the code that adds
`Win16` / `Win32` / `Win64` / `Mac` / `VBA6` / `VBA7`) down to the call.

```
h = 0x0DEADBEE
for b in name_bytes:                     # code-page bytes, not characters
    h = (37 * h + LOOKUP[b]) & 0xFFFFFFFF
id = (h % 65599) & 0xFFFF                # unsigned, low 16 bits
```

Two things about this are easy to get wrong, and both were wrong in an
earlier revision of this document:

- **The hash runs over code-page bytes, not characters.** Identifiers are
  stored MBCS-encoded, and each byte indexes the table.
- **`LOOKUP` is a real 384-byte table, not "uppercase ASCII".** Its ASCII
  half folds case and -- a genuine quirk -- maps `W` to `V` and `Y` to
  `U`. Its upper half folds accented Latin-1 onto base letters, so
  `0xC0`-`0xC5` (`A` with any accent) all hash as `A`. Treating high
  bytes as identity gets 4/20 accented identifiers right; the real table
  gets 20/20.

Which table applies is **fixed**, which is what makes this reliable.
`LHashValOfNameSysA` selects by `PRIMARYLANGID(lcid)`, and [MS-OVBA]
requires `PROJECTLCID` to be `0x00000409` (English US) -- confirmed in
every fixture on disk -- so the selection always lands on the `default`
branch, `Lookup_16`. VBA uses that one table whatever the project's code
page or the author's language. `syskind` matters only for Mac, where
`SYS_MAC` sets a mask shifting high bytes into the table's third section.

Verified exact on **6,825 ASCII names and 20 accented names**, with every
one of **14,878 compact records reproduced byte for byte**. That makes a
new identifier record fully generatable:
`pcode_hash.encode_identifier_record(name, type_byte, code_page=...)`
emits the exact `<u8 len><u8 type><name bytes><u16 hash><10 00>`.

(An earlier revision modelled this as `37*h + charval` from a per-length
"seed". There is no seed -- that was an artifact of missing the
`0x0DEADBEE` initialiser and using a signed rather than unsigned
reduction. The pieces that model did recover -- multiplier 37, modulus
65599, the 16-bit field, the `W`/`Y` fold -- are exactly the pieces of
`LHashValOfNameSysA`.)

Reference: ReactOS / Wine `dll/win32/oleaut32/hash.c`. Note that
ReactOS's `Lookup_64` (Japanese) is missing its first 13 entries
upstream; `Lookup_16`, the only table VBA reaches, is intact.

## 5. Name operand resolution (key finding)

Compiled `name` operands do **not** carry the record's `id_value`.
They are a linear function of the record's **ordinal index**:

```
name_operand     = 0x20E + 2 * identifier_index
identifier_index = (name_operand - 0x20E) / 2
```

Verified across multiple projects and hosts (Excel `.xlsm` / `.xlsb`,
PowerPoint `.pptm`). Worked example -- one module referencing five
identifiers, compiled by Excel:

| operand | index | identifier |
|---------|-------|------------|
| `0x0230` | 17 | `alpha` |
| `0x0232` | 18 | `beta` |
| `0x0234` | 19 | `MsgBox` |
| `0x0236` | 20 | `Helper` |
| `0x0238` | 21 | `gamma` |

**OPEN:** the constant `0x20E` held on every sample tested, but its
derivation is unknown; it presumably reserves operand space for
built-in / predeclared slots. Treat it as calibrated, not proven
universal.

### 5.1 The built-in identifier table (second operand space)

Operands **below `0x20E`** address a different space: a fixed table of
built-in names held by the **VBA runtime itself**, not stored in the
file. `Left`, for instance, appears nowhere in either the module stream
or `_VBA_PROJECT`, yet is called by operand `0x00DC`.

A user identifier that **collides with a built-in name reuses that
built-in's operand** instead of gaining a project-table entry. Declaring
all 26 single letters and checking which reached the project table
isolates this exactly: every letter appears *except* `b` and `f`, which
instead compile to the fixed operands `0x0018` and `0x00A4`. The same
`0x00A4` is emitted for a user function named `F`, so matching is
case-insensitive.

The table is **ordered alphabetically**, which makes it mappable by
probing and verifiable: sorting confirmed entries by operand must
reproduce alphabetical order.

| operand | name | operand | name |
|---------|------|---------|------|
| `0x0012` | `Array` | `0x00A4` | `f` |
| `0x0018` | `b` | `0x00AC` | `Format` |
| `0x0034` | `CDec` | `0x00B0` | `FreeFile` |
| `0x003A` | `ChDir` | `0x00C8` | `Input` |
| `0x004C` | `CurDir` | `0x00DC` | `Left` |
| `0x005A` | `Date` | `0x00FA` | `Mid` |
| `0x007E` | `Dir` | `0x00FE` | `MidB` |
| `0x0084` | `DoEvents` | `0x0134` | `Randomize` |
| `0x009A` | `Error` | `0x0140` | `RGB` |
| | | `0x0146` | `Seek` |
| | | `0x015C` | `String` |

All twenty are strictly alphabetical by operand, which is the self-check
`builtin_table_is_ordered()` enforces on additions.

Most VBA library functions do **not** live here. Probing 141 built-ins
one per source line showed the great majority -- `Abs`, `Chr`, `MsgBox`,
`UCase`, `Now`, `Split`, and so on -- resolving through ordinary
*project*-table operands, exactly like user names. What lands in the
runtime space is a small set that VBA's parser treats specially, largely
the names that double as statements or as file-I/O keywords.

Each entry was confirmed by compiling a probe with Excel and reading the
operand back. The map is **partial** -- it covers what has been probed,
not the whole runtime table -- and extends by probing further names;
`BUILTIN_OPERANDS` in `pcode_names.py` carries it.

**OPEN:** the table's full contents and its index origin. Note also that
names resolved this way lose their original casing, since the runtime
table supplies the spelling.

## 6. Declaration operands (`func_` / `var_`)

`FuncDefn` / `VarDefn` operands are **u32 offsets from a per-module
declaration base** to a `u16` which is itself a *name operand*, resolved
per section 5. Two hops:

```
decl_operand --(+ DECL_BASE)--> u16 name_operand --(section 5)--> name
```

Worked example (same module, `DECL_BASE = 450`):

| instruction | operand | u16 at base+operand | resolves to |
|---|---|---|---|
| `FuncDefn` | `func_00000000` | `0x022E` | `S` |
| `VarDefn`  | `var_00000058`  | `0x0230` | `alpha` |
| `VarDefn`  | `var_00000070`  | `0x0232` | `beta` |
| `FuncDefn` | `func_00000088` | `0x0236` | `Helper` |
| `VarDefn`  | `var_000000E0`  | `0x0238` | `gamma` |

`DECL_BASE` varies per module -- 450 is typical for a standard module,
but a module declaring types moves it (496, 542, 588 all observed), and
document modules such as `Sheet1` / `ThisWorkbook` sit at 546. No header
field holding it has been located, so it is **calibrated** against two
independent constraints: the type-reference table must validate at the
implied position (section 6.3), and every `func_` / `var_` / `rec_`
operand must land on a resolvable identifier, scored so procedure names
and typed parameters outweigh incidental hits.

Both constraints are needed. Name resolution alone settles on bases
shifted by `0x2C` or `0x58` that still resolve *some* operands, and a
module whose only declaration is an `Enum` has no `func_` or `var_`
operand at all -- `rec_` has to count too, or calibration returns
nothing.

### 6.1 The declaration record

Every declaration -- procedure, parameter, local, module-level variable,
UDT member, `Type` / `Enum` -- is a record in the same table, introduced
by a `u16` **tag** in the two bytes immediately before it. The fields
that decode are:

| offset | size | meaning |
|--------|------|---------|
| `-0x02` | u16 | record tag; bit `0x20` = the source wrote an `As` clause |
| `+0x00` | u16 | name operand (section 5). **Bit 0 is a flag**, set on object / class-typed declarations; mask it off before resolving |
| `+0x06` | i16 | stack-frame offset (negative for locals, positive for parameters and UDT members) |
| `+0x0E` | u8 | VARTYPE, or the low half of a descriptor offset |
| `+0x0F` | u8 | flags: `0x01` ByRef, `0x10` the declaration is a UDT member |
| `+0x10` | u16 | **discriminator**: `0xFFFF` = the type is the plain VARTYPE at `+0x0E`; anything else = `+0x0E` is a `u16` offset to a type descriptor |

The bit-0 name flag is easy to miss and produced a real defect before it
was found: `Dim c As Collection` stored `0x0233` where `c` is `0x0232`,
so the name silently failed to resolve. Name operands are always even,
which is what frees the bit.

The frame offset is a useful cross-check on the type: a `Dim a(3) As
Long` local sits at `-32`, matching a 64-bit `SAFEARRAY` header plus one
`SAFEARRAYBOUND`; a two-dimensional array sits at `-40`, one bound more.

### 6.2 Declared types: the plain form

When the discriminator at `+0x10` is `0xFFFF`, the byte at `+0x0E` is a
standard OLE Automation **VARTYPE**:

| byte | VBA type | byte | VBA type |
|------|----------|------|----------|
| `0x02` | `Integer` (VT_I2) | `0x08` | `String` (VT_BSTR) |
| `0x03` | `Long` (VT_I4) | `0x09` | `Object` (VT_DISPATCH) |
| `0x04` | `Single` (VT_R4) | `0x0B` | `Boolean` (VT_BOOL) |
| `0x05` | `Double` (VT_R8) | `0x0C` | `Variant` (VT_VARIANT) |
| `0x06` | `Currency` (VT_CY) | `0x11` | `Byte` (VT_UI1) |
| `0x07` | `Date` (VT_DATE) | `0x14` | `LongLong` (VT_I8) |

Bit `0x40` above the VARTYPE marks a **constant** (`0x43` = Const +
Long, `0x48` = Const + String).

A declaration always carries a type even when the source wrote none:
`Dim x` records `Variant`, and under a `DefLng L` statement `Dim Lx`
records `Long`. Only the record tag's `0x20` bit says whether an `As`
clause was actually written, so it is what decides whether the
decompiler emits one.

### 6.3 Declared types: the `type_` indirect table

When the discriminator is **not** `0xFFFF`, `+0x0E` is a `u16`
`DECL_BASE`-relative offset to a **type descriptor**, whose `u16` tag
sits in the two bytes immediately before it. This is the mechanism
behind arrays, fixed-length strings, and every named type.

```
tag = kind | flags << 8
```

| kind | body | meaning |
|------|------|---------|
| `0x1B` | `<u32 array_info><u32 element>` | array |
| `0x1D` | `<u16 target><u32 0x25>` | named type reference |
| `0x20` | `<u16 length><u32>` | `String * length` |

| flag | meaning |
|------|---------|
| `0x08` | dynamic array (`Dim a()`), no static bounds |
| `0x10` | the declaration is a UDT member |
| `0x40` | the reference names a module-local `Type` |
| `0x60` | the reference names a module-local `Enum` |
| none | on a `0x1D` tag, the type comes from a referenced type library (`Collection`, `Worksheet`) |

An array's `element` word is either a VARTYPE outright or, when its low
byte is one of the three kinds, a **nested descriptor tag** whose body
follows immediately -- which is how `Dim p(3) As Ea` records both
"array" and "of enum `Ea`" (`1B 00 | <info> | 1D 60 | <index>`).

The `0x1D` `target` is `8 * index` into the module's **type-reference
table**, a run of 10-byte entries ending `26` bytes before `DECL_BASE`,
preceded by a 6-byte header whose `u32` repeats the table's own byte
length -- which makes the table self-describing and therefore a sharp
validity test for a candidate `DECL_BASE`:

```
<u16 tag><u16><u16 name_operand><u16><u16>
```

Entry tags seen so far: `0x0448` / `0x0440` a module-local `Type`,
`0x0048` a module-local `Enum`, `0x9428` / `0x9420` a type-library
class. The `name_operand` resolves through the project identifier table
like any other, so `Dim v As Alpha` recovers the literal name `Alpha`.

The decisive evidence that `target` identifies the type, rather than
merely correlating with it: compiling `Dim p As Alpha : Dim q As Bravo`
and the same module with the two declarations **swapped** swaps the two
target values, and declaring both variables as `Alpha` makes both
targets equal.

Verified across the whole fixture corpus: **834 modules, 5,287 p-code
lines, zero undecoded type descriptors.**

### 6.4 Procedure signatures

A procedure record links to its own signature, so nothing has to be
inferred from slot strides:

| offset | size | meaning |
|--------|------|---------|
| `+0x2C` | -- | a type field of the same shape as a variable's (VARTYPE at `+14`, discriminator at `+16`): the **return type** |
| `+0x36` | u32 | offset of the **first parameter**, `0xFFFFFFFF` when there are none |
| `+0x51` | u8 | bit `0x02` = the procedure is effectively `Public` |

and each parameter record adds:

| offset | size | meaning |
|--------|------|---------|
| `+0x16` | u32 | offset of the **next parameter**, `0xFFFFFFFF` at the end |
| `+0x1A` | u16 | low byte `0x04` = `ByVal`, `0x02` = an explicit `ByRef` keyword; high byte `0x02` = `Optional`, `0x04` = has a default value |

The head-and-link pair is what separates parameters from locals: both
live in the same slot region, and a fixed `+0x58 + k * 0x20` scan reads
`Dim`-declared locals as extra parameters. It also handles the case that
breaks a fixed stride outright -- a parameter with a default value has a
**longer record**, so the following parameter is not one stride away.

Default values are pushed onto the stack before the `FuncDefn` by a
`ConstFuncExpr` marker; the parameters whose flags say they have one
consume them in order.

`Sub` / `Function` / `Property` comes from the closing opcode
(`EndSub`, `EndFunc`, `EndProp`) paired with its `FuncDefn`; the
`FuncDefn`'s own `op_type` refines it:

| bit | meaning |
|-----|---------|
| `0x02` | the procedure returns a value (`Function`, `Property Get`) |
| `0x04` | the source wrote an explicit `Public` keyword |

So `Property Get` **is** distinguishable from `Property Let` / `Set`
after all (an earlier revision of this document said otherwise): Get
carries `0x02`, Let and Set do not. Let and Set remain
indistinguishable from each other; an object-typed final parameter is
the only available heuristic.

Visibility takes both fields: `op_type & 0x04` renders `Public`,
otherwise a clear `0x02` in the record byte at `+0x51` renders
`Private`, otherwise no keyword. `Static Sub` compiles to p-code
identical to a plain `Sub`; the distinction lives in record flags that
are not yet decoded.

`Friend`, and any signature containing `ParamArray`, are not encoded at
all: the compiler stores the original line as a `Reparse` payload
instead (section 6.6).

### 6.5 Declaration flags

The declaration keyword lives in the **`Dim` opcode's `op_type`**, and
whether an entry is a constant in the **`VarDefn`'s `op_type`**:

| `Dim` op_type | keyword | `VarDefn` op_type | meaning |
|---|---|---|---|
| `0x00` | `Dim` | `0x01` | variable |
| `0x01` | `Const` | `0x02` | constant |
| `0x08` | `Public` | | |
| `0x10` | `Private` | | |
| `0x20` | `Static` | | |

This also disambiguates two cases that otherwise look identical, since
both push literals before `VarDefn`: an array's bounds
(`Dim a(1 To 5)`) and a constant's value (`Const K = 7`). A third case
joins them -- `Dim tag As String * 8` pushes its length the same way --
and is told apart by the type descriptor, which already carries the
length.

### 6.6 Types, enums, and Reparse

`Type` opens both user-defined types and enums; the closing opcode
distinguishes them (`EndType` vs `EndEnum`), and the opcode's `op_type`
carries `0x02` for `Enum` plus `0x01` when a visibility keyword was
written. Which keyword is in the record's `u16` at `+0x10`: `1` for
`Public`, `0` for `Private`. The name comes from the `rec_` operand,
resolved exactly like `func_` / `var_`.

Members are declared with `DimImplicit` + `VarDefn`. An enum member
written without a value (`e1` rather than `e1 = 1`) compiles to an
argument-less `ArgsCall` on its own name.

```vba
Type T          Enum E
    a As Long       E1 = 1
    b As String     E2
End Type        End Enum
```

Constructs the p-code compiler does not encode are kept verbatim: the
`Reparse` opcode carries the **original source line** as its payload.
Observed for `Friend`, `ParamArray` signatures, `Tab(n)` and `Spc(n)`.
It is a gift to a decompiler -- those lines reproduce exactly -- and a
warning to an assembler, which must reproduce the text byte for byte.

### 6.7 Statement opcodes

The statement forms the decompiler renders, with the opcodes that carry
them:

| VBA | opcodes |
|---|---|
| `Set x = e` | `SetStmt` (marker), then `Set` / `MemSet` |
| `a(i) = e` | `ArgsSt` (value pushed first, then indices) |
| `o.M(i)` / `o.M(i) = e` | `ArgsMemLd` / `ArgsMemSt`; **push order is value (if any), then arguments, then the object last** -- so unwind object, arguments, value. `ArgsDict*` are the `!` forms, `Args*With` drop the object |
| `Dim a(3)` vs `Dim a(1 To 5)` | an **`OptionBase` precedes each dimension whose lower bound is implicit**. Without it two pushed literals are ambiguous -- one explicit range or two implicit-lower dimensions -- so `Dim a(3, 4)` and `Dim a(1 To 5)` are indistinguishable by literal count alone. `ReDim` uses the same markers; its `0x` operand counts dimensions, not literals |
| `o.m = e` / `.m = e` | `MemSt` / `MemStWith` |
| `x!key` | `DictLd` / `DictSt` and their `With` variants |
| `With o` ... `End With` | `StartWithExpr`, `With`, `EndWith` |
| `ReDim a(n)` | `Redim` (with a `type_` operand) |
| `For Each v In c` | `StartForVariable`, `EndForVariable`, `ForEach` |
| `New Class1` | `New` (its `imp_` operand indexes the type-reference table, `8 * index`) |
| `Call X(a)` vs `X a` | `ArgsCall`; `op_type` bit `0x10` = written **without** the `Call` keyword |
| `On Error GoTo L` | `OnError`; `op_type` `0x01` = `Resume Next`, `0x02` = `GoTo 0` |
| `Resume` | `Resume`; `op_type` `0x08` = bare, `0x01` = `Next`, `0` = a label |
| `L:` / `GoSub L` / `Return` | `Label`, `GoSub`, `Return` |
| `Exit Sub` / `Exit For` / ... | `ExitSub`, `ExitFor`, `ExitDo`, `ExitFunc` |
| `Erase a` | `Erase` |
| `Mid(s, 1, 2) = "x"` | `Mid` (value pushed first) |
| `LSet` / `RSet` | `LSet`, `RSet` |
| `Error 5` | `Error` |
| `Stop` | `Stop` |
| `x = 1: x = 2` | `BoS` between the statements (`BoSImplicit` inside a single-line `If`) |
| `If c Then a Else b` | `If`, `Else`, `EndIf` (the block form uses `IfBlock` / `ElseBlock` / `EndIfBlock`) |
| `#If` / `#Else` / `#End If` | `LbMark`, `LbIf`, `LbElse`, `LbEndIf` -- **both branches are compiled**, so the inactive one is still recoverable |
| `' comment` | `QuoteRem`, whose `0x` operand is the column it starts at (`0` = the whole line, otherwise it trails a statement) |
| `Rem comment` | `Rem` |
| line continuation | `LineCont`, payload = the continuation columns |

**File I/O** is a small language of its own:

| VBA | opcodes |
|---|---|
| `Open p For Output As #f` | `Sharp`, `LitDefault`, `Open` (mode in the operand: 1 Input, 2 Output, 4 Append, 8 Binary, 16 Random) |
| `Print #f, x` | `PrintChan`, then `PrintItemNL` / `PrintItemComma` / `PrintItemSemiColon` |
| `Write #f, x` | `WriteChan`, then the same item opcodes |
| `Input #f, x` | `Input`, `InputItem`, `InputDone` |
| `Line Input #f, s` | `LineInput` |
| `Close #f` / `Close` | `Close` / `CloseAll` |
| `Name a As b` | `Name` |
| `Debug.Print` / `Debug.Assert` | `Debug` + `PrintObj`, `Assert` |

**Statements and options** carry their argument in `op_type`:

| opcode | `op_type` | meaning |
|--------|-----------|---------|
| `Option` | `0x01` / `0x02` / `0x04` / `0x05` | `Option Base 1`, `Option Compare Text`, `Option Explicit`, `Option Private Module` |
| `DefType` | a VARTYPE | `DefInt` / `DefLng` / ...; the two operands are a 64-bit bitmap of the letters covered, bit 0 = `A` |
| `Coerce` | target type | `CInt` `0x02`, `CLng` `0x03`, `CSng` `0x04`, `CDbl` `0x05`, `CCur` `0x06`, `CDate` `0x07`, `CStr` `0x08`, `CBool` `0x0B`, `CLngLng` `0x0D`, `CByte` `0x11`, `CVar` `0x00` |
| `CoerceVar` | -- | `CVErr` |
| `LitVarSpecial` | `0` / `1` / `2` / `3` | `False`, `True`, `Null`, `Empty` |
| `LitSmallI2` | the value | small integer literal, no operand |

**Literals** encode their value across `u16` operand words, least
significant first: `LitDI2` / `LitDI4` / `LitDI8` decimal, `LitHI*`
`&H`, `LitOI*` `&O`, `LitR4` / `LitR8` IEEE floats, `LitCy` a currency
scaled by 10,000, `LitDate` an OLE automation date (days since
1899-12-30).

A handful of functions compile to **dedicated opcodes** rather than a
call, and must be rendered by name: `FnAbs`, `FnFix`, `FnInt`, `FnSgn`,
`FnLen`, `FnLenB`, `FnInStr` / `FnInStr3` / `FnInStr4` (and the `B`
variants), `FnStrComp` / `FnStrComp3`, `FnLBound` / `FnUBound` (whose
`0x` operand is the dimension), `FnMid`, `FnCurDir`, `FnDir`,
`FnError`, `FnFormat`, `FnFreeFile`, `FnStringVar`, `FnStringStr`.

## 7. Assembler status (p-code writing)

Verified milestones, all **byte-exact against Office-compiled output**:

1. **Instruction encoder** -- re-emit any decoded instruction to its
   on-disk bytes. *278/278 instructions byte-exact across the Office
   fixture set (Excel, Word, PowerPoint).*
2. **Full CAFE-region regeneration** -- rebuild the entire region from
   captured structure (record table, offsets, reserved / gap bytes, all
   bodies). *Exact on a 230-line / 20,960-byte module.*
3. **Same-length semantic edit** -- change a `LitStr` payload; the
   region re-disassembles to the new value with only the target bytes
   altered.
4. **Variable-length edit with reflow** -- grow or shrink an
   instruction, bump that line's `line_length` and every later
   `line_offset`, rebuild. *Re-disassembles correctly with all other
   lines preserved.*

Not yet built: emitting **new identifiers** (requires writing the
`_VBA_PROJECT` table and the module's identifier hash buckets -- the
`x` to `y` differential showed entries moving between `0xFFFFFFFF`
slots, i.e. a hash-addressed table whose hash function is **OPEN**), and
a full source-to-p-code compiler (lexer, parser, codegen).

---

## 8. Decompiler status (p-code reading)

Given a module stream plus its `_VBA_PROJECT`, every `name`, `func_`,
`var_` and `rec_` operand resolves to a source identifier, every
declared type resolves to a type name, and the stack machine replays
into expressions. The measured results:

- **Round-trip corpus: 37 of 39 entries reproduce the original source
  character for character**, 0 failing. The other two are cases VBA
  itself cannot round-trip (below).
- **Coverage sweep: 834 modules, 5,287 p-code lines across every
  fixture in the repository -- zero unmapped opcodes, zero undecoded
  type descriptors**, two unresolved names (the same two cases).

- **Semantic round-trip: 95 of 95 modules equivalent**, including an
  8,060-opcode real-world module. This is the strongest of the three,
  because it needs no original source: decompile a compiled module,
  recompile the result with Excel, and compare the opcode streams. If
  they agree, the reconstruction *means* the same thing, whatever the
  original text looked like.

The corpus is `docs/research/pcode/roundtrip.py`, the sweep is
`docs/research/pcode/sweep.py`, and the semantic gate is
`docs/research/pcode/semantic_roundtrip.py`. All three are runnable, and
the round-trip prints a unified diff for anything that is not
byte-identical rather than counting it as a pass.

The semantic gate earned its place immediately by catching a defect the
text round-trip structurally could not: every `ArgsMem*` / `ArgsDict*`
opcode pushes its **object last**, above the arguments, with the assigned
value pushed first. The decompiler popped arguments before the object, so
`ws.Cells(r, c).Interior.Color = HelperValue(idx)` came back as
`Cells(c, ws).Interior.Color = r` -- the object consumed as an argument
and the value expression dropped. Nothing in a source-first corpus
exercises that, because the bug only appears in code the corpus did not
write. Operand order is now measured, not assumed (section 6.7).

Line boundaries are deliberately excluded from the comparison: `Case 0:
x = f()` and the same statements on separate lines compile to identical
opcodes, so a pure reflow is reported as equivalent rather than as a
failure.

What round-trips exactly: expressions with precedence and parentheses;
`If` / `ElseIf` / `Else` and the single-line `If ... Then ... Else`;
`For` / `Next` with `Step`; `For Each`; `Do While` / `Loop`; `While` /
`Wend`; `Select Case`; procedure signatures including `Public` /
`Private`, `ByVal` / `ByRef`, `Optional` with defaults, parameter types
and return types; `Property Get` / `Let`; `Dim` / `Const` / `Public` /
`Private` / `Static`; arrays, dynamic arrays and `ReDim`; fixed-length
strings; `Type` and `Enum` including implicit member values; UDT-,
enum- and class-typed variables; `Set` and `New`; member access and
`With` blocks; `On Error` in all three forms, labels, `Resume`,
`GoSub` / `Return`; file I/O (`Open`, `Print #`, `Write #`, `Input #`,
`Line Input #`, `Close`); `Debug.Print` / `Debug.Assert`; `Mid` /
`LSet` / `RSet` statements; conversions; `Option` statements;
`DefType`; conditional compilation; comments in all three positions;
colon-separated statements; and numeric literals in decimal, hex and
octal.

### 8.1 What p-code does not preserve

Three things are genuinely unrecoverable, and it is worth being precise
about which, because they look like decoder gaps and are not:

1. **Identifier casing, when two spellings collide.** VBA folds
   identifiers case-insensitively into one project-table entry. A module
   with `Sub S()` and `Dim s As String` keeps a single spelling for
   both, and it is not always the one the declaration used. Likewise a
   name that lands in the runtime operand table (section 5.1) takes that
   table's spelling. These are the two `LOSSY` entries in the corpus.
2. **`Property Let` vs `Property Set`.** Both are value-consuming
   properties with the same `op_type`; only an object-typed final
   parameter hints at `Set`.
3. **`Declare`'s `Alias`.** A `Declare` compiles to a `FuncDefn` with no
   closing opcode. Its `Lib` string is a normal project identifier and
   so is recoverable, but the `Alias` string is not in the module stream
   or `_VBA_PROJECT` -- it lives in the `__SRP_*` caches.

Two further things are lossy only in the sense that both source forms
compile identically, so either is a correct decompilation: `Dim x` and
`Dim x As Variant` differ only in the record tag's `As` bit (which *is*
decoded), while `$` type suffixes and `Static Sub` versus a `Sub` whose
locals are all `Static` are not distinguished at all.

### 8.2 Expressions and control flow

P-code is a **stack machine**, so expressions reconstruct by simulating
it: `Ld b | Ld c | LitDI2 2 | Mul | Add | St a` pops back to
`a = b + c * 2`, with precedence and parentheses (`Paren`) preserved.

Binary operators map directly: `Add` `+`, `Sub` `-`, `Mul` `*`, `Div`
`/`, `IDiv`, `Mod`, `Pwr` `^`, `Concat` `&`, `Eq` `=`, `Ne` `<>`,
`Lt` / `Gt` / `Le` / `Ge`, `Is`, `Like`, and
`And` / `Or` / `Xor` / `Eqv` / `Imp`. Unary: `Not`, `UMi` (negation).

Control flow is explicit in the opcode stream, which makes block
structure and indentation recoverable:

| construct | opcodes |
|---|---|
| `If` / `ElseIf` / `Else` / `End If` | `IfBlock`, `ElseIfBlock`, `ElseBlock`, `EndIfBlock` |
| `If c Then a Else b` (single line) | `If`, `Else`, `EndIf` |
| `For` / `Next` | `StartForVariable`, `EndForVariable`, `For` / `ForStep`, `Next` / `NextVar` |
| `For Each` / `Next` | `StartForVariable`, `EndForVariable`, `ForEach` |
| `Do While` / `Loop` | `DoWhile`, `Loop` (also `DoUntil`, `LoopWhile`, `LoopUntil`) |
| `While` / `Wend` | `While`, `Wend` |
| `Select Case` | `SelectCase`, `CaseDone`, `CaseElse`, `EndSelect` |

Source lines map one-to-one onto p-code lines, including blank ones, so
layout is preserved by emitting an empty line for every p-code line with
no instructions. Statements *within* a line are separated by `BoS`.

---

## 9. Verification methodology

Two independent oracles, both **dev-time only** -- pyOpenVBA itself
remains pure-Python, with no COM and no Office dependency:

1. **Compile oracle.** pyOpenVBA writes VBA source into a workbook
   (pure Python); Excel opens it, compiles, and saves; pyOpenVBA reads
   the compiled `prefix_bytes` back. Excel is thus the *reference
   implementation* an assembler can be diffed against.
2. **Differential corpus.** Compile variants differing in exactly one
   identifier (`Hello` to `Hellp`, `Dim x` to `Dim y`, `"hi"` to
   `"ho"`) and diff the bytes. This localised the identifier hash
   buckets, the content-hash string, and ultimately the operand
   mapping.

3. **Round-trip gate** (`docs/research/pcode/roundtrip.py`). A corpus
   of 37 sources, each compiled by Excel, read back, decompiled, and
   compared to the original text character for character. Anything not
   byte-identical prints a unified diff; the two entries VBA itself
   cannot round-trip are listed explicitly rather than tolerated
   silently.
4. **Coverage sweep** (`docs/research/pcode/sweep.py`). Decompiles every
   module in every fixture on disk and reports what the renderer could
   not map -- opcodes, names, type descriptors. The round-trip proves
   the decompiler reproduces known sources; the sweep proves it never
   *silently* drops p-code from unknown ones.

Excel automation runs through `pyvbaharness`, which is hang-safe
(bounded, popup-aware, hard process reaping) -- necessary because a
compile error otherwise blocks on a modal dialog.

The identifier hash was pinned differently -- by static analysis of VBE7.DLL rather than differential probing. Locating the intern routine's call to `LHashValOfNameSysA` turned a months-of-probing problem into a one-line answer, after differential fitting had produced only an approximate per-length model. When a value looks like a bespoke hash, check whether it is a documented OS API first.

A note on method that cost real time: several findings looked settled
after one probe and were wrong. `ByVal` / `ByRef` was first recorded as
unencoded, `Property Get` / `Let` as indistinguishable, and the built-in
operand table as containing `Beep` at `0x18`. Each fell to the same
technique -- compile two modules differing in exactly one thing and diff
the bytes -- which is worth reaching for before concluding that
something is not encoded at all.

---

## 10. Open questions

Solved since earlier revisions, and no longer open: the `type_` indirect
table and the type-reference table (6.3), procedure signatures including
`Optional` and defaults (6.4), `Property Get` vs `Let`/`Set`, explicit
vs implicit `ByRef`, and `Dim x` vs `Dim x As Variant`.

Still open:

- Derivation of the `0x20E` name-operand base.
- Location of `DECL_BASE` (a module-header field?), currently
  calibrated -- reliably, but by search rather than by reading a field.
- The identifier hash is fully solved (section 4.1): it is
  `LHashValOfNameSysA`, a fixed global function, and both the hash and a
  complete compact identifier record are now byte-generatable. The one
  loose end is whether the same value also keys the module's in-memory
  identifier bucket table (VBE7 builds that at load time from the stored
  records, so it does not gate on-disk writing).
- The full contents of the runtime built-in identifier table
  (section 5.1); the mechanism is understood, the map covers 20 entries.
  In particular, what the single-letter entries `b` and `f` name.
- `Static Sub` versus a `Sub` whose locals are all `Static`: the record
  flags differ, the opcodes do not.
- `Declare`'s `Alias` string, which lives in the `__SRP_*` caches.
- Source-to-p-code compilation: lexer, parser, codegen, and the slot
  allocation Office's compiler performs -- including the frame offsets
  and the record links that section 6 documents reading.

---

## 11. Related

- [`docs/msaccess_lessons_learned.md`](msaccess_lessons_learned.md) --
  why Access VBA writing is hard and what the p-code layer gates.
- [`docs/access_pcode_re.md`](access_pcode_re.md) -- Access-specific
  p-code storage (`rU@` rows, `CAFE` rows, plaintext B9 / E3 stores).
- `src/pyopenvba/vba_pcode.py` -- the shipped disassembler.
- `docs/research/pcode/` -- assembler and decompiler prototypes backing
  this document.
