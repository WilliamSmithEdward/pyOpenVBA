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
| `0x0012` | `Array` | `0x00AC` | `Format` |
| `0x0018` | `b` | `0x00DC` | `Left` |
| `0x005A` | `Date` | `0x00FA` | `Mid` |
| `0x007E` | `Dir` | `0x015C` | `String` |
| `0x00A4` | `f` | | |

Each entry was confirmed by compiling a probe with Excel and reading the
operand back. The map is **partial** -- it covers what has been probed,
not the whole runtime table -- and extends by probing further names;
`BUILTIN_OPERANDS` in `pcode_names.py` carries it, with
`builtin_table_is_ordered()` as a self-check on additions.

**OPEN:** the table's full contents and its index origin (operands are
`2 * index` into an alphabetical list, but the complete name list lives
in the VBA runtime, not the file). Note also that names resolved this
way lose their original casing, since the runtime table supplies the
spelling.

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

`DECL_BASE` varies per module: 450 for the standard modules tested, 546
for document modules such as `Sheet1` / `ThisWorkbook`.

### 6.1 Declared types

The declaration record also carries the variable's **declared type**, as
a single byte at:

```
DECL_BASE + var_operand + 14
```

The value is a standard OLE Automation **VARTYPE** code, which makes the
mapping self-evident and complete:

| byte | VBA type | byte | VBA type |
|------|----------|------|----------|
| `0x02` | `Integer` (VT_I2) | `0x08` | `String` (VT_BSTR) |
| `0x03` | `Long` (VT_I4) | `0x09` | `Object` (VT_DISPATCH) |
| `0x04` | `Single` (VT_R4) | `0x0B` | `Boolean` (VT_BOOL) |
| `0x05` | `Double` (VT_R8) | `0x0C` | `Variant` (VT_VARIANT) |
| `0x06` | `Currency` (VT_CY) | `0x11` | `Byte` (VT_UI1) |
| `0x07` | `Date` (VT_DATE) | `0x14` | `LongLong` (VT_I8) |

Verified 12/12 against Excel-compiled modules, one per declared type.

**OPEN:** no module-header field was found holding `DECL_BASE`. It is
currently *calibrated* -- the unique base at which every `func_` /
`var_` operand in the module resolves to a valid identifier.
Deterministic and reliable in practice, but a proper header field may
exist.

---

### 6.2 Procedure signatures

A `FuncDefn` operand is an offset into the declaration table, and the
procedure's whole signature is reachable from it:

```
base + func_operand + 0x00              the procedure's own name entry
base + func_operand + 0x58 + k * 0x20   parameter k (VARTYPE at +14)
```

The **return type** is a second entry repeating the procedure name with
a VARTYPE set; its offset is not fixed, so it is found by scanning
forward for that name with a type.

`Sub` vs `Function` vs `Property` comes from the closing opcode --
`EndSub`, `EndFunc`, `EndProp` -- paired with its `FuncDefn` in order.

Two traps worth recording, because both produced wrong output before
being fixed:

- **Parameters and locals share the slot region.** A naive scan reads
  `Dim`-declared locals as extra parameters (`Sub S()` became
  `Sub S(r As Long)`). The `var_` operands of `VarDefn` mark exactly
  which slots are locals; stop the parameter scan there.
- **`DECL_BASE` must be calibrated on procedure names, not just on
  "something resolves".** Since a `FuncDefn` operand points at its own
  name, requiring every procedure name to land is a far stronger
  constraint. Calibrating loosely settles on a base shifted by `0x2C`
  that still resolves other operands. Scoring candidate bases by
  resolved names *plus* typed parameters pins it: a `Sub B` whose name
  collides with the built-in `b` (section 5.1) otherwise defeats a
  strict project-table-only rule, and a base shifted by `0x58` can
  satisfy the names alone by landing on the parameter slots. With
  scoring, every standard module tested calibrates to **450**.

Recovered signatures are exact for `Sub` / `Function` / `Property`,
parameter names and types, and return types:

```vba
Sub A(p1 As Long, p2 As String)
Function B(q1 As Double) As Boolean
Property X(v As Long) As Long
```

`ByVal` **is** encoded, in the byte after the VARTYPE
(``+15``): `0x00` = ByVal, `0x01` = ByRef. Verified 10/10 across mixed
signatures. Explicit `ByRef` cannot be told from the implicit default,
since both compile to `0x01`.

Not recoverable: `Optional` and default values.

**OPEN:** `Optional` parameters with defaults are not located by the
`+0x58` stride, and `Property Get` / `Let` / `Set` are not yet
distinguished from one another.

### 6.3 Declaration flags

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
(`Dim a(1 To 5)`) and a constant's value (`Const K = 7`).

The declared-type byte carries flags above the VARTYPE:

- `0x40` marks a **constant** -- `0x43` is Const + Long, `0x48` is
  Const + String.
- Arrays encode differently: both `As Long` and `As String` arrays
  leave `0x10` in the low bits, so the **element type is not in this
  byte** (it lives in the `type_` indirect table -- **OPEN**).

### 6.4 Types and enums

`Type` opens both user-defined types and enums; only the closing
opcode distinguishes them (`EndType` vs `EndEnum`). The name comes from
the opcode's `rec_` operand, resolved exactly like `func_` / `var_`.
Members are declared with `DimImplicit` + `VarDefn` and carry no
keyword of their own:

```vba
Type T          Enum E
    a As Long       E1 = 1
    b As String     E2 = 2
End Type        End Enum
```

**OPEN:** a variable declared *as* a UDT or enum (`Dim v As E`) has no
VARTYPE for that type, so the annotation is not recovered.

### 6.5 Statement opcodes

The statement forms the decompiler renders, with the opcodes that carry
them:

| VBA | opcodes |
|---|---|
| `Set x = e` | `SetStmt` (marker), then `Set` / `MemSet` |
| `a(i) = e` | `ArgsSt` (value pushed first, then indices) |
| `o.m = e` / `.m = e` | `MemSt` / `MemStWith` |
| `With o` ... `End With` | `StartWithExpr`, `With`, `EndWith` |
| `ReDim a(n)` | `Redim` (with a `type_` operand) |
| `On Error GoTo L` | `OnError` |
| `L:` | `Label` |
| `Resume Next` | `Resume` |
| `Exit Sub` / `Exit For` / ... | `ExitSub`, `ExitFor`, `ExitDo`, `ExitFunc` |
| `Erase a` | `Erase` |
| `Option Explicit` | `Option` |

## 7. Assembler status (p-code writing)

Verified milestones, all **byte-exact against Office-compiled output**:

1. **Instruction encoder** -- re-emit any decoded instruction to its
   on-disk bytes. *278/278 instructions byte-exact across 19/20 Office
   fixtures (Excel, Word, PowerPoint).*
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

Name resolution makes real decompilation possible. Given a module stream
plus its `_VBA_PROJECT`, every `name`, `func_`, and `var_` operand
resolves to a source identifier.

Round-trip on an Excel-compiled module:

```
; line   0: FuncDefn S
; line   1: Dim | VarDefn alpha
; line   3: LitDI2 0x1 | St alpha
; line   5: LitStr "x" | ArgsCall MsgBox 0x1
; line   6: ArgsCall Helper 0x0
; line   7: EndSub
```

reconstructing to:

```vba
Sub S()
    Dim alpha
    Dim beta
    alpha = 1
    beta = 2
    MsgBox "x"
    Helper
End Sub
```

With declared types resolved (section 6.1) the reconstruction is an
**exact line-by-line match** against the original source:

```vba
Sub S()
    Dim alpha As Long
    Dim beta As Long
    alpha = 1
    beta = 2
    MsgBox "x"
    Helper
End Sub
Sub Helper()
    Dim gamma As Long
    gamma = 3
End Sub
```

The only difference from the input is the optional `Call` keyword
(`Call Helper` vs `Helper`), which is syntactic sugar the compiler
discards -- both forms emit identical p-code, so it is unrecoverable by
construction rather than a gap in decoding.

### 8.1 Expressions and control flow

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
| `For` / `Next` | `StartForVariable`, `EndForVariable`, `For` / `ForStep`, `Next` / `NextVar` |
| `Do While` / `Loop` | `DoWhile`, `Loop` (also `DoUntil`, `LoopWhile`, `LoopUntil`) |
| `While` / `Wend` | `While`, `Wend` |
| `Select Case` | `SelectCase`, `CaseDone`, `CaseElse`, `EndSelect` |

Verified against Excel-compiled modules, the reconstruction is
**character-for-character identical** to the original source for
conditionals, both loop forms, and `Select Case`:

```vba
Sub S()
    Dim i As Long, t As Long
    For i = 1 To 10
        t = t + i
    Next i
    For i = 10 To 1 Step -2
        t = t - 1
    Next
End Sub
```

Round-trip fidelity, measured over a corpus compiled by Excel and
decompiled back: **19/20 exact**, 1 structural (an enum-typed variable
loses its `As E` annotation), 0 wrong. Normalisation covers only what
p-code genuinely does not encode -- the optional `Call` keyword, `$`
type suffixes, and the casing of names that resolve through the runtime
built-in table.

Covered: expressions with precedence and parentheses; `If` / `ElseIf` /
`Else`; `For` / `Next` with `Step`; `Do While` / `Loop`; `While` /
`Wend`; `Select Case`; procedure signatures including `ByVal`,
parameter types and return types; `Dim` / `Const` / `Public` /
`Private` / `Static`; arrays and `ReDim`; `Set`; member access; `With`
blocks; `On Error` / labels / `Resume`; `Type` and `Enum`.

Across every committed Office fixture, **54/54 modules decompile with
no unmapped opcodes**, including a 230-line real-world module whose
comment wall reproduces exactly.

Coverage note: array *element* types and UDT/enum-typed variables are
the remaining annotation gaps (both need the `type_` indirect table).

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

Excel automation runs through `pyvbaharness`, which is hang-safe
(bounded, popup-aware, hard process reaping) -- necessary because a
compile error otherwise blocks on a modal dialog.

---

## 10. Open questions

- Derivation of the `0x20E` name-operand base.
- Location of `DECL_BASE` (a module-header field?), currently
  calibrated.
- The identifier **hash function** used by the module's bucket table --
  required before new identifiers can be written.
- The full contents of the runtime built-in identifier table
  (section 5.1); the mechanism is understood, the map is partial.
- The `type_` indirect table: array **element** types and variables
  declared as a UDT or enum (sections 6.3, 6.4). Scalar types are solved
  (6.1), as are signatures (6.2) and declaration flags (6.3).
- `Optional` parameters with defaults are not located by the `+0x58`
  stride.
- `Property Get` / `Let` / `Set` are **not distinguished in p-code** --
  all three compile to `FuncDefn` + `EndProp`. Recovering which is which
  would need another source.
- Source-to-p-code compilation: lexer, parser, codegen, and the slot
  allocation Office's compiler performs.

---

## 11. Related

- [`docs/msaccess_lessons_learned.md`](msaccess_lessons_learned.md) --
  why Access VBA writing is hard and what the p-code layer gates.
- [`docs/access_pcode_re.md`](access_pcode_re.md) -- Access-specific
  p-code storage (`rU@` rows, `CAFE` rows, plaintext B9 / E3 stores).
- `src/pyopenvba/vba_pcode.py` -- the shipped disassembler.
- `docs/research/pcode/` -- assembler and decompiler prototypes backing
  this document.
