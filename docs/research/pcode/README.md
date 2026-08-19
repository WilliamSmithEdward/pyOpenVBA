# P-code assembler / decompiler prototypes

Research prototypes backing [`docs/pcode_reference.md`](../../pcode_reference.md).

**Not part of the shipped library.** pyOpenVBA's public API is unchanged
and remains pure-Python with no Office dependency; nothing here is
imported by `src/pyopenvba`.

| File | Role |
|------|------|
| `pcode_asm.py` | Instruction encoder (inverse of the disassembler). Byte-exact on 278/278 instructions across Excel / Word / PowerPoint fixtures. |
| `pcode_names.py` | `_VBA_PROJECT` identifier-table parser and name-operand resolution (`0x20E + 2*index`), plus the runtime built-in operand map. Host-agnostic. |
| `pcode_types.py` | Declared-type decoding: the `type_` indirect descriptors (arrays, fixed-length strings, UDTs, Enums, type-library classes) and the module's type-reference table. |
| `pcode_decompile.py` | Declaration-table access: base calibration, name and type resolution, procedure signatures, annotated disassembly. |
| `pcode_source.py` | Full source reconstruction: stack-machine expression rebuilding plus control-flow and indentation rendering. |
| `pcode_hash.py` | The `_VBA_PROJECT` identifier hash = OLE `LHashValOfNameSysA` (`h=0x0DEADBEE`, `37*h+LOOKUP[b]`, `% 65599 & 0xFFFF`) with the real 384-byte `Lookup_16` table, plus a byte-exact compact-record encoder. Exact on 6,825 ASCII and 20 accented names. |
| `hash_probe.py` | Measures that hash against Excel and checks the model; `identifier_hashes.json` caches measured samples. |
| `roundtrip.py` | Acceptance gate: compile a 37-entry corpus with Excel, decompile it, diff against the original source. |
| `semantic_roundtrip.py` | Stronger gate for modules whose source is unknown: decompile, recompile with Excel, compare the opcode streams. 19/19 equivalent, including an 8,060-opcode real-world module. |
| `sweep.py` | Coverage check: decompile every module in every fixture on disk and report anything unmapped. |
| `compile_oracle.py` | Dev-time oracle: pyOpenVBA writes source, Excel compiles and saves, pyOpenVBA reads the compiled p-code back. |
| `batch.py` | Compiles many source variants through one Excel session, for differential analysis. |

## Running them

`pcode_asm`, `pcode_names`, `pcode_types`, `pcode_decompile`,
`pcode_source` and `sweep` need only `src/` on the path:

```bash
python docs/research/pcode/sweep.py demo tests
```

`roundtrip`, `compile_oracle` and `batch` drive Excel through
`pyvbaharness` and are **development tools only**: they additionally
require Windows and desktop Excel, and write their workbooks to
`PYOPENVBA_PCODE_SCRATCH` or a temp directory.

```bash
python docs/research/pcode/roundtrip.py /path/to/scratch
```
