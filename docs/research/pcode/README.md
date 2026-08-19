# P-code assembler / decompiler prototypes

Research prototypes backing [`docs/pcode_reference.md`](../../pcode_reference.md).

**Not part of the shipped library.** pyOpenVBA's public API is unchanged
and remains pure-Python with no Office dependency; nothing here is
imported by `src/pyopenvba`.

| File | Role |
|------|------|
| `pcode_asm.py` | Instruction encoder (inverse of the disassembler). Byte-exact on 278/278 instructions across Excel / Word / PowerPoint fixtures. |
| `pcode_names.py` | `_VBA_PROJECT` identifier-table parser and name-operand resolution (`0x20E + 2*index`). Host-agnostic. |
| `pcode_decompile.py` | Annotated disassembly and best-effort source reconstruction, including `func_` / `var_` declaration-name resolution. |
| `compile_oracle.py` | Dev-time oracle: pyOpenVBA writes source, Excel compiles and saves, pyOpenVBA reads the compiled p-code back. |
| `batch.py` | Compiles many source variants through one Excel session, for differential analysis. |

The two oracle modules drive Excel through `pyvbaharness` and are
**development tools only**.

## Running them

```bash
python -c "import sys; sys.path[:0]=['docs/research/pcode','src']; \
from pcode_decompile import annotate"
```

`pcode_asm` and `pcode_names` need only `src/` on the path; the oracle
modules additionally require Windows, desktop Excel, and `pyvbaharness`.
