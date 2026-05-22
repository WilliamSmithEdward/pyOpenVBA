# pyOpenVBA

Pure-Python library for reading and writing VBA source code embedded in
Excel workbooks (`.xlsm`, `.xlsb`, `.xls`) — no external dependencies.

## Installation

```bash
pip install pyOpenVBA           # once published to PyPI
# or, for development:
pip install -e ".[dev]"
```

## Quick start

```python
from pyopenvba import ExcelFile

with ExcelFile("workbook.xlsm") as wb:
    # List all VBA module names
    print(wb.module_names())

    # Read a module's source code
    src = wb.get_module("Module1")
    print(src)

    # Inspect all modules at once
    for name, source in wb.vba_modules().items():
        print(f"--- {name} ---")
        print(source)
```

## Supported formats

| Extension | Description                          | Read | Write |
|-----------|--------------------------------------|:----:|:-----:|
| `.xlsm`   | OOXML macro-enabled workbook          |  yes |  soon |
| `.xlsb`   | Binary workbook                       |  yes |  soon |
| `.xlam`   | OOXML macro-enabled add-in            |  yes |  soon |
| `.xls`    | Legacy BIFF8 workbook                 |  yes |  soon |

## Architecture

```
src/pyopenvba/
  __init__.py   public API surface
  exceptions.py custom exception hierarchy
  cfb.py        Compound File Binary (MS-CFB) parser
  vba.py        VBA project / dir-stream parser + MS-OVBA decompressor
  excel.py      ExcelFile facade (ZIP / CFB dispatch)
```

## Roadmap

- [ ] Write-back: re-compress and patch the CFB (issue #1)
- [ ] Enumerate forms, class modules, and document modules separately
- [ ] CLI: `pyopenvba extract book.xlsm --out src/`
- [ ] CLI: `pyopenvba inject  book.xlsm --src src/`

## Development

```bash
pip install -e ".[dev]"
pytest
pyright src/ tests/
```

## License

MIT
