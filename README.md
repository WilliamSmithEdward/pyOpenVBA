# pyOpenVBA

Pure-Python library for reading and writing VBA source code embedded in
Excel workbooks (`.xlsm`, `.xlsb`, `.xls`, `.xlam`) — no external
dependencies, Python 3.10+.

> **Scope:** "MS-OVBA round-tripper with module source reader/writer
> (.xlsm focus)."  Module source can be added, replaced, renamed, and
> deleted end-to-end through `ExcelFile.save()`.  Everything outside the
> module-source surface (UserForm layout, ActiveX licenses, project
> password, digital signatures) is preserved verbatim through the CFB
> round-trip but not interpreted.  See [docs/roadmap.md](docs/roadmap.md)
> for the per-gate status.

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

    # Edit a module and save in place
    wb.set_module("Module1", "Sub Hello()\r\n    MsgBox \"hi\"\r\nEnd Sub\r\n")
    wb.save()
```

### Add, rename, delete modules

```python
from pyopenvba import ExcelFile, VBAModuleKind

with ExcelFile("workbook.xlsm") as wb:
    proj = wb.vba_project()
    proj.add_module("NewModule", "' new module\r\n", kind=VBAModuleKind.standard)
    proj.rename_module("OldModule", "Renamed")
    proj.delete_module("Obsolete")
    wb.save("out.xlsm")
```

## Disk-based push / pull workflow

Export every module to a `.bas` / `.cls` / `.frm` file for use with any
text editor or version control, then push edits back:

```bash
python -m pyopenvba pull workbook.xlsm ./vba/
# ...edit ./vba/Module1.bas in your editor...
python -m pyopenvba push workbook.xlsm ./vba/
python -m pyopenvba ls   workbook.xlsm
```

## Supported formats

| Extension | Description                  | Read | Write |
|-----------|------------------------------|:----:|:-----:|
| `.xlsm`   | OOXML macro-enabled workbook |  yes |  yes  |
| `.xlsb`   | Binary workbook              |  yes |  yes  |
| `.xlam`   | OOXML macro-enabled add-in   |  yes |  yes  |
| `.xls`    | Legacy BIFF8 workbook        |  yes |  yes  |

## Safety guards

`ExcelFile.save()` refuses to silently produce a broken workbook:

- A password-protected project raises `VBAProjectError` on mutation
  unless `save(allow_protected=True)` is passed.  (The password material
  is preserved verbatim; the resulting workbook may be inconsistent.)
- A digitally-signed project has its stale signature streams dropped on
  mutation and emits a `UserWarning`.  Pass
  `save(allow_invalidate_signature=True)` to silence.

## Architecture

```
src/pyopenvba/
  __init__.py    public API surface (ExcelFile, pull, push, exceptions)
  exceptions.py  custom exception hierarchy
  cfb.py         Compound File Binary (MS-CFB) parser/writer
  vba.py         VBA project / dir-stream parser + MS-OVBA codec
  excel.py       ExcelFile facade (ZIP / CFB dispatch, pull/push helpers)
  __main__.py    `python -m pyopenvba {pull,push,ls}` CLI
```

## License

See LICENSE.
