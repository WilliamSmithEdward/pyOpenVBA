# pyOpenVBA

[![PyPI version](https://img.shields.io/pypi/v/pyOpenVBA.svg)](https://pypi.org/project/pyOpenVBA/)
[![Python versions](https://img.shields.io/pypi/pyversions/pyOpenVBA.svg)](https://pypi.org/project/pyOpenVBA/)
[![CI](https://github.com/WilliamSmithEdward/pyOpenVBA/actions/workflows/ci.yml/badge.svg)](https://github.com/WilliamSmithEdward/pyOpenVBA/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE.md)
[![Downloads](https://img.shields.io/pypi/dm/pyOpenVBA.svg)](https://pypi.org/project/pyOpenVBA/)

**Read and write the VBA macros inside Excel workbooks, in pure Python.**

No external dependencies. No Excel install required. Works on Windows,
macOS, and Linux. Python 3.10 or newer.

  
Supports `.xlsm`, `.xlsb`, `.xlam`, and legacy `.xls`.

<a href="https://github.com/sponsors/WilliamSmithEdward"><img src="https://img.shields.io/badge/Sponsor-%E2%9D%A4-pink?style=for-the-badge" alt="Sponsor WilliamSmithEdward"></a>

---

## Why use this?

Several excellent Python tools already exist for **reading** VBA out of
Excel files (oletools, olefile, and friends), and they remain the right
choice for forensics, malware analysis, and audit use-cases. pyOpenVBA
focuses on the next step: safely **writing** changes back so the
workbook still opens cleanly in Excel.

The write path is the whole point of the library:

- **Replace** a module's source in place.
- **Add** a new standard module, class module, or document/UserForm
  code-behind.
- **Rename** any module (the CFB stream, `dir` record, `PROJECT`
  declaration, `PROJECTwm` name map, and `Attribute VB_Name` are all
  updated in lockstep).
- **Delete** a module cleanly.
- **Save** the workbook and have it reopen in Excel with no "we found
  a problem with some content" repair dialog. Every supported format
  (`.xlsm`, `.xlsb`, `.xlam`, `.xls`) is verified against live Excel.
- **Safely refuse** to corrupt password-protected or digitally signed
  projects unless you explicitly opt in.

That makes it a good fit for:

- **Version-controlling your VBA** in git like normal source code, then
  pushing edits back without ever opening Excel.
- **Diffing** two workbooks to see what changed in `Module1`.
- **Generating or updating macros from a script** without scripting
  Excel through COM automation.
- **Reading and writing macros on a server** (Linux / CI) where Excel
  is not installed.

pyOpenVBA is a complete read-and-write library, so it covers the full
lifecycle of a VBA project in one place: extract, edit, version, write
back, and verify. If you only ever need to read, the existing tools
remain a solid choice; if you might ever need to write, start here and
you will not have to switch later.

## Installation

From PyPI:

```bash
pip install pyOpenVBA
```

Requires Python 3.10 or newer. There are no other dependencies.

After install, the CLI is available either as a module or as a script:

```bash
python -m pyopenvba --help
pyopenvba --help
```

From source (for development):

```bash
git clone https://github.com/WilliamSmithEdward/pyOpenVBA
cd pyOpenVBA
pip install -e ".[dev]"
```

---

## 30-second tour

```python
from pyopenvba import ExcelFile

with ExcelFile("workbook.xlsm") as wb:
    # 1. List all VBA modules in the workbook.
    print(wb.module_names())
    # ['ThisWorkbook', 'Sheet1', 'Module1']

    # 2. Read a module's source as a string.
    source = wb.get_module("Module1")
    print(source)

    # 3. Edit a module and save the workbook.
    wb.set_module("Module1", 'Sub Hello()\r\n    MsgBox "hi"\r\nEnd Sub\r\n')
    wb.save()                       # overwrites the original file
    # wb.save("edited.xlsm")        # ...or save to a new file
```

That is the entire core API. Three methods.

---

## Add, rename, or delete a module

```python
from pyopenvba import ExcelFile, VBAModuleKind

with ExcelFile("workbook.xlsm") as wb:
    project = wb.vba_project()

    project.add_module(
        "NewModule",
        "Sub Hi()\r\n    MsgBox \"hi\"\r\nEnd Sub\r\n",
        kind=VBAModuleKind.standard,
    )
    project.rename_module("OldName", "NewName")
    project.delete_module("Obsolete")

    wb.save("out.xlsm")
```

---

## Edit your macros as files on disk (recommended workflow)

This is the easiest way to manage VBA in a git repo. Export every
module to a folder, edit the files in any text editor, then push the
changes back into the workbook.

From the command line:

```bash
# 1. Pull every module out of the workbook into ./vba/
python -m pyopenvba pull workbook.xlsm ./vba

# 2. ...edit ./vba/Module1.bas in your editor of choice...

# 3. Push your edits back into the workbook
python -m pyopenvba push ./vba workbook.xlsm

# Bonus: list modules without extracting anything
python -m pyopenvba ls workbook.xlsm
```

From Python:

```python
from pyopenvba import pull, push

pull("workbook.xlsm", "./vba")
# ...edit files...
push("./vba", "workbook.xlsm")                       # in place
push("./vba", "workbook.xlsm", out="edited.xlsm")    # to a new file
```

Module files use the extensions VBA already uses: `.bas` for standard
modules, `.cls` for class modules, `.frm` for UserForm code-behind.

---

## Supported formats

| Extension | What it is                   | Read | Write |
|-----------|------------------------------|:----:|:-----:|
| `.xlsm`   | Macro-enabled workbook       |  yes |  yes  |
| `.xlsb`   | Binary workbook              |  yes |  yes  |
| `.xlam`   | Macro-enabled add-in         |  yes |  yes  |
| `.xls`    | Legacy (Excel 97-2003)       |  yes |  yes  |

Every save is verified to reopen in Excel **without** the "we found a
problem with some content" repair dialog.

---

## Safety guards

`save()` refuses to silently produce a broken workbook.

### Password-protected projects

If the VBA project is password-protected, any mutation will raise
`VBAProjectError` unless you explicitly opt in:

```python
wb.save(allow_protected=True)
```

The library never tries to decrypt or change the password - it just
preserves the existing protection bytes verbatim. The resulting
workbook still requires the original password to open the VBE.

### Digitally-signed projects

A digital signature is invalidated by *any* change to the macros. On
mutation, the library drops the stale signature streams and emits a
`UserWarning` so you know trust has been removed:

```python
import warnings
warnings.filterwarnings("error", category=UserWarning)   # treat as fatal

# ...or silence the warning if you accept the consequence:
wb.save(allow_invalidate_signature=True)
```

---

## What's out of scope

This library is intentionally focused on **module source code**. The
following are preserved byte-for-byte but not interpreted:

- UserForm **layout** (controls, properties, positions). Editing the
  **code-behind** of a UserForm works fine; editing the design surface
  does not.
- VBA project password decryption / re-encryption.
- Re-signing digitally signed projects.
- ActiveX license editing.

See [docs/roadmap.md](docs/roadmap.md) for the full feature matrix.

---

## Architecture

```
src/pyopenvba/
  __init__.py     public API (ExcelFile, pull, push, exceptions)
  excel.py        ExcelFile facade (ZIP / CFB dispatch, pull/push helpers)
  vba.py          VBA project parser + MS-OVBA codec
  cfb.py          MS-CFB (Compound File Binary) parser/writer
  exceptions.py   custom exception hierarchy
  __main__.py     `python -m pyopenvba {pull,push,ls}` CLI
```

For deeper documentation:

- [docs/architecture.md](docs/architecture.md) - internal module layout.
- [docs/ms-ovba-implementation-guide_v2.md](docs/ms-ovba-implementation-guide_v2.md) -
  language-agnostic guide for re-implementing MS-OVBA in another
  language.
- [docs/roadmap.md](docs/roadmap.md) - per-feature implementation status.

---

## Contributing

Bug reports, weird workbooks that break the library, and PRs are all
welcome. Please include the workbook (or a minimal redacted version)
when filing a parsing bug.

Run the full local check (same as CI):

```bash
pip install -e ".[dev]"
pyright src tests
pytest -p no:randomly
```

CI runs the test matrix on Python 3.10 / 3.11 / 3.12 / 3.13 across
Linux, plus 3.12 on Windows and macOS, on every push and pull request.
Releases are published to PyPI automatically when a `v*.*.*` tag is
pushed.

---

## License

[MIT](LICENSE.md).
