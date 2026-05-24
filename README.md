# pyOpenVBA

[![PyPI version](https://img.shields.io/pypi/v/pyOpenVBA.svg)](https://pypi.org/project/pyOpenVBA/)
[![Python versions](https://img.shields.io/pypi/pyversions/pyOpenVBA.svg)](https://pypi.org/project/pyOpenVBA/)
[![CI](https://github.com/WilliamSmithEdward/pyOpenVBA/actions/workflows/ci.yml/badge.svg)](https://github.com/WilliamSmithEdward/pyOpenVBA/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/LICENSE.md)
[![Downloads](https://img.shields.io/pypi/dm/pyOpenVBA.svg)](https://pypi.org/project/pyOpenVBA/)

**Read and write VBA macros inside Excel, Word, and PowerPoint files, in pure Python.**

No external dependencies. No Office install required. Works on Windows,
macOS, and Linux. Python 3.10 or newer.

Supports:

* Excel (`.xlsm`, `.xlsb`, `.xlam`, `.xls`)
* PowerPoint (`.pptm`, `.potm`, `.ppt`)
* Word (`.docm`, `.dotm`, `.doc`)

<a href="https://github.com/sponsors/WilliamSmithEdward"><img src="https://img.shields.io/badge/Sponsor-%E2%9D%A4-pink?style=for-the-badge" alt="Sponsor WilliamSmithEdward"></a>

---

## Why use this?

Several excellent Python tools already exist for **reading** VBA out of
Office files (oletools, olefile, and friends), and they remain a strong
choice for forensics, malware analysis, and audit use-cases. pyOpenVBA
focuses on the next step: safely **writing** changes back so the file
still opens cleanly in the host application.

The write path is the whole point of the library:

- **Modify** a module's source in place.
- **Add** a new standard module, class module, or document/UserForm
  code-behind.
- **Rename** any module (the CFB stream, `dir` record, `PROJECT`
  declaration, `PROJECTwm` name map, and `Attribute VB_Name` are all
  updated in lockstep).
- **Delete** a module cleanly.
- **Save** the file and have it reopen in the host application with no
  repair dialog. Every supported format is verified against live Office.
- **Create** new `.xlsm`, `.xlsb`, `.docm`, or `.pptm` files on the
  fly, and inject VBA code into them.

That makes it a good fit for:

- **Version-controlling your VBA** in git like normal source code, then
  pushing edits back without ever opening Office.
- **Diffing** two workbooks or documents to see what changed in a module.
- **Generating or updating macros from a script** without scripting
  Office through COM automation.
- **Reading and writing macros on a server** (Linux / CI) where Office
  is not installed.
- **Agentic AI Integration** - allow your AI agent easy access to
  both push and pull VBA code in your Office files.

pyOpenVBA is a complete read-and-write library, so it covers the full
lifecycle of a VBA project in one place: extract, edit, version, write
back, and verify.

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

### Excel

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

### Word

```python
from pyopenvba import WordFile

with WordFile("document.docm") as doc:
    print(doc.module_names())
    # ['ThisDocument', 'Module1']

    doc.set_module("Module1", 'Sub Hello()\r\n    MsgBox "hi"\r\nEnd Sub\r\n')
    doc.save()
```

### PowerPoint

```python
from pyopenvba import PowerPointFile

with PowerPointFile("presentation.pptm") as prs:
    print(prs.module_names())
    # ['Module1']

    prs.set_module("Module1", 'Sub Hello()\r\n    MsgBox "hi"\r\nEnd Sub\r\n')
    prs.save()
```

The API is identical across all three hosts: `module_names()`, `get_module()`,
`set_module()`, `save()`.

---

## Create a brand-new file from scratch

Need a fresh macro-enabled file without launching Office? Use
`create_new()` on any of the three file classes. The extension in the
path controls the format:

```python
from pyopenvba import ExcelFile, WordFile, PowerPointFile

# Excel - macro-enabled workbook (.xlsm) or binary workbook (.xlsb)
with ExcelFile.create_new("new_book.xlsm") as wb:
    wb.set_module("Module1", 'Sub Hello()\r\n    MsgBox "xlsm"\r\nEnd Sub\r\n')
    wb.save()

with ExcelFile.create_new("new_book.xlsb") as wb:
    wb.set_module("Module1", 'Sub Hello()\r\n    MsgBox "xlsb"\r\nEnd Sub\r\n')
    wb.save()

# Word - macro-enabled document (.docm)
with WordFile.create_new("new_doc.docm") as doc:
    doc.set_module("Module1", 'Sub Hello()\r\n    MsgBox "docm"\r\nEnd Sub\r\n')
    doc.save()

# PowerPoint - macro-enabled presentation (.pptm)
with PowerPointFile.create_new("new_prs.pptm") as prs:
    prs.set_module("Module1", 'Sub Hello()\r\n    MsgBox "pptm"\r\nEnd Sub\r\n')
    prs.save()
```

Each new file is built from a baked-in template captured from a
freshly Office-authored file, so it opens cleanly with no repair prompt.

---

## Add, rename, or delete a module

The same `vba_project()` API works for all three hosts:

```python
from pyopenvba import ExcelFile, VBAModuleKind

with ExcelFile("workbook.xlsm") as wb:
    project = wb.vba_project()

    # Add a standard module
    project.add_module(
        "NewModule",
        'Sub Hi()\r\n    MsgBox "hi"\r\nEnd Sub\r\n',
        kind=VBAModuleKind.standard,
    )

    # Add a class module (header is synthesized automatically)
    project.add_module(
        "MyClass",
        "Option Explicit\r\n",
        kind=VBAModuleKind.other,
    )

    project.rename_module("OldName", "NewName")
    project.delete_module("Obsolete")

    wb.save("out.xlsm")
```

---

## Edit your macros as files on disk (recommended workflow)

This is the easiest way to manage VBA in a git repo. Export every
module to a folder, edit the files in any text editor, then push the
changes back.

### Excel

From the command line:

```bash
# Pull every module out of the workbook into ./vba/
python -m pyopenvba pull workbook.xlsm ./vba

# ...edit ./vba/Module1.bas in your editor of choice...

# Push your edits back into the workbook
python -m pyopenvba push ./vba workbook.xlsm

# List modules without extracting
python -m pyopenvba ls workbook.xlsm
```

From Python:

```python
from pyopenvba import pull, push

pull("workbook.xlsm", "./vba")
push("./vba", "workbook.xlsm")                    # in place
push("./vba", "workbook.xlsm", out="edited.xlsm") # to a new file
```

### Word

```python
from pyopenvba import pull_word, push_word

pull_word("document.docm", "./vba")
push_word("./vba", "document.docm")
push_word("./vba", "document.docm", out="edited.docm")
```

### PowerPoint

```python
from pyopenvba import pull_ppt, push_ppt

pull_ppt("presentation.pptm", "./vba")
push_ppt("./vba", "presentation.pptm")
push_ppt("./vba", "presentation.pptm", out="edited.pptm")
```

Module files use the extensions VBA already uses: `.bas` for standard
modules, `.cls` for class modules and code-behind.

---

## Supported formats

### Excel

| Extension | What it is                   | Read | Write | create_new |
|-----------|------------------------------|:----:|:-----:|:----------:|
| `.xlsm`   | Macro-enabled workbook       |  yes |  yes  |    yes     |
| `.xlsb`   | Binary workbook              |  yes |  yes  |    yes     |
| `.xlam`   | Macro-enabled add-in         |  yes |  yes  |    no      |
| `.xls`    | Legacy (Excel 97-2003)       |  yes |  yes  |    no      |

### Word

| Extension | What it is                   | Read | Write | create_new |
|-----------|------------------------------|:----:|:-----:|:----------:|
| `.docm`   | Macro-enabled document       |  yes |  yes  |    yes     |
| `.dotm`   | Macro-enabled template       |  yes |  yes  |    no      |
| `.doc`    | Legacy (Word 97-2003)        |  yes |  yes  |    no      |

### PowerPoint

| Extension | What it is                   | Read | Write | create_new |
|-----------|------------------------------|:----:|:-----:|:----------:|
| `.pptm`   | Macro-enabled presentation   |  yes |  yes  |    yes     |
| `.potm`   | Macro-enabled template       |  yes |  yes  |    no      |
| `.ppt`    | Legacy (PowerPoint 97-2003)  |  yes |  yes  |    no      |

Every save is verified to reopen in the host application **without** the
"we found a problem with some content" repair dialog.

---

## Safety guards

`save()` refuses to silently produce a broken file.

### Password-protected projects

If the VBA project is password-protected, any mutation will raise
`VBAProjectError` unless you explicitly opt in:

```python
wb.save(allow_protected=True)
```

The library never tries to decrypt or change the password - it just
preserves the existing protection bytes verbatim. The resulting file
still requires the original password to open the VBE.

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

See [docs/roadmap.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/roadmap.md) for the full feature matrix.

---

## Architecture

```
src/pyopenvba/
  __init__.py     public API (ExcelFile, WordFile, PowerPointFile,
                              pull/push, pull_word/push_word, pull_ppt/push_ppt,
                              VBAModuleKind, synthesize_class_header, exceptions)
  excel.py        ExcelFile facade (ZIP / CFB dispatch, pull/push helpers)
  word.py         WordFile facade
  powerpoint.py   PowerPointFile facade
  vba.py          VBA project parser + MS-OVBA codec
  cfb.py          MS-CFB (Compound File Binary) parser/writer
  exceptions.py   custom exception hierarchy
  _templates/     baked-in empty .xlsm/.xlsb/.docm/.pptm bytes for create_new()
  __main__.py     `python -m pyopenvba {pull,push,ls}` CLI
```

For deeper documentation:

- [docs/architecture.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/architecture.md) - internal module layout.
- [docs/ms-ovba-implementation-guide_v2.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/ms-ovba-implementation-guide_v2.md) -
  language-agnostic guide for re-implementing MS-OVBA in another language.
- [docs/roadmap.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/roadmap.md) - per-feature implementation status.

---

## Contributing

Bug reports, weird files that break the library, and PRs are all
welcome. Please include the file (or a minimal redacted version) when
filing a parsing bug.

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

[MIT](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/LICENSE.md).
