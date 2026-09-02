# pyOpenVBA

[![PyPI version](https://img.shields.io/pypi/v/pyOpenVBA.svg)](https://pypi.org/project/pyOpenVBA/)
[![Python versions](https://img.shields.io/pypi/pyversions/pyOpenVBA.svg)](https://pypi.org/project/pyOpenVBA/)
[![CI](https://github.com/WilliamSmithEdward/pyOpenVBA/actions/workflows/ci.yml/badge.svg)](https://github.com/WilliamSmithEdward/pyOpenVBA/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/LICENSE.md)
[![Downloads](https://static.pepy.tech/badge/pyOpenVBA/month)](https://pepy.tech/project/pyOpenVBA)

**Read and write VBA macros inside Office 365 files, in pure Python.**

No external dependencies. No Office install required. Works on Windows,
macOS, and Linux. Python 3.10 or newer.

Supports:

* Excel (`.xlsm`, `.xlsb`, `.xlam`, `.xls`)
* PowerPoint (`.pptm`, `.potm`, `.ppt`)
* Word (`.docm`, `.dotm`, `.doc`)
* Access (`.accdb`) - **read-only**

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
- **Design** UserForms, not just their code-behind: read a form's control
  tree and its properties, edit them, add and remove controls, Frames,
  MultiPages and pages, or compose a whole form from nothing.
- **Save** the file and have it reopen in the host application with no
  repair dialog. Every supported format is verified against live Office.
- **Create** new `.xlsm`, `.xlsb`, `.xlam`, `.docm`, or `.pptm` files on
  the fly, and inject VBA code into them.

That makes it a good fit for:

- **Version-controlling your VBA** in git like normal source code, then
  pushing edits back without ever opening Office.
- **Diffing** two workbooks or documents to see what changed in a module --
  or in a form's design, down to which properties an author actually set.
- **Building UserForms programmatically**, on a machine with no Office at
  all.
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

### Access (read-only)

```python
from pyopenvba import AccessReader

with AccessReader("database.accdb") as db:
    # 1. List all VBA modules in the database.
    modules = db.vba_modules()
    print(list(modules))
    # ['Module1', 'Form_Form1']

    # 2. Read a module's source as a string.
    source = db.get_module("Module1")
    print(source)
```

Excel, Word, and PowerPoint share the same read/write API:
`module_names()`, `get_module()`, `set_module()`, `save()`. Access VBA is
currently read-only and exposes `vba_modules()` and `get_module()`.

### Access tables, no Office required

`AccessDatabase` is a pure-Python implementation of the Jet 4 / ACE
storage engine: it reads and writes the database itself, not just its
VBA. Every edit reproduces what the engine writes, page for page.

```python
import datetime as dt
from pyopenvba import AccessDatabase, ColumnSpec, IndexSpec

with AccessDatabase.create_new("orders.accdb") as db:
    orders = db.create_table(
        "Orders",
        [
            ColumnSpec("Id", "Long", autonumber=True),
            ColumnSpec("Customer", "Text", size=80),
            ColumnSpec("Placed", "DateTime"),
            ColumnSpec("Total", "Currency"),
            ColumnSpec("Notes", "Memo"),
        ],
        [IndexSpec("PrimaryKey", ("Id",), primary=True), IndexSpec("ByCustomer", ("Customer",))],
    )
    orders.insert_row({"Customer": "Ada", "Placed": dt.datetime(2026, 9, 2, 9, 30), "Total": 19.99})
    row_id, row = next(orders.rows_with_ids())
    orders.update_row(row_id, {"Notes": "first order"})
    db.save()

with AccessDatabase("orders.accdb") as db:
    print(db.table_names())                       # ['Orders']
    for row in db.table("Orders").index("ByCustomer").rows():
        print(row["Customer"], row["Total"], row["Notes"])
```

Every column type is covered (Boolean through BigInt, Decimal, GUID,
Memo and OLE), indexes are maintained on every write, relationships
are created with `db.create_relationship(...)` and read with
`db.relationships()`, table and column properties (Description,
Caption, Format, ...) are read and set through `table.properties()` and
`table.set_properties(...)`, and files grow past their first 512 pages
the way the engine grows them. What it does not do yet is listed in
[docs/access_engine.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/access_engine.md),
along with every format rule and how it was measured.

---

## Create a brand-new file from scratch

Need a fresh macro-enabled file without launching Office? Use
`create_new()` on any of the three file classes. The extension in the
path controls the format:

```python
from pyopenvba import ExcelFile, WordFile, PowerPointFile

# Excel - macro-enabled workbook (.xlsm), binary workbook (.xlsb),
# or add-in (.xlam)
with ExcelFile.create_new("new_book.xlsm") as wb:
    wb.set_module("Module1", 'Sub Hello()\r\n    MsgBox "xlsm"\r\nEnd Sub\r\n')
    wb.save()

with ExcelFile.create_new("new_book.xlsb") as wb:
    wb.set_module("Module1", 'Sub Hello()\r\n    MsgBox "xlsb"\r\nEnd Sub\r\n')
    wb.save()

with ExcelFile.create_new("new_addin.xlam") as wb:
    wb.set_module("Module1", 'Sub Hello()\r\n    MsgBox "xlam"\r\nEnd Sub\r\n')
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

Class sources are accepted in any form: a bare body (the header is
synthesized), a `.cls` file exported straight from the VBE (the
`VERSION ... CLASS` preamble is stripped and the required
`Attribute VB_Base` line is added automatically), or a full
stream-form source.

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

## Creating and editing a UserForm's design

A form's *code* is a module like any other. Its *design* -- which controls
exist, how they nest, and what their properties are -- lives in separate
streams that no module source carries. `forms()` reads them, with no
Office installed:

```python
import pyopenvba

with pyopenvba.ExcelFile("book.xlsm") as wb:
    for form in wb.forms():
        print(form.name, len(form.walk()), "controls")
        for control in form.walk():
            print(f"  {control.name:<16} {control.kind:<22} "
                  f"{control.properties()}")
```

And edits them:

```python
with pyopenvba.ExcelFile("book.xlsm") as wb:
    form = wb.forms()[0]
    form.control("OkButton").set_property("Caption", "Save")
    form.control("NameBox").set_property("MaxLength", 40)
    form.add_control("Label", "Hint", left=12, top=120, width=200)
    form.remove_control("OldCheckbox")
    wb.save()
```

Containers work too. Each gets a storage of its own, and removing one
takes its children with it:

```python
form.add_control("Frame", "Shipping", left=12, top=160, width=200, height=80)
form.add_control("OptionButton", "Ground", container="Shipping")
form.remove_control("OldFrame")     # and everything inside it
```

A `MultiPage` arrives with the two pages Excel gives it, and pages are
added and removed through it, because a page is also a *tab*:

```python
form.add_control("MultiPage", "Wizard", left=12, top=40, width=300, height=200)
form.add_page("Wizard", name="Review", caption="Review && confirm")
form.add_control("Label", "Summary", container="Review")
form.remove_page("Page2", multipage="Wizard")
```

Page names are scoped to their MultiPage rather than to the form, which is
why `remove_page` takes an optional `multipage` to disambiguate.

And a form can be built from nothing -- `add_form` creates the designer
storage and the code-behind module together, because a storage without a
module is not a component the host will show:

```python
with pyopenvba.ExcelFile("book.xlsm") as wb:
    form = wb.add_form("Wizard", caption="Setup", width=300, height=200)
    form.add_control("Label", "Prompt", left=12, top=12, width=200)
    form.add_control("TextBox", "Answer", left=12, top=40, width=200)
    form.add_control("CommandButton", "Ok", left=12, top=80)
    wb.set_module("Wizard", "Private Sub Ok_Click()\r\n    Me.Hide\r\nEnd Sub\r\n")
    wb.save()
```

Geometry is in points, the unit the designer shows. `set_property(name,
None)` clears a property, which is how a control goes back to inheriting
the default.

Or from the command line:

```bash
python -m pyopenvba forms book.xlsm
```

```
FrmNested  (13 controls)
  TopLabel             MSForms.Label          id=1    set=0x00000028
  GroupBox             MSForms.Frame          id=2    set=0x0c0a0c48
    OptOne               MSForms.OptionButton   id=3    set=0x0000000180c00146
  Pages                MSForms.MultiPage      id=6    set=0x0c000c48
    (unnamed)            MSForms.TabStrip       id=7    set=0x00fa8031
    Page1                MSForms.Form           id=8    set=0x0c000c48
      PageOneCheck         MSForms.CheckBox       id=10   set=0x0000000080c00146
```

Containers nest: a `Frame`'s children and a `MultiPage`'s `Page`s live in
storages of their own, and MSForms sites an unnamed `TabStrip` beside the
pages. `walk()` flattens the tree depth-first; `form.controls` gives just
the top level.

### Only what the developer set

MSForms writes a property into a control's record **only when it differs
from that control's default**, so `properties()` returns the set the
developer chose -- not every property the control has. That is not
something a live host can tell you: a sited control reports inherited,
default, and chosen values indistinguishably.

`properties_set` is the same information as a raw bit mask, if you want to
diff two files without comparing names. Read a bit index together with
`property_mask_width`: MorphData controls (`TextBox`, `ListBox`,
`ComboBox`, `CheckBox`, `OptionButton`, `ToggleButton`) carry an 8-byte
mask and everything else carries 4.

Writing is lossless. An unedited form saves back byte for byte, because
alignment padding, string bytes, pictures, and anything the property
tables do not model are all replayed as they were read. If a form's
streams do not reconcile, this raises `FormParseError` rather than
returning a partly-guessed control list.

---

## Supported formats

### Excel

| Extension | What it is                   | Read | Write | create_new |
|-----------|------------------------------|:----:|:-----:|:----------:|
| `.xlsm`   | Macro-enabled workbook       |  yes |  yes  |    yes     |
| `.xlsb`   | Binary workbook              |  yes |  yes  |    yes     |
| `.xlam`   | Macro-enabled add-in         |  yes |  yes  |    yes     |
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

### Access

| Extension | What it is                   | Read | Write | create_new |
|-----------|------------------------------|:----:|:-----:|:----------:|
| `.accdb`  | Access database (ACE engine) | tables, indexes, VBA | tables, indexes, rows (`AccessDatabase`); VBA no | yes |
| `.mdb`    | Access database (Jet 4)      | tables, indexes, VBA | tables, indexes, rows | no |

The VBA side stays read-only. Access stores compiled VBA p-code (the `rU@` + `CAFE` rows in the LVAL
catalog) separately from the OVBA source cache. The compiled p-code is
authoritative for the Access GUI; mutations to the source cache do not
survive reload because Access never recompiles from the cache. After
extensive reverse-engineering experiments we concluded that a
production-quality writer would require a complete VBA7 p-code
assembler, which is out of scope. See
[docs/msaccess_lessons_learned.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/msaccess_lessons_learned.md)
for the full chronicle.

What `AccessReader` does support:

- `AccessReader(path)` / `vba_module_names()` / `read_vba_module(name)`
- `read_vba_module_with_attributes(name)`
- `vba_modules()` (dict of name -> source)
- `iter_vba_modules()` (rich `VBAModule` records)
- `export_module()` / `export_modules()` / `pull_modules()` (write `.bas` / `.cls` to disk)
- `read_project_info()`, `identifiers()`, `find_interned_strings()`,
  `find_module_streams()`, `iter_pcode_streams()`, `disassemble_module()`
- `iter_msys_objects()` / `msys_objects()` / `iter_msys_modules()` /
  `find_msys_module()` (MSysObjects catalog inspection)
- Top-level helper: `pyopenvba.pull_access(database, dest_dir)`

```python
from pyopenvba import AccessReader, pull_access

with AccessReader("database.accdb") as db:
    for name, source in db.vba_modules().items():
        print(name, len(source))

pull_access("database.accdb", "./vba_src")   # export every module to .bas / .cls
```

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

A project's code and its UserForm designs are read and written. The
following are preserved byte-for-byte but not interpreted:

- VBA project password decryption / re-encryption.
- Re-signing digitally signed projects.
- ActiveX license editing.

See [docs/roadmap.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/roadmap.md) for the full feature matrix.

---

## Architecture

```
src/pyopenvba/
  __init__.py        public API (ExcelFile, WordFile, PowerPointFile,
                                 AccessReader, VBAForm, FormControl, Size,
                                 pull/push, pull_word/push_word,
                                 pull_ppt/push_ppt, pull_access,
                                 VBAModuleKind, synthesize_class_header,
                                 exceptions)
  _host.py           VBAHostFile: shared open/edit/pull/push/save pipeline
  excel.py           ExcelFile (thin VBAHostFile subclass + create_new template)
  word.py            WordFile (thin VBAHostFile subclass + create_new template)
  powerpoint.py      PowerPointFile (thin subclass; .ppt overrides the two
                     container hooks)
  access_read.py     AccessReader (read-only ACE/Jet page + LVAL reader)
  access/            the Jet 4 / ACE storage engine, in progress: reads
                     every table, index and long value; inserts, updates
                     and deletes rows and creates and drops tables and
                     indexes the way the engine does (docs/access_engine.md)
  vba.py             VBA project parser + MS-OVBA codec
  vba_pcode.py       VBA7 p-code disassembler
  cfb.py             MS-CFB (Compound File Binary) parser/writer
  forms.py           UserForm designer streams: control tree, read and write
  _oforms_records.py [MS-OFORMS] property table, one per control class
  _oforms_pages.py   a MultiPage's tabs and page bookkeeping
  _ppt_container.py  the VBA project a binary .ppt hides in its document stream
  exceptions.py      custom exception hierarchy
  _templates/        baked-in empty .xlsm/.xlsb/.xlam/.docm/.pptm/.accdb bytes
                     for create_new()
  __main__.py        `python -m pyopenvba {pull,push,ls,forms,disasm,access-*}`
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

On a Windows machine with desktop Excel installed you can additionally run
the live compile-and-run gate (skipped by default and in CI). It builds a
workbook with pyOpenVBA, runs its macro in real Excel under a popup-aware
harness, and fails on any VBE dialog:

```powershell
$env:RUN_LIVE_EXCEL = "1"; pytest tests/test_live_excel_gate.py
```

CI runs the test matrix on Python 3.10 / 3.11 / 3.12 / 3.13 across
Linux, plus 3.12 on Windows and macOS, on every push and pull request.
Releases are published to PyPI automatically when a `v*.*.*` tag is
pushed.

---

## License

[MIT](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/LICENSE.md).

---

## Support Open Source

pyOpenVBA is open-source software. If it saves you time or helps your team keep
VBA workbooks maintainable, support helps keep the project moving.

- [GitHub Sponsors](https://github.com/sponsors/WilliamSmithEdward)
- [PayPal](https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F&currency_code=USD)
- [Cash App](https://cash.app/$williamesmithjcil)
