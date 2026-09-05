# pyOpenVBA

[![PyPI version](https://img.shields.io/pypi/v/pyOpenVBA.svg)](https://pypi.org/project/pyOpenVBA/)
[![Python versions](https://img.shields.io/pypi/pyversions/pyOpenVBA.svg)](https://pypi.org/project/pyOpenVBA/)
[![CI](https://github.com/WilliamSmithEdward/pyOpenVBA/actions/workflows/ci.yml/badge.svg)](https://github.com/WilliamSmithEdward/pyOpenVBA/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/LICENSE.md)
[![Downloads](https://static.pepy.tech/badge/pyOpenVBA/month)](https://pepy.tech/project/pyOpenVBA)

**Read and write the code inside Office files, in pure Python: VBA macros
in four hosts, and Power Query in Excel.**

No dependencies beyond the standard library. No Office install needed.
Works on Windows, macOS and Linux. Python 3.10 or newer.

VBA, four hosts and one API:

* Excel (`.xlsm`, `.xlsb`, `.xlam`, `.xls`)
* Word (`.docm`, `.dotm`, `.doc`)
* PowerPoint (`.pptm`, `.potm`, `.ppt`)
* Access (`.accdb`, `.mdb`)

Power Query, in any Excel package (`.xlsx` included):

```python
from pyopenvba import PowerQueryWorkbook

with PowerQueryWorkbook("orders.xlsx") as book:
    print(book.query_names())
    book.query("Orders").formula = "let Source = Excel.CurrentWorkbook() in Source"
    book.save()
```

---

## Why use this?

Good Python tools exist for reading VBA out of Office files (oletools,
olefile and friends), and they remain the right choice for forensics,
malware analysis and audits. pyOpenVBA covers the next step: writing
changes back so the file still opens cleanly in the host application.

The write path is the point of the library:

- **Modify** a module's source in place.
- **Add** a standard module, a class module, or code behind a form.
- **Rename** a module everywhere its name lives, in one step.
- **Delete** a module cleanly.
- **Design** forms as well as their code: read a form's controls and
  properties, edit them, add and remove controls, or build a form from
  nothing.
- **Query** as well as code: read, edit, add, group and load Power
  Queries, or write a whole workbook of them from nothing.
- **Create** a new `.xlsx`, `.xlsm`, `.xlsb`, `.xlam`, `.docm`, `.pptm`
  or `.accdb` file and put code in it.
- **Save**, and have the file reopen in the host with no repair dialog.

Every format is verified against live Office: the saved file reopens
without a repair prompt, and the code in it runs. Two parts are held to a
stricter bar. Access edits are compared byte for byte with the same edit
made by Access or its database engine, and the Power Query writer is
compared with Microsoft's own packaging assemblies.

That makes it a good fit for:

- Version-controlling VBA and M in git like any other source, then
  pushing edits back without opening Office.
- Diffing two files to see what changed in a module, a form's design or a
  query.
- Building forms, macros and queries from a script on a machine without
  Office.
- Reading and writing Office code on a server or in CI.
- Letting an AI agent read and change the code in your Office files.

---

## Installation

```bash
pip install pyOpenVBA
```

Requires Python 3.10 or newer. There are no other dependencies.

After installing, the CLI is available as a module or as a script:

```bash
python -m pyopenvba --help
pyopenvba --help
```

From source, for development:

```bash
git clone https://github.com/WilliamSmithEdward/pyOpenVBA
cd pyOpenVBA
pip install -e ".[dev]"
```

---

## 30-second tour

The four host classes share the same module API: `module_names()`,
`get_module()`, `set_module()`, `save()`.

### Excel

```python
from pyopenvba import ExcelFile

with ExcelFile("workbook.xlsm") as wb:
    print(wb.module_names())        # ['ThisWorkbook', 'Sheet1', 'Module1']
    source = wb.get_module("Module1")
    wb.set_module("Module1", 'Sub Hello()\r\n    MsgBox "hi"\r\nEnd Sub\r\n')
    wb.save()                       # in place
    # wb.save("edited.xlsm")        # or to a new file
```

### Word and PowerPoint

The same three calls, on the class for the host:

```python
from pyopenvba import PowerPointFile, WordFile

with WordFile("document.docm") as doc:
    doc.set_module("Module1", 'Sub Hello()\r\n    MsgBox "hi"\r\nEnd Sub\r\n')
    doc.save()

with PowerPointFile("presentation.pptm") as prs:
    prs.set_module("Module1", 'Sub Hello()\r\n    MsgBox "hi"\r\nEnd Sub\r\n')
    prs.save()
```

### Access

```python
from pyopenvba import AccessDatabase

with AccessDatabase("database.accdb") as db:
    print(db.module_names())        # ['Module1', 'Form_Orders']
    source = db.get_module("Module1")
    db.set_module("Module1", "Option Compare Database\r\n\r\nPublic Sub Hello()\r\n    MsgBox \"hi\"\r\nEnd Sub")
    db.save()
```

---

## Create a new file

`create_new()` builds a fresh macro-enabled file from a template Office
authored itself, so it opens with no repair prompt. The extension picks
the format:

```python
from pyopenvba import AccessDatabase, ExcelFile, PowerPointFile, WordFile

with ExcelFile.create_new("new_book.xlsm") as wb:        # also .xlsb, .xlam
    wb.set_module("Module1", 'Sub Hello()\r\n    MsgBox "xlsm"\r\nEnd Sub\r\n')
    wb.save()

with WordFile.create_new("new_doc.docm") as doc:
    doc.set_module("Module1", 'Sub Hello()\r\n    MsgBox "docm"\r\nEnd Sub\r\n')
    doc.save()

with PowerPointFile.create_new("new_prs.pptm") as prs:
    prs.set_module("Module1", 'Sub Hello()\r\n    MsgBox "pptm"\r\nEnd Sub\r\n')
    prs.save()

with AccessDatabase.create_new("new_db.accdb") as db:
    db.set_module("Module1", "Option Compare Database\r\n\r\nPublic Sub Hello()\r\nEnd Sub")
    db.save()
```

A workbook with no macros in it comes from `PowerQueryWorkbook`, which
adds the Power Query package the first time a query is written:

```python
from pyopenvba import PowerQueryWorkbook

with PowerQueryWorkbook.create_new("new_book.xlsx") as book:
    book.add_query("Numbers", "let Source = {1..10} in Source")
    book.save()
```

---

## Add, rename or delete a module

`vba_project()` gives the project, and the same three calls work on every
host:

```python
from pyopenvba import ExcelFile, VBAModuleKind

with ExcelFile("workbook.xlsm") as wb:
    project = wb.vba_project()
    project.add_module("NewModule", 'Sub Hi()\r\n    MsgBox "hi"\r\nEnd Sub\r\n')
    project.add_module("MyClass", "Option Explicit\r\n", kind=VBAModuleKind.other)
    project.rename_module("OldName", "NewName")
    project.delete_module("Obsolete")
    wb.save("out.xlsm")
```

```python
from pyopenvba import AccessDatabase, VBAModuleKind

with AccessDatabase("database.accdb") as db:
    project = db.vba_project()
    project.add_module("Helpers", "Option Compare Database\r\n\r\nPublic Function Twice(n As Long) As Long\r\n    Twice = n * 2\r\nEnd Function")
    project.add_module("Widget", "Option Compare Database", kind=VBAModuleKind.other)
    project.rename_module("Helpers", "Tools")
    project.delete_module("Widget")
    db.save()
```

A class source is accepted in any form: a bare body (the header is
synthesized), a `.cls` file exported from the VBE (the `VERSION ... CLASS`
preamble is stripped and the `Attribute VB_Base` line restored), or a
full stream-form source. `db.references()`, `db.add_reference(...)` and
`db.drop_reference(...)` manage the libraries an Access project points
at.

---

## Edit your macros as files on disk

The easiest way to keep VBA in a git repo: export every module to a
folder, edit the files in any editor, push the changes back.

```bash
python -m pyopenvba pull workbook.xlsm ./vba     # every module to ./vba/*.bas and *.cls
python -m pyopenvba push ./vba workbook.xlsm     # edits back into the workbook
python -m pyopenvba ls workbook.xlsm             # list modules without extracting

python -m pyopenvba access-pull database.accdb ./vba
python -m pyopenvba access-push ./vba database.accdb
python -m pyopenvba access-ls database.accdb
```

The same from Python, one pair per host:

```python
from pyopenvba import pull, push, pull_word, push_word, pull_ppt, push_ppt, pull_access, push_access

pull("workbook.xlsm", "./vba")
push("./vba", "workbook.xlsm", out="edited.xlsm")   # omit out= to save in place

pull_word("document.docm", "./vba")
push_word("./vba", "document.docm")

pull_ppt("presentation.pptm", "./vba")
push_ppt("./vba", "presentation.pptm")

pull_access("database.accdb", "./vba")
push_access("./vba", "database.accdb")
```

Module files use the extensions VBA already uses: `.bas` for standard
modules, `.cls` for class modules and code-behind. `push` replaces the
source of every module that has a file of its name; a file that matches
no module is skipped, or refused with `strict=True`.

---

## Forms

A form's code is a module like any other. Its design, which controls
exist, how they nest and what their properties are, lives beside it and
is read and written with the same calls on every host.

### UserForms in Excel, Word and PowerPoint

```python
import pyopenvba

with pyopenvba.ExcelFile("book.xlsm") as wb:
    for form in wb.forms():
        print(form.name, len(form.walk()), "controls")
        for control in form.walk():
            print(f"  {control.name:<16} {control.kind:<22} {control.properties()}")

    form = wb.forms()[0]
    form.control("OkButton").set_property("Caption", "Save")
    form.control("NameBox").set_property("MaxLength", 40)
    form.add_control("Label", "Hint", left=12, top=120, width=200)
    form.remove_control("OldCheckbox")
    wb.save()
```

Containers work too. A `Frame` gets a storage of its own and removing it
takes its children; a `MultiPage` arrives with the two pages Excel gives
it, and pages are added and removed through it:

```python
form.add_control("Frame", "Shipping", left=12, top=160, width=200, height=80)
form.add_control("OptionButton", "Ground", container="Shipping")
form.add_control("MultiPage", "Wizard", left=12, top=40, width=300, height=200)
form.add_page("Wizard", name="Review", caption="Review && confirm")
form.remove_page("Page2", multipage="Wizard")
```

A form can be built from nothing. `add_form` creates the designer storage
and the code-behind module together:

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
None)` clears a property, so the control goes back to its default.
MSForms stores a property only when it differs from the control's
default, so `properties()` returns what the developer set, which no live
host can tell you. Writing is lossless: an unedited form saves back byte
for byte.

The command line shows the tree:

```bash
python -m pyopenvba forms book.xlsm
```

### Forms and reports in Access

Access forms and reports read and edit through the same surface. Sizes
are in twips, the unit Access keeps, and a report takes `kind="report"`:

```python
from pyopenvba import AccessDatabase

with AccessDatabase("app.accdb") as db:
    for form in db.forms():
        print(form.name, [s.name for s in form.sections])
        for control in form.walk():
            print("  ", control.name, control.kind, control.properties().get("Caption"))

    form = db.add_form("Summary", caption="Totals", width=8000, height=3000)
    form.add_control("Label", "Title", left=240, top=240, width=2000, height=300, caption="Hello")
    form.add_control("TextBox", "Total", top=700, caption="=1+1")
    form.control("Title").set_property("FontSize", 14)
    form.remove_control("Total")
    form.set_code("Option Compare Database\r\n\r\nPrivate Sub Form_Load()\r\n    Me.Caption = \"Loaded\"\r\nEnd Sub")

    report = db.add_report("Monthly")
    report.add_control("Label", "Banner", section="PageHeaderSection", caption="Header band")
    db.delete_form("Old")
    db.save()
```

[examples/access_form_demo.py](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/examples/access_form_demo.py)
builds a working order calculator this way: a form whose buttons call a
standard module and keep their running total in a class module, laid out
and coloured through `set_property` (fonts, fills, borders, hover
colours, a currency format) and opened with the database through
`db.set_database_properties({"StartUpForm": "Calculator"})`, with the
database it produces beside it. Twenty-three control types can be
written, including a tab control and its pages
(`form.add_control("Page", "First", parent="Tabs")`); a navigation
control is read but not written. Each control gets only the
properties Access's own designs give its type, and `set_property` refuses
a name the type does not have. A live gate opens every written design in
Access's designer and reads back each control, measurement, caption and
tab index.

---

## Power Query

Excel keeps its Get and Transform queries in one custom XML part of the
workbook: an M section document, a metadata document beside it, and a
permission list, all inside a base64 blob. `PowerQueryWorkbook` reads
that, edits it and writes it back, for any Excel package including plain
`.xlsx`.

```python
from pyopenvba import PowerQueryWorkbook

with PowerQueryWorkbook("orders.xlsx") as book:
    for query in book.queries():
        print(query.name, query.load_target, query.steps)

    book.query("Orders").formula = 'let\r\n    Source = Csv.Document(File.Contents("o.csv"))\r\nin\r\n    Source'
    book.query("Orders").description = "Every order, straight from the export."
    book.add_query("Totals", "let Source = Table.RowCount(Orders) in Source")
    book.rename_query("Orders", "Order Lines")     # rewrites the queries that name it
    book.remove_query("Scratch")
    book.save()
```

Groups are the folders in the Queries pane, and a query can be loaded
onto a sheet or taken back off it:

```python
with PowerQueryWorkbook.create_new("report.xlsx") as book:
    staging = book.add_group("Staging")
    book.add_query("Raw", "let Source = 1 in Source", group=staging)
    book.add_query("Report", "let Source = Raw in Source")
    book.load_to_sheet("Report", ["Value"], cell="A1")
    book.save()
```

`load_to_sheet` writes the connection, the query table, the table and the
sheet's reference to it, because the metadata alone does not make Excel
load anything. The column names are yours to give: knowing them means
running the query, and Excel settles them against the real result on its
first refresh. `unload()` takes every piece back out.

A loaded query also carries the settings behind Excel's Connection
Properties dialog:

```python
settings = book.query("Report").refresh
settings.on_open = True             # refresh when the workbook opens
settings.interval_minutes = 60      # and every hour after that
settings.background = False         # in the foreground, so a macro can wait
settings.keep_data = False          # save the query, not its rows
settings.in_refresh_all = False     # leave it out of Refresh All
settings.enabled = True             # or False to stop it refreshing at all
```

"Enable Fast Data Load" is missing from that list on purpose. Excel's
object model does not expose it, so where Excel writes it could not be
measured, and nothing here is written on a guess.

Queries go to disk and back like modules do:

```bash
python -m pyopenvba pq-ls   workbook.xlsx        # name, where it loads, its group
python -m pyopenvba pq-pull workbook.xlsx ./queries
python -m pyopenvba pq-push ./queries workbook.xlsx
```

Each query becomes one `.m` file, beside a `queries.json` manifest that
carries what a file name cannot: the real name, the description and the
group. A query called `Sales/EU` survives the round trip.

Three examples ship with the library, and every query in each was
refreshed in Excel before it was committed:

- [power_query_demo.py](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/examples/power_query_demo.py)
  builds twelve queries in four groups, six of them loaded onto a sheet,
  reading JSON from four public APIs (PokeAPI, USGS, Frankfurter, Hacker
  News) in five different shapes.
- [power_query_steps.py](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/examples/power_query_steps.py)
  builds a sales pipeline whose main query has eleven applied steps. It
  keeps the steps in a list and writes `#PREV` where a step means the one
  before it, so inserting or removing a step is a list operation and the
  references rewire themselves.
- [power_query_refresh.py](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/examples/power_query_refresh.py)
  gives three loaded queries three refresh profiles: one that refreshes on
  open and keeps its rows, one on an hourly timer in the foreground, and
  one that keeps no rows in the file at all.

What the format is, and how each rule was measured, is in
[docs/power_query.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/power_query.md).

---

## Supported formats

### Excel

| Extension | What it is                   | VBA | Power Query | create_new |
|-----------|------------------------------|:---:|:-----------:|:----------:|
| `.xlsx`   | Workbook                     |  -  |     yes     |    yes     |
| `.xlsm`   | Macro-enabled workbook       | yes |     yes     |    yes     |
| `.xlsb`   | Binary workbook              | yes |     yes     |    yes     |
| `.xlam`   | Macro-enabled add-in         | yes |     yes     |    yes     |
| `.xls`    | Legacy (Excel 97-2003)       | yes |      -      |    no      |

A `.xlsx` file has no VBA project by design, and Power Query lives
outside the project, so it is read and written in every package above.

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
| `.accdb`  | Access database (ACE)        |  yes |  yes  |    yes     |
| `.mdb`    | Access database (Jet 4)      |  yes |  yes  |    no      |

An Access file keeps its VBA project inside the database itself, in the
system tables Access uses for its own objects, so writing a module means
writing rows, long values and index entries the way the database engine
does. `AccessDatabase` does that with a pure-Python implementation of the
Jet 4 / ACE storage engine, documented rule by rule with how each was
measured in
[docs/access_engine.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/access_engine.md).
Two things follow. A written module has no compiled copy until Access
recompiles the project on its next open, which Access does on its own.
And `AccessReader`, the older read-only class, is still there for
inspecting a database: `vba_modules()`, `read_project_info()`,
`identifiers()`, `disassemble_module()` and the `MSysObjects` catalog.

Every save is verified to reopen in the host application without the
"we found a problem with some content" repair dialog.

---

## Safety guards

`save()` refuses to silently produce a broken file.

### Password-protected projects

A mutation to a password-protected project raises `VBAProjectError`
unless you opt in:

```python
wb.save(allow_protected=True)
```

`AccessDatabase` does the same: `db.vba_is_protected()` says whether the
project carries a password, and `db.save()` refuses a VBA change to a
protected project without `allow_protected=True`. The library never
decrypts or changes the password; the protection bytes are preserved and
the file still asks for the original password in the VBE.

### Digitally signed projects

Any change to the macros invalidates a digital signature. On mutation the
library drops the stale signature streams and emits a `UserWarning`:

```python
import warnings
warnings.filterwarnings("error", category=UserWarning)   # treat as fatal

wb.save(allow_invalidate_signature=True)                 # or accept it
```

---

## Out of scope

Preserved byte for byte but not interpreted:

- VBA project password decryption or re-encryption.
- Re-signing digitally signed projects.
- ActiveX license editing.

[docs/roadmap.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/roadmap.md)
has the feature matrix.

---

## Architecture

```
src/pyopenvba/
  __init__.py        public API: ExcelFile, WordFile, PowerPointFile,
                     AccessDatabase, AccessReader, VBAForm, FormControl,
                     pull/push for each host, VBAModuleKind, exceptions
  _host.py           VBAHostFile: shared open/edit/pull/push/save pipeline
  excel.py           ExcelFile (VBAHostFile subclass, create_new template)
  word.py            WordFile
  powerpoint.py      PowerPointFile (.ppt overrides the container hooks)
  access/            AccessDatabase: the VBA project, forms and reports,
                     and the Jet 4 / ACE storage engine they live in
  access_read.py     AccessReader: the older read-only inspector
  powerquery/        PowerQueryWorkbook: the DataMashup blob, the M section
                     document, the metadata beside it, and worksheet loads
  _deflate.py        classic zlib's deflate, for the bytes Office writes
  vba.py             VBA project parser and MS-OVBA codec
  vba_pcode.py       VBA7 p-code disassembler
  cfb.py             MS-CFB (Compound File Binary) parser/writer
  forms.py           UserForm designer streams: control tree, read and write
  _oforms_records.py [MS-OFORMS] property table, one per control class
  _oforms_pages.py   a MultiPage's tabs and page bookkeeping
  _ppt_container.py  the VBA project a binary .ppt hides in its document stream
  exceptions.py      exception hierarchy
  _templates/        empty .xlsx/.xlsm/.xlsb/.xlam/.docm/.pptm/.accdb bytes for create_new()
  __main__.py        python -m pyopenvba {pull,push,ls,forms,disasm,access-*,pq-ls,pq-pull,pq-push}
```

For more:

- [docs/architecture.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/architecture.md): internal layout and conventions.
- [docs/access_engine.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/access_engine.md): the Access file format as measured, and what the engine reproduces.
- [docs/power_query.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/power_query.md): how Excel stores Power Query, and how each rule was measured.
- [docs/ms-ovba-implementation-guide_v2.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/ms-ovba-implementation-guide_v2.md): a language-agnostic guide to re-implementing MS-OVBA.
- [docs/roadmap.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/roadmap.md): per-feature status.

---

## Contributing

Bug reports, files that break the library, and pull requests are welcome.
Please include the file, or a minimal redacted version, when filing a
parsing bug.

Run the same checks as CI:

```bash
pip install -e ".[dev]"
pyright src tests
pytest -p no:randomly
```

On Windows with desktop Office installed you can also run the live gates,
which are skipped by default and in CI. Each builds a file with pyOpenVBA
and has the real application open it, run its code, or perform the same
edit for a byte-for-byte comparison:

```powershell
$env:RUN_LIVE_EXCEL = "1"; pytest tests/test_live_excel_gate.py
$env:RUN_LIVE_ACCESS = "1"; pytest tests/test_live_access_engine_gate.py
$env:RUN_LIVE_ACCESS_VBA = "1"; pytest tests/test_live_access_design_gate.py
$env:RUN_LIVE_POWER_QUERY = "1"; pytest tests/test_live_powerquery_gate.py
```

CI runs the test matrix on Python 3.10 through 3.14 on Linux, plus 3.12
on Windows and macOS, on every push and pull request. Releases go to
PyPI when a `v*.*.*` tag is pushed.

---

## License

[MIT](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/LICENSE.md).

---

## Support open source

If pyOpenVBA saves you time or helps your team keep VBA maintainable,
support keeps the project moving.

- [GitHub Sponsors](https://github.com/sponsors/WilliamSmithEdward)
- [PayPal](https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F&currency_code=USD)
- [Cash App](https://cash.app/$williamesmithjcil)
