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

Power Query, in any Excel package:

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

They also share `references()`, `add_reference()` and `remove_reference()`.
These edit the VBA project's library references, as in Tools > References:

```python
with ExcelFile("workbook.xlsm") as wb:
    wb.add_reference("Word")  # Excel, Word, PowerPoint and Access presets
    wb.add_reference("Scripting", "420B2830-E718-11CF-893D-00A0C9054228",
                     path="C:/Windows/System32/scrrun.dll")
    for reference in wb.references():
        print(reference.name, reference.guid, reference.version)
    wb.remove_reference("Word")  # also accepts a GUID; returns whether removed
    wb.save()
```

Adding an existing GUID or removing an absent reference is a no-op.
The existing Access `drop_reference()` spelling remains available and raises
if the reference is absent. VBA and the file's own host library are implicit;
they are not listed or added. MSForms cannot be removed while UserForms exist.
Removing a library does not rewrite code that names its types. Library paths
are resolution hints; adding a reference does not install the library or add
its implementation to the headless runtime. Use these methods to persist edits,
rather than mutating `VBAProject.references` directly.

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

Deleting a UserForm's module takes the form's design with it, as the
editor does.

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
and the code-behind module together. A project with no Microsoft Forms
reference gets the one the editor adds with a first form, so code that
names `MSForms` types compiles:

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
    db.rename_form("Draft", "Invoice")
    db.delete_form("Old")
    db.save()
```

Renaming and deleting reach the code behind a design as well as the
design itself. A form's module is bound to it by name, so `Form_Draft`
becomes `Form_Invoice` and a deleted form takes its module with it.

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

## Running the macros

Reading and writing VBA is one thing; this runs it. `ExcelApplication`
is an Excel instance in memory, with a VBA interpreter attached: load a
workbook, execute a macro, look at what it did, write the file back.

```python
from pyopenvba.apps.excel import ExcelApplication

app = ExcelApplication.open("report.xlsm")   # sheets, names, queries, macros
app.run("BuildReport")
print(app.describe())                        # a readable dump of the grid
app.save("report_out.xlsm")
```

State can also be made up from nothing, with no file at all:

```python
app = ExcelApplication()
app.add_workbook()
app.add_module('Sub Main()\n    Range("A1").Value = 42\nEnd Sub\n', name="Module1")
app.run("Main")
app.sheet(1).value("A1")          # 42
app.evaluate('Range("A1").Value * 2')   # 84
```

Failures are three separate exceptions, because they mean different
things: `VBACompileError` for VBA that does not compile,
`VBARuntimeError` for an error VBA itself would raise, carrying the
number `Err` would hold, and `VBAUnsupportedError` for real VBA that
pyOpenVBA does not implement. The third is never trappable by
`On Error`: a gap here must not be swallowed and reported as a result.

Formulas are calculated, not just carried: write one from a macro and
read the answer back.

```python
app.add_module('''
Sub Total()
    Range("A1:A3").Value = 10
    Range("B1").Formula = "=SUM(A1:A3)"
End Sub
''', name="Module1")
app.run("Total")
app.sheet(1).value("B1")          # 30.0
```

About a hundred worksheet functions are implemented; anything else
Excel has says so by name rather than quietly answering `#NAME?`.

Power Query is evaluated too. Refreshing a query works out its M and
writes the rows to the sheet it loads to, so the table, its queryTable
and the saved file all follow:

```python
app = ExcelApplication.open("sales.xlsx")
app.refresh_query("Summary")      # the rows, headers first
app.save("sales_out.xlsx")
```

Queries can name each other, `Excel.CurrentWorkbook()` reads the
workbook's own tables and named ranges, and a source this cannot reach
without a network or a driver reports itself rather than guessing.

Word and PowerPoint run their macros too, each with its own object
model:

```python
from pyopenvba.apps.word import WordApplication
from pyopenvba.apps.powerpoint import PowerPointApplication

doc = WordApplication.open("report.docm")
doc.run("Relabel")
doc.save("report_out.docm")

deck = PowerPointApplication.open("deck.pptm")
deck.run("Rebuild")
deck.slide(1).shapes()            # what is on the first slide
deck.save("deck_out.pptm")
```

Shapes are read and written on all three surfaces: adding, moving,
resizing, renaming, retyping and deleting, with the macro a click runs
attached or removed. In Excel that is `Shape.OnAction`, form-control
buttons included; in PowerPoint it is
`ActionSettings(ppMouseClick).Run`; Word has no macro on a shape and
says so rather than pretending.

```python
app.add_module('''
Sub Draw()
    Dim sh As Object
    Set sh = ActiveSheet.Shapes.AddFormControl(0, 100, 50, 90, 30)
    sh.Name = "Go"
    sh.OnAction = "Refresh"
    sh.TextFrame.Characters.Text = "Refresh"
End Sub
''', name="Module1")
app.run("Draw")
app.save("with_button.xlsm")      # Excel opens it as a button
```

The same worksheet shapes can be edited directly from Python:

```python
app = ExcelApplication.open("report.xlsm", with_vba=False)
sheet = app.sheet(1)
sheet.add_button(name="Refresh", left=100, top=50, width=90, height=30,
                 text="Refresh", macro="RefreshReport")
sheet.update_shape("Refresh", left=120, text="Refresh report")
print(sheet.shape("Refresh").control)  # control details
sheet.update_shape("Refresh", macro="")  # unlink its macro
sheet.remove_shape("Refresh")
app.save("report_out.xlsm")
```

`Range.Formula` and `Range.FormulaR1C1` support scalar and array writes, with
shared bounds, error and merge handling. `FormulaR1C1` converts relative,
absolute and mixed references. Both properties return scalar or
two-dimensional array results, including requested trailing blank cells.
References survive save/reopen as ordinary A1 formulas. Array writes support
custom lower bounds, single-axis expansion, mixed values and formulas, and
Excel's `#N/A` filling for uncovered cells.

VBA `Range.Find`, `FindNext` and `FindPrevious` support single-area searches
over formula text or displayed values, with whole/partial matching,
wildcards, case sensitivity, traversal order and wraparound. Whole-sheet
searches keep empty cells sparse. Multi-area, comment and format searches,
and `MatchByte=True`, remain unsupported.

One `ExcelApplication` can hold multiple workbooks. Workbook handles let Python
callers select and save a particular book without depending on the active book:

```python
source_book = app.add_workbook()
source = app.sheet(1, workbook=source_book)
source.set_value("A1", 10)
destination_book = app.open_workbook("destination.xlsx")
destination = app.sheet(1, workbook=destination_book)
source.copy_range("A1:B5", destination, "D2", name_conflict="rename")
copied_sheet = source.copy(after=destination)  # omit before/after for a new book
copied_sheet.move(before=destination)
app.save("destination.xlsx", workbook=destination_book)
```

VBA supports `Workbooks.Add`, `Open`, `Activate`, `Close`, and `Worksheet.Copy`
before/after a sheet in another open workbook. Copies preserve modeled cells,
merges and relevant names. Range copies import names used by formulas, including
aliases. `name_conflict="reuse"` (the default) uses the destination definition,
matching unattended Excel; `"rename"` imports names with unused `_2`, `_3`, etc.
suffixes and rewrites copied formulas; `"error"` rejects collisions before making
changes. Names referencing the source sheet retain their cell addresses on the
destination sheet; they do not shift with the paste offset. External-link range
copies, relative name imports, sheet-copy name conflicts, drawings,
controls and advanced worksheet metadata are explicitly unsupported. Opening
additional workbooks does not create independent VBA execution projects.

VBA `Worksheet.Move Before:=...` / `After:=...` uses the same anchors as Copy.
Omitting both moves the sheet into a new workbook. Cross-book moves invalidate
held VBA worksheet, range and local-name objects (error 424), matching Excel;
reacquire them from the destination. Python `move` returns the new sheet view.
Same-book reordering preserves object identity. Moving the last sheet to an existing
workbook closes the empty source; moving its only sheet to a new workbook raises
1004. Moves that need external links or unsupported sheet-content transfer fail
before changing either workbook.

Named ranges support create/read/update/delete through VBA `Names` collections
and shared Python APIs. `app` operates on workbook names; `app.sheet(...)`
operates on worksheet-local names. Snapshots include scope, reference, visibility
and comments; renaming updates dependent formulas.

```python
app.add_named_range("Revenue", "=Sheet1!$B$2:$B$10", comment="Sales totals")
app.sheet(1).add_named_range("Revenue", "=$C$2:$C$10", visible=False)
print(app.named_ranges())                  # workbook and local definitions
print(app.named_range("Revenue"))         # workbook definition
app.update_named_range("Revenue", new_name="Sales", refers_to="=Sheet1!$B$2:$B$20")
app.sheet(1).update_named_range("Revenue", comment="Local override")
app.remove_named_range("Sales")
app.sheet(1).remove_named_range("Revenue")
app.save("named_ranges.xlsm")
```

Whole-row/column `Insert` and `Delete` update cell references, including absolute
and cross-sheet references, range boundaries and defined names, and move row
heights and column widths. Deleted targets become `#REF!`. Partial-cell reference
updates and full formatting/metadata movement remain incomplete; sheets with
merges or shapes are explicitly unsupported.

`Range.Copy Destination:=...` handles overlapping copies, blank source cells,
repeated destination blocks and relative formula shifts, and copies each
cell's format. Merged-cell and multi-area copies, large destinations and
clipboard pastes remain incomplete.

`Range.Merge`, `UnMerge`, `MergeCells` and single-cell `MergeArea` support
merged-region editing and save/reopen. Ordinary overlapping merges expand
to include the existing regions; `Across=True` merges each row separately.
Excel's first nonempty value or formula is retained, and partial clears raise
error 1004. Multi-area merges, across merges over existing multi-row regions,
and full merge-formatting parity remain outside the measured support.

Cell formats read and write the way Excel does: `Font`, `Interior`,
`Borders`, `BorderAround`, the alignment and protection properties,
`NumberFormat` and `ClearFormats`. A read over several cells answers Null
when they differ, theme colours take their tints the way Excel computes
them, and a border between two cells is stored and read the way Excel
shares it. A workbook Excel formatted reads back exactly as Excel reads
it, and a save adds only the stylesheet entries the file lacks, spelled
as Excel spells them. Formatting whole rows, whole columns or the whole
sheet keeps the row, column and sheet formats Excel keeps, with the
cells Excel makes where they cross, and saves them as Excel does. Named
cell styles and conditional formats are not implemented yet.

Row heights, column widths and hidden rows and columns behave as they do
in Excel on a 96-DPI display, the one the model emulates: `RowHeight = 20`
reads back 20, draws 19.5pt tall and saves as Excel saves it; `Hidden`,
`UseStandardHeight`, `UseStandardWidth`, `Height`, `Width`, `Left`, `Top`
and a sheet's standard sizes follow, and all of it is read from and saved
to the file. A row grows or shrinks with the fonts its cells, its own
format and its columns show, as Excel's does, from heights measured in
live Excel for the common Office and Windows fonts; a row
whose height rests on an unmeasured font, a mix of fonts, or wrapped
text reports itself unsupported, and column `AutoFit` measures no text
yet.

`shapes()` and `shape(name)` return detached snapshots, including each
control's linked cell, list range and value. Use `update_shape` to make
an edit. `add_shape` creates a measured AutoShape and `add_textbox`
creates a horizontal text box; coordinates and sizes are nonnegative
points. `update_control(name, linked_cell="$H$1", list_range="$J$1:$J$3")`
edits saved form-control bindings (list ranges apply to dropdowns and
list boxes); use an empty string to disconnect a binding. References
may be A1 cells/ranges or workbook and worksheet-local names in this
workbook. These binding edits leave cell
values untouched. `set_control_value(name, 1)` checks a checkbox; use
`-4146` or `0` to uncheck it and `2` for mixed. Changed states write
True, False or #N/A to the linked cell. Direct cell writes and calculated
formulas also update linked checkboxes, including links across sheets.
VBA exposes the same state through `Shape.ControlFormat.Value`; its
`LinkedCell` setter follows Excel's rebinding behavior.
For single-selection dropdowns and list boxes backed by ranges,
`set_control_value(name, 2)` selects the second row; zero clears the
selection. These controls synchronize numeric indexes with linked cells.
VBA also supports `ControlFormat.ListFillRange` and `ListCount`.
Changing the source range clamps the selection or restores it from the
linked cell without overwriting that cell.

Named bindings support aliases, local-name precedence and changes to a
name's target. A linked name covering a rectangle writes its top-left
cell. Missing names remain saved; they provide no linked target and an
empty list source. Shrinking or deleting a named list source adjusts its
selection. Named `CHOOSE`, `IF`, `INDEX`, `OFFSET` and `INDIRECT` formulas support dynamic
targets, including cell-driven list sizes and addresses. `INDEX` supports
single-area sources and zero indices selecting whole rows or columns.
Invalid `INDEX` and `OFFSET` ranges provide empty sources.
`CHOOSE` and `IF` resolve only the selected branch; selected scalar values
or errors provide empty sources, and cell-driven conditions redirect bindings.
`INDIRECT` accepts A1 and absolute R1C1 addresses, including quoted sheet names.
Other reference formulas, multi-area `INDEX`, relative R1C1 `INDIRECT`,
cyclic aliases and external workbook bindings remain explicitly unsupported.

Inline lists support `control_items`, `add_control_item`,
`update_control_item`, `remove_control_item`, and `clear_control_items`.
Their VBA equivalents are `List(index)`, `AddItem`, `RemoveItem`, and
`RemoveAllItems`. Item edits adjust selection indexes without changing
the linked cell. Use `set_control_selection_mode(name, 2)` and
`set_control_selection(name, [1, 3])` for multi-selection; mode 3 is
extended selection. Snapshots expose `items`, `selection_mode`, and
`selected_indices`. VBA supports `ControlFormat.MultiSelect` and
`Shape.DrawingObject.Selected(index)`. A multi-selection list's scalar
`Value` read raises Excel error 1004.

`set_control_items(name, ["one", "two"])` replaces the entire list after
validating every string. VBA supports whole-list `ControlFormat.List`
reads and assignments: reads return a detached one-based Variant array,
or Null when empty. VBA assignments follow Excel's error behavior,
including retaining a successfully written prefix when a later array
element is invalid. Python replacement validates before making changes.

As in Excel, adding or replacing an item in a range-backed list
disconnects the source and starts a new inline list; it does not edit
the source cells. Clearing disconnects it too, while removing individual
range-backed items raises an error. These conversions reset selection
and can write zero to the linked cell. Edit the source cells directly
when the control should remain range-backed.
Spinners and scroll bars support `set_control_value` and VBA
`ControlFormat.Value`, `LinkedCell`, `Min`, `Max` and `SmallChange`.
Scroll bars also expose `LargeChange`; Excel rejects that member on
spinners. Direct values must fit the bounds. Numeric linked-cell values
clamp the control while preserving the cell, text preserves its state,
and empty/error cells reset it to its minimum. Snapshots retain numeric
bounds and increments, and saved controls preserve their type and state.
Excel can reconcile a saved control with its linked cell when opening
the file; its stored value alone does not determine the displayed state.
`add_form_control(control_type, name="Choices", left=12, top=20)` creates
buttons (0), checkboxes (1), dropdowns (2), group boxes (4), labels (5),
list boxes (6), option buttons (7), scroll bars (8), and spinners (9).
It uses the same geometry, caption, macro and name arguments as `add_button`.
Multi/extended lists retain selections when switching between those two
modes. Changing their source keeps only selected indexes that fit the new
range; switching to single selection clears them. VBA rebinding writes the
stored scalar index when the destination cell differs, even though a
multi-selection list does not expose a scalar `Value` getter.
Radio groups support off (`-4146` or `0`) and on (`1`) through
`set_control_value` and VBA `ControlFormat.Value`. Selecting a button clears
its peers and writes its one-based group index to the shared linked cell;
cell edits select that index without rewriting the cell. Text preserves
selection; empty/error cells and out-of-range indexes clear it. Group
membership stays stable during ordinary movement and is reconstructed from
group-box geometry on reopening, matching Excel. Non-overlapping boxes
support interleaved creation and regrouping, including a separate unboxed
group. Splitting preserves selected buttons; merging keeps the first
selected button in control order. Adding a box to split interleaved radios
preserves the linked cell's previous index. Radios can also be created
inside existing overlapping or nested boxes: earlier boxes win partial
overlaps, while a strictly nested inner box wins regardless of creation
order. Adding and deleting those boxes transfers group links and selections
with Excel's measured behavior. A late nested box leaves the link with its
surviving outer group. An identical box can transfer the original link to
an unlinked unboxed group; an already linked unboxed group keeps its link.
Deleting a radio leaves the linked cell
unchanged; deleting the leader retains the link for a boxed group and
clears it for an unboxed group.
Creating pictures, charts and shape groups remains tracked work.
Removing a control cleans up its worksheet record, properties
part, relationship and VML while preserving other controls and notes.

Every answer the interpreter, the calculation engine and the three
object models give was measured in the application itself rather than
assumed, and a file nobody changed saves back byte for byte. What is
implemented, what is not, and how it was measured is in
[docs/vba_runtime.md](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/vba_runtime.md).

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

### Files with no VBA project

A file saved before its first macro has no VBA project: no
`vbaProject.bin` in a `.xlsm`, `.docm` or `.pptm`, and no project storage
in a binary file. That is its normal shape, and it opens like any other.
`has_vba_project()` says False, listing reads such as `module_names()`,
`vba_modules()`, `forms()` and `references()` answer empty, and a read or
write that needs the project raises `NoVBAProjectError`, a kind of
`VBAProjectError`:

```python
with ExcelFile("book.xlsm") as wb:
    if wb.has_vba_project():
        print(wb.module_names())
```

`add_vba_project()` gives a `.xlsm`, `.docm` or `.pptm` a project, the
one its application makes for a first macro. In Excel that is a
document module for the workbook and one for each sheet, named as Excel
names them; in Word it is `ThisDocument`; in PowerPoint it is empty. Add
modules to it and save:

```python
with ExcelFile("book.xlsm") as wb:
    if not wb.has_vba_project():
        wb.add_vba_project().add_module("Module1", "Public Sub Hello()\r\nEnd Sub\r\n")
    wb.save()
```

A project its application would not write is not written here either:
Excel writes only the code names for one with no code in it, and
PowerPoint nothing for one with no module.

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

Any change to the macros invalidates a digital signature. Excel, Word and
PowerPoint keep the signature of a `.xlsm`, `.docm` or `.pptm` in parts
beside `vbaProject.bin`, and `vba_signature()` reads it from there:

```python
info = wb.vba_signature()
info.present, info.kinds, info.parts
# (True, ['v3', 'agile', 'legacy'], ['xl/vbaProjectSignatureV3.bin', ...])
```

On a save that changes the project, the library takes the signature out
as the applications do: the parts, their relationships and their entries
in `[Content_Types].xml`. It then emits a `UserWarning`:

```python
import warnings
warnings.filterwarnings("error", category=UserWarning)   # treat as fatal

wb.save(allow_invalidate_signature=True)                 # or accept it
```

A save that changes nothing keeps the signature. A signature in a binary
`.xls`, `.doc` or `.ppt` file is not looked for: there it lives in the
file's property sets or a string table, and no signed binary file has
been measured.

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
- [docs/host_completeness.md](docs/host_completeness.md): whole-host completion criteria and the Excel coverage audit; Excel, then Word, then PowerPoint.

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

## Built with this

[xlide-mcp](https://github.com/WilliamSmithEdward/xlide_mcp)
is an MCP server built on this. It gives an AI agent the VBA, the
UserForms and the Power Query inside an Office file, so a model can read
a macro, rewrite it and save it back without anyone exporting modules by
hand. Every container this library opens is one that server can reach.

## License

[MIT](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/LICENSE.md).

---

## Support open source

If pyOpenVBA saves you time or helps your team keep VBA maintainable,
support keeps the project moving.

- [GitHub Sponsors](https://github.com/sponsors/WilliamSmithEdward)
- [PayPal](https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F&currency_code=USD)
- [Cash App](https://cash.app/$williamesmithjcil)
