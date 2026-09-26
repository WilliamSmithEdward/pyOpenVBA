"""What Excel's hyperlinks are and do: Hyperlinks.Add, the Hyperlink object, the collections and the file.

Hyperlinks.Add over empty cells and cells holding a number, a formula or
text, with every mix of address, subaddress, screen tip and text to show,
over one cell and several, over a link already there and over part of
one, on merged cells, with arguments of the wrong kind, and with
addresses of every kind -- which Excel tidies as a browser would, and
writes to the file tidied again; what the link and its cell read after;
setting each property on each kind of link, EmailSubject in an Excel of
its own; the ways a link goes -- Delete on a link, on a range's links
and on a sheet's, ClearHyperlinks, Clear, ClearContents, a new value,
deleting the cell -- and what the cell keeps; the order the collections
list links in, and which a range's collection holds; links moving with
inserted and deleted rows, columns and cells, copies, pastes, a cut, a
sort and a copy of the sheet; protected sheets; the HYPERLINK function
and the style it gives; the Hyperlink style, where the stylesheet puts
it and what a style of that name made first does; the relationship ids
a sheet's links take beside a note and a shape. Workbooks are saved
along the way, and some are opened again and read.

The probe is one module, which the record keeps, each section run in an
Excel of its own with a folder to save workbooks in; the replay runs the
same module in the model.

    python scripts/measure_hyperlinks.py

writes tests/fixtures/hyperlinks/hyperlinks.json and the workbooks Excel
saved, which tests/test_excel_hyperlinks.py replays.
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from fixture_workbook import copy_saved
from measure_cell_styles import HELPER, STYLE_READS, function, reads, statements

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "hyperlinks"

#: What a Hyperlink answers.
LINK_READS = ["h.Address", "h.SubAddress", "h.ScreenTip", "h.TextToDisplay", "h.Name", "h.Type", "h.Range.Address",
              "h.EmailSubject", "TypeName(h.Parent)", "TypeName(h.Application)", "h.Creator", "h.Shape.Name"]
#: What a cell answers once a link lands on it or goes.
CELL_READS = ["c.Value", "c.Formula", "c.NumberFormat", "c.Style.Name", "c.Font.Bold", "c.Font.Underline",
              "c.Font.Color", "c.Font.ThemeColor", "c.Interior.Color", "c.Hyperlinks.Count"]

#: (cell, statements before, the Add): links over empty cells and full ones, and every kind of address.
ADDS = [
    ("A1", [], 'ws.Hyperlinks.Add Anchor:=c, Address:="http://example.com/"'),
    ("A2", [], 'ws.Hyperlinks.Add c, "", "Sheet1!C3"'),
    ("A3", [], 'ws.Hyperlinks.Add c, "other.xlsx", "Sheet2!B2"'),
    ("A4", [], 'ws.Hyperlinks.Add c, "http://example.com/", ScreenTip:="Tip"'),
    ("A5", [], 'ws.Hyperlinks.Add c, "http://example.com/", TextToDisplay:="Shown"'),
    ("A6", ["c.Value = 5"], 'ws.Hyperlinks.Add c, "http://example.com/"'),
    ("A7", ['c.Formula = "=1+2"'], 'ws.Hyperlinks.Add c, "http://example.com/"'),
    ("A8", ['c.Value = "old"'], 'ws.Hyperlinks.Add c, "http://example.com/", TextToDisplay:="new"'),
    ("A9", ['c.Value = "old"'], 'ws.Hyperlinks.Add c, "http://example.com/"'),
    ("A12", ['ws.Hyperlinks.Add c, "http://one.example/"'], 'ws.Hyperlinks.Add c, "http://two.example/"'),
    ("A13", [], 'ws.Hyperlinks.Add c, "mailto:x@example.com?subject=Hi"'),
    ("A14", [], 'ws.Hyperlinks.Add c, "#Sheet1!A1"'),
    ("A15", [], 'ws.Hyperlinks.Add c, "http://example.com/a b"'),
    ("A16", [], 'ws.Hyperlinks.Add c, "HTTP://EXAMPLE.COM/Path"'),
    ("A17", [], 'ws.Hyperlinks.Add c, "www.example.com"'),
    ("A18", [], 'ws.Hyperlinks.Add c, "..\\up.xlsx"'),
    ("A19", [], 'ws.Hyperlinks.Add c, "sub/file.txt"'),
    ("A20", ["c.Font.Bold = True"], 'ws.Hyperlinks.Add c, "http://example.com/"'),
    ("A21", ['c.Style = "Good"'], 'ws.Hyperlinks.Add c, "http://example.com/"'),
    ("A22", [], 'ws.Hyperlinks.Add c, "http://example.com/", TextToDisplay:=""'),
    ("A23", [], 'ws.Hyperlinks.Add c, "http://example.com/", "", "", ""'),
    ("A24", ['c.NumberFormat = "0.00"', "c.Value = 1.5"], 'ws.Hyperlinks.Add c, "http://example.com/"'),
    ("A25", [], 'ws.Hyperlinks.Add c, "http://example.com/#frag"'),
    ("A26", [], 'ws.Hyperlinks.Add c, "", "Sheet1!C3", TextToDisplay:="Inside"'),
    ("A27", [], 'ws.Hyperlinks.Add c, "", "Nowhere!Z99"'),
    ("A28", [], 'ws.Hyperlinks.Add c, "file.xlsx#Sheet1!A1"'),
]
#: A link over several cells, read cell by cell.
SPREAD = 'ws.Hyperlinks.Add ws.Range("A10:B11"), "http://example.com/"'

#: (cell, what is done to its link): the ways a link goes and what the cell keeps.
REMOVALS = [("B1", "c.Hyperlinks(1).Delete"), ("B2", "c.Hyperlinks.Delete"), ("B3", "c.ClearHyperlinks"),
            ("B4", "c.Clear"), ("B5", "c.ClearContents"), ("B6", "c.ClearFormats"), ("B7", 'c.Value = "x"'),
            ("B8", 'c.Formula = "=2*3"'), ("B9", "c.Hyperlinks(1).TextToDisplay = \"Other\""),
            ("B10", 'c.Style = "Normal"'), ("B11", "c.Font.Underline = -4142")]

#: Statements setting a link's properties, each followed by the link's and its cell's reads. EmailSubject is
#: measured on its own (SUBJECTS): set on a link that is not a mail link, it leaves Excel refusing every
#: change after it.
SETS = ['h.Address = "http://changed.example/"', 'h.SubAddress = "Sheet1!D4"', 'h.ScreenTip = "T"',
        'h.TextToDisplay = "Display"', 'h.Address = ""', 'h.SubAddress = ""', 'h.Name = "Renamed"']
#: The same for a mail link.
MAIL_SETS = ['h.Address = "mailto:y@example.com"', 'h.Address = "mailto:z@example.com?subject=Third"']
#: (cell, statements making its link, the statement setting it): one property set on each kind of link.
SETTINGS = [
    ("A4", ['ws.Hyperlinks.Add c, "http://example.com/", TextToDisplay:="Shown"'],
     'c.Hyperlinks(1).Address = "http://other.example/"'),
    ("A5", ['ws.Hyperlinks.Add c, "other.xlsx"'], 'c.Hyperlinks(1).SubAddress = "Sheet2!B2"'),
    ("A6", ['ws.Hyperlinks.Add c, "", "Sheet1!C3"'], 'c.Hyperlinks(1).Address = "http://example.com/"'),
    ("A7", ['ws.Hyperlinks.Add c, "http://example.com/", ScreenTip:="Tip"'], 'c.Hyperlinks(1).ScreenTip = ""'),
    ("A8", ['ws.Hyperlinks.Add c, "http://example.com/"'], 'c.Hyperlinks(1).TextToDisplay = ""'),
    ("A9", ['c.Formula = "=1+2"', 'ws.Hyperlinks.Add c, "http://example.com/"'],
     'c.Hyperlinks(1).TextToDisplay = "Over"'),
    ("A12", ['ws.Hyperlinks.Add c, "http://example.com/"'], 'c.Hyperlinks(1).Address = "#Sheet1!A1"'),
    ("A13", ['ws.Hyperlinks.Add c, "http://example.com/"'], 'c.Hyperlinks(1).Address = "HTTP://UPPER.EXAMPLE/X"'),
    ("A14", ['ws.Hyperlinks.Add c, "http://example.com/"'], 'c.Hyperlinks(1).SubAddress = "frag"'),
    ("A15", ['ws.Hyperlinks.Add c, "other.xlsx", "Sheet2!B2"'], 'c.Hyperlinks(1).Address = ""'),
    ("A16", ['ws.Hyperlinks.Add c, "http://example.com/"'],
     'c.Hyperlinks(1).ScreenTip = "A ""quoted"" <tip> & more"'),
    ("A17", ["c.Value = 7", 'ws.Hyperlinks.Add c, "http://example.com/"'], 'c.Hyperlinks(1).TextToDisplay = "Seven"'),
    ("A18", ['ws.Hyperlinks.Add c, "http://example.com/"'], "c.Hyperlinks(1).TextToDisplay = 5"),
    ("A19", ['ws.Hyperlinks.Add c, "mailto:x@example.com"'], 'c.Hyperlinks(1).Address = "mailto:y@example.com"'),
    ("A20", ['ws.Hyperlinks.Add c, "http://example.com/", TextToDisplay:="Shown"'], 'c.Value = "Changed"'),
    ("A21", ['ws.Hyperlinks.Add c, "http://example.com/"'], 'c.Hyperlinks(1).Address = "other.xlsx#Sheet2!B2"'),
    ("A22", ['ws.Hyperlinks.Add c, "http://example.com/"'], 'c.Hyperlinks(1).ScreenTip = 5'),
]
#: A link over four cells given a text to show, and the cells read after.
WIDE_TEXT = ['ws.Hyperlinks.Add ws.Range("A23:B24"), "http://example.com/"',
             'ws.Range("A23").Hyperlinks(1).TextToDisplay = "Wide"']

#: (section, statements making the link, the EmailSubject set): each run in an Excel of its own.
SUBJECTS = {
    "SubjectMail": ['ws.Hyperlinks.Add c, "mailto:x@example.com"', 'c.Hyperlinks(1).EmailSubject = "Hi there"'],
    "SubjectReplace": ['ws.Hyperlinks.Add c, "mailto:x@example.com?subject=Old&cc=y@example.com"',
                       'c.Hyperlinks(1).EmailSubject = "New"'],
    "SubjectEmpty": ['ws.Hyperlinks.Add c, "mailto:x@example.com?subject=Old"', 'c.Hyperlinks(1).EmailSubject = ""'],
    "SubjectWeb": ['ws.Hyperlinks.Add c, "http://example.com/"', 'c.Hyperlinks(1).EmailSubject = "Hi"'],
    "SubjectInternal": ['ws.Hyperlinks.Add c, "", "Sheet1!C3"', 'c.Hyperlinks(1).EmailSubject = "Hi"'],
}

#: Ranges asked for their links, around a link over A10:B11 and two of one cell.
RANGE_COUNTS = ["A10:A11", "A10", "B10", "A11", "B10:B11", "A9:A11", "A10:B11", "A10:C12", "A9:C12", "B11:C12",
                "A1:A10", "A10,B11", "A9:B10", "A1:C5", "10:10", "A:A", "A1:C12", "B11:B12"]

#: (cell or block, statements before, what takes the link off): what a cell keeps once its link goes.
DELETIONS = [
    ("A1", ["c.Font.Bold = True", "c.Interior.Color = 255", 'c.NumberFormat = "0.00"', "c.Borders(9).LineStyle = 1"],
     "c.Hyperlinks(1).Delete"),
    ("A2", ['c.Style = "Good"'], "c.Hyperlinks(1).Delete"),
    ("A4", ["c.Font.Italic = True"], "c.Hyperlinks(1).Delete"),
    ("A5", ["c.ClearFormats"], "c.Hyperlinks(1).Delete"),
    ("A6", ["c.Font.Underline = -4142"], "c.Hyperlinks.Delete"),
]
#: The same where the style came first.
STYLED_FIRST = ("A3", ['c.Style = "Good"'], "c.Hyperlinks(1).Delete")
#: (block, the cell acted on, what is done): a link over four cells losing part of itself.
WIDE_DELETIONS = [("A10:B11", "B11", "c.Hyperlinks.Delete"), ("D10:E11", "E11", "c.ClearHyperlinks"),
                  ("G10:H11", "H11", "c.ClearContents"), ("G13:H14", "G13", 'c.Value = "x"'),
                  ("J10:K11", "K11", "c.Clear"), ("J13:K14", "K14", "c.Hyperlinks(1).Delete")]

#: Add with arguments of the wrong kind, one at a time.
ARGUMENTS = ['ws.Hyperlinks.Add ws.Range("D6"), "http://example.com/", 7',
             'ws.Hyperlinks.Add ws.Range("D7"), "http://example.com/", "", 8',
             'ws.Hyperlinks.Add ws.Range("D8"), "http://example.com/", "", "", 9',
             'ws.Hyperlinks.Add ws.Range("D9"), ""',
             'ws.Hyperlinks.Add ws.Range("D10"), "", ""',
             'ws.Hyperlinks.Add ws.Range("D11"), Null',
             'ws.Hyperlinks.Add ws.Range("D12"), "http://example.com/", Null',
             'ws.Hyperlinks.Add ws.Range("D13"), True',
             'ws.Hyperlinks.Add Nothing, "http://example.com/"',
             'ws.Hyperlinks.Add ws.Range("D14"), "http://example.com/", , , "Text", "extra"',
             'ws.Hyperlinks.Add ws.Range("D15"), "http://example.com/", ScreenTip:=Null',
             'ws.Hyperlinks.Add ws.Range("D16"), "http://example.com/", TextToDisplay:=Null']

#: Addresses of every kind, each added to a cell of its own, read, saved and read again.
ADDRESSES = ["Http://Example.COM/Path?Q=1", "HTTPS://A.Example/C", "ftp://X.Example/", "MAILTO:X@Example.COM",
             "mailto:y@example.com?Subject=Hi&body=x", "http://example.com/100%", "http://example.com/%41",
             "http://example.com/a+b", "http://example.com/a#b c", "file with space.xlsx", "C:\\docs\\file.xlsx",
             "\\\\server\\share\\file.xlsx", "file:///C:/docs/file.xlsx", "news:comp.lang", "#'My Sheet'!A1",
             "Sheet1!A1", "http://example.com:8080/", "http://user@example.com/", "HTTP://EXAMPLE.COM",
             "http://example.com", "http://example.com/path/", "http://example.com/?", "https://example.com/a/../b",
             "www.Example.com/Path", "example.com", "  http://example.com/  ", "http://example.com/<>",
             'http://example.com/"q"', "..\\Folder\\File.xlsx", "sub\\File.XLSX#Sheet1!A1", "http://example.com/#",
             "#", "http://example.com/a%20b", "C:/docs/file.xlsx", "http://example.com/a\\b",
             # A DOS path on a drive other than the one the workbook is saved on.
             "Q:\\docs\\file.xlsx", "Q:/docs/file.xlsx", "file:///Q:/docs/file.xlsx"]
#: The same with characters a module cannot spell, built with ChrW: (cell text, VBA expression).
WIDE_ADDRESSES = [("umlaut path", '"http://example.com/" & ChrW(228)'),
                  ("umlaut host", '"http://" & ChrW(228) & ".example/"'),
                  ("umlaut query", '"http://example.com/?x=" & ChrW(252)'),
                  ("tab", '"http://example.com/a" & Chr(9) & "b"')]

#: (cell, statements): the HYPERLINK function entered each way, over cells of each kind.
FORMULAS = [
    ("B1", ['c.Formula = "=HYPERLINK(""http://example.com/"",""go"")"']),
    ("B2", ['c.Formula = "=IF(1,HYPERLINK(""http://example.com/""))"']),
    ("B3", ['c.Formula = "=HYPERLINK(""http://example.com/"")&""x"""']),
    ("B4", ['c.Style = "Good"', 'c.Formula = "=HYPERLINK(""http://example.com/"")"']),
    ("B5", ["c.Font.Bold = True", 'c.Formula = "=HYPERLINK(""http://example.com/"")"']),
    ("B6", ['c.Value = "=HYPERLINK(""http://example.com/"")"']),
    ("B7", ['c.FormulaR1C1 = "=HYPERLINK(""http://example.com/"")"']),
    ("B10", ['c.Formula = "=hyperlink(""http://example.com/"")"']),
    ("B11", ['c.Formula = "=HYPERLINK(""http://example.com/"")"', 'c.Formula = "=1"']),
    ("B12", ['ws.Hyperlinks.Add c, "http://example.com/"', 'c.Formula = "=HYPERLINK(""http://example.com/x"")"']),
    ("B13", ['c.Formula2 = "=HYPERLINK(""http://example.com/"")"']),
    ("B14", ['c.Formula = "=+HYPERLINK(""http://example.com/"")"']),
    ("B15", ['c.Formula = "=(HYPERLINK(""http://example.com/""))"']),
    ("B16", ['c.Formula = "=HYPERLINK(A1)"']),
    ("B17", ['c.NumberFormat = "0.00"', 'c.Formula = "=HYPERLINK(""http://example.com/"",5)"']),
    ("B18", ['c.Formula = "=SUM(1)"', 'c.Formula = "=HYPERLINK(""http://example.com/"")"']),
    ("B19", ['c.Interior.Color = 255', 'c.Formula = "=HYPERLINK(""http://example.com/"")"']),
]
#: What else a cell answers once its link goes.
FORMAT_READS = ["c.Font.Italic", "c.Borders(9).LineStyle", "c.Interior.Pattern"]

#: Links of each kind that are saved and opened again, then given a new address.
REOPENED = ['ws.Hyperlinks.Add ws.Range("A1"), "http://example.com/"',
            'ws.Hyperlinks.Add ws.Range("A2"), "http://example.com/", TextToDisplay:="Shown"',
            'ws.Range("A3").Value = 5', 'ws.Hyperlinks.Add ws.Range("A3"), "http://example.com/"',
            'ws.Hyperlinks.Add ws.Range("A4"), "http://example.com/"', 'ws.Range("A4").Value = "x"',
            'ws.Hyperlinks.Add ws.Range("A5"), "", "Sheet1!C3"',
            'ws.Hyperlinks.Add ws.Range("A6"), "other.xlsx", "Sheet2!B2"',
            'ws.Hyperlinks.Add ws.Range("A7:B8"), "http://example.com/"',
            'ws.Hyperlinks.Add ws.Range("A9"), "http://example.com/", ScreenTip:="Tip"']
REOPENED_CELLS = ["A1", "A2", "A3", "A4", "A5", "A6", "A7", "A9"]

#: Links added over links, each group of cells read or listed after.
OVERLAPS = ['ws.Range("A1").Value = "keep"', 'ws.Range("B2").Value = 5',
            'ws.Hyperlinks.Add ws.Range("A1:B2"), "http://example.com/AB"',
            *(f'out = out & CellProps(ws.Range("{one}")) & "^"' for one in ("A1", "B1", "B2")), 'out = out & "~"',
            'ws.Hyperlinks.Add ws.Range("A10:B11"), "http://example.com/wide"',
            'ws.Hyperlinks.Add ws.Range("B11"), "http://example.com/B11"', 'out = out & Listing(ws) & "~"',
            *(f'out = out & CellProps(ws.Range("{one}")) & "^"' for one in ("A10", "B11")), 'out = out & "~"',
            'ws.Hyperlinks.Add ws.Range("C1"), "http://example.com/C1"',
            'ws.Hyperlinks.Add ws.Range("C1:D2"), "http://example.com/CD"', 'out = out & Listing(ws) & "~"',
            'ws.Hyperlinks.Add ws.Range("E1"), "http://example.com/E1"',
            'ws.Hyperlinks.Add ws.Range("E2"), "http://example.com/E2"',
            'ws.Hyperlinks.Add ws.Range("E1"), "http://again.example/"', 'out = out & Listing(ws) & "~"',
            'ws.Hyperlinks.Add ws.Range("F1:G2"), "http://example.com/FG"',
            'ws.Hyperlinks.Add ws.Range("F1:G2"), "http://example.com/FG2"', 'out = out & Listing(ws) & "~"',
            'ws.Hyperlinks.Add ws.Range("H1:I2"), "http://example.com/HI"',
            'ws.Hyperlinks.Add ws.Range("H2:I3"), "http://example.com/HI2"', 'out = out & Listing(ws) & "~"']

#: Adds, and the edits after each of which the sheet's links are listed.
STRUCTURE = ['ws.Hyperlinks.Add ws.Range("A10:B11"), "http://example.com/wide"',
             'ws.Hyperlinks.Add ws.Range("D3"), "http://example.com/D3"',
             "ws.Rows(11).Insert", "ws.Rows(12).Delete", "ws.Rows(11).Delete", "ws.Columns(1).Delete",
             'ws.Hyperlinks.Add ws.Range("C20:D21"), "http://example.com/C20"', 'ws.Range("C20").Delete -4162',
             'ws.Hyperlinks.Add ws.Range("F20:G21"), "http://example.com/F20"', 'ws.Range("F20:G20").Insert -4121',
             'ws.Hyperlinks.Add ws.Range("J20:K21"), "http://example.com/J20"', 'ws.Range("J20").Copy ws.Range("M1")',
             'ws.Range("J20:K21").Copy ws.Range("M5")', 'ws.Range("J20:K21").Cut ws.Range("P1")',
             'ws.Rows("20:21").Delete']

#: A protected sheet whose cells A1:A6 are unlocked, A3:A6 holding links.
UNLOCKED = ["ws.Protect", 'ws.Hyperlinks.Add ws.Range("A1"), "http://example.com/A1"',
            'ws.Range("A3").Hyperlinks(1).Address = "http://changed.example/"', 'ws.Range("A3").Hyperlinks.Delete',
            "ws.Unprotect", "ws.Protect AllowInsertingHyperlinks:=True",
            'ws.Hyperlinks.Add ws.Range("A2"), "http://example.com/A2"', 'ws.Range("A4").ClearHyperlinks',
            'ws.Range("A5").Hyperlinks(1).Delete', 'ws.Range("A6").Hyperlinks.Delete', "ws.Unprotect"]

#: Sheets with links beside a note and a shape, in each order.
RELS = ['ws.Range("A1").AddComment "note"', 'ws.Hyperlinks.Add ws.Range("B1"), "http://example.com/B1"',
        'ws.Hyperlinks.Add ws.Range("B2"), "http://example.com/B2"',
        "Set ws = wb.Worksheets.Add(After:=ws)", 'ws.Hyperlinks.Add ws.Range("B1"), "http://example.com/B1"',
        'ws.Hyperlinks.Add ws.Range("B2"), "http://example.com/B2"', 'ws.Range("A1").AddComment "note"',
        "Set ws = wb.Worksheets.Add(After:=ws)", 'ws.Hyperlinks.Add ws.Range("B1"), "http://example.com/B1"',
        "ws.Shapes.AddShape 1, 100, 100, 50, 20", 'ws.Hyperlinks.Add ws.Range("B2"), "http://example.com/B2"',
        "Set ws = wb.Worksheets.Add(After:=ws)", 'ws.Hyperlinks.Add ws.Range("B1"), "", "Sheet1!A1"',
        'ws.Hyperlinks.Add ws.Range("B2"), "http://example.com/B2"', 'ws.Range("A1").AddComment "note"',
        'ws.Hyperlinks.Add ws.Range("B3"), "", "Sheet1!A2"']


#: Styles given before and after a link, for the place the Hyperlink style takes among the built-in ones.
STYLE_ORDER = ['ws.Range("A1").Style = "Title"', 'ws.Hyperlinks.Add ws.Range("A2"), "http://example.com/"',
               'ws.Range("A3").Style = "Percent"', 'ws.Range("A4").Style = "Good"', 'ws.Range("A5").Style = "Accent1"',
               'ws.Range("A6").Style = "Note"']
#: A style a macro makes before the link, and the last standard style after it: whether Hyperlink comes among
#: the built-in styles or where it was made.
STYLE_ORDER_MADE = ['wb.Styles.Add "Mine"', 'wb.Styles("Mine").Font.Bold = True', 'ws.Range("A1").Style = "Mine"',
                    'ws.Hyperlinks.Add ws.Range("A2"), "http://example.com/"',
                    'ws.Range("A3").Style = "60% - Accent6"', 'ws.Range("A4").Style = "Good"']
#: The Hyperlink style asked for before any link makes it.
STYLE_FIRST = ['v = wb.Styles("Hyperlink").Name', 'ws.Range("B1").Style = "Hyperlink"',
               'ws.Range("B2").Style = "Followed Hyperlink"', 'wb.Styles.Add "Hyperlink"',
               'ws.Hyperlinks.Add ws.Range("B3"), "http://example.com/"']

#: Links copied and pasted every way, and copies landing on links.
PASTES = ['ws.Hyperlinks.Add ws.Range("A1"), "http://example.com/A1"', 'ws.Range("A1").Font.Bold = True',
          'ws.Range("A1").Copy', 'ws.Range("C1").PasteSpecial -4104', 'ws.Range("A1").Copy',
          'ws.Range("C2").PasteSpecial -4163', 'ws.Range("A1").Copy', 'ws.Range("C3").PasteSpecial -4122',
          'ws.Range("A1").Copy', 'ws.Range("C4").Select', "ws.Paste", "Application.CutCopyMode = False",
          'ws.Hyperlinks.Add ws.Range("E1"), "http://example.com/E1"', 'ws.Range("A1").Copy ws.Range("E1")',
          'ws.Hyperlinks.Add ws.Range("E2"), "http://example.com/E2"', 'ws.Range("A5").Value = "plain"',
          'ws.Range("A5").Copy ws.Range("E2")', 'ws.Hyperlinks.Add ws.Range("E3:F4"), "http://example.com/E3"',
          'ws.Range("A1").Copy ws.Range("F4")']

#: Edges of which links a range holds and which a clear takes.
EDGES = ['ws.Hyperlinks.Add ws.Range("A10:B11"), "http://example.com/wide"',
         'out = out & ws.Range("A10:A11,B10:B11").Hyperlinks.Count & "`"',
         'ws.Range("A9:A11").ClearContents', 'out = out & ws.Hyperlinks.Count & "`"',
         'ws.Range("A10").Value = Empty', 'out = out & ws.Hyperlinks.Count & "`"',
         'ws.Hyperlinks.Add ws.Range("D1"), 5', 'out = out & CellProps(ws.Range("D1")) & "`"',
         'ws.Hyperlinks.Add ws.Range("D2"), 2.5', 'out = out & CellProps(ws.Range("D2")) & "`"',
         'ws.Range("D3").Value = "old"', 'ws.Hyperlinks.Add ws.Range("D3"), 7',
         'out = out & CellProps(ws.Range("D3")) & "`"',
         'ws.Rows(10).Delete', 'out = out & Listing(ws) & "`"',
         'ws.Hyperlinks.Add ws.Range("F1:F3"), "http://example.com/F"', 'ws.Range("F1").Value = 3',
         'ws.Range("F2").Value = 1', 'ws.Range("F3").Value = 2', 'ws.Range("F1:F3").Sort ws.Range("F1"), 1',
         'out = out & Listing(ws) & "`"',
         'ws.Range("H1").Value = "x"', 'ws.Range("H2").Value = "y"',
         'ws.Hyperlinks.Add ws.Range("H1"), "http://example.com/H1"', 'ws.Range("H1:H2").Sort ws.Range("H1"), 2',
         'out = out & Listing(ws) & "`" & CellProps(ws.Range("H2")) & "`"',
         'ws.Range("J1").Hyperlinks.Add ws.Range("J1"), "http://example.com/J1"', 'out = out & Listing(ws) & "`"',
         'ws.Hyperlinks.Add ws.Range("L1"), "http://example.com/L1"', 'ws.Range("L1").Value = ""',
         'ws.Hyperlinks.Add ws.Range("L2"), "http://example.com/L2"', 'ws.Range("L2").Formula = ""',
         'ws.Hyperlinks.Add ws.Range("L3"), "http://example.com/L3"', 'ws.Range("L3").Value = 0',
         'out = out & Listing(ws) & "`"',
         'ws.Hyperlinks.Delete', 'out = out & ws.Hyperlinks.Count & "`" & CellProps(ws.Range("J1")) & "`"']


def _cells(block: str) -> list[str]:
    """Each cell of a block such as A10:B11, row by row."""
    first, _, last = block.partition(":")
    top, bottom = int(first[1:]), int((last or first)[1:])
    left, right = ord(first[0]), ord((last or first)[0])
    return [f"{chr(column)}{row}" for row in range(top, bottom + 1) for column in range(left, right + 1)]

#: Reads that go wrong, and the collection's own.
ERRORS = ["ws.Hyperlinks.Count", "ws.Hyperlinks(0).Address", "ws.Hyperlinks(3).Address",
          'ws.Hyperlinks("http://example.com/").Address', "TypeName(ws.Hyperlinks)", "TypeName(ws.Hyperlinks(1))",
          "TypeName(ws.Hyperlinks.Parent)", "ws.Hyperlinks.Creator", 'ws.Range("A1").Hyperlinks.Count',
          'TypeName(ws.Range("A1:B2").Hyperlinks)', 'ws.Range("C1").Hyperlinks(1).Address', "ws.Hyperlinks(1)"]
ERROR_STATEMENTS = ['ws.Hyperlinks.Add "D1", "http://example.com/"',
                    'ws.Hyperlinks.Add ws.Range("D2")',
                    'ws.Hyperlinks.Add ws2.Range("A1"), "http://example.com/"',
                    'ws.Hyperlinks.Add ws.Range("D3:D4,F3"), "http://example.com/"',
                    'ws.Hyperlinks.Add ws.Range("D5"), 5',
                    'ws.Hyperlinks.Add ws.Range("D6"), "http://example.com/", 7, 8, 9']


def _quoted(text: str) -> str:
    return text.replace('"', '""')


def module() -> str:
    """The whole probe: a function per section, fields apart by `, items by ^ and groups by ~."""
    parts = [HELPER]
    parts += ["Private Function LinkProps(h As Object) As String", "Dim out As String, v As Variant",
              "On Error Resume Next", *reads(LINK_READS), "LinkProps = out", "End Function"]
    parts += ["Private Function CellProps(c As Object) As String", "Dim out As String, v As Variant",
              "On Error Resume Next", *reads(CELL_READS),
              'If c.Hyperlinks.Count > 0 Then out = out & "|" & LinkProps(c.Hyperlinks(1))', "CellProps = out",
              "End Function"]
    parts += ["Private Function StyleProps(s As Object) As String", "Dim out As String, v As Variant",
              "On Error Resume Next", *reads(STYLE_READS), "StyleProps = out", "End Function"]
    parts += ["Private Function Listing(ws As Object) As String", "Dim out As String, h As Object",
              "For Each h In ws.Hyperlinks", 'out = out & h.Range.Address(False, False) & "`"', "Next",
              "Listing = out", "End Function"]
    # Links added over empty cells and full ones, of every kind; the workbook is saved and opened again.
    for number, batch in enumerate((ADDS[:13], ADDS[13:])):
        body = ["Set wb = book", "Set ws = wb.Worksheets(1)"]
        for cell, before, add in batch:
            body += [f'Set c = ws.Range("{cell}")', *statements([*before, add]), 'out = out & CellProps(c) & "^"']
        parts += function(f"Adding{number}", "book As Object", body)
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", "wb.Worksheets.Add After:=ws",
            'out = out & Adding0(wb) & "~" & Adding1(wb) & "~"', *statements([SPREAD]),
            *(f'out = out & CellProps(ws.Range("{cell}")) & "^"' for cell in ("A10", "B10", "A11", "B11")),
            'out = out & "~" & Listing(ws) & "~"',
            *reads(["wb.Styles.Count", 'wb.Styles("Hyperlink").Name', 'wb.Styles("Followed Hyperlink").Name']),
            'out = out & "~" & StyleProps(wb.Styles("Hyperlink")) & "~"',
            'wb.SaveAs Filename:=target & "links.xlsx"', "wb.Close False",
            'Set wb = Workbooks.Open(target & "links.xlsx")', "Set ws = wb.Worksheets(1)",
            'out = out & Listing(ws) & "~"']
    for cell, _, _ in ADDS:
        body += [f'out = out & CellProps(ws.Range("{cell}")) & "^"']
    body += ['wb.SaveAs Filename:=target & "links_again.xlsx"', "wb.Close False"]
    parts += function("Adds", "target As String", body)
    # The ways a link goes, each on its own cell.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1").Value = 1']
    for cell, action in REMOVALS:
        body += [f'Set c = ws.Range("{cell}")', f'ws.Hyperlinks.Add c, "http://example.com/{cell}"',
                 *statements([action]), 'out = out & CellProps(c) & "^"']
    body += ['ws.Hyperlinks.Add ws.Range("C1"), "http://example.com/C1"',
             'ws.Hyperlinks.Add ws.Range("C2"), "http://example.com/C2"',
             *statements(['ws.Range("C1:C2").Hyperlinks.Delete']),
             'out = out & CellProps(ws.Range("C1")) & "^" & CellProps(ws.Range("C2")) & "~"',
             *reads(["ws.Hyperlinks.Count"]), 'out = out & "~" & Listing(ws)',
             'wb.SaveAs Filename:=target & "removed.xlsx"', "wb.Close False"]
    parts += function("Removing", "target As String", body)
    parts += ["Private Function FormatProps(c As Object) As String", "Dim out As String, v As Variant",
              "On Error Resume Next", *reads(FORMAT_READS), "FormatProps = out", "End Function"]
    # A link's properties set one after another, then one set on each kind of link.
    body = ["Dim h As Object", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            'Set c = ws.Range("A1")', 'Set h = ws.Hyperlinks.Add(c, "http://example.com/")',
            'out = out & TypeName(h) & "~"']
    for line in SETS:
        body += [*statements([line]), 'out = out & CellProps(c) & "^"']
    body += ['Set c = ws.Range("A3")', 'Set h = ws.Hyperlinks.Add(c, "mailto:x@example.com")', 'out = out & "~"']
    for line in MAIL_SETS:
        body += [*statements([line]), 'out = out & CellProps(c) & "^"']
    body += ['out = out & "~"']
    for cell, before, change in SETTINGS:
        body += [f'Set c = ws.Range("{cell}")', *statements([*before, change]), 'out = out & CellProps(c) & "^"']
    body += ['out = out & "~"', *statements(WIDE_TEXT),
             *(f'out = out & CellProps(ws.Range("{cell}")) & "^"' for cell in _cells("A23:B24")),
             'wb.SaveAs Filename:=target & "set.xlsx"', "wb.Close False"]
    parts += function("Setting", "target As String", body)
    # Links saved and opened again: what they read, and what a new address does to them.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", *REOPENED,
            'wb.SaveAs Filename:=target & "reopen.xlsx"', "wb.Close False",
            'Set wb = Workbooks.Open(target & "reopen.xlsx")', "Set ws = wb.Worksheets(1)",
            *(f'out = out & CellProps(ws.Range("{cell}")) & "^"' for cell in REOPENED_CELLS), 'out = out & "~"']
    for cell in REOPENED_CELLS:
        body += [*statements([f'ws.Range("{cell}").Hyperlinks(1).Address = "http://after.example/"']),
                 f'out = out & CellProps(ws.Range("{cell}")) & "^"']
    body += ['wb.SaveAs Filename:=target & "reopened.xlsx"', "wb.Close False"]
    parts += function("Reopen", "target As String", body)
    # Which links a range's collection holds.
    body = ["Dim h As Object", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            'ws.Hyperlinks.Add ws.Range("A10:B11"), "http://example.com/wide"',
            'ws.Hyperlinks.Add ws.Range("C5"), "http://example.com/C5"',
            'ws.Hyperlinks.Add ws.Range("A1"), "http://example.com/A1"',
            *reads([f'ws.Range("{one}").Hyperlinks.Count' for one in RANGE_COUNTS]), 'out = out & "~"',
            *reads(['ws.Range("A10:A11").Hyperlinks(1).Range.Address', 'ws.Range("A10").Hyperlinks(1).Range.Address',
                    'TypeName(ws.Range("A1:C12").Hyperlinks.Parent)', 'ws.Range("A1:C12").Hyperlinks.Parent.Address',
                    'ws.Range("A1:C12").Hyperlinks(0).Address', 'ws.Range("A1:C12").Hyperlinks(4).Address',
                    'ws.Range("A1:C12").Hyperlinks("http://example.com/C5").Range.Address']),
            'out = out & "~"', 'For Each h In ws.Range("A1:C12").Hyperlinks',
            'out = out & h.Range.Address(False, False) & "`"', "Next", "wb.Close False"]
    parts += function("Ranges", "", body)
    # What a cell keeps once its link goes, and a link over four cells losing part of itself.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)"]
    for cell, before, action in DELETIONS:
        body += [f'Set c = ws.Range("{cell}")', f'ws.Hyperlinks.Add c, "http://example.com/{cell}"',
                 *statements([*before, action]), 'out = out & CellProps(c) & FormatProps(c) & "^"']
    cell, before, action = STYLED_FIRST
    body += [f'Set c = ws.Range("{cell}")', *statements(before), f'ws.Hyperlinks.Add c, "http://example.com/{cell}"',
             'out = out & CellProps(c) & "^"', *statements([action]), 'out = out & CellProps(c) & "^~"']
    for block, cell, action in WIDE_DELETIONS:
        body += [f'ws.Hyperlinks.Add ws.Range("{block}"), "http://example.com/{block}"', f'Set c = ws.Range("{cell}")',
                 *statements([action]), *(f'out = out & CellProps(ws.Range("{one}")) & "^"' for one in _cells(block)),
                 'out = out & "~"']
    body += ['out = out & Listing(ws)', 'wb.SaveAs Filename:=target & "deleted.xlsx"', "wb.Close False"]
    parts += function("Deleting", "target As String", body)
    # Add given arguments of the wrong kind, and what it returns.
    body = ["Dim h As Object", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            *statements(ARGUMENTS), 'out = out & "~"',
            *(f'out = out & CellProps(ws.Range("D{row}")) & "^"' for row in range(6, 17)), 'out = out & "~"',
            'Set h = ws.Hyperlinks.Add(ws.Range("F1:F2,H1"), "http://example.com/")',
            *reads(["TypeName(h)", "h.Range.Address",
                    'ws.Hyperlinks.Add(ws.Range("F5"), "http://example.com/").Range.Address']),
            'out = out & "~" & Listing(ws)', "wb.Close False"]
    parts += function("Arguments", "", body)
    # Links added over links.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", *OVERLAPS,
            *reads([f"ws.Hyperlinks({i}).Address" for i in range(1, 13)]),
            'wb.SaveAs Filename:=target & "overlaps.xlsx"', "wb.Close False"]
    parts += function("Overlaps", "target As String", body)
    # A link over several cells through inserts, deletes, copies, a cut and a copy of its sheet.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)"]
    for line in STRUCTURE:
        body += [*statements([line]), 'out = out & Listing(ws) & "~"'] if not line.startswith("ws.Hyperlinks.Add") \
            else [line]
    body += [*statements(["ws.Copy After:=ws"]), 'out = out & Listing(wb.Worksheets(2)) & "~"',
             *reads(["wb.Worksheets(2).Hyperlinks(1).Address", "wb.Worksheets(2).Hyperlinks.Count"]),
             'wb.SaveAs Filename:=target & "structure.xlsx"', "wb.Close False"]
    parts += function("Structure", "target As String", body)
    # Addresses of every kind: read, saved, opened and read again.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)"]
    body += statements([f'ws.Hyperlinks.Add ws.Range("A{row}"), "{_quoted(address)}"'
                        for row, address in enumerate(ADDRESSES, start=1)])
    body += statements([f'ws.Hyperlinks.Add ws.Range("A{row}"), {expression}'
                        for row, (_, expression) in enumerate(WIDE_ADDRESSES, start=len(ADDRESSES) + 1)])
    total = len(ADDRESSES) + len(WIDE_ADDRESSES)
    body += ['out = out & "~"', f"For i = 1 To {total}", 'out = out & CellProps(ws.Cells(i, 1)) & "^"', "Next",
             'out = out & "~"', 'wb.SaveAs Filename:=target & "canon.xlsx"', "wb.Close False",
             'Set wb = Workbooks.Open(target & "canon.xlsx")', "Set ws = wb.Worksheets(1)",
             f"For i = 1 To {total}", 'out = out & CellProps(ws.Cells(i, 1)) & "^"', "Next", "wb.Close False"]
    parts += function("Canon", "target As String", body)
    # A protected sheet's unlocked cells.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", 'ws.Range("A1:A6").Locked = False',
            *(f'ws.Hyperlinks.Add ws.Range("A{row}"), "http://example.com/A{row}"' for row in (3, 4, 5, 6)),
            *statements(UNLOCKED), 'out = out & "~" & Listing(ws)', "wb.Close False"]
    parts += function("Unlocked", "", body)
    # The HYPERLINK function entered each way.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            'ws.Range("A1").Value = "http://example.com/a1"']
    for cell, lines in FORMULAS:
        body += [f'Set c = ws.Range("{cell}")', *statements(lines), 'out = out & CellProps(c) & "^"']
    body += [*statements(['ws.Range("B8:B9").Formula = "=HYPERLINK(""http://example.com/"")"']),
             *(f'out = out & CellProps(ws.Range("{one}")) & "^"' for one in ("B8", "B9")),
             'wb.SaveAs Filename:=target & "formulas.xlsx"', "wb.Close False"]
    parts += function("Formulas", "target As String", body)
    # Links beside a note and a shape, for the ids their relationships get.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", *RELS,
            'wb.SaveAs Filename:=target & "rels.xlsx"', "wb.Close False"]
    parts += function("Rels", "target As String", body)
    # Links on merged cells.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            'ws.Range("A30:B30").Merge', 'ws.Range("A31:B31").Merge', 'ws.Range("A32:B32").Merge',
            *statements(['ws.Hyperlinks.Add ws.Range("A30"), "http://example.com/30"',
                         'ws.Hyperlinks.Add ws.Range("A31:B31"), "http://example.com/31"',
                         'ws.Hyperlinks.Add ws.Range("B32"), "http://example.com/32"']),
            'out = out & "~" & Listing(ws) & "~"',
            *(f'out = out & CellProps(ws.Range("{one}")) & "^"' for one in ("A30", "B30", "A31", "A32", "B32")),
            'wb.SaveAs Filename:=target & "merged.xlsx"', 'out = out & "~"',
            *statements(['ws.Range("A31").Hyperlinks(1).Delete']),
            *reads(['ws.Range("A31").MergeCells', 'ws.Range("A31").MergeArea.Address', 'ws.Range("A31").Style.Name',
                    "ws.Hyperlinks.Count"]),
            "wb.Close False"]
    parts += function("Merged", "target As String", body)
    # The place the Hyperlink style takes among the built-in ones, and the style asked for before any link.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", *statements(STYLE_ORDER),
            'wb.SaveAs Filename:=target & "style_order.xlsx"', "wb.Close False",
            "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", *statements(STYLE_ORDER_MADE),
            'wb.SaveAs Filename:=target & "style_order_made.xlsx"', "wb.Close False", 'out = out & "~"']
    for number, line in enumerate(STYLE_FIRST[:4]):
        body += ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", *statements([line]),
                 *statements(['ws.Hyperlinks.Add ws.Range("B3"), "http://example.com/"'] if number == 3 else []),
                 *reads(["wb.Styles.Count", 'wb.Styles("Hyperlink").BuiltIn', 'wb.Styles("Hyperlink").IncludeNumber',
                         'wb.Styles("Hyperlink").Font.Underline', 'ws.Range("B1").Style.Name',
                         'ws.Range("B2").Style.Name', 'ws.Range("B3").Style.Name', 'ws.Range("B3").Font.Underline']),
                 f'wb.SaveAs Filename:=target & "style_first{number}.xlsx"', "wb.Close False", 'out = out & "~"']
    parts += function("StyleOrder", "target As String", body)
    # Links copied and pasted every way, and copies landing on links.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)", *statements(PASTES),
            'out = out & "~" & Listing(ws) & "~"',
            *(f'out = out & CellProps(ws.Range("{one}")) & "^"' for one in ("C1", "C2", "C3", "C4", "E1", "E2", "E3",
                                                                            "F4")),
            'wb.SaveAs Filename:=target & "pastes.xlsx"', "wb.Close False"]
    parts += function("Pastes", "target As String", body)
    # Which links a range of two areas holds, what a clear and a sort do at the edges, and the whole collection's
    # Delete.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)"]
    for line in EDGES:
        body += [line] if line.startswith("out =") else statements([line])
    parts += function("Edges", "", [*body, "wb.Close False"])
    # EmailSubject set on each kind of link, and whether Excel takes a change after it.
    for name, lines in SUBJECTS.items():
        parts += function(name, "", ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
                                     'Set c = ws.Range("A1")', *statements(lines), 'out = out & CellProps(c) & "~"',
                                     *statements(['ws.Range("B1").Value = 1']), *reads(['ws.Range("B1").Value']),
                                     "wb.Close False"])
    # The order the collections list links in, and which a range's collection holds.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)"]
    for cell in ("C5", "A1", "B3", "A2", "E1"):
        body += [f'ws.Hyperlinks.Add ws.Range("{cell}"), "http://example.com/{cell}"']
    body += ['ws.Hyperlinks.Add ws.Range("A10:B11"), "http://example.com/spread"', 'out = out & Listing(ws) & "~"',
             *reads(["ws.Hyperlinks(1).Range.Address", "ws.Hyperlinks(6).Range.Address",
                     'ws.Range("A1:B3").Hyperlinks.Count', 'ws.Range("A1:B3").Hyperlinks(1).Range.Address',
                     'ws.Range("A1:B3").Hyperlinks(3).Range.Address', 'ws.Range("B11").Hyperlinks.Count',
                     'ws.Range("A1:A10").Hyperlinks.Count', 'ws.Range("B10:C11").Hyperlinks.Count',
                     'ws.Range("A1,C5").Hyperlinks.Count', 'ws.Range("D1:D9").Hyperlinks.Count',
                     'ws.Range("B11").Hyperlinks(1).Range.Address']),
             *statements(['ws.Range("B11").Hyperlinks.Delete']), *reads(["ws.Hyperlinks.Count"]),
             'out = out & "~" & Listing(ws)', "wb.Close False"]
    parts += function("Order", "", body)
    # Links moving with rows and columns, a copy, a cut and a sort.
    body = ["Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)"]
    for cell in ("A2", "B2", "C5", "G1", "G2", "G3"):
        body += [f'ws.Hyperlinks.Add ws.Range("{cell}"), "http://example.com/{cell}"']
    body += ['ws.Range("G1").Value = 2', 'ws.Range("G2").Value = 3', 'ws.Range("G3").Value = 1',
             *statements(["ws.Rows(1).Insert"]), 'out = out & Listing(ws) & "~"',
             *statements(['ws.Columns("B").Delete']), 'out = out & Listing(ws) & "~"',
             *statements(['ws.Range("A3").Copy ws.Range("D1")']), 'out = out & Listing(ws) & "~"',
             *statements(['ws.Range("B6").Cut ws.Range("D8")']), 'out = out & Listing(ws) & "~"',
             *statements(['ws.Range("F2:F4").Sort Key1:=ws.Range("F2"), Order1:=1, Header:=2']),
             'out = out & Listing(ws) & "~"',
             *(f'out = out & CellProps(ws.Range("{cell}")) & "^"' for cell in ("F2", "F3", "F4", "D1", "D8")),
             *statements(['ws.Range("A3").Delete -4162']), 'out = out & "~" & Listing(ws) & "~"',
             'wb.SaveAs Filename:=target & "moved.xlsx"', "wb.Close False"]
    parts += function("Moving", "target As String", body)
    # The HYPERLINK function, a shape's link, and the errors.
    body = ["Dim shp As Object", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            'ws.Range("A1").Formula = "=HYPERLINK(""http://example.com/"",""go"")"',
            'out = out & CellProps(ws.Range("A1")) & "~"', *reads(["ws.Hyperlinks.Count"]), 'out = out & "~"',
            "Set shp = ws.Shapes.AddShape(1, 100, 100, 50, 20)",
            *statements(['ws.Hyperlinks.Add shp, "http://example.com/shape"']),
            *reads(["ws.Hyperlinks.Count", "ws.Hyperlinks(1).Type", "ws.Hyperlinks(1).Address",
                    "ws.Hyperlinks(1).Shape.Name", "ws.Hyperlinks(1).Range.Address", "shp.Hyperlink.Address",
                    "ws.Hyperlinks(1).TextToDisplay", "ws.Hyperlinks(1).Name"]),
            "wb.Close False"]
    parts += function("Other", "", body)
    body = ["Dim ws2 As Object", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Set ws = wb.Worksheets(1)",
            "Set ws2 = wb.Worksheets.Add(After:=ws)", 'ws.Hyperlinks.Add ws.Range("A1"), "http://example.com/"',
            'ws.Hyperlinks.Add ws.Range("B1"), "http://example.com/B1"', *statements(ERROR_STATEMENTS),
            'out = out & "~"', *reads(ERRORS), 'out = out & "~" & Listing(ws) & "~" & Listing(ws2) & "~"',
            *statements(["ws.Protect", 'ws.Hyperlinks.Add ws.Range("E1"), "http://example.com/"',
                         'ws.Range("A1").Hyperlinks(1).Address = "http://x.example/"',
                         'ws.Range("B1").Hyperlinks.Delete', "ws.Unprotect",
                         "ws.Protect AllowInsertingHyperlinks:=True",
                         'ws.Hyperlinks.Add ws.Range("E2"), "http://example.com/"', "ws.Unprotect"]),
            'out = out & "~" & Listing(ws)', "wb.Close False"]
    parts += function("Errors", "", body)
    return "\n".join(parts) + "\n"


#: Each section and whether it takes the folder to save in.
SECTIONS = {"Adds": True, "Removing": True, "Setting": True, "Order": False, "Moving": True, "Other": False,
            "Errors": False, "Reopen": True, "Ranges": False, "Deleting": True, "Arguments": False, "Overlaps": True,
            "Structure": True, "Canon": True, "Unlocked": False, "Formulas": True, "Rels": True, "Merged": True,
            "StyleOrder": True, "Pastes": True, "Edges": False, **dict.fromkeys(SUBJECTS, False)}


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    saved = str(folder) + "\\"
    code = module()
    runs: dict[str, str] = {}
    for name, saves in SECTIONS.items():
        # Each section in an Excel of its own: once a link's EmailSubject is set, Excel refuses what comes after.
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(code, name, args=(saved,) if saves else (), timeout=900.0)
        if name in SUBJECTS and not result.ok:
            runs[name] = f"!{result.outcome}: {result.message}"
            continue
        assert result.ok, f"{name}: {result.outcome}: {result.message} {result.error}"
        runs[name] = str(result.value)
    OUT.mkdir(parents=True, exist_ok=True)
    for path in folder.glob("*.xlsx"):
        copy_saved(path, OUT / path.name)
    record = {"module": code, "sections": SECTIONS, "link_reads": LINK_READS, "cell_reads": CELL_READS,
              "format_reads": FORMAT_READS, "adds": ADDS, "removals": REMOVALS, "sets": SETS, "mail_sets": MAIL_SETS,
              "settings": SETTINGS, "subjects": SUBJECTS, "range_counts": RANGE_COUNTS, "deletions": DELETIONS,
              "wide_deletions": WIDE_DELETIONS, "arguments": ARGUMENTS, "addresses": ADDRESSES,
              "wide_addresses": WIDE_ADDRESSES, "formulas": FORMULAS, "errors": ERRORS,
              "error_statements": ERROR_STATEMENTS, "runs": runs}
    (OUT / "hyperlinks.json").write_text(json.dumps(record, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    for name, answer in runs.items():
        print(f"{name}: {len(answer)} characters")


if __name__ == "__main__":
    main()
