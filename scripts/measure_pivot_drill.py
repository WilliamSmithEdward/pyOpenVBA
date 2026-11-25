"""Range.ShowDetail on a pivot table's value: the sheet of records Excel shows for it (drill-through). First look.

Each case makes a workbook with one sheet holding a list in A1:C5, a pivot table over it, and runs its statements with
a sink class holding the workbook WithEvents. It records the events, r, and for every sheet but the source: its name and
place, the active cell and selection, each table's name, range, style and AutoFilter, every cell's value, type and
number format, the widths of columns A to D, and the sheet's StandardHeight, since widths depend on the display's DPI.
Each case is saved in a folder of TEMP and its table and sheet parts kept. No row label is ever double-clicked: that
opens a dialog.

Predictions, written before measuring:
- D1: ShowDetail = True on a value puts a new sheet before the active one, named as Worksheets.Add names one, and
  activates it. The sheet holds a table at A1 named Table1 in the workbook's default style, TableStyleMedium2. Its
  headers are the cache fields' names in the source's order, and its rows are the records behind the value, in the
  source's order. The table's whole range is selected. Rival D2: the cells without a table. Predict D1.
- F1: each column takes the number format of its source column's first record. Rival F2: General. Predict F1.
- W1: the columns are fitted to their cells, header and filter button counted. Recorded at this DPI, not predicted.
- E1: ShowDetail = False on a value, or read on one, fails with 1004. Predict E1.
- Z1: with EnableDrilldown off, ShowDetail = True on a value fails with 1004 and makes no sheet. Predict Z1.
- H1: an item hidden by a filter, or a page field's choice, keeps its records out of the detail. Predict H1.

What look 1 (cases 0 to 15) showed:
- D1 held: a new sheet before the active one, Table1 at A1 in TableStyleMedium2 with the cache fields' names, and
  its whole range selected. Events: SelectionChange on the value's cell, NewSheet, Deactivate, Activate, then
  SelectionChange on the table. The columns keep the standard width. A second detail is Table2 on Sheet2.
- F1 was wrong. Every cell is General: the source's "$#,##0.00" and "@" are not carried.
- E1 was half right. Reading ShowDetail fails with 1004, "Unable to get the ShowDetail property of the Range
  class", but setting it False does nothing at all.
- Z1 held, failing with 1004, "Unable to set the ShowDetail property of the Range class".
- H1 held for a page field. Case 14 aimed at the wrong cell, outside the table, and failed with 1004.
- Several cells take the first one's detail. DoubleClick on a value is ShowDetail.
- The records' order is not the source's. A grand total listed a,x / a,y / b,y / b,x with rows k (source order
  inside each item), but b,x before b,y with a second data field. Look 2 (cases 16 to 25), written before it ran,
  splits the rival orders over eight records:
  - O1: grouped by the row fields' items, then the column fields', as the table shows them, in source order within.
  - O2: as O1, in reverse source order within.
  - O3: reverse source order overall.
  Predict O1.

What look 2 showed: O1 held for every grand total. Records are grouped by the row fields' items, then the column
fields', in the order the table shows them (a descending sort included), in source order within. One item's detail
did not follow it: a's records (sources 2 x, 4 y, 7 x) came 7, 2, 4. Look 3 (cases 26 to 31), written before it
ran, splits the rival orders for one item's detail:
- S1: grouped by the fields not on the table's axes, in their items' order (x, y, z), in reverse source order within.
- S2: as S1, in source order within.
- S3: reverse source order overall.
Predict S1.

What look 3 showed: S1, S2 and S3 all wrong. a's detail with k on the rows came in source order this time (case 26),
as it did with a page field (29) and with no field on an axis (30). The grand total held O1 again (31). But (a, x)
came 6, 2, 4 (27), and x with g on the rows came 7, 2, 6, 4 (28). Over looks 1 to 3 the scrambled lists are 2, 4, 7
as 7, 2, 4; 2, 4, 6 as 6, 2, 4; 2, 4, 6, 7 as 7, 2, 6, 4; and 3, 5 as 5, 3. Source order held where the item drilled
was the first record's. No rule fitted to these four predicts the rest, so a part's detail is not built: the order
of its records is not worked out.

Held-out round A (cases 32 to 36, the grand total), predicted on 89a6e51d (sha256 07e9fa9c7b27c572): 4 of 5. The miss
had two fields on the rows, k outside g: Excel grouped the records by g first, then k. Within an axis the innermost
field now leads. Held-out round B (cases 37 to 41) tests that on new layouts, predicted by the model before it ran.

Round B: 5 of 5. Look 4 (cases 42 to 51), written before it ran, settles one item's order where nothing else groups
the records: k on the rows, g the same throughout, v each record's number, a's own value drilled. The rivals, over
a's source places p1 < ... < pn:
- R1: source order (the item first in the source keeps it).
- R2: the last first, then the rest in source order (7, 2, 4 from 2, 4, 7).
- R3: a max-heap's array after inserting them in source order.
- R4: the order depends on whether a is the first record's item: source order if so, R2 if not.
Predict R4.

What look 4 showed: R1 to R4 all wrong. Every order, these ten and the scrambled ones before, is what a removal
by swapping leaves: Excel goes through the records from the front, and each one the value does not count it
replaces with the last record, shortening the list, then looks again at that place. b a b a b a b a gives 8, 2, 6, 4;
the earlier 2, 4, 7 from b a c a b c a b gives 7, 2, 4. What is left is grouped as the grand total's records are.
Held-out rounds C (52 to 56) and D (57 to 61) tested that on ten shuffled records, predicted by the model on 2c727c73
(sha256 debdee64fac19104) before they ran: 10 of 10.

Look 5 (cases 62 to 71) measures what the model refuses: hidden items, a page field showing several items, label and
value filters, a collapsed item, a grouped number field. Predictions, written before it ran:
- H1: a hidden item's records, a page field's left-out items' and those a label or value filter leaves out are
  taken away as the value's other records are, by swapping, so the rest keep the order that leaves. Rival H2:
  they are left out first, the order of what is left then made as if they had never been there. H1 and H2 can
  give one order; case 63 (b's value with c hidden) and 71 are where they may part.
- C1: a collapsed item changes nothing about which records its value stands for, and the records still group by
  the field inside it. Rival C2: with the inner field's lines gone, they group by the outer field alone. Predict C1.
- G1: a field grouped by number ranges groups its records by range, in the ranges' order, and the detail's
  columns are the source's, with no column for the groups. Predict G1.

What look 5 showed: H1 held for hidden items, a page field's left-out items and a label filter. A value filter
leaves out nothing: the grand total's detail under one held all ten records, the hidden c's among them, in source
order. C1 held, and so did G1; the second range's value gave 10, 9, 8, 7, 6, the removal by swapping again. Held-out
rounds E (72 to 76) and F (77 to 81) test these on twelve new records, predicted by the model before they ran.

Rounds E and F, predicted on 1f15c729 (sha256 b668f9c3ee96d0a3): 9 of 10. The miss, a top-2 value filter's grand
total, grouped every item in the field's own order, a, b, c, d, the two the filter hides among them, where the model
ranked the items by the lines the table shows. A value filter's field now ranks by its items' order. Held-out rounds
G (82 to 86) and H (87 to 91) tested it on eleven new records, predicted by the model on 629eecb4 (sha256
a11de309d46c3bbc) before they ran: 10 of 10.

    python scripts/measure_pivot_drill.py

writes tests/fixtures/pivot_drill.json.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivot_drill.json'
SINK = '''Public WithEvents Book As Workbook

Private Sub Book_NewSheet(ByVal Sh As Object)
    Log = Log & "[NewSheet " & Sh.Name & "]"
End Sub

Private Sub Book_SheetActivate(ByVal Sh As Object)
    Log = Log & "[Activate " & Sh.Name & "]"
End Sub

Private Sub Book_SheetDeactivate(ByVal Sh As Object)
    Log = Log & "[Deactivate " & Sh.Name & "]"
End Sub

Private Sub Book_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Log = Log & "[Change " & Sh.Name & " " & Target.Address(False, False) & "]"
End Sub

Private Sub Book_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Log = Log & "[SelectionChange " & Sh.Name & " " & Target.Address(False, False) & "]"
End Sub

Private Sub Book_SheetPivotTableUpdate(ByVal Sh As Object, ByVal Target As PivotTable)
    Log = Log & "[Update " & Sh.Name & " " & Target.Name & "]"
End Sub
'''
LOG = 'Public Log As String\n'
ROWS = 'Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:C5")).CreatePivotTable(ws.Range("E1"), "P"): '
K_ROWS = ROWS + 'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
K_G = (ROWS + 'pt.PivotFields("k").Orientation = xlRowField: pt.PivotFields("g").Orientation = xlColumnField: '
       'pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ')
#: GO turns events on, after a case has set its sheet up with them off.
GO = 'Application.EnableEvents = True: '
def ordered(keys: str) -> str:
    """Look 4's case: records whose k are ``keys``, g x throughout and v each record's number, a pivot table with k on
    the rows, and a's own value drilled."""
    count = len(keys)
    values = ';'.join(f'""{key}"",""x"",{number}' for number, key in enumerate(keys, start=1))
    return (f'ws.Range("A1:C5").ClearContents: ws.Range("A1:C1").Value = Array("k", "g", "v"): '
            f'ws.Range("A2:C{count + 1}").Value = ws.Evaluate("{{{values}}}"): '
            f'Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:C{count + 1}")).CreatePivotTable(ws.Range("E1"), '
            f'"P"): pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", '
            f'xlSum: {GO}r = "{keys}": pt.PivotFields("k").PivotItems("a").DataRange.Cells(1).ShowDetail = True')


#: The table's last value, its grand total's.
TOTAL = 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True'
#: Eleven records shuffled, for held-out rounds G and H.
ELEVEN = ('ws.Range("A2:C12").Value = ws.Evaluate("{""b"",""y"",1;""c"",""x"",2;""a"",""x"",3;""b"",""x"",4;'
          '""a"",""y"",5;""c"",""y"",6;""b"",""x"",7;""a"",""x"",8;""c"",""x"",9;""b"",""y"",10;""a"",""y"",11}"): ')
#: Twelve records shuffled, four items, for held-out rounds E and F.
TWELVE = ('ws.Range("A2:C13").Value = ws.Evaluate("{""d"",""x"",1;""a"",""y"",2;""b"",""x"",3;""c"",""y"",4;'
          '""a"",""x"",5;""d"",""y"",6;""b"",""y"",7;""c"",""x"",8;""a"",""x"",9;""b"",""x"",10;""d"",""y"",11;'
          '""c"",""y"",12}"): ')
#: Ten records shuffled, for held-out rounds C and D.
SHUF = ('ws.Range("A2:C11").Value = ws.Evaluate("{""c"",""y"",1;""a"",""x"",2;""b"",""x"",3;""a"",""y"",4;'
        '""c"",""x"",5;""b"",""y"",6;""a"",""x"",7;""c"",""y"",8;""b"",""x"",9;""a"",""y"",10}"): ')
#: Nine records, seven of them a's, for look 3.
MANY = ('ws.Range("A2:C10").Value = ws.Evaluate("{""a"",""y"",1;""a"",""x"",2;""a"",""z"",3;""a"",""x"",4;'
        '""a"",""y"",5;""a"",""x"",6;""b"",""x"",7;""a"",""z"",8;""b"",""y"",9}"): ')
#: Eight records in no order, for look 2.
LONG = ('ws.Range("A2:C9").Value = ws.Evaluate("{""b"",""y"",1;""a"",""x"",2;""c"",""y"",3;""a"",""y"",4;'
        '""b"",""x"",5;""c"",""x"",6;""a"",""x"",7;""b"",""y"",8}"): ')
CASES = [
    K_ROWS + GO + 'ws.Range("F2").ShowDetail = True',
    K_ROWS + GO + 'ws.Range("F4").ShowDetail = True',
    K_G + GO + 'ws.Range("F3").ShowDetail = True',
    K_G + GO + 'ws.Range("H3").ShowDetail = True',
    K_G + GO + 'ws.Range("G5").ShowDetail = True',
    K_G + GO + 'ws.Range("H5").ShowDetail = True',
    K_ROWS + 'pt.AddDataField pt.PivotFields("v"), "Count of v", xlCount: ' + GO + 'ws.Range("G3").ShowDetail = True',
    ROWS + 'pt.PivotFields("g").Orientation = xlPageField: pt.PivotFields("g").CurrentPage = "x": '
    'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
    + GO + 'r = pt.DataBodyRange.Address(False, False): pt.DataBodyRange.Cells(1).ShowDetail = True',
    K_ROWS + GO + 'ws.Range("F2").ShowDetail = True: ws.Activate: ws.Range("F3").ShowDetail = True',
    K_ROWS + GO + 'ws.Range("F2").Select: Application.DoubleClick',
    'ws.Range("C2:C5").NumberFormat = "$#,##0.00": ws.Range("B2:B5").NumberFormat = "@": ' + K_ROWS + GO
    + 'ws.Range("F2").ShowDetail = True',
    K_ROWS + GO + 'ws.Range("F2").ShowDetail = False',
    K_ROWS + GO + 'r = ws.Range("F2").ShowDetail',
    K_ROWS + 'pt.EnableDrilldown = False: ' + GO + 'ws.Range("F2").ShowDetail = True',
    K_ROWS + 'pt.PivotFields("g").Orientation = xlRowField: pt.PivotFields("g").PivotItems("y").Visible = False: '
    + GO + 'r = pt.TableRange1.Address(False, False): ws.Range("G2").ShowDetail = True',
    K_ROWS + GO + 'ws.Range("F2:F3").ShowDetail = True',
    # Look 2: the order of the records, over eight of them; LONG replaces the list first.
    LONG + K_ROWS.replace('A1:C5', 'A1:C9') + GO + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + K_ROWS.replace('A1:C5', 'A1:C9') + 'pt.AddDataField pt.PivotFields("v"), "Count of v", xlCount: ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + ROWS.replace('A1:C5', 'A1:C9') + 'pt.PivotFields("g").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + K_G.replace('A1:C5', 'A1:C9') + GO + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + K_ROWS.replace('A1:C5', 'A1:C9') + 'pt.PivotFields("k").AutoSort xlDescending, "k": ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + K_ROWS.replace('A1:C5', 'A1:C9') + GO + 'pt.DataBodyRange.Cells(1).ShowDetail = True',
    LONG + K_ROWS.replace('A1:C5', 'A1:C9') + 'pt.AddDataField pt.PivotFields("v"), "Count of v", xlCount: ' + GO
    + 'pt.DataBodyRange.Cells(2).ShowDetail = True',
    LONG + ROWS.replace('A1:C5', 'A1:C9') + 'pt.PivotFields("k").Orientation = xlColumnField: '
    'pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + ROWS.replace('A1:C5', 'A1:C9') + 'pt.PivotFields("v").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("k"), "Count of k", xlCount: ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + K_ROWS.replace('A1:C5', 'A1:C9') + GO + 'ws.Range("F2").ShowDetail = True: ws.Activate: '
    'ws.Range("F2").ShowDetail = True',
    # Look 3: one item's detail, over nine records seven of which are a's; MANY replaces the list first.
    MANY + K_ROWS.replace('A1:C5', 'A1:C10') + GO + 'pt.DataBodyRange.Cells(1).ShowDetail = True',
    MANY + K_G.replace('A1:C5', 'A1:C10') + GO + 'pt.DataBodyRange.Cells(1).ShowDetail = True',
    MANY + ROWS.replace('A1:C5', 'A1:C10') + 'pt.PivotFields("g").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO + 'pt.DataBodyRange.Cells(1).ShowDetail = True',
    MANY + K_ROWS.replace('A1:C5', 'A1:C10') + 'pt.PivotFields("g").Orientation = xlPageField: ' + GO
    + 'pt.DataBodyRange.Cells(1).ShowDetail = True',
    MANY + ROWS.replace('A1:C5', 'A1:C10') + 'pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.DataBodyRange.Cells(1).ShowDetail = True',
    MANY + K_ROWS.replace('A1:C5', 'A1:C10') + GO + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    # Held-out round A, the grand total's detail, predicted by the model before it ran.
    LONG + ROWS.replace('A1:C5', 'A1:C9') + 'pt.PivotFields("g").Orientation = xlRowField: '
    'pt.PivotFields("k").Orientation = xlColumnField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    MANY + ROWS.replace('A1:C5', 'A1:C10') + 'pt.PivotFields("k").Orientation = xlRowField: '
    'pt.PivotFields("g").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + K_ROWS.replace('A1:C5', 'A1:C9') + 'pt.PivotFields("g").Orientation = xlPageField: ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + ROWS.replace('A1:C5', 'A1:C9') + 'pt.AddDataField pt.PivotFields("v"), "Count of v", xlCount: ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + K_G.replace('A1:C5', 'A1:C9') + 'pt.PivotFields("k").AutoSort xlDescending, "k": '
    'pt.PivotFields("g").AutoSort xlDescending, "g": ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).Select: Application.DoubleClick',
    # Held-out round B, the order of two fields on one axis, predicted by the model before it ran.
    MANY + ROWS.replace('A1:C5', 'A1:C10') + 'pt.PivotFields("g").Orientation = xlRowField: '
    'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + ROWS.replace('A1:C5', 'A1:C9') + 'pt.PivotFields("k").Orientation = xlColumnField: '
    'pt.PivotFields("g").Orientation = xlColumnField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + ROWS.replace('A1:C5', 'A1:C9') + 'pt.PivotFields("k").Orientation = xlRowField: '
    'pt.PivotFields("g").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
    'pt.PivotFields("k").AutoSort xlDescending, "k": ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    LONG + ROWS.replace('A1:C5', 'A1:C9') + 'pt.PivotFields("g").Orientation = xlRowField: '
    'pt.PivotFields("k").Orientation = xlColumnField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
    'pt.PivotFields("k").AutoSort xlDescending, "k": ' + GO
    + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    MANY + ROWS.replace('A1:C5', 'A1:C10') + 'pt.PivotFields("k").Orientation = xlRowField: '
    'pt.PivotFields("g").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
    'pt.RowAxisLayout xlTabularRow: ' + GO + 'pt.DataBodyRange.Cells(pt.DataBodyRange.Count).ShowDetail = True',
    # Look 4: one item's detail where nothing else groups the records (g the same throughout, v the record's number),
    # a's place among them varied.
    *(ordered(keys) for keys in ('baaaa', 'babababa', 'abab', 'bbaaaaaa', 'ba', 'baa', 'baaa', 'cbaaaa', 'abba',
                                 'bcabaca')),
    # Held-out round C, a part's detail, predicted by the model before it ran.
    SHUF + K_ROWS.replace('A1:C5', 'A1:C11') + GO + 'pt.PivotFields("k").PivotItems("b").DataRange.Cells(1).ShowDetail = True',
    SHUF + K_G.replace('A1:C5', 'A1:C11') + GO + 'Application.Intersect(pt.PivotFields("k").PivotItems("c").DataRange, '
    'pt.PivotFields("g").PivotItems("y").DataRange).ShowDetail = True',
    SHUF + ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("g").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.PivotFields("g").PivotItems("y").DataRange.Cells(1).ShowDetail = True',
    SHUF + ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("k").Orientation = xlColumnField: '
    'pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.PivotFields("k").PivotItems("a").DataRange.Cells(1).ShowDetail = True',
    SHUF + ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("g").Orientation = xlPageField: '
    'pt.PivotFields("g").CurrentPage = "y": pt.PivotFields("k").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.PivotFields("k").PivotItems("a").DataRange.Cells(1).ShowDetail = True',
    # Held-out round D, a part's detail, predicted by the model before it ran.
    SHUF + ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("g").Orientation = xlRowField: '
    'pt.PivotFields("k").Orientation = xlColumnField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'Application.Intersect(pt.PivotFields("g").PivotItems("x").DataRange, '
    'pt.PivotFields("k").PivotItems("b").DataRange).ShowDetail = True',
    SHUF + K_ROWS.replace('A1:C5', 'A1:C11') + GO + 'pt.PivotFields("k").PivotItems("c").DataRange.Cells(1).Select: '
    'Application.DoubleClick',
    SHUF + ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("g").Orientation = xlRowField: '
    'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.PivotFields("k").PivotItems("a").DataRange.Cells(2).ShowDetail = True',
    SHUF + K_ROWS.replace('A1:C5', 'A1:C11') + 'pt.AddDataField pt.PivotFields("v"), "Count of v", xlCount: ' + GO
    + 'pt.PivotFields("k").PivotItems("c").DataRange.Cells(2).ShowDetail = True',
    SHUF + K_ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("k").AutoSort xlDescending, "k": ' + GO
    + 'pt.PivotFields("k").PivotItems("a").DataRange.Cells(1).ShowDetail = True',
    # Look 5: what the model refuses, a table's filters, hidden and collapsed items and a grouped field.
    SHUF + K_ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("k").PivotItems("c").Visible = False: ' + GO + TOTAL,
    SHUF + K_ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("k").PivotItems("c").Visible = False: ' + GO
    + 'pt.PivotFields("k").PivotItems("b").DataRange.Cells(1).ShowDetail = True',
    SHUF + ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("g").Orientation = xlPageField: '
    'pt.PivotFields("g").EnableMultiplePageItems = True: pt.PivotFields("g").PivotItems("y").Visible = False: '
    'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + TOTAL,
    SHUF + K_ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("k").PivotFilters.Add Type:=xlCaptionEquals, '
    'Value1:="a": ' + GO + TOTAL,
    SHUF + K_ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("k").PivotFilters.Add Type:=xlValueIsGreaterThan, '
    'DataField:=pt.DataFields(1), Value1:=15: ' + GO + TOTAL,
    SHUF + ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("k").Orientation = xlRowField: '
    'pt.PivotFields("g").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
    'pt.PivotFields("k").PivotItems("a").ShowDetail = False: ' + GO + TOTAL,
    SHUF + ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("k").Orientation = xlRowField: '
    'pt.PivotFields("g").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
    'pt.PivotFields("k").PivotItems("a").ShowDetail = False: ' + GO
    + 'pt.PivotFields("k").PivotItems("a").DataRange.Cells(1).ShowDetail = True',
    SHUF + ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("v").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("k"), "Count of k", xlCount: '
    'pt.PivotFields("v").DataRange.Cells(1).Group Start:=1, End:=10, By:=5: ' + GO + TOTAL,
    SHUF + ROWS.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("v").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("k"), "Count of k", xlCount: '
    'pt.PivotFields("v").DataRange.Cells(1).Group Start:=1, End:=10, By:=5: ' + GO
    + 'pt.DataBodyRange.Cells(2).ShowDetail = True',
    SHUF + K_G.replace('A1:C5', 'A1:C11') + 'pt.PivotFields("g").PivotItems("y").Visible = False: ' + GO
    + 'pt.PivotFields("k").PivotItems("a").DataRange.Cells(pt.PivotFields("k").PivotItems("a").DataRange.Count)'
    '.Offset(0, 1).ShowDetail = True',
    # Held-out round E, filters and the rest beside a part, predicted by the model before it ran.
    TWELVE + K_ROWS.replace('A1:C5', 'A1:C13') + 'pt.PivotFields("k").PivotItems("b").Visible = False: ' + GO
    + 'pt.PivotFields("k").PivotItems("a").DataRange.Cells(1).ShowDetail = True',
    TWELVE + K_ROWS.replace('A1:C5', 'A1:C13') + 'pt.PivotFields("k").PivotFilters.Add Type:=xlCaptionDoesNotEqual, '
    'Value1:="c": ' + GO + TOTAL,
    TWELVE + K_ROWS.replace('A1:C5', 'A1:C13') + 'pt.PivotFields("k").PivotFilters.Add Type:=xlValueIsGreaterThan, '
    'DataField:=pt.DataFields(1), Value1:=17: ' + GO
    + 'pt.PivotFields("k").PivotItems("b").DataRange.Cells(1).ShowDetail = True',
    TWELVE + ROWS.replace('A1:C5', 'A1:C13') + 'pt.PivotFields("k").Orientation = xlRowField: '
    'pt.PivotFields("g").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
    'pt.PivotFields("k").PivotItems("b").ShowDetail = False: ' + GO
    + 'pt.PivotFields("k").PivotItems("b").DataRange.Cells(1).ShowDetail = True',
    TWELVE + ROWS.replace('A1:C5', 'A1:C13') + 'pt.PivotFields("v").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("k"), "Count of k", xlCount: '
    'pt.PivotFields("v").DataRange.Cells(1).Group Start:=1, End:=12, By:=4: ' + GO
    + 'pt.DataBodyRange.Cells(3).ShowDetail = True',
    # Held-out round F, predicted by the model before it ran.
    TWELVE + ROWS.replace('A1:C5', 'A1:C13') + 'pt.PivotFields("g").Orientation = xlPageField: '
    'pt.PivotFields("g").EnableMultiplePageItems = True: pt.PivotFields("g").PivotItems("x").Visible = False: '
    'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ' + GO
    + 'pt.PivotFields("k").PivotItems("c").DataRange.Cells(1).ShowDetail = True',
    TWELVE + K_G.replace('A1:C5', 'A1:C13') + 'pt.PivotFields("k").PivotItems("c").Visible = False: ' + GO
    + 'Application.Intersect(pt.PivotFields("k").PivotItems("a").DataRange, '
    'pt.PivotFields("g").PivotItems("y").DataRange).ShowDetail = True',
    TWELVE + K_ROWS.replace('A1:C5', 'A1:C13') + 'pt.PivotFields("k").PivotFilters.Add Type:=xlCaptionEquals, '
    'Value1:="d": ' + GO + 'pt.DataBodyRange.Cells(1).Select: Application.DoubleClick',
    TWELVE + K_ROWS.replace('A1:C5', 'A1:C13') + 'pt.PivotFields("k").PivotFilters.Add Type:=xlTopCount, '
    'DataField:=pt.DataFields(1), Value1:=2: ' + GO + TOTAL,
    TWELVE + ROWS.replace('A1:C5', 'A1:C13') + 'pt.PivotFields("g").Orientation = xlRowField: '
    'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
    'pt.PivotFields("g").PivotItems("x").ShowDetail = False: ' + GO + TOTAL,
    # Held-out round G, predicted by the model before it ran.
    ELEVEN + K_ROWS.replace('A1:C5', 'A1:C12') + 'pt.PivotFields("k").PivotFilters.Add Type:=xlBottomCount, '
    'DataField:=pt.DataFields(1), Value1:=1: ' + GO + TOTAL,
    ELEVEN + K_ROWS.replace('A1:C5', 'A1:C12') + 'pt.PivotFields("k").PivotFilters.Add Type:=xlValueIsLessThan, '
    'DataField:=pt.DataFields(1), Value1:=20: ' + GO + 'pt.DataBodyRange.Cells(1).ShowDetail = True',
    ELEVEN + K_G.replace('A1:C5', 'A1:C12') + 'pt.PivotFields("g").PivotFilters.Add Type:=xlCaptionEquals, '
    'Value1:="y": ' + GO + TOTAL,
    ELEVEN + ROWS.replace('A1:C5', 'A1:C12') + 'pt.PivotFields("k").Orientation = xlRowField: '
    'pt.PivotFields("g").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
    'pt.PivotFields("k").PivotItems("a").ShowDetail = False: pt.PivotFields("g").PivotItems("x").Visible = False: '
    + GO + TOTAL,
    ELEVEN + ROWS.replace('A1:C5', 'A1:C12') + 'pt.PivotFields("v").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("k"), "Count of k", xlCount: '
    'pt.PivotFields("v").DataRange.Cells(1).Group Start:=1, End:=11, By:=3: ' + GO
    + 'pt.DataBodyRange.Cells(2).ShowDetail = True',
    # Held-out round H, predicted by the model before it ran.
    ELEVEN + ROWS.replace('A1:C5', 'A1:C12') + 'pt.PivotFields("g").Orientation = xlPageField: '
    'pt.PivotFields("g").CurrentPage = "x": pt.PivotFields("k").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: pt.PivotFields("k").PivotItems("a").Visible = False: '
    + GO + TOTAL,
    ELEVEN + ROWS.replace('A1:C5', 'A1:C12') + 'pt.PivotFields("g").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: pt.PivotFields("g").PivotFilters.Add Type:=xlTopCount, '
    'DataField:=pt.DataFields(1), Value1:=1: ' + GO + TOTAL,
    ELEVEN + K_ROWS.replace('A1:C5', 'A1:C12') + 'pt.PivotFields("k").PivotFilters.Add Type:=xlValueIsGreaterThan, '
    'DataField:=pt.DataFields(1), Value1:=15: ' + GO + 'pt.DataBodyRange.Cells(1).Select: Application.DoubleClick',
    ELEVEN + ROWS.replace('A1:C5', 'A1:C12') + 'pt.PivotFields("k").Orientation = xlRowField: '
    'pt.PivotFields("g").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
    'pt.PivotFields("k").PivotItems("a").ShowDetail = False: ' + GO
    + 'pt.PivotFields("k").PivotItems("b").DataRange.Cells(1).ShowDetail = True',
    ELEVEN + K_G.replace('A1:C5', 'A1:C12') + 'pt.PivotFields("g").PivotItems("y").Visible = False: ' + GO + TOTAL,
]
PROBE = '''Private Function Digest(wb As Workbook) As String
    Dim s As String, sh As Worksheet, lo As ListObject, cell As Range, c As Long
    On Error Resume Next
    s = "active " & ActiveSheet.Name & "!" & ActiveCell.Address(False, False) & " sel " & Selection.Address(False, False)
    For Each sh In wb.Worksheets
        s = s & " <" & sh.Index & " " & sh.Name
        If sh.Name <> "Data" Then
            s = s & " sh " & sh.StandardHeight & " widths"
            For c = 1 To 4
                s = s & " " & sh.Columns(c).ColumnWidth
            Next
            For Each lo In sh.ListObjects
                s = s & " {" & lo.Name & " " & lo.Range.Address(False, False) & " " & lo.TableStyle & " " & lo.ShowAutoFilter
                s = s & " " & lo.ShowTotals & " " & TypeName(lo.QueryTable) & "}"
                Err.Clear
            Next
            For Each cell In sh.Range("A1:D14")
                If Not IsEmpty(cell.Value) Then s = s & " " & cell.Address(False, False) & "=" & TypeName(cell.Value) & ":" & cell.Text & ":" & cell.NumberFormat
            Next
        End If
        s = s & ">"
    Next
    Digest = s
End Function

Function Probe(ByVal folder As String, ByVal which As Long) As String
    Dim sink As New ExcelSink, wb As Workbook, ws As Worksheet, pt As PivotTable, r As Variant, failure As String
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"
    ws.Range("A1:C5").Value = ws.Evaluate("{""k"",""g"",""v"";""a"",""x"",1;""b"",""y"",2;""a"",""y"",3;""b"",""x"",4}")
    ws.Activate
    ws.Range("A1").Select
    Set sink.Book = wb
    Log = ""
    On Error GoTo Failed
    Select Case which
{cases}
    End Select
Answered:
    On Error GoTo 0
    Application.EnableEvents = False
    Probe = Log & "|" & CStr(r) & failure & "|" & Digest(wb)
    On Error Resume Next
    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook
    wb.Close False
    Application.EnableEvents = True
    Exit Function
Failed:
    failure = "!E" & Err.Number & " " & Err.Description
    Resume Answered
End Function
'''


def probe() -> str:
    cases = '\n'.join(f'    Case {index}\n        {statement}' for index, statement in enumerate(CASES))
    return PROBE.replace('{cases}', cases)


def install() -> str:
    lines = ['Function Install() As String', 'Dim text As String, part As Object',
             'Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)', 'part.Name = "ExcelSink"', 'text = ""']
    lines += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in SINK.splitlines()]
    lines += ['part.CodeModule.AddFromString text', 'Install = "ok"', 'End Function']
    return '\n'.join(lines)


def parts(path: Path) -> dict[str, str]:
    """The table parts and sheets a saved case holds, with the folder the workbook was saved in, which names the user,
    taken out."""
    if not path.exists():
        return {}
    keep = re.compile(r'xl/tables/[^/]*\.xml|xl/worksheets/sheet\d+\.xml|xl/workbook\.xml')
    with zipfile.ZipFile(path) as package:
        return {name: scrubbed(package.read(name).decode('utf-8')) for name in package.namelist() if keep.fullmatch(name)}


def scrubbed(text: str) -> str:
    return re.sub(r'(absPath url=")[^"]*"', r'\1&lt;folder&gt;"', re.sub(r'refreshedBy="[^"]*"', 'refreshedBy="&lt;user&gt;"', text))


def main() -> None:
    chosen = [int(one) for one in sys.argv[1:]] or list(range(len(CASES)))
    source = probe()
    rows: list[list[object]] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        for index in chosen:
            excel.new_document()
            excel.add_module('EventLog', LOG)
            installed = excel.run_vba(install(), 'Install', timeout=60.0, line_numbers=False)
            assert installed.ok, installed
            done = excel.run_vba(source, 'Probe', (folder, index), timeout=120.0, line_numbers=False)
            answer = (str(done.value) if done.ok and not done.dialogs
                      else f'STOPPED {done.outcome} {done.error} {done.message} {done.dialogs}')
            rows.append([CASES[index], answer.replace(folder, '<folder>'), parts(Path(folder) / f'c{index}.xlsx')])
            print(index, CASES[index][-60:], '=>', answer, flush=True)
    if len(chosen) == len(CASES):
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'log': LOG, 'sink': SINK,
                                   'source': source, 'cases': rows}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
