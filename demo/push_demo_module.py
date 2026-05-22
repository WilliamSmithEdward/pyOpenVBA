"""
Pushes a demo VBA module into test_macro_workbook.xlsm.
Contains three subs:
  - RainbowGrid    : 20x20 RGB colour-gradient heatmap painted into cells
  - FibSeries      : Fibonacci sequence with in-cell bar charts (REPT)
  - BubbleSortRace : Bubble-sort a shuffled column while colouring passes
"""

from pyopenvba import ExcelFile, VBAModuleKind

# VBA source: use \r\n line endings (required by the VBA storage format).
# LongLong is 64-bit VBA only, so use Long throughout (Fib(20)=6765, safe).
_VBA_SRC_LF = '''\
Attribute VB_Name = "DemoShowcase"
' ============================================================
' DemoShowcase -- pyOpenVBA write demo
' Run any of the three subs from the VBA IDE or a button.
' ============================================================

' ------------------------------------------------------------
' RainbowGrid
' Paints a 20x20 block of cells with a smooth RGB gradient.
' Top-left  = pure red,  top-right  = pure green,
' bottom-left = pure blue, bottom-right = white.
' ------------------------------------------------------------
Sub RainbowGrid()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    Const ROWS As Long = 20
    Const COLS As Long = 20

    Dim r As Long, c As Long
    Dim red As Long, grn As Long, blu As Long

    ws.Cells.Clear
    ws.Columns.ColumnWidth = 4
    ws.Rows.RowHeight = 18

    For r = 1 To ROWS
        For c = 1 To COLS
            red = CLng(255 * (1 - (r - 1) / (ROWS - 1)) * (1 - (c - 1) / (COLS - 1)) + _
                       255 * ((r - 1) / (ROWS - 1)) * ((c - 1) / (COLS - 1)))
            grn = CLng(255 * (1 - (r - 1) / (ROWS - 1)) * ((c - 1) / (COLS - 1)))
            blu = CLng(255 * ((r - 1) / (ROWS - 1)) * (1 - (c - 1) / (COLS - 1)))

            With ws.Cells(r, c).Interior
                .Color = RGB(red, grn, blu)
            End With
        Next c
    Next r

    ws.Range("A1").Select
    MsgBox "RainbowGrid done -- " & ROWS * COLS & " cells painted!", vbInformation
End Sub

' ------------------------------------------------------------
' FibSeries
' Writes the first 20 Fibonacci numbers with a text bar chart.
' Col A = index, Col B = value, Col C = bar (pipe chars).
' ------------------------------------------------------------
Sub FibSeries()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear
    ws.Columns.ColumnWidth = 14

    Const N As Long = 20
    Dim a As Long, b As Long, tmp As Long
    Dim i As Long
    Dim maxVal As Long

    ws.Cells(1, 1).Value = "n"
    ws.Cells(1, 2).Value = "Fib(n)"
    ws.Cells(1, 3).Value = "Bar"
    ws.Rows(1).Font.Bold = True

    a = 0: b = 1
    Dim vals(1 To N) As Long
    For i = 1 To N
        vals(i) = a
        tmp = a + b: a = b: b = tmp
    Next i
    maxVal = vals(N)

    For i = 1 To N
        Dim barLen As Long
        If maxVal > 0 Then
            barLen = CLng(40 * vals(i) / maxVal)
        Else
            barLen = 0
        End If
        ws.Cells(i + 1, 1).Value = i - 1
        ws.Cells(i + 1, 2).Value = vals(i)
        ws.Cells(i + 1, 3).Value = String(barLen, "|")
        ws.Cells(i + 1, 3).Font.Color = RGB(0, 112, 192)
    Next i

    ws.Columns("A:C").AutoFit
    MsgBox "FibSeries done -- first " & N & " Fibonacci numbers written.", vbInformation
End Sub

' ------------------------------------------------------------
' BubbleSortRace
' Fills column A with 20 shuffled integers (1-20),
' then bubble-sorts them, painting each pass a new colour
' so you can see the sweeps in the cell backgrounds.
' ------------------------------------------------------------
Sub BubbleSortRace()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear

    Const N As Long = 20
    Dim arr(1 To N) As Long
    Dim i As Long, j As Long, tmp As Long

    For i = 1 To N: arr(i) = i: Next i

    Randomize
    For i = N To 2 Step -1
        j = CLng(Int(Rnd() * i) + 1)
        tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
    Next i

    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 36
    For i = 1 To N
        ws.Cells(i, 1).Value = arr(i)
    Next i

    Dim palette(0 To 9) As Long
    palette(0) = RGB(255, 182, 193)
    palette(1) = RGB(255, 218, 185)
    palette(2) = RGB(255, 255, 153)
    palette(3) = RGB(180, 255, 180)
    palette(4) = RGB(173, 216, 230)
    palette(5) = RGB(216, 191, 216)
    palette(6) = RGB(255, 160, 122)
    palette(7) = RGB(144, 238, 144)
    palette(8) = RGB(135, 206, 235)
    palette(9) = RGB(255, 228, 181)

    Dim pass As Long: pass = 0
    Dim swapped As Boolean
    Dim passCol As Long
    Do
        swapped = False
        passCol = palette(pass Mod 10)
        For i = 1 To N - pass - 1
            If arr(i) > arr(i + 1) Then
                tmp = arr(i): arr(i) = arr(i + 1): arr(i + 1) = tmp
                ws.Cells(i, 1).Value = arr(i)
                ws.Cells(i + 1, 1).Value = arr(i + 1)
                ws.Cells(i, 1).Interior.Color = passCol
                ws.Cells(i + 1, 1).Interior.Color = passCol
                swapped = True
            End If
        Next i
        pass = pass + 1
        If pass > N Then Exit Do
    Loop While swapped

    For i = 1 To N
        ws.Cells(i, 2).Value = "Sorted in " & pass & " pass(es): " & arr(i)
    Next i

    ws.Columns("A:B").AutoFit
    MsgBox "BubbleSortRace done -- sorted in " & pass & " pass(es).", vbInformation
End Sub
'''

# VBA storage requires CRLF line endings; normalise from the Python string.
VBA_SRC = _VBA_SRC_LF.replace("\r\n", "\n").replace("\n", "\r\n")

with ExcelFile("test_macro_workbook.xlsm") as wb:
    project = wb.vba_project()

    # Remove old copy if re-running
    if "DemoShowcase" in wb.module_names():
        project.delete_module("DemoShowcase")

    project.add_module("DemoShowcase", VBA_SRC, kind=VBAModuleKind.standard)
    wb.save()

print("Done. Verifying...")

with ExcelFile("test_macro_workbook.xlsm") as wb:
    assert "DemoShowcase" in wb.module_names()
    src = wb.get_module("DemoShowcase")
    subs = [line.strip() for line in src.splitlines() if line.strip().startswith("Sub ")]
    print(f"Module 'DemoShowcase' written successfully.")
    print(f"Subs found: {subs}")
