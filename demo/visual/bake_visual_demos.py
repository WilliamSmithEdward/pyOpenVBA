"""Bake a small gallery of demo .xlsm files showcasing ExcelFile.create_new()
together with visual / animated VBA macros.

Run from the repo root:

    python scripts/bake_visual_demos.py

Output goes to ``demo/visual/`` and overwrites any existing files there.
Open any of the resulting workbooks in Excel, enable macros, and run the
``Run`` macro (Alt+F8) to watch the show.
"""

from __future__ import annotations

from pathlib import Path

from pyopenvba import ExcelFile

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / "demo" / "visual"


# ---------------------------------------------------------------------------
# Demo 1 - Rainbow grid: paint a 20x20 block of cells in a smooth HSV sweep.
# ---------------------------------------------------------------------------
RAINBOW_GRID = """Attribute VB_Name = "Module1"
Option Explicit

' Paint a 20x20 grid of cells with a smooth rainbow gradient.
' Run this from Excel with Alt+F8 -> Run.
Public Sub Run()
    Const SIZE As Long = 20
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Cells.Clear
    ws.Cells.ColumnWidth = 3
    ws.Cells.RowHeight = 18

    Dim r As Long, c As Long
    For r = 1 To SIZE
        For c = 1 To SIZE
            Dim hue As Double
            hue = ((r + c) Mod (SIZE * 2)) / (SIZE * 2#)
            ws.Cells(r, c).Interior.Color = HsvToRgb(hue, 1#, 1#)
            DoEvents
        Next c
    Next r
End Sub

' h,s,v in [0,1] -> Long RGB usable as Interior.Color
Private Function HsvToRgb(ByVal h As Double, ByVal s As Double, ByVal v As Double) As Long
    Dim i As Long, f As Double, p As Double, q As Double, t As Double
    Dim red As Double, green As Double, blue As Double

    If s = 0 Then
        red = v: green = v: blue = v
    Else
        h = h * 6#
        i = Int(h)
        f = h - i
        p = v * (1# - s)
        q = v * (1# - s * f)
        t = v * (1# - s * (1# - f))
        Select Case i Mod 6
            Case 0: red = v: green = t: blue = p
            Case 1: red = q: green = v: blue = p
            Case 2: red = p: green = v: blue = t
            Case 3: red = p: green = q: blue = v
            Case 4: red = t: green = p: blue = v
            Case 5: red = v: green = p: blue = q
        End Select
    End If

    HsvToRgb = RGB(CInt(red * 255), CInt(green * 255), CInt(blue * 255))
End Function
"""


# ---------------------------------------------------------------------------
# Demo 2 - Bouncing ball: animate a shape across the worksheet.
# ---------------------------------------------------------------------------
BOUNCING_BALL = """Attribute VB_Name = "Module1"
Option Explicit

' Bounce a coloured circle around the worksheet for a few seconds.
' Run this from Excel with Alt+F8 -> Run.
Public Sub Run()
    Const FRAMES As Long = 240
    Const W As Single = 600
    Const H As Single = 400
    Const D As Single = 30

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Wipe any previous balls.
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "ball" Then shp.Delete
    Next shp

    Set shp = ws.Shapes.AddShape(msoShapeOval, 0, 0, D, D)
    shp.Name = "ball"
    shp.Fill.ForeColor.RGB = RGB(220, 40, 60)
    shp.Line.Visible = msoFalse

    Dim x As Single, y As Single, vx As Single, vy As Single
    x = 50: y = 50: vx = 6: vy = 4

    Dim i As Long
    For i = 1 To FRAMES
        x = x + vx
        y = y + vy
        If x < 0 Or x > W - D Then vx = -vx: x = x + vx
        If y < 0 Or y > H - D Then vy = -vy: y = y + vy
        shp.Left = x
        shp.Top = y
        shp.Fill.ForeColor.RGB = RGB( _
            (i * 3) Mod 256, _
            (i * 5) Mod 256, _
            (i * 7) Mod 256)
        DoEvents
    Next i
End Sub
"""


# ---------------------------------------------------------------------------
# Demo 3 - Sine wave: render a sine curve via cell colours.
# ---------------------------------------------------------------------------
SINE_WAVE = """Attribute VB_Name = "Module1"
Option Explicit

' Render an animated sine wave by colouring cells in a 40x20 grid.
' Run this from Excel with Alt+F8 -> Run.
Public Sub Run()
    Const COLS As Long = 60
    Const ROWS As Long = 20
    Const FRAMES As Long = 80
    Const PI As Double = 3.14159265358979

    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Cells.Clear
    ws.Cells.ColumnWidth = 2
    ws.Cells.RowHeight = 14

    Dim f As Long, c As Long, midRow As Long
    midRow = ROWS \\ 2

    For f = 1 To FRAMES
        Dim phase As Double
        phase = f / 6#
        ws.Range(ws.Cells(1, 1), ws.Cells(ROWS, COLS)).Interior.ColorIndex = xlNone
        For c = 1 To COLS
            Dim y As Double, row_ As Long
            y = Sin((c / 4#) + phase)
            row_ = midRow + CLng(y * (midRow - 1))
            If row_ < 1 Then row_ = 1
            If row_ > ROWS Then row_ = ROWS
            ws.Cells(row_, c).Interior.Color = RGB(40, 120 + CInt(y * 80), 220)
        Next c
        DoEvents
    Next f
End Sub
"""


# ---------------------------------------------------------------------------
# Demo 4 - Mandelbrot: render the Mandelbrot set in cell colours.
# ---------------------------------------------------------------------------
MANDELBROT = """Attribute VB_Name = "Module1"
Option Explicit

' Render the Mandelbrot set as cell colours on the active sheet.
' Run this from Excel with Alt+F8 -> Run. Takes a few seconds.
Public Sub Run()
    Const COLS As Long = 80
    Const ROWS As Long = 50
    Const MAX_ITER As Long = 60

    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Cells.Clear
    ws.Cells.ColumnWidth = 2
    ws.Cells.RowHeight = 12
    Application.ScreenUpdating = False

    Dim r As Long, c As Long
    For r = 1 To ROWS
        For c = 1 To COLS
            Dim x0 As Double, y0 As Double
            x0 = (c - 1) / (COLS - 1) * 3.5 - 2.5
            y0 = (r - 1) / (ROWS - 1) * 2# - 1#

            Dim x As Double, y As Double, iter As Long
            x = 0: y = 0: iter = 0
            Do While (x * x + y * y) <= 4# And iter < MAX_ITER
                Dim xt As Double
                xt = x * x - y * y + x0
                y = 2# * x * y + y0
                x = xt
                iter = iter + 1
            Loop

            If iter = MAX_ITER Then
                ws.Cells(r, c).Interior.Color = RGB(0, 0, 0)
            Else
                Dim t As Double
                t = iter / MAX_ITER
                ws.Cells(r, c).Interior.Color = RGB( _
                    CInt(9 * (1 - t) * t * t * t * 255), _
                    CInt(15 * (1 - t) * (1 - t) * t * t * 255), _
                    CInt(8.5 * (1 - t) * (1 - t) * (1 - t) * t * 255))
            End If
        Next c
    Next r

    Application.ScreenUpdating = True
End Sub
"""


DEMOS: list[tuple[str, str]] = [
    ("01_rainbow_grid.xlsm", RAINBOW_GRID),
    ("02_bouncing_ball.xlsm", BOUNCING_BALL),
    ("03_sine_wave.xlsm", SINE_WAVE),
    ("04_mandelbrot.xlsm", MANDELBROT),
]


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for filename, source in DEMOS:
        target = OUT_DIR / filename
        with ExcelFile.create_new(target) as wb:
            project = wb.vba_project()
            module = project.get_module("Module1")
            module.source = source
            module.dirty = True
            wb.save()
        print(f"baked {target.relative_to(ROOT)} ({target.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
