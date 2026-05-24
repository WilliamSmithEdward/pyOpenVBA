"""
Generate live PowerPoint test fixtures for pyOpenVBA.

Each fixture is derived from tests/live_powerpoint_testing/Presentation1.pptm
and given a specific VBA payload.

Run from the repo root:
    python scripts/bake_powerpoint_test_fixtures.py
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.powerpoint import PowerPointFile  # noqa: E402

BASE_PPTM = ROOT / "tests" / "live_powerpoint_testing" / "Presentation1.pptm"
OUT_DIR   = ROOT / "tests" / "live_powerpoint_testing"

# ---------------------------------------------------------------------------
# VBA payloads
# ---------------------------------------------------------------------------

_SIMPLE_MACROS = """\
Option Explicit

' Simple self-contained macros for basic round-trip testing.

Public Sub SayHello()
    MsgBox "Hello from pyOpenVBA!", vbInformation, "Simple Macros"
End Sub

Public Function AddIntegers(a As Long, b As Long) As Long
    AddIntegers = a + b
End Function

Public Function IsEven(n As Long) As Boolean
    IsEven = (n Mod 2 = 0)
End Function

Public Function ReverseString(s As String) As String
    Dim i As Long
    Dim result As String
    For i = Len(s) To 1 Step -1
        result = result & Mid(s, i, 1)
    Next i
    ReverseString = result
End Function

Public Function Factorial(n As Long) As Long
    If n <= 1 Then
        Factorial = 1
    Else
        Factorial = n * Factorial(n - 1)
    End If
End Function

Public Sub PrintPrimes(limit As Long)
    Dim i As Long, j As Long
    Dim isPrime As Boolean
    For i = 2 To limit
        isPrime = True
        For j = 2 To CLng(Sqr(i))
            If i Mod j = 0 Then
                isPrime = False
                Exit For
            End If
        Next j
        If isPrime Then Debug.Print i
    Next i
End Sub
"""

_LARGE_MODULE = """\
Option Explicit

' =============================================================================
' LargeModule -- exercises the compressor / decompressor with a long source.
' =============================================================================

Private Const MAX_ITEMS   As Long = 512
Private Const APP_NAME    As String = "pyOpenVBA PowerPoint Fixture"

Private m_Initialised As Boolean
Private m_ItemCount   As Long
Private m_Items()     As String
Private m_Keys()      As Long
Private m_Log         As String

' =============================================================================
' Initialisation
' =============================================================================

Public Sub Initialise()
    If m_Initialised Then Exit Sub
    ReDim m_Items(0 To MAX_ITEMS - 1)
    ReDim m_Keys(0 To MAX_ITEMS - 1)
    m_ItemCount   = 0
    m_Log         = ""
    m_Initialised = True
End Sub

Public Sub Reset()
    m_Initialised = False
    Initialise
End Sub

' =============================================================================
' Item management
' =============================================================================

Public Function AddItem(key As Long, value As String) As Boolean
    If Not m_Initialised Then Initialise
    If m_ItemCount >= MAX_ITEMS Then
        AppendLog "AddItem: capacity exceeded"
        AddItem = False
        Exit Function
    End If
    m_Keys(m_ItemCount)  = key
    m_Items(m_ItemCount) = value
    m_ItemCount = m_ItemCount + 1
    AddItem = True
End Function

Public Function FindItem(key As Long) As String
    Dim i As Long
    For i = 0 To m_ItemCount - 1
        If m_Keys(i) = key Then
            FindItem = m_Items(i)
            Exit Function
        End If
    Next i
    FindItem = ""
End Function

Public Function RemoveItem(key As Long) As Boolean
    Dim i As Long, j As Long
    For i = 0 To m_ItemCount - 1
        If m_Keys(i) = key Then
            For j = i To m_ItemCount - 2
                m_Keys(j)  = m_Keys(j + 1)
                m_Items(j) = m_Items(j + 1)
            Next j
            m_ItemCount = m_ItemCount - 1
            RemoveItem = True
            Exit Function
        End If
    Next i
    RemoveItem = False
End Function

' =============================================================================
' Sorting
' =============================================================================

Public Sub SortByKey()
    Dim i As Long, j As Long
    Dim tmpKey As Long, tmpVal As String
    For i = 1 To m_ItemCount - 1
        tmpKey = m_Keys(i)
        tmpVal = m_Items(i)
        j = i - 1
        Do While j >= 0 And m_Keys(j) > tmpKey
            m_Keys(j + 1)  = m_Keys(j)
            m_Items(j + 1) = m_Items(j)
            j = j - 1
        Loop
        m_Keys(j + 1)  = tmpKey
        m_Items(j + 1) = tmpVal
    Next i
End Sub

Public Sub SortByValue()
    Dim i As Long, j As Long
    Dim tmpKey As Long, tmpVal As String
    For i = 1 To m_ItemCount - 1
        tmpKey = m_Keys(i)
        tmpVal = m_Items(i)
        j = i - 1
        Do While j >= 0 And m_Items(j) > tmpVal
            m_Keys(j + 1)  = m_Keys(j)
            m_Items(j + 1) = m_Items(j)
            j = j - 1
        Loop
        m_Keys(j + 1)  = tmpKey
        m_Items(j + 1) = tmpVal
    Next i
End Sub

' =============================================================================
' String utilities
' =============================================================================

Public Function PadLeft(s As String, width As Long, ch As String) As String
    Dim needed As Long
    needed = width - Len(s)
    If needed <= 0 Then PadLeft = s Else PadLeft = String(needed, ch) & s
End Function

Public Function PadRight(s As String, width As Long, ch As String) As String
    Dim needed As Long
    needed = width - Len(s)
    If needed <= 0 Then PadRight = s Else PadRight = s & String(needed, ch)
End Function

Public Function WrapText(s As String, maxWidth As Long) As String
    Dim result As String, remaining As String, lineEnd As Long
    remaining = s
    result    = ""
    Do While Len(remaining) > maxWidth
        lineEnd = maxWidth
        Do While lineEnd > 0 And Mid(remaining, lineEnd, 1) <> " "
            lineEnd = lineEnd - 1
        Loop
        If lineEnd = 0 Then lineEnd = maxWidth
        result    = result & Left(remaining, lineEnd) & vbCrLf
        remaining = LTrim(Mid(remaining, lineEnd + 1))
    Loop
    WrapText = result & remaining
End Function

Public Function TitleCase(s As String) As String
    Dim words() As String
    Dim i As Long
    words = Split(s, " ")
    For i = LBound(words) To UBound(words)
        If Len(words(i)) > 0 Then
            words(i) = UCase(Left(words(i), 1)) & LCase(Mid(words(i), 2))
        End If
    Next i
    TitleCase = Join(words, " ")
End Function

' =============================================================================
' Numeric utilities
' =============================================================================

Public Function GCD(a As Long, b As Long) As Long
    Dim t As Long
    Do While b <> 0
        t = b : b = a Mod b : a = t
    Loop
    GCD = Abs(a)
End Function

Public Function IsPrime(n As Long) As Boolean
    Dim i As Long
    If n < 2 Then IsPrime = False : Exit Function
    If n = 2 Then IsPrime = True : Exit Function
    If n Mod 2 = 0 Then IsPrime = False : Exit Function
    i = 3
    Do While i * i <= n
        If n Mod i = 0 Then IsPrime = False : Exit Function
        i = i + 2
    Loop
    IsPrime = True
End Function

Public Function Fibonacci(n As Long) As Long
    If n <= 0 Then Fibonacci = 0 : Exit Function
    If n = 1 Then Fibonacci = 1 : Exit Function
    Dim a As Long, b As Long, tmp As Long, i As Long
    a = 0 : b = 1
    For i = 2 To n
        tmp = a + b : a = b : b = tmp
    Next i
    Fibonacci = b
End Function

Public Function BinomialCoeff(n As Long, k As Long) As Double
    If k < 0 Or k > n Then BinomialCoeff = 0 : Exit Function
    If k = 0 Or k = n Then BinomialCoeff = 1 : Exit Function
    If k > n - k Then k = n - k
    Dim result As Double, i As Long
    result = 1
    For i = 0 To k - 1
        result = result * (n - i) / (i + 1)
    Next i
    BinomialCoeff = result
End Function

Public Function Clamp(value As Double, lo As Double, hi As Double) As Double
    If value < lo Then Clamp = lo _
    ElseIf value > hi Then Clamp = hi _
    Else Clamp = value
End Function

' =============================================================================
' Bit manipulation
' =============================================================================

Public Function PopCount(value As Long) As Long
    Dim count As Long, v As Long
    v = value
    Do While v <> 0
        count = count + (v And 1)
        v = v \\ 2
    Loop
    PopCount = count
End Function

' =============================================================================
' Log
' =============================================================================

Private Sub AppendLog(msg As String)
    m_Log = m_Log & Now() & ": " & msg & vbCrLf
End Sub

Public Function GetLog() As String
    GetLog = m_Log
End Function

Public Sub ClearLog()
    m_Log = ""
End Sub

' =============================================================================
' Self-test
' =============================================================================

Public Sub SelfTest()
    Dim ok As Boolean : ok = True

    If GCD(48, 18) <> 6 Then AppendLog "GCD fail" : ok = False
    If Not IsPrime(97) Then AppendLog "IsPrime(97) fail" : ok = False
    If IsPrime(91) Then AppendLog "IsPrime(91) fail" : ok = False
    If Fibonacci(10) <> 55 Then AppendLog "Fibonacci(10) fail" : ok = False
    If PadLeft("7", 3, "0") <> "007" Then AppendLog "PadLeft fail" : ok = False
    If TitleCase("hello world") <> "Hello World" Then AppendLog "TitleCase fail" : ok = False

    If ok Then
        Debug.Print APP_NAME & " SelfTest PASSED"
    Else
        Debug.Print APP_NAME & " SelfTest FAILED"
        Debug.Print GetLog()
    End If
End Sub

Public Function AddIntegers(a As Long, b As Long) As Long
    AddIntegers = a + b
End Function
"""

_SLIDE_UTILS = """\
Option Explicit
' SlideUtils -- helpers for working with presentation slides.
' These reference the PowerPoint object model so they require an open
' presentation at runtime; they compile cleanly in the VBE regardless.

Public Function SlideCount(prs As Object) As Long
    SlideCount = prs.Slides.Count
End Function

Public Sub SetSlideTitle(sld As Object, title As String)
    On Error Resume Next
    sld.Shapes.Title.TextFrame.TextRange.Text = title
    On Error GoTo 0
End Sub

Public Function GetSlideTitle(sld As Object) As String
    On Error Resume Next
    GetSlideTitle = sld.Shapes.Title.TextFrame.TextRange.Text
    On Error GoTo 0
End Function

Public Sub AddTextBox(sld As Object, text As String, _
                      left As Single, top As Single, _
                      width As Single, height As Single)
    Dim shp As Object
    Set shp = sld.Shapes.AddTextbox(1, left, top, width, height)
    shp.TextFrame.TextRange.Text = text
End Sub

Public Function CountShapes(sld As Object) As Long
    CountShapes = sld.Shapes.Count
End Function

Public Sub DeleteAllTextBoxes(sld As Object)
    Dim i As Long
    For i = sld.Shapes.Count To 1 Step -1
        If sld.Shapes(i).Type = 17 Then   ' msoTextBox = 17
            sld.Shapes(i).Delete
        End If
    Next i
End Sub
"""

_MATH_UTILS = """\
Option Explicit
' MathUtils -- numeric helpers (no object model dependency).

Public Function Primes(limit As Long) As Long()
    Dim sieve() As Boolean
    Dim result() As Long
    Dim i As Long, j As Long, count As Long

    If limit < 2 Then
        Primes = result
        Exit Function
    End If

    ReDim sieve(2 To limit)
    For i = 2 To limit : sieve(i) = True : Next i

    For i = 2 To CLng(Sqr(limit))
        If sieve(i) Then
            j = i * i
            Do While j <= limit
                sieve(j) = False
                j = j + i
            Loop
        End If
    Next i

    count = 0
    For i = 2 To limit
        If sieve(i) Then count = count + 1
    Next i

    ReDim result(0 To count - 1)
    Dim idx As Long : idx = 0
    For i = 2 To limit
        If sieve(i) Then
            result(idx) = i
            idx = idx + 1
        End If
    Next i
    Primes = result
End Function

Public Function LinearInterp(x0 As Double, y0 As Double, _
                              x1 As Double, y1 As Double, _
                              x  As Double) As Double
    If x1 = x0 Then
        LinearInterp = y0
    Else
        LinearInterp = y0 + (y1 - y0) * (x - x0) / (x1 - x0)
    End If
End Function

Public Function RoundToNearest(value As Double, step As Double) As Double
    RoundToNearest = Int(value / step + 0.5) * step
End Function
"""

# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------

def _bake(name: str, modules: dict[str, str]) -> Path:
    out = OUT_DIR / name
    shutil.copy(BASE_PPTM, out)

    with PowerPointFile(out) as prs:
        for mod_name, src in modules.items():
            try:
                prs.set_module(mod_name, src)
            except KeyError:
                proj = prs.vba_project()
                proj.add_module(mod_name, src)
        prs.save()

    print(f"  wrote {out.relative_to(ROOT)}")
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    if not BASE_PPTM.exists():
        print(f"ERROR: base fixture not found: {BASE_PPTM}")
        raise SystemExit(1)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    print("Baking PowerPoint live test fixtures ...")

    _bake("simple_macros.pptm", {"Module1": _SIMPLE_MACROS})
    _bake("large_vba_module.pptm", {"Module1": _LARGE_MODULE})
    _bake(
        "multi_module.pptm",
        {
            "Module1":    "'Entry point -- delegates to SlideUtils and MathUtils\r\n",
            "SlideUtils": _SLIDE_UTILS,
            "MathUtils":  _MATH_UTILS,
        },
    )

    print("Done.")


if __name__ == "__main__":
    main()
