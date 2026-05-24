"""
Generate live Word test fixtures for pyOpenVBA.

Each fixture is derived from tests/live_word_testing/Doc1.docm (the smallest
valid Word+VBA document we have on hand) and given a specific VBA payload to
exercise a different aspect of the reader/writer pipeline.

Run from the repo root:
    python scripts/bake_word_test_fixtures.py
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from pyopenvba.word import WordFile  # noqa: E402

BASE_DOCM = ROOT / "tests" / "live_word_testing" / "Doc1.docm"
OUT_DIR = ROOT / "tests" / "live_word_testing"

# ---------------------------------------------------------------------------
# VBA payloads
# ---------------------------------------------------------------------------

_SIMPLE_MACROS = """\
Option Explicit

' A handful of simple, self-contained macros suitable for basic round-trip
' testing.  Nothing here calls the Word object model so it can be compiled
' without an open document.

Public Sub SayHello()
    MsgBox "Hello from pyOpenVBA!", vbInformation, "Simple Macros"
End Sub

Public Sub SayGoodbye()
    MsgBox "Goodbye!", vbInformation, "Simple Macros"
End Sub

Public Function AddIntegers(a As Long, b As Long) As Long
    AddIntegers = a + b
End Function

Public Function MultiplyDoubles(x As Double, y As Double) As Double
    MultiplyDoubles = x * y
End Function

Public Function IsEven(n As Long) As Boolean
    IsEven = (n Mod 2 = 0)
End Function

Public Sub PrintRange(lo As Long, hi As Long)
    Dim i As Long
    For i = lo To hi
        Debug.Print i
    Next i
End Sub

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
"""

# ---- large / complex module (data structures, sorting, string ops) ---------

_LARGE_MODULE = """\
Option Explicit

' =============================================================================
' LargeModule -- stress-tests the compressor / decompressor with a long,
' varied source stream containing many repeated and non-repeated byte patterns.
' =============================================================================

' --- Constants ---------------------------------------------------------------

Private Const MAX_ITEMS     As Long = 1024
Private Const BUFFER_SIZE   As Long = 4096
Private Const VERSION_STR   As String = "1.0.0"
Private Const APP_NAME      As String = "pyOpenVBA Word Fixture"

' --- Module-level state ------------------------------------------------------

Private m_Initialised As Boolean
Private m_ItemCount   As Long
Private m_Items()     As String
Private m_Keys()      As Long
Private m_ErrorLog    As String

' =============================================================================
' Initialisation
' =============================================================================

Public Sub Initialise()
    If m_Initialised Then Exit Sub
    ReDim m_Items(0 To MAX_ITEMS - 1)
    ReDim m_Keys(0 To MAX_ITEMS - 1)
    m_ItemCount   = 0
    m_ErrorLog    = ""
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
        LogError "AddItem: capacity exceeded (MAX_ITEMS=" & MAX_ITEMS & ")"
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

Public Function ItemCount() As Long
    ItemCount = m_ItemCount
End Function

' =============================================================================
' Sorting (insertion sort -- stable, good for nearly-sorted data)
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

Public Function Trim2(s As String) As String
    ' Trim leading and trailing whitespace including tabs and non-breaking spaces.
    Dim i As Long, j As Long
    i = 1
    j = Len(s)
    Do While i <= j
        Select Case Asc(Mid(s, i, 1))
            Case 9, 10, 13, 32, 160
                i = i + 1
            Case Else
                Exit Do
        End Select
    Loop
    Do While j >= i
        Select Case Asc(Mid(s, j, 1))
            Case 9, 10, 13, 32, 160
                j = j - 1
            Case Else
                Exit Do
        End Select
    Loop
    If j < i Then
        Trim2 = ""
    Else
        Trim2 = Mid(s, i, j - i + 1)
    End If
End Function

Public Function SplitCSV(line As String) As String()
    ' Split a CSV line respecting double-quoted fields.
    Dim fields() As String
    Dim fieldCount As Long
    Dim inQuote As Boolean
    Dim current As String
    Dim ch As String
    Dim i As Long
    Dim dq As String
    dq = Chr(34)

    fieldCount = 0
    inQuote    = False
    current    = ""
    ReDim fields(0)

    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        If inQuote Then
            If ch = dq Then
                If Mid(line, i + 1, 1) = dq Then
                    current = current & dq
                    i = i + 1
                Else
                    inQuote = False
                End If
            Else
                current = current & ch
            End If
        Else
            If ch = dq Then
                inQuote = True
            ElseIf ch = "," Then
                ReDim Preserve fields(fieldCount)
                fields(fieldCount) = current
                fieldCount = fieldCount + 1
                current = ""
                ReDim Preserve fields(fieldCount)
            Else
                current = current & ch
            End If
        End If
    Next i

    ReDim Preserve fields(fieldCount)
    fields(fieldCount) = current
    SplitCSV = fields
End Function

Public Function PadLeft(s As String, totalWidth As Long, padChar As String) As String
    Dim needed As Long
    needed = totalWidth - Len(s)
    If needed <= 0 Then
        PadLeft = s
    Else
        PadLeft = String(needed, padChar) & s
    End If
End Function

Public Function PadRight(s As String, totalWidth As Long, padChar As String) As String
    Dim needed As Long
    needed = totalWidth - Len(s)
    If needed <= 0 Then
        PadRight = s
    Else
        PadRight = s & String(needed, padChar)
    End If
End Function

Public Function WrapText(s As String, maxWidth As Long) As String
    ' Word-wrap s at maxWidth columns, breaking on spaces.
    Dim result As String
    Dim remaining As String
    Dim lineEnd As Long

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

' =============================================================================
' Numeric utilities
' =============================================================================

Public Function GCD(a As Long, b As Long) As Long
    Dim t As Long
    Do While b <> 0
        t = b
        b = a Mod b
        a = t
    Loop
    GCD = Abs(a)
End Function

Public Function LCM(a As Long, b As Long) As Long
    If a = 0 Or b = 0 Then
        LCM = 0
    Else
        LCM = Abs(a \\ GCD(a, b)) * Abs(b)
    End If
End Function

Public Function IsPrime(n As Long) As Boolean
    Dim i As Long
    If n < 2 Then
        IsPrime = False
        Exit Function
    End If
    If n = 2 Then
        IsPrime = True
        Exit Function
    End If
    If n Mod 2 = 0 Then
        IsPrime = False
        Exit Function
    End If
    i = 3
    Do While i * i <= n
        If n Mod i = 0 Then
            IsPrime = False
            Exit Function
        End If
        i = i + 2
    Loop
    IsPrime = True
End Function

Public Function SieveOfEratosthenes(limit As Long) As Long()
    Dim composite() As Boolean
    Dim primes()    As Long
    Dim primeCount  As Long
    Dim i As Long, j As Long

    ReDim composite(0 To limit)
    primeCount = 0

    For i = 2 To limit
        If Not composite(i) Then
            primeCount = primeCount + 1
            j = i * 2
            Do While j <= limit
                composite(j) = True
                j = j + i
            Loop
        End If
    Next i

    ReDim primes(0 To primeCount - 1)
    Dim idx As Long
    idx = 0
    For i = 2 To limit
        If Not composite(i) Then
            primes(idx) = i
            idx = idx + 1
        End If
    Next i

    SieveOfEratosthenes = primes
End Function

Public Function MatrixMult2x2(a11 As Double, a12 As Double, _
                               a21 As Double, a22 As Double, _
                               b11 As Double, b12 As Double, _
                               b21 As Double, b22 As Double, _
                               ByRef r11 As Double, ByRef r12 As Double, _
                               ByRef r21 As Double, ByRef r22 As Double)
    r11 = a11 * b11 + a12 * b21
    r12 = a11 * b12 + a12 * b22
    r21 = a21 * b11 + a22 * b21
    r22 = a21 * b12 + a22 * b22
End Function

' =============================================================================
' Bit manipulation
' =============================================================================

Public Function BitSet(value As Long, bitPos As Integer) As Long
    BitSet = value Or (1 And &HFFFFFFFF&) * (2 ^ bitPos)
End Function

Public Function BitClear(value As Long, bitPos As Integer) As Long
    BitClear = value And Not ((1 And &HFFFFFFFF&) * (2 ^ bitPos))
End Function

Public Function BitTest(value As Long, bitPos As Integer) As Boolean
    BitTest = ((value And ((1 And &HFFFFFFFF&) * (2 ^ bitPos))) <> 0)
End Function

Public Function PopCount(value As Long) As Long
    Dim count As Long
    Dim v As Long
    v = value
    count = 0
    Do While v <> 0
        count = count + (v And 1)
        v = v \ 2
    Loop
    PopCount = count
End Function

' =============================================================================
' Error log
' =============================================================================

Private Sub LogError(msg As String)
    m_ErrorLog = m_ErrorLog & Now() & ": " & msg & vbCrLf
End Sub

Public Function GetErrorLog() As String
    GetErrorLog = m_ErrorLog
End Function

Public Sub ClearErrorLog()
    m_ErrorLog = ""
End Sub

' =============================================================================
' Self-test (no Word object model required)
' =============================================================================

Public Sub SelfTest()
    Dim ok As Boolean
    ok = True

    ' AddIntegers equivalent
    If AddIntegers_local(3, 4) <> 7 Then
        LogError "SelfTest: AddIntegers failed"
        ok = False
    End If

    ' GCD
    If GCD(48, 18) <> 6 Then
        LogError "SelfTest: GCD(48,18) expected 6"
        ok = False
    End If

    ' LCM
    If LCM(4, 6) <> 12 Then
        LogError "SelfTest: LCM(4,6) expected 12"
        ok = False
    End If

    ' IsPrime
    If Not IsPrime(97) Then
        LogError "SelfTest: IsPrime(97) expected True"
        ok = False
    End If
    If IsPrime(91) Then
        LogError "SelfTest: IsPrime(91) expected False (7x13)"
        ok = False
    End If

    ' PadLeft
    If PadLeft("42", 5, "0") <> "00042" Then
        LogError "SelfTest: PadLeft failed"
        ok = False
    End If

    ' Trim2
    If Trim2("  hello  ") <> "hello" Then
        LogError "SelfTest: Trim2 failed"
        ok = False
    End If

    If ok Then
        Debug.Print "SelfTest PASSED"
    Else
        Debug.Print "SelfTest FAILED -- see GetErrorLog()"
    End If
End Sub

' Local helper (avoids circular dep on AddIntegers in separate module)
Private Function AddIntegers_local(a As Long, b As Long) As Long
    AddIntegers_local = a + b
End Function

Public Function AddIntegers(a As Long, b As Long) As Long
    AddIntegers = a + b
End Function
"""

# ---- document events (ThisDocument body only) -------------------------------

_DOCUMENT_EVENTS_BODY = """\
Option Explicit

' Document-level event handlers.  These fire automatically when the
' corresponding Word events occur.

Private m_OpenTime As Date
Private m_ChangeCount As Long

Private Sub Document_Open()
    m_OpenTime   = Now()
    m_ChangeCount = 0
    Application.StatusBar = "Document opened at " & Format(m_OpenTime, "hh:mm:ss")
End Sub

Private Sub Document_Close()
    Application.StatusBar = False
End Sub

Private Sub Document_ContentControlOnEnter(ByVal ContentControl As ContentControl)
    Application.StatusBar = "Entering: " & ContentControl.Tag
End Sub

Private Sub Document_ContentControlOnExit( _
        ByVal ContentControl As ContentControl, _
        Cancel As Boolean)
    Application.StatusBar = "Leaving: " & ContentControl.Tag
End Sub

' ---------------------------------------------------------------------------
' Helpers accessible from other modules
' ---------------------------------------------------------------------------

Public Function OpenTime() As Date
    OpenTime = m_OpenTime
End Function

Public Function ChangeCount() As Long
    ChangeCount = m_ChangeCount
End Function

Public Sub IncrementChangeCount()
    m_ChangeCount = m_ChangeCount + 1
End Sub
"""

# ---- multi-module: Module1, TextUtils, MathUtils ---------------------------

_MULTI_MOD1 = """\
Option Explicit
' Module1 -- entry points that delegate to the utility modules.

Public Sub RunAll()
    Dim s As String
    s = TextUtils.RepeatStr("ab", 5)
    Debug.Print "RepeatStr: " & s

    Dim primes() As Long
    Dim i As Long
    primes = MathUtils.Primes(50)
    Dim out As String
    For i = LBound(primes) To UBound(primes)
        out = out & CStr(primes(i)) & " "
    Next i
    Debug.Print "Primes <= 50: " & out
End Sub
"""

_MULTI_TEXT_UTILS = """\
Option Explicit
' TextUtils -- reusable string helpers.

Public Function RepeatStr(s As String, n As Long) As String
    Dim i As Long, result As String
    For i = 1 To n
        result = result & s
    Next i
    RepeatStr = result
End Function

Public Function CountOccurrences(haystack As String, needle As String) As Long
    If Len(needle) = 0 Then
        CountOccurrences = 0
        Exit Function
    End If
    Dim count As Long
    Dim pos As Long
    count = 0
    pos   = 1
    Do
        pos = InStr(pos, haystack, needle)
        If pos = 0 Then Exit Do
        count = count + 1
        pos   = pos + Len(needle)
    Loop
    CountOccurrences = count
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

Public Function ReplaceAll(s As String, findStr As String, replaceStr As String) As String
    ReplaceAll = s
    If Len(findStr) = 0 Then Exit Function
    Dim pos As Long
    pos = InStr(ReplaceAll, findStr)
    Do While pos > 0
        ReplaceAll = Left(ReplaceAll, pos - 1) & replaceStr & Mid(ReplaceAll, pos + Len(findStr))
        pos = InStr(pos + Len(replaceStr), ReplaceAll, findStr)
    Loop
End Function

Public Function Slugify(s As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    result = LCase(s)
    For i = 1 To Len(result)
        ch = Mid(result, i, 1)
        Select Case ch
            Case "a" To "z", "0" To "9"
                Slugify = Slugify & ch
            Case " ", "-", "_"
                Slugify = Slugify & "-"
            Case Else
                ' drop
        End Select
    Next i
    ' Collapse consecutive hyphens
    Do While InStr(Slugify, "--") > 0
        Slugify = ReplaceAll(Slugify, "--", "-")
    Loop
    Slugify = Trim(Slugify, "-")
End Function

Private Function Trim(s As String, ch As String) As String
    Trim = s
    Do While Left(Trim, 1) = ch
        Trim = Mid(Trim, 2)
    Loop
    Do While Right(Trim, 1) = ch
        Trim = Left(Trim, Len(Trim) - 1)
    Loop
End Function
"""

_MULTI_MATH_UTILS = """\
Option Explicit
' MathUtils -- numeric helpers.

Public Function Primes(limit As Long) As Long()
    Dim sieve() As Boolean
    Dim result() As Long
    Dim i As Long, j As Long, count As Long

    ReDim sieve(2 To limit)
    For i = 2 To limit
        sieve(i) = True
    Next i

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

    If count = 0 Then
        Primes = result
        Exit Function
    End If

    ReDim result(0 To count - 1)
    Dim idx As Long
    idx = 0
    For i = 2 To limit
        If sieve(i) Then
            result(idx) = i
            idx = idx + 1
        End If
    Next i

    Primes = result
End Function

Public Function Fibonacci(n As Long) As Long
    If n <= 0 Then
        Fibonacci = 0
        Exit Function
    End If
    If n = 1 Then
        Fibonacci = 1
        Exit Function
    End If
    Dim a As Long, b As Long, tmp As Long, i As Long
    a = 0
    b = 1
    For i = 2 To n
        tmp = a + b
        a   = b
        b   = tmp
    Next i
    Fibonacci = b
End Function

Public Function BinomialCoeff(n As Long, k As Long) As Double
    ' Compute C(n, k) using the multiplicative formula to avoid overflow.
    If k < 0 Or k > n Then
        BinomialCoeff = 0
        Exit Function
    End If
    If k = 0 Or k = n Then
        BinomialCoeff = 1
        Exit Function
    End If
    If k > n - k Then k = n - k
    Dim result As Double
    Dim i As Long
    result = 1
    For i = 0 To k - 1
        result = result * (n - i)
        result = result / (i + 1)
    Next i
    BinomialCoeff = result
End Function

Public Function RoundToNearest(value As Double, step As Double) As Double
    RoundToNearest = Int(value / step + 0.5) * step
End Function

Public Function Clamp(value As Double, lo As Double, hi As Double) As Double
    If value < lo Then
        Clamp = lo
    ElseIf value > hi Then
        Clamp = hi
    Else
        Clamp = value
    End If
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
"""

# ---------------------------------------------------------------------------
# Helper: write a fixture
# ---------------------------------------------------------------------------

def _bake(
    name: str,
    modules: dict[str, str],
    *,
    document_body: str | None = None,
) -> Path:
    """
    Copy Doc1.docm to OUT_DIR/<name>, set the given modules, and save.
    ``modules`` maps logical module name -> full source (without attribute header).
    ``document_body`` replaces ThisDocument's body only (header is preserved).
    """
    out = OUT_DIR / name
    shutil.copy(BASE_DOCM, out)

    with WordFile(out) as doc:
        if document_body is not None:
            doc.set_module("ThisDocument", document_body)

        for mod_name, src in modules.items():
            try:
                doc.set_module(mod_name, src)
            except KeyError:
                # Module doesn't exist in the base; add it.
                proj = doc.vba_project()
                proj.add_module(mod_name, src)

        doc.save()

    print(f"  wrote {out.relative_to(ROOT)}")
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    if not BASE_DOCM.exists():
        print(f"ERROR: base fixture not found: {BASE_DOCM}")
        raise SystemExit(1)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    print("Baking Word live test fixtures ...")

    # 1. Simple macros only (Module1 replaced, ThisDocument untouched)
    _bake("simple_macros.docm", {"Module1": _SIMPLE_MACROS})

    # 2. Large single module with lots of varied code
    _bake("large_vba_module.docm", {"Module1": _LARGE_MODULE})

    # 3. Document events in ThisDocument, standard module left bare
    _bake(
        "document_events.docm",
        {"Module1": "'Placeholder -- see ThisDocument for event handlers\r\n"},
        document_body=_DOCUMENT_EVENTS_BODY,
    )

    # 4. Multiple modules (Module1 + TextUtils + MathUtils)
    _bake(
        "multi_module_doc.docm",
        {
            "Module1":   _MULTI_MOD1,
            "TextUtils": _MULTI_TEXT_UTILS,
            "MathUtils": _MULTI_MATH_UTILS,
        },
    )

    print("Done.")


if __name__ == "__main__":
    main()
