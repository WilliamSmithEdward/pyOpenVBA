Attribute VB_Name = "Matrix"
Option Explicit

Private Type Point2D
    X As Long
    Y As Long
End Type

Private Const KMax As Long = 10

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Public Function Arith(ByVal a As Long, ByVal b As Long) As Long
    Dim r As Long
    r = a + b
    r = a - b
    r = a * b
    r = a \ b
    r = a Mod b
    r = -a
    r = a ^ b
    r = (a + b) * a
    Arith = r
End Function

Public Function Logic(ByVal a As Long, ByVal b As Long) As Boolean
    Dim t As Boolean
    t = a = b
    t = a <> b
    t = a < b
    t = a > b
    t = a <= b
    t = a >= b
    t = Not t
    t = a > 0 And b > 0
    t = a > 0 Or b > 0
    t = a > 0 Xor b > 0
    Logic = t
End Function

Public Function Text(ByVal s As String) As String
    Dim u As String
    u = s & "x"
    u = Left(s, 2)
    u = Mid(s, 1, 3)
    u = UCase(s)
    u = Trim(s)
    Text = u
End Function

Public Sub Branching(ByVal n As Long)
    Dim r As Long
    If n > 0 Then
        r = 1
    ElseIf n < 0 Then
        r = 2
    Else
        r = 3
    End If
    If n = 0 Then r = 4
    Select Case n
        Case 0
            r = 5
        Case 1, 2
            r = 6
        Case 3 To 5
            r = 9
        Case Is > 9
            r = 7
        Case Is < 0
            r = 10
        Case Is <> 11
            r = 11
        Case Else
            r = 8
    End Select
End Sub

Public Sub Looping()
    Dim i As Long
    Dim total As Long
    For i = 1 To 10
        total = total + i
    Next i
    For i = 10 To 1 Step -1
        total = total - i
    Next i
    Do While total > 0
        total = total - 1
    Loop
    Do Until total > 5
        total = total + 1
    Loop
    Do
        total = total + 1
    Loop While total < 20
    While total < 30
        total = total + 1
    Wend
    For i = 1 To 3
        If i = 2 Then Exit For
    Next i
End Sub

Public Sub Arrays()
    Dim fixedArr(1 To 5) As Long
    Dim dynArr() As Long
    Dim v As Variant
    ReDim dynArr(1 To 3)
    fixedArr(1) = 10
    dynArr(1) = fixedArr(1)
    v = Array(1, 2, 3)
    Erase dynArr
End Sub

Public Sub Objects()
    Dim d As Object
    Dim p As Point2D
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "k", 1
    d.CompareMode = 1
    p.X = 1
    p.Y = 2
    With d
        .Add "j", 2
    End With
    Set d = Nothing
End Sub

Public Function Guarded(ByVal n As Long) As Long
    On Error GoTo Failed
    If n = 0 Then Err.Raise 5
    Guarded = 100 \ n
    Exit Function
Failed:
    Guarded = -1
    Resume Next
End Function

Public Sub Specials()
    Dim v As Variant
    Dim o As Object
    v = Empty
    v = Null
    Set o = Nothing
    v = True
    v = False
    v = 3.5
    v = 100000
    v = #1/2/2003#
    v = GetTickCount()
    Debug.Print "x"
    DoCmd.Beep
End Sub

Public Sub MemberForms()
    Dim d As Object
    Dim x As Variant
    Set d = CreateObject("Scripting.Dictionary")
    x = d.Item("k")
    d.Item("k") = 5
    x = d.Count
    d.CompareMode = 1
    x = d.Exists("k")
End Sub
