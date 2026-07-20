Attribute VB_Name = "Module1"

'with_class fixture: exercises a class module and writes a sentinel file.
Sub RunFixture()
    On Error GoTo ErrorHandler
    
    Dim obj As New Class1
    obj.Message = "it works"

    Dim outPath As String
    outPath = ThisWorkbook.Path & "\" & Replace(ThisWorkbook.Name, ".xlsm", "") & ".txt"

    If outPath = "\" Then
        ' Fallback: use Environ("TEMP") if ThisWorkbook.Path is empty
        outPath = Environ("TEMP") & "\" & Replace(ThisWorkbook.Name, ".xlsm", "") & ".txt"
    End If

    On Error Resume Next
    Open outPath For Output As #1
    If Err.Number <> 0 Then
        Err.Clear
        ' Try alternate path
        outPath = Environ("TEMP") & "\" & Replace(ThisWorkbook.Name, ".xlsm", "") & ".txt"
        Open outPath For Output As #1
    End If
    On Error GoTo ErrorHandler
    
    Print #1, obj.Greet()
    Close #1
    Exit Sub
    
ErrorHandler:
    ' If error, at least write a file to indicate the macro ran
    Dim errPath As String
    errPath = Environ("TEMP") & "\" & Replace(ThisWorkbook.Name, ".xlsm", "") & "_error.txt"
    On Error Resume Next
    Open errPath For Output As #2
    Print #2, "Error in RunFixture: " & Err.Description
    Close #2
End Sub

