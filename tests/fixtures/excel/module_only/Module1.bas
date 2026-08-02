Attribute VB_Name = "Module1"

'Simple fixture: writes a sentinel file to confirm macro execution.
Sub RunFixture()
    Dim outPath As String
    outPath = ThisWorkbook.Path & "\" & Replace(ThisWorkbook.Name, ".xlsm", "") & ".txt"

    Open outPath For Output As #1
    Print #1, "simple: OK"
    Close #1
End Sub