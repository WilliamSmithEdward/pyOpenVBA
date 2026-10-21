Option Explicit
Public src As Worksheet, calc As Worksheet, held As Range, localName As Name
Function Observe() As String
Dim n As Name, result As String
result="sheet=" & src.Name & ";code=" & src.CodeName & ";held=" & held.Address(External:=True) & ";formula=" & calc.Range("A1").Formula & ";value=" & CStr(calc.Range("A1").Value2) & ";self=" & src.Range("B1").Formula & ";selfValue=" & CStr(src.Range("B1").Value2)
For Each n In ActiveWorkbook.Names
result=result & ";name=[" & n.Name & ":" & n.RefersTo & ":" & CStr(n.Visible) & ":" & n.Comment & "]"
Next
Observe=result
End Function
Function TryRename(ws As Worksheet, target As String) As String
On Error GoTo Failed
ws.Name=target
TryRename="OK"
Exit Function
Failed:
TryRename="ERR:" & CStr(Err.Number) & ":" & Err.Description
End Function
Function Reopened() As String
Set src=ActiveWorkbook.Worksheets(1)
Set calc=ActiveWorkbook.Worksheets(2)
Set held=src.Range("A1")
Reopened=Observe()
End Function