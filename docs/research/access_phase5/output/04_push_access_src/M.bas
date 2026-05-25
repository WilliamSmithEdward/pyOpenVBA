Option Explicit

Sub Pushed()
    MsgBox "pushed via pyopenvba.push_access()"
End Sub

Sub Pushed2()
    Dim i As Long
    For i = 1 To 3
        Debug.Print "i=" & i
    Next i
End Sub
