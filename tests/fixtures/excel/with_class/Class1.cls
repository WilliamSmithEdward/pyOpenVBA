VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mMessage As String

Public Property Get Message() As String
    Message = mMessage
End Property

Public Property Let Message(v As String)
    mMessage = v
End Property

Public Function Greet() As String
    Greet = "Class1.Greet: " & mMessage
End Function
