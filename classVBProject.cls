VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classVBProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name$
Public Modules As Collection

Property Get MyName()
    MyName = Name
End Property

Property Let MyName(InputName)
    Name = InputName
End Property

Private Sub Class_Initialize()
    Set Modules = New Collection
End Sub

Public Sub AddModule(InputModule As classModule)
    Modules.Add InputModule
End Sub

Public Function Item(Index&) As classModule
    Set Item = Modules(Index)
End Function

