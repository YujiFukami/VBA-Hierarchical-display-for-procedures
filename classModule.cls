VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name$
Public Procedures As Collection
Public VBProjectName$
Public ModuleType$ 'モジュールのタイプ、標準モジュール、ユーザーフォーム、クラスモジュールなど
Public BookName$
Public Sengen$ '冒頭の宣言文

'名前
Property Get MyName()
    MyName = Name
End Property

Property Let MyName(InputName)
    Name = InputName
End Property

'初期化時
Private Sub Class_Initialize()
    Set Procedures = New Collection
End Sub

Public Sub AddProcedure(InputProcedure As classProcedure)
    Procedures.Add InputProcedure
End Sub

'プロシージャ取得
Public Function Item(Index&) As classProcedure
    Set Item = Procedures(Index)
End Function

'VBProject名
Property Get MyVBProjectName()
    MyVBProjectName = VBProjectName
End Property

Property Let MyVBProjectName(InputVBProjectName)
    VBProjectName = InputVBProjectName
End Property

'ModuleType名
Property Get MyModuleType()
    MyModuleType = ModuleType
End Property

Property Let MyModuleType(InputModuleType)
    ModuleType = InputModuleType
End Property

'ブックの名前
Property Get MyBookName()
    MyBookName = BookName
End Property

Property Let MyBookName(InputBookName)
    BookName = InputBookName
End Property

'冒頭の宣言文
Property Get MySengen()
    MySengen = Sengen
End Property

Property Let MySengen(InputSengen)
    Sengen = InputSengen
End Property

