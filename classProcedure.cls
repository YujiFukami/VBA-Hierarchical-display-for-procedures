VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name$
Public Code
Public KensakuCode As Object '使用プロシージャ検索用連想配列
Public UseProcedure As Collection
Public ModuleName$
Public VBProjectName$
Public RangeOfUse$ 'プロシージャの使用可能範囲(PublicかPrivateか）
Public ProcedureType$ 'プロシージャのタイプ（Sub,Functionなど）

'初期化時
Private Sub Class_Initialize()
    Set UseProcedure = New Collection
End Sub

'名前
Property Get MyName()
    MyName = Name
End Property

Property Let MyName(InputName)
    Name = InputName
End Property

'コード
Property Get MyCode()
    MyCode = Code
End Property

Property Let MyCode(InputCode)
    Code = InputCode
End Property

'使用プロシージャ検索用連想配列
Property Get MyKensakuCode()
    Set MyKensakuCode = KensakuCode
End Property

Property Set MyKensakuCode(InputKensakuCode As Object)
    Set KensakuCode = InputKensakuCode
End Property

'使用プロシージャ
Public Sub AddUseProcedure(InputUseProcedure As ClassProcedure)
    UseProcedure.Add InputUseProcedure
End Sub

Public Function Item(Index&) As ClassProcedure
    Set Item = UseProcedure(Index)
End Function

'モジュール
Property Get MyModuleName()
    MyModuleName = ModuleName
End Property

Property Let MyModuleName(InputModuleName)
    ModuleName = InputModuleName
End Property

'VBProject
Property Get MyVBProjectName()
    MyVBProjectName = VBProjectName
End Property

Property Let MyVBProjectName(InputVBProjectName)
    VBProjectName = InputVBProjectName
End Property


'プロシージャの使用可能範囲(PublicかPrivateか）
Property Get MyRangeOfUse()
    MyRangeOfUse = RangeOfUse
End Property

Property Let MyRangeOfUse(InputRangeOfUse)
    RangeOfUse = InputRangeOfUse
End Property

'プロシージャのタイプ（Sub,Functionなど）
Property Get MyProcedureType()
    MyProcedureType = ProcedureType
End Property

Property Let MyProcedureType(InputProcedureType)
    ProcedureType = InputProcedureType
End Property

