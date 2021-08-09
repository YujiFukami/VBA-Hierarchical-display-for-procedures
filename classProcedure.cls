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
Public KensakuCode As Object '�g�p�v���V�[�W�������p�A�z�z��
Public UseProcedure As Collection
Public ModuleName$
Public VBProjectName$
Public RangeOfUse$ '�v���V�[�W���̎g�p�\�͈�(Public��Private���j
Public ProcedureType$ '�v���V�[�W���̃^�C�v�iSub,Function�Ȃǁj
Public BookName$

'��������
Private Sub Class_Initialize()
    Set UseProcedure = New Collection
End Sub

'���O
Property Get MyName()
    MyName = Name
End Property

Property Let MyName(InputName)
    Name = InputName
End Property

'�R�[�h
Property Get MyCode()
    MyCode = Code
End Property

Property Let MyCode(InputCode)
    Code = InputCode
End Property

'�g�p�v���V�[�W�������p�A�z�z��
Property Get MyKensakuCode()
    Set MyKensakuCode = KensakuCode
End Property

Property Set MyKensakuCode(InputKensakuCode As Object)
    Set KensakuCode = InputKensakuCode
End Property

'�g�p�v���V�[�W��
Public Sub AddUseProcedure(InputUseProcedure As ClassProcedure)
    UseProcedure.Add InputUseProcedure
End Sub

Public Function Item(Index&) As ClassProcedure
    Set Item = UseProcedure(Index)
End Function

'���W���[��
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


'�v���V�[�W���̎g�p�\�͈�(Public��Private���j
Property Get MyRangeOfUse()
    MyRangeOfUse = RangeOfUse
End Property

Property Let MyRangeOfUse(InputRangeOfUse)
    RangeOfUse = InputRangeOfUse
End Property

'�v���V�[�W���̃^�C�v�iSub,Function�Ȃǁj
Property Get MyProcedureType()
    MyProcedureType = ProcedureType
End Property

Property Let MyProcedureType(InputProcedureType)
    ProcedureType = InputProcedureType
End Property

'�u�b�N�̖��O
Property Get MyBookName()
    MyBookName = BookName
End Property

Property Let MyBookName(InputBookName)
    BookName = InputBookName
End Property

