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
Public ModuleType$ '���W���[���̃^�C�v�A�W�����W���[���A���[�U�[�t�H�[���A�N���X���W���[���Ȃ�
Public BookName$

'���O
Property Get MyName()
    MyName = Name
End Property

Property Let MyName(InputName)
    Name = InputName
End Property

'��������
Private Sub Class_Initialize()
    Set Procedures = New Collection
End Sub

Public Sub AddProcedure(InputProcedure As ClassProcedure)
    Procedures.Add InputProcedure
End Sub

'�v���V�[�W���擾
Public Function Item(Index&) As ClassProcedure
    Set Item = Procedures(Index)
End Function

'VBProject��
Property Get MyVBProjectName()
    MyVBProjectName = VBProjectName
End Property

Property Let MyVBProjectName(InputVBProjectName)
    VBProjectName = InputVBProjectName
End Property

'ModuleType��
Property Get MyModuleType()
    MyModuleType = ModuleType
End Property

Property Let MyModuleType(InputModuleType)
    ModuleType = InputModuleType
End Property

'�u�b�N�̖��O
Property Get MyBookName()
    MyBookName = BookName
End Property

Property Let MyBookName(InputBookName)
    BookName = InputBookName
End Property

