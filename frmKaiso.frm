VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKaiso 
   Caption         =   "�K�w���\���t�H�[��"
   ClientHeight    =   9072
   ClientLeft      =   36
   ClientTop       =   408
   ClientWidth     =   15396
   OleObjectBlob   =   "frmKaiso.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmKaiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�F�t���p�̗񋓌^(���W���[��)
Private Enum ModuleColor
    Module_ = rgbBlue
    Document_ = rgbGreen
    Class_ = rgbLightPink
    Form_ = rgbRed
    ActiveX_ = rgbLightGray
End Enum

'�F�t���p�̗񋓌^(�v���V�[�W��)
Private Enum ProcedureColor
    SubColor = rgbOrange
    FunctionColor = rgbGreen
    GetColor = rgbBlue
    SetColor = rgbRed
    LetColor = rgbPink
    
End Enum

'�N����VBProject�̑S���
Private PriVBProjectNameList
Private PriVBProjectList() As classVBProject
Private PriModuleList() As classModule
Private PriProcedureList() As classProcedure
Private PriUseProcedureList() As classProcedure
Private PriProcedure As classProcedure
Private PriShowProcedure As classProcedure
Private PriTreeProcedureList() As classProcedure
Private PriSearchProcedureList() As classProcedure
Private PriExtProcedureList() As classProcedure

'�e���X�g�r���[�̃w�b�_�[�̏�����
Private PriListViewColumnWidthModuleList&(1 To 2)
Private PriListViewColumnWidthProcedureList&(1 To 6)
Private PriListViewColumnWidthUseProcedureList&(1 To 6)
Private PriListViewColumnWidthCodeList&(1 To 2)

'���[�U�[�t�H�[���̏����T�C�Y
Private PriIniFormHeight#
Private PriIniFormWidth#

Private Sub Cmb�����T�C�Y_Change()
    
    '�e���X�g�\���̃R���g���[���̃t�H���g�T�C�Y�ύX
    Dim SelectedFontSize&
    SelectedFontSize = Me.Cmb�����T�C�Y.Text
    
    With Me
        .ListVBProject.Font.Size = SelectedFontSize
        .ListViewModule.Font.Size = SelectedFontSize
        .ListViewProcedure.Font.Size = SelectedFontSize
        .ListViewUseProcedure.Font.Size = SelectedFontSize
        .ListViewCode.Font.Size = SelectedFontSize
        .TreeProcedure.Font.Size = SelectedFontSize
    End With
    
End Sub

Private Sub CmdAllCodeCopy_Click()
    
    Call �R�[�h�̎g�p��܂ߑS���R�s�[

End Sub

Private Sub CmdCodeCopy_Click()
    
    Call �R�[�h�̃R�s�[
    
End Sub

Private Sub CmdExt_Click()
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim ExtProcedureList
    ExtProcedureList = �O���Q�ƃv���V�[�W�����X�g�쐬(PriVBProjectList)
    
    If Me.ListVBProject.ListIndex < 0 Then
        Exit Sub
    End If
    
    PriExtProcedureList = ExtProcedureList(Me.ListVBProject.ListIndex + 1)
    
    PriProcedureList = PriExtProcedureList
    
    Call �v���V�[�W�����X�g�r���[������
    Call �g�p�v���V�[�W�����X�g�r���[������
    
    If Not PriExtProcedureList(1) Is Nothing Then
        
        For I = 1 To UBound(PriExtProcedureList, 1)
                        
            With Me.ListViewProcedure.ListItems.Add
                .Text = PriExtProcedureList(I).Name '�v���V�[�W����
                .SubItems(1) = PriExtProcedureList(I).UseProcedure.Count '�g�p�v���V�[�W����
                .SubItems(2) = PriExtProcedureList(I).VBProjectName 'VBProject��
                .SubItems(3) = PriExtProcedureList(I).ModuleName '���W���[����
                .SubItems(4) = PriExtProcedureList(I).RangeOfUse '�v���V�[�W���g�p�\�͈�
                .SubItems(5) = PriExtProcedureList(I).ProcedureType '�v���V�[�W�����
            End With
            
            Me.ListViewProcedure.ListItems(I).ForeColor = �v���V�[�W����ނł̐F�擾(PriExtProcedureList(I).ProcedureType)
        
        Next I
        
        Me.CmdExtCodeCopy.Enabled = True
    Else
        MsgBox ("�O���Q�Ƃ��Ă���v���V�[�W���͌�����܂���ł���")
    End If

End Sub

Private Sub �O���Q�ƃv���V�[�W���R�[�h�R�s�[()
    
    If IsEmpty(PriExtProcedureList) Then Exit Sub
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim TmpCode$, TmpProcedureName$
    Dim TmpProcedureDict As Object
    Set TmpProcedureDict = CreateObject("Scripting.Dictionary")
    Dim TmpClassProcedure As classProcedure
    
    Dim SengenList
    SengenList = ���W���[���̐錾�����擾(PriExtProcedureList)
    
    Dim ProcedureItiran$
    ProcedureItiran = �v���V�[�W���ꗗ���쐬(PriExtProcedureList)
    
    TmpCode = ""
    TmpCode = "Option Explicit" & vbLf & vbLf
    TmpCode = TmpCode & ProcedureItiran & vbLf
    TmpCode = TmpCode & "'---------------------------------" & vbLf
    
    For I = 1 To UBound(SengenList, 1)
        TmpCode = TmpCode & SengenList(I) & vbLf
        TmpCode = TmpCode & "'---------------------------------" & vbLf
    Next I
    
    For I = 1 To UBound(PriExtProcedureList)
        Set TmpClassProcedure = PriProcedureList(I)
        TmpProcedureName = TmpClassProcedure.Name
        
        If TmpProcedureDict.Exists(TmpProcedureName) = False Then
            TmpProcedureDict.Add TmpProcedureName, ""
            TmpCode = TmpCode & vbLf & vbLf & TmpClassProcedure.Code
        End If
    Next I
    
    Call ClipboardCopy(TmpCode)
    MsgBox ("�O���Q�ƃv���V�[�W���̑S�R�[�h���N���b�v�{�[�h�ɃR�s�[���܂����B" & vbLf & _
            "�v���V�[�W�����F" & UBound(PriExtProcedureList, 1) & vbLf & _
            "�S�R�[�h���F" & UBound(Split(TmpCode, vbLf), 1) & vbLf & _
            "�������F" & Len(TmpCode))

End Sub

Private Sub CmdExtCodeCopy_Click()

    Call �O���Q�ƃv���V�[�W���R�[�h�R�s�[

End Sub

Private Sub CmdSwitch_Click()
    
    If Me.CmdSwitch.Caption = "�K�w�\��" Then
        '�K�w�\�����[�h�ɐ؂�ւ�
        Me.CmdSwitch.Caption = "�R�[�h�\��"
'        Me.ListViewCode.Visible = False
'        Me.ListViewCode.Height = 300
        Me.ListViewCode.Top = Me.TreeProcedure.Top + Me.TreeProcedure.Height + 1
        Me.ListViewCode.Height = Me.Height - Me.ListViewCode.Top - 32
        Me.TreeProcedure.Visible = True
        If Not PriShowProcedure Is Nothing Then
            Call �v���V�[�W���R�[�h�\��(PriShowProcedure)
            Call �c���[�r���[�Ƀv���V�[�W���̊K�w�\��(PriShowProcedure)
        End If
    Else
        '�R�[�h�\�����[�h�ɐ؂�ւ�
        Me.CmdSwitch.Caption = "�K�w�\��"
'        Me.ListViewCode.Visible = True
        Me.ListViewCode.Top = Me.CmdSwitch.Top + Me.CmdSwitch.Height + 1
        Me.ListViewCode.Height = Me.Height - Me.ListViewCode.Top - 32
        Me.TreeProcedure.Visible = False
        If Not PriShowProcedure Is Nothing Then
            Call �v���V�[�W���R�[�h�\��(PriShowProcedure)
        End If
    End If

End Sub


Private Sub listVBProject_Click()

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim TmpModuleName As String
    Dim TmpProcedureList
    Dim TmpProcedureKosu As Integer
    
    For I = 1 To Me.ListViewModule.ListItems.Count
        Me.ListViewModule.ListItems.Remove (1)
    Next I
    For I = 1 To Me.ListViewProcedure.ListItems.Count
        Me.ListViewProcedure.ListItems.Remove (1)
    Next I
    For I = 1 To Me.ListViewUseProcedure.ListItems.Count
        Me.ListViewUseProcedure.ListItems.Remove (1)
    Next I
    
    Me.txtVBProject.Text = ""
    Me.txtModule.Text = ""
    Me.txtKensaku.Text = ""
    For I = 1 To Me.ListViewCode.ListItems.Count
        Me.ListViewCode.ListItems.Remove (1)
    Next I
    
    For I = 1 To UBound(PriVBProjectList, 1)
        Select Case Me.ListVBProject.List(Me.ListVBProject.ListIndex)
            Case PriVBProjectList(I).Name
                
                ReDim PriModuleList(PriVBProjectList(I).Modules.Count)
                
                For J = 1 To UBound(PriModuleList, 1)
                    
                    Set PriModuleList(J) = PriVBProjectList(I).Modules(J)
                    
                    With Me.ListViewModule.ListItems.Add
                        .Text = PriModuleList(J).Name '���W���[����
                        .SubItems(1) = PriModuleList(J).Procedures.Count '�v���V�[�W���̌�
                        .SubItems(2) = PriModuleList(J).ModuleType '���W���[�����
                        
                    End With
                    
                    Me.ListViewModule.ListItems(J).ForeColor = ���W���[����ނł̐F�擾(PriModuleList(J).ModuleType)
                Next J
                
                Exit For
                
        End Select
    Next I

End Sub

Private Sub listViewModule_Click()
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    For I = 1 To Me.ListViewProcedure.ListItems.Count
        Me.ListViewProcedure.ListItems.Remove (1)
    Next I
    For I = 1 To Me.ListViewUseProcedure.ListItems.Count
        Me.ListViewUseProcedure.ListItems.Remove (1)
    Next I
    Me.txtVBProject.Text = ""
    Me.txtModule.Text = ""
    Me.txtKensaku.Text = ""
    
    On Error GoTo ErrorEscape
    
    For I = 1 To UBound(PriModuleList, 1)
        Select Case Me.ListViewModule.SelectedItem
            Case PriModuleList(I).Name
                If PriModuleList(I).Procedures.Count <> 0 Then
                    ReDim PriProcedureList(1 To PriModuleList(I).Procedures.Count)
                    
                    For J = 1 To UBound(PriProcedureList, 1)
                        
                        Set PriProcedureList(J) = PriModuleList(I).Procedures(J)
                        
                        With Me.ListViewProcedure.ListItems.Add
                            .Text = PriProcedureList(J).Name '�v���V�[�W����
                            .SubItems(1) = PriProcedureList(J).UseProcedure.Count '�g�p�v���V�[�W����
                            .SubItems(2) = PriProcedureList(J).VBProjectName 'VBProject��
                            .SubItems(3) = PriProcedureList(J).ModuleName '���W���[����
                            .SubItems(4) = PriProcedureList(J).RangeOfUse '�v���V�[�W���g�p�\�͈�
                            .SubItems(5) = PriProcedureList(J).ProcedureType '�v���V�[�W�����
                        End With
                        
                        Me.ListViewProcedure.ListItems(J).ForeColor = �v���V�[�W����ނł̐F�擾(PriProcedureList(J).ProcedureType)
                    
                    Next J
                    
                    Exit For
                Else
                    ReDim PriProcedureList(0 To 0)
                End If
        End Select
    Next I
      
ErrorEscape:
    Debug.Print Err.Number
    On Error GoTo 0
    
End Sub

Private Sub ListViewModule_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
            
    With ListViewModule
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = .SortOrder Xor lvwDescending
        .Sorted = True
    End With

End Sub

Private Sub ListViewProcedure_Click()
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    
    For I = 1 To Me.ListViewUseProcedure.ListItems.Count
        Me.ListViewUseProcedure.ListItems.Remove (1)
    Next I
    Me.txtVBProject.Text = ""
    Me.txtModule.Text = ""
    
    For I = 1 To Me.ListViewCode.ListItems.Count
        Me.ListViewCode.ListItems.Remove (1)
    Next I
    
    On Error GoTo ErrorEscape
    If UBound(PriProcedureList, 1) <= 0 Then
        Exit Sub
    End If
    
    For I = 1 To UBound(PriProcedureList, 1)
        Select Case Me.ListViewProcedure.SelectedItem
            Case PriProcedureList(I).Name
                
                Set PriShowProcedure = PriProcedureList(I)
                If Me.CmdSwitch.Caption = "�K�w�\��" Then
                    Call �v���V�[�W���R�[�h�\��(PriShowProcedure)
                Else
                    Call �c���[�r���[�Ƀv���V�[�W���̊K�w�\��(PriShowProcedure)
                    Call �v���V�[�W���R�[�h�\��(PriShowProcedure)
                End If
                
                If PriProcedureList(I).UseProcedure.Count <> 0 Then
                    ReDim PriUseProcedureList(1 To PriProcedureList(I).UseProcedure.Count)
                    
                    For J = 1 To UBound(PriUseProcedureList, 1)
                        
                        Set PriUseProcedureList(J) = PriProcedureList(I).UseProcedure(J)
                        
                        With Me.ListViewUseProcedure.ListItems.Add
                            .Text = PriUseProcedureList(J).Name '�v���V�[�W����
                            .SubItems(1) = PriUseProcedureList(J).UseProcedure.Count '�g�p�v���V�[�W����
                            .SubItems(2) = PriUseProcedureList(J).VBProjectName 'VBProject��
                            .SubItems(3) = PriUseProcedureList(J).ModuleName '���W���[����
                            .SubItems(4) = PriUseProcedureList(J).RangeOfUse '�v���V�[�W���g�p�\�͈�
                            .SubItems(5) = PriUseProcedureList(J).ProcedureType '�v���V�[�W�����
                        End With
                        
                        Me.ListViewUseProcedure.ListItems(J).ForeColor = �v���V�[�W����ނł̐F�擾(PriUseProcedureList(J).ProcedureType)
                    
                    Next J
                    
                    Exit For
                Else
                    ReDim PriUseProcedureList(0 To 0)
                End If
                
        End Select
    Next I
    
    Exit Sub
    
ErrorEscape:
    Debug.Print Err.Number
    On Error GoTo 0
    
End Sub

Private Sub ListViewProcedure_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With ListViewProcedure
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = .SortOrder Xor lvwDescending
        .Sorted = True
    End With
    
End Sub

Private Sub ListViewProcedure_DblClick()

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    
    On Error GoTo ErrorEscape
    If UBound(PriProcedureList, 1) <= 0 Then
        Exit Sub
    End If
    
    For I = 1 To UBound(PriProcedureList, 1)
        Select Case Me.ListViewProcedure.SelectedItem
            Case PriProcedureList(I).Name
                
                Set PriShowProcedure = PriProcedureList(I)
                
                Call �w��v���V�[�W��VBE��ʕ\��(PriShowProcedure)
                                
        End Select
    Next I
    
    
ErrorEscape:
    Debug.Print Err.Number
    On Error GoTo 0

End Sub

Private Sub ListViewUseProcedure_Click()

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    
    Me.txtVBProject.Text = ""
    Me.txtModule.Text = ""
    
    If UBound(PriUseProcedureList, 1) <= 0 Then
        Exit Sub
    End If
    
    For I = 1 To UBound(PriUseProcedureList, 1)
        Select Case Me.ListViewUseProcedure.SelectedItem
            Case PriUseProcedureList(I).Name
                
                Set PriShowProcedure = PriUseProcedureList(I)
                If Me.CmdSwitch.Caption = "�K�w�\��" Then
                    Call �v���V�[�W���R�[�h�\��(PriShowProcedure)
                Else
                    Call �c���[�r���[�Ƀv���V�[�W���̊K�w�\��(PriShowProcedure)
                    Call �v���V�[�W���R�[�h�\��(PriShowProcedure)
                End If
                
        End Select
    Next I

End Sub
Private Sub ListViewUseProcedure_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With ListViewUseProcedure
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = .SortOrder Xor lvwDescending
        .Sorted = True
    End With
    
End Sub

Private Sub ListViewUseProcedure_DblClick()

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
       
    If UBound(PriProcedureList, 1) <= 0 Then
        Exit Sub
    End If
    
    For I = 1 To UBound(PriUseProcedureList, 1)
        Select Case Me.ListViewUseProcedure.SelectedItem
            Case PriUseProcedureList(I).Name
                
                Set PriShowProcedure = PriUseProcedureList(I)
                
                Call �w��v���V�[�W��VBE��ʕ\��(PriShowProcedure)
                                
        End Select
    Next I

End Sub

Private Sub TreeProcedure_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim TmpProcedure As classProcedure
    Set TmpProcedure = PriTreeProcedureList(Node.Index)
    
    Call �v���V�[�W���R�[�h�\��(TmpProcedure)
    
End Sub


Private Sub UserForm_Activate()
    Call SetFormEnableResize
End Sub

Private Sub UserForm_Initialize()

    PriVBProjectList = �t�H�[���pVBProject�쐬
    
    Dim AllProcedureList
    AllProcedureList = �S�v���V�[�W���ꗗ�쐬(PriVBProjectList)
    Call �v���V�[�W�����̎g�p�v���V�[�W���擾(PriVBProjectList, AllProcedureList)
    
    '�t�H�[���p�p�u���b�N�ϐ��ݒ�
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    N = UBound(PriVBProjectList, 1)
    ReDim PbVBProjectNameList(1 To N)
    For I = 1 To N
        PbVBProjectNameList(I) = PriVBProjectList(I).Name
    Next I
        
    '�t�H�[���ݒ�
    Me.ListVBProject.List = PbVBProjectNameList
    
    With Me.ListViewModule '���W���[���̃��X�g�r���[�̃^�u�ݒ�
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add , "���W���[����", "���W���[����"
        .ColumnHeaders.Add , "��", "��"
        .ColumnHeaders.Add , "���", "���"
        .ColumnHeaders(2).Width = 16
        
        '�e�w�b�_�[�̕������T�C�Y�p�ɂƂ��Ă���
        PriListViewColumnWidthModuleList(1) = .ColumnHeaders(1).Width
        PriListViewColumnWidthModuleList(2) = .ColumnHeaders(2).Width
        
    End With
    
    With Me.ListViewProcedure '�v���V�[�W���̃��X�g�r���[�̃^�u�ݒ�
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add , "�v���V�[�W����", "�v���V�[�W����"
        .ColumnHeaders.Add , "��", "��"
        .ColumnHeaders.Add , "VBProject", "VBProject"
        .ColumnHeaders.Add , "���W���[��", "���W���[��"
        .ColumnHeaders.Add , "�͈�", "�͈�"
        .ColumnHeaders.Add , "���", "���"
        .ColumnHeaders(2).Width = 16
        .ColumnHeaders(3).Width = 20
        .ColumnHeaders(4).Width = 35
        .ColumnHeaders(5).Width = 25
        .ColumnHeaders(6).Width = 25
    
        '�e�w�b�_�[�̕������T�C�Y�p�ɂƂ��Ă���
        PriListViewColumnWidthProcedureList(1) = .ColumnHeaders(1).Width
        PriListViewColumnWidthProcedureList(2) = .ColumnHeaders(2).Width
        PriListViewColumnWidthProcedureList(3) = .ColumnHeaders(3).Width
        PriListViewColumnWidthProcedureList(4) = .ColumnHeaders(4).Width
        PriListViewColumnWidthProcedureList(5) = .ColumnHeaders(5).Width
        PriListViewColumnWidthProcedureList(6) = .ColumnHeaders(6).Width
    
    End With
   
    With Me.ListViewUseProcedure '�g�p�v���V�[�W���̃��X�g�r���[�̃^�u�ݒ�
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add , "�v���V�[�W����", "�v���V�[�W����"
        .ColumnHeaders.Add , "��", "��"
        .ColumnHeaders.Add , "VBProject", "VBProject"
        .ColumnHeaders.Add , "���W���[��", "���W���[��"
        .ColumnHeaders.Add , "�͈�", "�͈�"
        .ColumnHeaders.Add , "���", "���"
        .ColumnHeaders(2).Width = 16
        .ColumnHeaders(3).Width = 20
        .ColumnHeaders(4).Width = 35
        .ColumnHeaders(5).Width = 25
        .ColumnHeaders(6).Width = 25

        '�e�w�b�_�[�̕������T�C�Y�p�ɂƂ��Ă���
        PriListViewColumnWidthUseProcedureList(1) = .ColumnHeaders(1).Width
        PriListViewColumnWidthUseProcedureList(2) = .ColumnHeaders(2).Width
        PriListViewColumnWidthUseProcedureList(3) = .ColumnHeaders(3).Width
        PriListViewColumnWidthUseProcedureList(4) = .ColumnHeaders(4).Width
        PriListViewColumnWidthUseProcedureList(5) = .ColumnHeaders(5).Width
        PriListViewColumnWidthUseProcedureList(6) = .ColumnHeaders(6).Width

    End With

    With Me.ListViewCode '�R�[�h�̃��X�g�r���[�̃^�u�ݒ�
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add , "�s", "�s"
        .ColumnHeaders.Add , "�R�[�h", "�R�[�h"
        .ColumnHeaders(1).Width = 16
        .ColumnHeaders(2).Width = 1000
        
        '�e�w�b�_�[�̕������T�C�Y�p�ɂƂ��Ă���
        PriListViewColumnWidthCodeList(1) = .ColumnHeaders(1).Width
        PriListViewColumnWidthCodeList(2) = .ColumnHeaders(2).Width
    
    End With
    
    Me.ListViewCode.Visible = True
    Me.TreeProcedure.Visible = False
    Me.CmdExtCodeCopy.Enabled = False
    
    With Me.Cmb�����T�C�Y '�����T�C�Y�̃R���{�{�b�N�X�̐ݒ�
        .List = Array(6, 7, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72)
        .Text = 8
        .ListWidth = 35
        .ColumnWidths = 35
    End With
    
    Call InitializeFormResize(Me)
    
    '���T�C�Y�����p�Ƀ��[�U�[�t�H�[���̃T�C�Y������Ă���
    PriIniFormHeight = Me.Height
    PriIniFormWidth = Me.Width

End Sub


Private Sub listUseProcedure_Click()

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim TmpModuleName As String
    Dim TmpProcedureList
    Dim TmpProcedureKosu As Integer
    
    Me.txtCode.Text = ""
    
    For I = 1 To UBound(PriUseProcedureList, 1)
        Select Case Me.listUseProcedure.List(Me.listUseProcedure.ListIndex)
            Case PriUseProcedureList(I).Name
                
                Me.txtVBProject.Text = PriUseProcedureList(I).VBProjectName
                Me.txtModule.Text = PriUseProcedureList(I).ModuleName
                Me.txtCode.Text = PriUseProcedureList(I).Code

        End Select
    Next I

End Sub


Private Sub Cmd����_Click()
    Call �R�[�h�������s(Me.txtKensaku.Text)
End Sub

Private Sub Cmd����_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call �R�[�h�������s(Me.txtKensaku.Text)
    End If
End Sub

Private Sub txtKensaku_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call �R�[�h�������s(Me.txtKensaku.Text)
    End If
End Sub

Sub �R�[�h�������s(SearchStr$)

    Dim TmpVBProject As classVBProject
    Dim TmpModule As classModule
    Dim TmpProcedure As classProcedure
    
    ReDim PriSearchProcedureList(1 To 1)
    
    Dim I%, J%, II%, K%, M%, N% '�����グ�p(Integer�^)
    For I = 1 To UBound(PriVBProjectList, 1)
        Set TmpVBProject = PriVBProjectList(I)
        For J = 1 To TmpVBProject.Modules.Count
            Set TmpModule = TmpVBProject.Modules(J)
            For II = 1 To TmpModule.Procedures.Count
                Set TmpProcedure = TmpModule.Procedures(II)
                If InStr(1, StrConv(TmpProcedure.Code, vbUpperCase), StrConv(SearchStr, vbUpperCase)) > 0 Then
                    If Not PriSearchProcedureList(1) Is Nothing Then
                        ReDim Preserve PriSearchProcedureList(1 To UBound(PriSearchProcedureList, 1) + 1)
                    End If
                    
                    Set PriSearchProcedureList(UBound(PriSearchProcedureList, 1)) = TmpProcedure
                End If
            Next II
        Next J
    Next I
    
    If Not PriSearchProcedureList(1) Is Nothing Then
        '�������ʂ���
        
        '���X�g�r���[������
        For I = 1 To Me.ListViewProcedure.ListItems.Count
            Me.ListViewProcedure.ListItems.Remove (1)
        Next I
        For I = 1 To Me.ListViewUseProcedure.ListItems.Count
            Me.ListViewUseProcedure.ListItems.Remove (1)
        Next I
        
        For I = 1 To UBound(PriSearchProcedureList, 1)
            
            With Me.ListViewProcedure.ListItems.Add
                .Text = PriSearchProcedureList(I).Name '�v���V�[�W����
                .SubItems(1) = PriSearchProcedureList(I).UseProcedure.Count '�g�p�v���V�[�W����
                .SubItems(2) = PriSearchProcedureList(I).VBProjectName 'VBProject��
                .SubItems(3) = PriSearchProcedureList(I).ModuleName '���W���[����
                .SubItems(4) = PriSearchProcedureList(I).RangeOfUse '�v���V�[�W���g�p�\�͈�
                .SubItems(5) = PriSearchProcedureList(I).ProcedureType '�v���V�[�W�����
            End With
            
            Me.ListViewProcedure.ListItems(I).ForeColor = �v���V�[�W����ނł̐F�擾(PriSearchProcedureList(I).ProcedureType)
        
        Next I
        
        PriProcedureList = PriSearchProcedureList
        
        Me.Cmd����.Caption = "����" & "(" & UBound(PriSearchProcedureList, 1) & ")"
        
    Else
        MsgBox ("�u" & SearchStr & "�v" & "�����Ō�����܂���ł���")
    End If
              
End Sub

Private Function ���W���[����ނł̐F�擾(ModuleType$)
    
    Dim Output&
    Select Case ModuleType
    Case "�W�����W���[��"
        Output = ModuleColor.Module_
    Case "�N���X���W���[��"
        Output = ModuleColor.Class_
    Case "���[�U�[�t�H�[��"
        Output = ModuleColor.Form_
    Case "ActiveX �f�U�C�i"
        Output = ModuleColor.ActiveX_
    Case "Document ���W���[��"
        Output = ModuleColor.Document_
    End Select
        
    ���W���[����ނł̐F�擾 = Output
    
End Function

Private Function �v���V�[�W����ނł̐F�擾(ProcedureType$)
    
    Dim Output&
    Select Case ProcedureType
    Case "Sub"
        Output = ProcedureColor.SubColor
    Case "Function"
        Output = ProcedureColor.FunctionColor
    Case "Property Get"
        Output = ProcedureColor.GetColor
    Case "Property Let"
        Output = ProcedureColor.LetColor
    Case "Property Set"
        Output = ProcedureColor.SetColor
    End Select
        
    �v���V�[�W����ނł̐F�擾 = Output
    
End Function

Private Sub �v���V�[�W���R�[�h�\��(ShowProcedure As classProcedure)
    
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim TmpCode
    
    '������
    For I = 1 To Me.ListViewCode.ListItems.Count
        Me.ListViewCode.ListItems.Remove (1)
    Next I
    
    TmpCode = Split(ShowProcedure.Code, vbLf)
    TmpCode = Application.Transpose(Application.Transpose(TmpCode))
        
    Me.txtVBProject.Text = ShowProcedure.VBProjectName
    Me.txtModule.Text = ShowProcedure.ModuleName
    
    For I = 1 To UBound(TmpCode)
        With Me.ListViewCode.ListItems.Add
            .Text = I
            .SubItems(1) = TmpCode(I)
            
            If Me.txtKensaku.Text <> "" Then
                If InStr(1, StrConv(.SubItems(1), vbUpperCase), StrConv(Me.txtKensaku.Text, vbUpperCase)) > 0 Then
                    .ForeColor = rgbRed
                    .Bold = True
                End If
            End If
            
        End With
    Next I
End Sub

Private Sub �R�[�h�̃R�s�[()

    If Not PriShowProcedure Is Nothing Then
        Call ClipboardCopy(PriShowProcedure.Code, False)
        
        MsgBox ("�u" & PriShowProcedure.Name & "�v" & vbLf & _
               "�̃R�[�h���N���b�v�{�[�h�ɃR�s�[���܂���" & vbLf & _
               "�R�[�h���F" & UBound(Split(PriShowProcedure.Code, vbLf)) & vbLf & _
               "�������F" & Len(PriShowProcedure.Code))
        
    End If

End Sub

Private Sub �R�[�h�̎g�p��܂ߑS���R�s�[()
    
    Dim TmpProcedureName$
    Dim CopyCode$
    
    If Not PriShowProcedure Is Nothing Then
        TmpProcedureName = PriShowProcedure.VBProjectName & "." & PriShowProcedure.ModuleName & "." & PriShowProcedure.Name
        CopyCode = GetProcedureAllCode(TmpProcedureName)
        
        Call ClipboardCopy(CopyCode, False)
        
        MsgBox ("�u" & PriShowProcedure.Name & "�v" & vbLf & _
               "�̎g�p��܂߂��S�R�[�h���N���b�v�{�[�h�ɃR�s�[���܂���" & vbLf & _
               "�R�[�h���F" & UBound(Split(CopyCode, vbLf)) & vbLf & _
               "�������F" & Len(CopyCode))
        
    End If

End Sub

Private Sub �c���[�r���[�Ƀv���V�[�W���̊K�w�\��(ShowProcedure As classProcedure)
    
    '������
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    For I = Me.TreeProcedure.Nodes.Count To 1 Step -1
        Me.TreeProcedure.Nodes.Remove (I)
    Next I
    ReDim PriTreeProcedureList(1 To 1)
    Set PriTreeProcedureList(1) = ShowProcedure
    
    With Me.TreeProcedure
        .Nodes.Add Key:=ShowProcedure.Name, Text:=ShowProcedure.Name & "(" & ShowProcedure.UseProcedure.Count & ")"
        .Nodes(1).Expanded = True
        .Nodes(1).ForeColor = �v���V�[�W����ނł̐F�擾(ShowProcedure.ProcedureType)
        
    End With
    
    Me.txtVBProject.Text = ShowProcedure.VBProjectName
    Me.txtModule.Text = ShowProcedure.ModuleName
    
    Call �ċA�^�c���[�r���[�Ƀv���V�[�W���̊K�w�\��(ShowProcedure, ShowProcedure.Name, 0)
    
End Sub

Private Sub �ċA�^�c���[�r���[�Ƀv���V�[�W���̊K�w�\��(ShowProcedure As classProcedure, ParentKey$, ByVal Depth&)
        
    '�ċA�֐��̐[���i���[�v�j�����ȏ㒴���Ȃ��悤�ɂ���B
    Depth = Depth + 1
    Debug.Print String(Depth, " ") & "��" & ShowProcedure.Name
    If Depth > 15 Then
'        Debug.Print "�K�萔�̊K�w�𒴂��܂���", ShowProcedure.Name
        Exit Sub
    End If
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim TmpKey$
    Dim TmpProcedure As classProcedure
    Dim DummyNum%
    With Me.TreeProcedure
        
        If ShowProcedure.UseProcedure.Count > 0 Then
            For I = 1 To ShowProcedure.UseProcedure.Count
                Set TmpProcedure = ShowProcedure.UseProcedure(I)
        
                TmpKey = ParentKey & "_" & TmpProcedure.Name & .Nodes.Count
                
                .Nodes.Add Relative:=ParentKey, _
                           Relationship:=tvwChild, Key:=TmpKey, _
                           Text:=TmpProcedure.Name & "(" & TmpProcedure.UseProcedure.Count & ")"
                
                ReDim Preserve PriTreeProcedureList(1 To UBound(PriTreeProcedureList, 1) + 1)
                Set PriTreeProcedureList(UBound(PriTreeProcedureList, 1)) = TmpProcedure
                .Nodes(TmpKey).ForeColor = �v���V�[�W����ނł̐F�擾(TmpProcedure.ProcedureType)
                
                Call �ċA�^�c���[�r���[�Ƀv���V�[�W���̊K�w�\��(TmpProcedure, TmpKey, Depth)
                .Nodes(TmpKey).Expanded = True
                
                
            Next I
        End If
    End With
    
End Sub

Private Sub ���W���[�����X�g�r���[������()
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    For I = 1 To Me.ListViewModule.ListItems.Count
        Me.ListViewModule.ListItems.Remove (1)
    Next I

End Sub

Private Sub �v���V�[�W�����X�g�r���[������()
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    For I = 1 To Me.ListViewProcedure.ListItems.Count
        Me.ListViewProcedure.ListItems.Remove (1)
    Next I

End Sub

Private Sub �g�p�v���V�[�W�����X�g�r���[������()
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    For I = 1 To Me.ListViewUseProcedure.ListItems.Count
        Me.ListViewUseProcedure.ListItems.Remove (1)
    Next I

End Sub

Private Sub �R�[�h�v���V�[�W�����X�g�r���[������()
        Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    For I = 1 To Me.ListViewCode.ListItems.Count
        Me.ListViewCode.ListItems.Remove (1)
    Next I

End Sub

Private Sub �w��v���V�[�W��VBE��ʕ\��(ShowProcedure As classProcedure)
'https://www.relief.jp/docs/excel-vba-application-goto-reference.html
    Dim ReferenceStr$
    With ShowProcedure
        ReferenceStr = .BookName & "!" & .ModuleName & "." & .Name
    End With
    
    On Error Resume Next
    Application.Goto Reference:=ReferenceStr
    On Error GoTo 0

End Sub

Private Sub ClipboardCopy(ByVal InputClipText, Optional MessageIrunaraTrue As Boolean = False)
'���̓e�L�X�g���N���b�v�{�[�h�Ɋi�[
'�z��Ȃ�Η������Tab�킯�A�s���������s����B
'20210719�쐬
    
    '���͂����������z�񂩁A�z��̏ꍇ��1�����z�񂩁A2�����z�񂩔���
    Dim HairetuHantei%
    Dim Jigen1%, Jigen2%
    If IsArray(InputClipText) = False Then
        '���͈������z��łȂ�
        HairetuHantei = 0
    Else
        On Error Resume Next
        Jigen2 = UBound(InputClipText, 2)
        On Error GoTo 0
        
        If Jigen2 = 0 Then
            HairetuHantei = 1
        Else
            HairetuHantei = 2
        End If
    End If
    
    '�N���b�v�{�[�h�Ɋi�[�p�̃e�L�X�g�ϐ����쐬
    Dim Output$
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    
    If HairetuHantei = 0 Then '�z��łȂ��ꍇ
        Output = InputClipText
    ElseIf HairetuHantei = 1 Then '1�����z��̏ꍇ
    
        If LBound(InputClipText, 1) <> 1 Then '�ŏ��̗v�f�ԍ���1�o�Ȃ��ꍇ�͍ŏ��̗v�f�ԍ���1�ɂ���
            InputClipText = Application.Transpose(Application.Transpose(InputClipText))
        End If
        
        N = UBound(InputClipText, 1)
        
        Output = ""
        For I = 1 To N
            If I = 1 Then
                Output = InputClipText(I)
            Else
                Output = Output & vbLf & InputClipText(I)
            End If
            
        Next I
    ElseIf HairetuHantei = 2 Then '2�����z��̏ꍇ
        
        If LBound(InputClipText, 1) <> 1 Or LBound(InputClipText, 2) <> 1 Then
            InputClipText = Application.Transpose(Application.Transpose(InputClipText))
        End If
        
        N = UBound(InputClipText, 1)
        M = UBound(InputClipText, 2)
        
        Output = ""
        
        For I = 1 To N
            For J = 1 To M
                If J < M Then
                    Output = Output & InputClipText(I, J) & Chr(9)
                Else
                    Output = Output & InputClipText(I, J)
                End If
                
            Next J
            
            If I < N Then
                Output = Output & Chr(10)
            End If
        Next I
    End If
    
    
    '�N���b�v�{�[�h�Ɋi�['�Q�l https://www.ka-net.org/blog/?p=7537
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = Output
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With

    '�i�[�����e�L�X�g�ϐ������b�Z�[�W�\��
    If MessageIrunaraTrue Then
        MsgBox ("�u" & Output & "�v" & vbLf & _
                "���N���b�v�{�[�h�ɃR�s�[���܂����B")
    End If
    
End Sub

Private Function ���W���[���̐錾�����擾(UseProcedureList() As classProcedure)

    Dim AllProcedureList, VBProjectList() As classVBProject
    AllProcedureList = �S�v���V�[�W���ꗗ�쐬(PriVBProjectList)
    VBProjectList = PriVBProjectList
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(UseProcedureList, 1)
        
    'VBProject�� & ���W���[���� �ŏd��������
    Dim ModuleNameDict As Object
    Set ModuleNameDict = CreateObject("Scripting.Dictionary")
    Dim TmpModuleName$
    For I = 1 To N
        TmpModuleName = UseProcedureList(I).VBProjectName & "." & UseProcedureList(I).ModuleName
        If ModuleNameDict.Exists(TmpModuleName) = False Then
            ModuleNameDict.Add TmpModuleName, ""
        End If
    Next I
    
    Dim ModuleNameList
    ModuleNameList = ModuleNameDict.Keys
    ModuleNameList = Application.Transpose(Application.Transpose(ModuleNameList))
    
    N = UBound(ModuleNameList, 1)
    
    '�錾�����擾
    Dim SengenList, TmpClassModule As classModule
    ReDim SengenList(1 To N)
    Dim Num1&, Num2&
    
    For I = 1 To N
        TmpModuleName = ModuleNameList(I)
        For J = 1 To UBound(AllProcedureList, 1)
            If AllProcedureList(J, 1) = Split(TmpModuleName, ".")(0) And _
               AllProcedureList(J, 2) = Split(TmpModuleName, ".")(1) Then
               
                Num1 = AllProcedureList(J, 4)
                Num2 = AllProcedureList(J, 5)
                
                Set TmpClassModule = VBProjectList(Num1).Modules(Num2)
                
                SengenList(I) = TmpClassModule.Sengen
                Exit For
            End If
        Next J
    Next I
    
    'Option Explicit����������
    For I = 1 To N
        SengenList(I) = Replace(SengenList(I), "Option Explicit", "")
    Next I
    
    ���W���[���̐錾�����擾 = SengenList
    
End Function

Private Function �v���V�[�W���ꗗ���쐬(classProcedureList() As classProcedure)

    Dim TmpClassProcedure As classProcedure
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(classProcedureList, 1)
    Dim StrProcedureNameList
    ReDim StrProcedureNameList(1 To N, 1 To 2)
    
    For I = 1 To N
        Set TmpClassProcedure = classProcedureList(I)
        
        StrProcedureNameList(I, 1) = "'" & TmpClassProcedure.Name
        StrProcedureNameList(I, 2) = "���ꏊ�F" & TmpClassProcedure.VBProjectName & "." & TmpClassProcedure.ModuleName
        
    Next I
    
    Dim OutputStr$
    OutputStr = MakeAligmentedArray(StrProcedureNameList, "�E�E�E")
    
    �v���V�[�W���ꗗ���쐬 = OutputStr

End Function


'MakeAligmentedArray�E�E�E���ꏊ�FFukamiAddins3.ModAlignmentArray
'------------------------------
'������z��𐮗񂳂���1�̕�����Ƃ��ďo�͂���
'------------------------------

Private Function MakeAligmentedArray(ByVal StrArray, Optional SikiriMoji$ = "�F")
    '20210916
    '������z��𐮗񂳂���1�̕�����Ƃ��ďo�͂���
    
    Dim I&, J&, K&, M&, N&                     '�����グ�p(Long�^)
    Dim TateMin&, TateMax&, YokoMin&, YokoMax& '�z��̏c���C���f�b�N�X�ő�ŏ�
    Dim WithTableArray                         '�e�[�u���t�z��c�C�~�f�B�G�C�g�E�B���h�E�ɕ\������ۂɃC���f�b�N�X�ԍ���\�������e�[�u����ǉ������z��
    Dim NagasaList, MaxNagasaList              '�e�����̕����񒷂����i�[�A�e��ł̕����񒷂��̍ő�l���i�[
    Dim NagasaOnajiList                        '" "�i���p�X�y�[�X�j�𕶎���ɒǉ����Ċe��ŕ����񒷂��𓯂��ɂ�����������i�[
    Dim OutputStr                              '��������i�[
    
    '������������������������������������������������������
    '���͈����̏���
    Dim Jigen2%
    On Error Resume Next
    Jigen2 = UBound(StrArray, 2)
    On Error GoTo 0
    If Jigen2 = 0 Then '1�����z���2�����z��ɂ���
        StrArray = Application.Transpose(StrArray)
    End If
    
    TateMin = LBound(StrArray, 1) '�z��̏c�ԍ��i�C���f�b�N�X�j�̍ŏ�
    TateMax = UBound(StrArray, 1) '�z��̏c�ԍ��i�C���f�b�N�X�j�̍ő�
    YokoMin = LBound(StrArray, 2) '�z��̉��ԍ��i�C���f�b�N�X�j�̍ŏ�
    YokoMax = UBound(StrArray, 2) '�z��̉��ԍ��i�C���f�b�N�X�j�̍ő�
    
    
    '������������������������������������������������������
    '�e��̕��𓯂��ɐ����邽�߂ɕ����񒷂��Ƃ��̊e��̍ő�l���v�Z����B
    N = UBound(StrArray, 1) '�uStrArray�v�̏c�C���f�b�N�X���i�s���j
    M = UBound(StrArray, 2) '�uStrArray�v�̉��C���f�b�N�X���i�񐔁j
    ReDim NagasaList(1 To N, 1 To M)
    ReDim MaxNagasaList(1 To M)
    
    Dim TmpStr$
    For J = 1 To M
        For I = 1 To N
        
'            If J > 1 And HyoujiMaxNagasa <> 0 Then
'                '�ő�\���������w�肳��Ă���ꍇ�B
'                '1��ڂ̃e�[�u���͂��̂܂܂ɂ���B
'                TmpStr = StrArray(I, J)
'                StrArray(I, J) = ��������w��o�C�g���������ɏȗ�(TmpStr, HyoujiMaxNagasa)
'            End If
            
            NagasaList(I, J) = LenB(StrConv(StrArray(I, J), vbFromUnicode)) '�S�p�Ɣ��p����ʂ��Ē������v�Z����B
            MaxNagasaList(J) = WorksheetFunction.Max(MaxNagasaList(J), NagasaList(I, J))
            
        Next I
    Next J
    
    '������������������������������������������������������
    '" "(���p�X�y�[�X)��ǉ����ĕ����񒷂��𓯂��ɂ���B
    ReDim NagasaOnajiList(1 To N, 1 To M)
    Dim TmpMaxNagasa&
    
    For J = 1 To M
        TmpMaxNagasa = MaxNagasaList(J) '���̗�̍ő啶���񒷂�
            For I = 1 To N
            'Rept�c�w�蕶������w����A�����ĂȂ�����������o�͂���B
            '�i�ő啶����-�������j�̕�" "�i���p�X�y�[�X�j�����ɂ�������B
            NagasaOnajiList(I, J) = StrArray(I, J) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(I, J))
       
        Next I
    Next J
    
    '������������������������������������������������������
    '��������쐬
    OutputStr = ""
    For I = 1 To N
        For J = 1 To M
            If J = 1 Then
                OutputStr = OutputStr & NagasaOnajiList(I, J)
            Else
                OutputStr = OutputStr & SikiriMoji & NagasaOnajiList(I, J)
            End If
        Next J
        
        If I < N Then
            OutputStr = OutputStr & vbLf
        End If
    Next I
    
    ''������������������������������������������������������
    '�o��
    MakeAligmentedArray = OutputStr
    
End Function

Private Sub UserForm_Resize()
    Call ResizeForm(Me)
    
    With Me 'VBProject�̃��X�g�̃T�C�Y�����ւ񂿂����Ȃ������������
        .ListVBProject.Height = .Label1.Top - .ListVBProject.Top
        .ListVBProject.Width = .ListViewModule.Width
    End With
    
    Call ���X�g�r���[�̃w�b�_�[������(Me.ListViewModule, PriListViewColumnWidthModuleList)
    Call ���X�g�r���[�̃w�b�_�[������(Me.ListViewProcedure, PriListViewColumnWidthProcedureList)
    Call ���X�g�r���[�̃w�b�_�[������(Me.ListViewUseProcedure, PriListViewColumnWidthUseProcedureList)
    Call ���X�g�r���[�̃w�b�_�[������(Me.ListViewCode, PriListViewColumnWidthCodeList)
    
End Sub


Private Sub ���X�g�r���[�̃w�b�_�[������(ListView As ListView, HeaderWidthList)
    
    Dim NowFormWidth&
    NowFormWidth = Me.Width
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    For I = 1 To UBound(HeaderWidthList, 1)
        ListView.ColumnHeaders(I).Width = HeaderWidthList(I) * NowFormWidth / PriIniFormWidth
    Next I
    
End Sub
