VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKaiso 
   Caption         =   "�K�w���\���t�H�[��"
   ClientHeight    =   9432
   ClientLeft      =   36
   ClientTop       =   408
   ClientWidth     =   15480
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
'�N����VBProject�̑S���
Public PbUnLockVBProjectList
Public PbUnLockVBProjectFileNameList
Public PbModuleList
Public PbProcedureList
Public PbProcedureNameList
Public PbProcedureCodeList
Public PbShiyoProcedureListList
Public PbShiyoSakiProcedureListList
Public PbKaisoList

Public PbAllProcedureNameList
Public PbAllInfoList

'VBProjectListBox�őI������VBProject�̏��݂̂��i�[
Public PbOutputModuleList
Public PbOutputProcedureList
Public PbOutputProcedureNameList
Public PbOutputProcedureCodeList
Public PbOutputShiyoProcedureListList
Public PbOutputShiyoSakiProcedureListList
Public PbOutputKaisoList

Public PbTmpKaisoList

'Private Sub CodeListBox_Click()
'
'    Dim TmpCodeItigyo
'    TmpCodeItigyo = CodeListBox.List(CodeListBox.ListIndex)
'
'    '�I�����X�g�S�̕\�����X�g�{�b�N�X�ɑI��������s��\��
'    TotemoNagaiListBox.Clear
'    TotemoNagaiListBox.AddItem TmpCodeItigyo
'
'End Sub

Private Sub KaisoHyouji_Change()
    Dim ListNo As Integer
    ListNo = KaisoHyouji.ListIndex
'    Stop
    
    KaisoListBox.List = �K�w���X�g���w��K�w�܂ł̃��X�g�擾(PbTmpKaisoList, ListNo)
    
End Sub
Sub �R�[�h�������s()
    Dim KensakuWord As String
    KensakuWord = txtKensaku.Value
    KensakuWord = StrConv(KensakuWord, vbLowerCase)
    
    If KensakuWord = "" Then
        Exit Sub
    End If
    
    Dim KensakuWordList
    KensakuWordList = Split(KensakuWord, " ")
    KensakuWordList = Application.Transpose(KensakuWordList)
    KensakuWordList = Application.Transpose(KensakuWordList)
    
    Dim I%, J%, I2%, N%, K%
    Dim KensakuKosu%
    KensakuKosu = UBound(KensakuWordList, 1)
    
    Dim TmpCodeListList
    Dim TmpCodeList
    TmpCodeListList = PbProcedureCodeList
    Dim TmpProcedureListList
    Dim TmpProcedureList
    TmpProcedureListList = PbProcedureList
    Dim TmpSiyosakiListList
    Dim TmpSiyosakiList
    TmpSiyosakiListList = PbShiyoSakiProcedureListList
    Dim TmpSiyoListList
    Dim TmpSiyoList
    TmpSiyoListList = PbShiyoProcedureListList
    Dim TmpKaisoListList
    Dim TmpKaisoList
    TmpKaisoListList = PbKaisoList
    
    Dim TmpCode
    Dim TmpItiGyo
    Dim TmpKensakuWord As String
    Dim TmpKensakuHanteiList
    
    Dim NaiNaraTrue As Boolean
    Dim AruNaraTrue As Boolean
    
    Dim KensakuProcedureList, KensakuCodeList, KensakuSiyosakiList, KensakuSiyoList, KensakuKaisoList
    K = 0
    ReDim KensakuProcedureList(1 To 1)
    ReDim KensakuCodeList(1 To 1)
    ReDim KensakuSiyosakiList(1 To 1)
    ReDim KensakuSiyoList(1 To 1)
    ReDim KensakuKaisoList(1 To 1)
    
    For I = 1 To UBound(TmpCodeListList, 1)
        TmpCodeList = TmpCodeListList(I)
        TmpSiyosakiList = TmpSiyosakiListList(I)
        TmpSiyoList = TmpSiyoListList(I)
        TmpKaisoList = TmpKaisoListList(I)
        
        For I2 = 1 To UBound(TmpCodeList, 1)
            TmpCode = TmpCodeList(I2)
            ReDim TmpKensakuHanteiList(1 To KensakuKosu)
            For J = 1 To KensakuKosu
                TmpKensakuWord = KensakuWordList(J)
                
                For Each TmpItiGyo In TmpCode
                    If InStr(1, TmpItiGyo, TmpKensakuWord) <> 0 Then
                        TmpKensakuHanteiList(J) = 1
                        Exit For
                    End If
                Next
            Next J
            
            If WorksheetFunction.Sum(TmpKensakuHanteiList) = KensakuKosu Then
                K = K + 1
                ReDim Preserve KensakuProcedureList(1 To K)
                ReDim Preserve KensakuCodeList(1 To K)
                ReDim Preserve KensakuSiyosakiList(1 To K)
                ReDim Preserve KensakuSiyoList(1 To K)
                ReDim Preserve KensakuKaisoList(1 To K)
                TmpProcedureList = TmpProcedureListList(I)
                KensakuProcedureList(K) = TmpProcedureList(I2, 2)
                KensakuCodeList(K) = TmpCode
                KensakuSiyosakiList(K) = TmpSiyosakiList(I2)
                KensakuSiyoList(K) = TmpSiyoList(I2)
                KensakuKaisoList(K) = TmpKaisoList(I2)
                
            End If
        Next I2
    Next I
    
    PbOutputProcedureNameList = KensakuProcedureList
    PbOutputProcedureCodeList = KensakuCodeList
    PbOutputShiyoProcedureListList = KensakuSiyoList
    PbOutputShiyoSakiProcedureListList = KensakuSiyosakiList
    PbOutputKaisoList = KensakuKaisoList
    
    ProcedureListBox.List = KensakuProcedureList
    
    
End Sub
Private Sub tgl����_Click()
    Call �R�[�h�������s
    tgl����.Value = False
End Sub
Private Sub tgl����_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call �R�[�h�������s
    tgl����.Value = False
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '�T�C�Y����
    Application.WindowState = xlMaximized
    Dim TakasaWariai As Double
    TakasaWariai = 0.9
    
    Me.Zoom = 100 * (Application.Height) / Me.Height
    Me.Height = (Application.Height * TakasaWariai)
    Me.Width = (Application.Width * TakasaWariai)
    Me.Zoom = 100 * (Application.Height) / Me.Height
    Me.Height = (Application.Height * TakasaWariai)
    Me.Width = (Application.Width * TakasaWariai)
    Me.Top = 20
    Me.Left = 20

End Sub
Private Sub UserForm_Initialize()

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Dummy1, Dummy2
    Dim TmpFileName As String
    Dim TmpVBProject As Object, VBProjectCount As Byte
    Dim TmpModuleList, TmpProcedureList, TmpProcedureNameList, TmpProcedureCodeList
    Dim TmpProcedureKosu As Integer
    Dim TmpCode, TmpProcedureName As String
    Dim TmpSiyoProcedureList
    Dim TmpSiyoProcedureListList, TmpSiyosakiProcedureListList
    Dim AllSiyoProcedureListList
    Dim TmpKaiso, TmpKaisoList
    
    '�N������VBProject�̃��X�g�擾'������������������������������������������������������
    PbUnLockVBProjectList = �񃍃b�N��VBProject���X�g�擾
    VBProjectCount = UBound(PbUnLockVBProjectList, 1)
    
    'VBProject�̃t�@�C�����̃��X�g���쐬���Ă���
    ReDim PbUnLockVBProjectFileNameList(1 To VBProjectCount)
    For I = 1 To VBProjectCount
        Set Dummy1 = PbUnLockVBProjectList(I)
        PbUnLockVBProjectFileNameList(I) = Dir(Dummy1.FileName) '�t�@�C�������o
    Next I

    '�e���W���[���A�v���V�[�W���̖��O�A�R�[�h���擾'������������������������������������������������������
    ReDim PbModuleList(1 To VBProjectCount)
    ReDim PbProcedureList(1 To VBProjectCount)
    ReDim PbProcedureNameList(1 To VBProjectCount)
    ReDim PbProcedureCodeList(1 To VBProjectCount)
    
    For I = 1 To VBProjectCount
        Set TmpVBProject = PbUnLockVBProjectList(I)
        TmpProcedureList = �v���V�[�W���ꗗ�擾(TmpVBProject) '1��ڃ��W���[���A2��ڃv���V�[�W�����A3��ڃv���V�[�W���R�[�h
        TmpModuleList = ���W���[���ꗗ�擾(TmpVBProject, TmpProcedureList) '1��ڃ��W���[��(�I�u�W�F�N�g�`��)�A2��ڃ��W���[�����̃v���V�[�W�����X�g
        
        TmpProcedureKosu = UBound(TmpProcedureList, 1)
        
        ReDim TmpProcedureNameList(1 To TmpProcedureKosu)
        ReDim TmpProcedureCodeList(1 To TmpProcedureKosu)
           
        For J = 1 To TmpProcedureKosu
            TmpProcedureNameList(J) = TmpProcedureList(J, 2)
            TmpProcedureCodeList(J) = TmpProcedureList(J, 3)
        Next J
        
        '�p�u���b�N�����Ɋi�[
        PbModuleList(I) = TmpModuleList
        PbProcedureList(I) = TmpProcedureList
        PbProcedureNameList(I) = TmpProcedureNameList
        PbProcedureCodeList(I) = TmpProcedureCodeList
        
    Next I
        
    '�SVBProject�̃v���V�[�W�����ꗗ���쐬
    PbAllProcedureNameList = ���d�z������ɂ܂Ƃ߂�(PbProcedureNameList)
    
    
    Dim KensakuProcedureNameList
    ReDim KensakuProcedureNameList(1 To UBound(PbAllProcedureNameList, 1))
    For I = 1 To UBound(PbAllProcedureNameList, 1)
        KensakuProcedureNameList(I) = StrConv(PbAllProcedureNameList(I), vbLowerCase)
    Next I
    
    '�e�v���V�[�W���̎g�p�֌W���R�[�h��ǂݎ���Ď擾'������������������������������������������������������
    ReDim PbShiyoProcedureListList(1 To VBProjectCount)
    ReDim PbShiyoSakiProcedureListList(1 To VBProjectCount)
    
    '�g�p�v���V�[�W���擾
    For I = 1 To VBProjectCount
        TmpProcedureNameList = PbProcedureNameList(I)
        TmpProcedureCodeList = PbProcedureCodeList(I)
        TmpProcedureKosu = UBound(TmpProcedureNameList, 1)
        
        ReDim TmpSiyoProcedureListList(1 To TmpProcedureKosu)
        
        For J = 1 To TmpProcedureKosu
            TmpCode = TmpProcedureCodeList(J)
            TmpProcedureName = TmpProcedureNameList(J)
            TmpSiyoProcedureList = �v���V�[�W�����̎g�p�v���V�[�W���̃��X�g�擾(TmpCode, PbAllProcedureNameList, KensakuProcedureNameList, TmpProcedureName)
            TmpSiyoProcedureListList(J) = TmpSiyoProcedureList '�g�p��̃��X�g�����X�g�Ɋi�[
        Next J
        
        PbShiyoProcedureListList(I) = TmpSiyoProcedureListList
    
    Next I
    
    '�SVBProject�̎g�p�v���V�[�W�����X�g�����ɂ܂Ƃ߂�B
    AllSiyoProcedureListList = ���d�z������ɂ܂Ƃ߂�(PbShiyoProcedureListList)
    
    
    '�g�p��v���V�[�W���擾
    For I = 1 To VBProjectCount
        TmpProcedureNameList = PbProcedureNameList(I)
        TmpProcedureKosu = UBound(TmpProcedureNameList, 1)
        
        TmpSiyosakiProcedureListList = �v���V�[�W���̎g�p��̃v���V�[�W���̃��X�g�擾(TmpProcedureNameList, AllSiyoProcedureListList)
        
        PbShiyoSakiProcedureListList(I) = TmpSiyosakiProcedureListList
    
    Next I

    '�S�����ЂƂ܂Ƃ߂ɂ��Ă���
    PbAllInfoList = �S�����ЂƂ܂Ƃ߂ɂ���(PbUnLockVBProjectFileNameList, PbProcedureList, PbShiyoSakiProcedureListList)

    '�v���V�[�W���̊K�w���X�g���擾����
    ReDim PbKaisoList(1 To VBProjectCount)
    
    For I = 1 To VBProjectCount
        PbOutputShiyoProcedureListList = PbShiyoProcedureListList(I)
        PbOutputProcedureNameList = PbProcedureNameList(I)
        TmpProcedureKosu = UBound(PbOutputProcedureNameList, 1)
        
        ReDim TmpKaisoList(1 To TmpProcedureKosu)
        
        For J = 1 To TmpProcedureKosu
            TmpProcedureName = PbOutputProcedureNameList(J)
            TmpKaiso = �v���V�[�W���̊K�w�\���擾(TmpProcedureName, AllSiyoProcedureListList, _
                                                    PbAllProcedureNameList)
            TmpKaisoList(J) = TmpKaiso
        Next J
        
        PbKaisoList(I) = TmpKaisoList
        
    Next I

    'ListBox�u�N����VBProjecct�v�ɋN������VBProject�̃��X�g�o��'������������������������������������������������������
    For I = 1 To UBound(PbUnLockVBProjectList, 1)
        VBProjectListBox.AddItem PbUnLockVBProjectFileNameList(I)
    Next I
    
    '�K�w�\���̃R���{�{�b�N�X�ɃA�C�e���ǉ�
    With KaisoHyouji
        .AddItem "�S���\��"
        .AddItem "��1�K�w�܂�"
        .AddItem "��2�K�w�܂�"
        .AddItem "��3�K�w�܂�"
        .AddItem "��4�K�w�܂�"
        .AddItem "��5�K�w�܂�"
    End With

    tgl����.Value = False
    
End Sub
Private Sub VBProjectListBox_Click()

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim TmpModuleName As String
    Dim TmpProcedureList
    Dim TmpProcedureKosu As Integer
    
    ModuleListBox.Clear
    ProcedureListBox.Clear
    KaisoListBox.Clear
    ProcedureListBox.Clear
    SiyosakiListBox.Clear
    txtListBox.Text = ""
    
      
    For I = 1 To UBound(PbUnLockVBProjectFileNameList, 1)
        Select Case VBProjectListBox.List(VBProjectListBox.ListIndex)
            Case PbUnLockVBProjectFileNameList(I)
                
                '�I������VBProject�̏����i�[
                PbOutputModuleList = PbModuleList(I)
                PbOutputProcedureList = PbProcedureList(I)
                PbOutputProcedureNameList = PbProcedureNameList(I)
                PbOutputProcedureCodeList = PbProcedureCodeList(I)
                PbOutputShiyoProcedureListList = PbShiyoProcedureListList(I)
                PbOutputShiyoSakiProcedureListList = PbShiyoSakiProcedureListList(I)
                PbOutputKaisoList = PbKaisoList(I)
                                
                For J = 1 To UBound(PbOutputModuleList, 1)
                    TmpModuleName = PbOutputModuleList(J, 1).Name
                    TmpProcedureList = PbOutputModuleList(J, 2)
                    
                    If IsEmpty(TmpProcedureList) Then
                        TmpProcedureKosu = 0
                    Else
                        TmpProcedureKosu = UBound(TmpProcedureList, 1)
                    End If
                    
                    ModuleListBox.AddItem TmpModuleName & "(" & TmpProcedureKosu & ")" '���W���[�����̃v���V�[�W���̌������ɂ���
                
                Next J
                
                Exit For
                
        End Select
    Next I
    
End Sub
Private Sub ModuleListBox_Click()

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim TmpModuleBango As Integer
    Dim TmpModuleName As String
    Dim TmpProcedureList
    Dim TmpProcedureName As String
    Dim TmpKaisoList
    Dim TmpKaisoKosu As Integer
    Dim TmpListAddName As String
    
'    ModuleListBox.Clear
    ProcedureListBox.Clear
    KaisoListBox.Clear
    ProcedureListBox.Clear
    SiyosakiListBox.Clear
    txtListBox.Text = ""
        
    TmpModuleBango = ModuleListBox.ListIndex + 1
    TmpModuleName = PbOutputModuleList(TmpModuleBango, 1).Name
    TmpProcedureList = PbOutputModuleList(TmpModuleBango, 2)
    
    On Error GoTo ErrorEscape
    
    If IsEmpty(TmpProcedureList) = False Then
    
        For I = 1 To UBound(TmpProcedureList, 1)
            TmpProcedureName = TmpProcedureList(I)
            
            '�e�v���V�[�W���ŃR�[�h���Ŏg�p���Ă���v���V�[�W���̐����擾����B
            For J = 1 To UBound(PbOutputProcedureList, 1)
                If TmpProcedureName = PbOutputProcedureList(J, 2) Then
                    TmpKaisoList = PbOutputKaisoList(J)
                    TmpKaisoKosu = UBound(TmpKaisoList, 1) - 1
                    TmpListAddName = TmpProcedureName & "(" & TmpKaisoKosu & ")" '�g�p�v���V�[�W���̌������ɂ���
                    Exit For
                End If
            Next J
            
            ProcedureListBox.AddItem TmpListAddName

        Next I
    End If
 
ErrorEscape:

End Sub
Private Sub ProcedureListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '���X�g�{�b�N�X�őI���_�u���N���b�N������VBE�N�����ăR�[�h��\������B
        
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    
    Dim TmpProcedureName As String

    TmpProcedureName = ProcedureListBox.List(ProcedureListBox.ListIndex)
    TmpProcedureName = ������؂�(TmpProcedureName, "(", 1)
    
'    On Error GoTo ErrorEscape
    On Error Resume Next
    Application.Goto Reference:=TmpProcedureName
    On Error GoTo 0
'ErrorEscape:

End Sub
Private Sub ProcedureListBox_Click()
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    
    Dim TmpProcedureName As String
    Dim TmpSiyoProcedureList
    Dim TmpSiyosakiProcedureList
    Dim TmpKaisoList
    Dim TmpCode
    Dim ProcedureCode As String
    
    TmpProcedureName = ProcedureListBox.List(ProcedureListBox.ListIndex)
    TmpProcedureName = ������؂�(TmpProcedureName, "(", 1)
    
    '�I�����X�g�S�̕\�����X�g�{�b�N�X�ɑI��������s��\��
    TotemoNagaiListBox.Clear
    TotemoNagaiListBox.AddItem ProcedureListBox.List(ProcedureListBox.ListIndex)
    
    For I = 1 To UBound(PbOutputProcedureNameList, 1)
        If TmpProcedureName = PbOutputProcedureNameList(I) Then
            
            'TxtLISTBOX�ɃR�[�h�̕\��
            txtListBox.Text = ""
    
            TmpCode = PbOutputProcedureCodeList(I)
            For J = 1 To UBound(TmpCode, 1)
                ProcedureCode = ProcedureCode & TmpCode(J) & Chr(10)
        '        CodeListBox.AddItem TmpCode(I)
            Next J
            txtListBox.Text = ProcedureCode
    
'            For J = 1 To UBound(TmpCode, 1)
'                CodeListBox.AddItem TmpCode(J)
'            Next J
            
            'KaisoListBox�ɊK�w�\���̕\��
            KaisoListBox.Clear
            TmpKaisoList = PbOutputKaisoList(I)
                                                        
            For J = 1 To UBound(TmpKaisoList, 1)
                KaisoListBox.AddItem TmpKaisoList(J)
            Next J
            
            'SiyosakiListBox�Ɏg�p��v���V�[�W���̕\��
            SiyosakiListBox.Clear
            TmpSiyosakiProcedureList = PbOutputShiyoSakiProcedureListList(I)
            If IsEmpty(TmpSiyosakiProcedureList) Then
                '�������Ȃ�
            Else
                For J = 1 To UBound(TmpSiyosakiProcedureList, 1)
                    SiyosakiListBox.AddItem TmpSiyosakiProcedureList(J)
                Next J
            End If
            
            PbTmpKaisoList = TmpKaisoList
            
            Exit For
        End If
    Next I

End Sub
Private Sub KaisoListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '���X�g�{�b�N�X�őI���_�u���N���b�N������VBE�N�����ăR�[�h��\������B
        
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    
    Dim TmpProcedureName As String
    
    TmpProcedureName = KaisoListBox.List(KaisoListBox.ListIndex)
    TmpProcedureName = ������؂�(TmpProcedureName, "(", 1)
    TmpProcedureName = Replace(TmpProcedureName, "�@", "")
    TmpProcedureName = Replace(TmpProcedureName, "��", "")
    
    On Error GoTo ErrorEscape
    Application.Goto Reference:=TmpProcedureName

ErrorEscape:
End Sub
Private Sub KaisoListBox_Click()
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim TmpCode
    Dim TmpProcedureName As String
    
    TmpProcedureName = KaisoListBox.List(KaisoListBox.ListIndex)
    
    '�I�����X�g�S�̕\�����X�g�{�b�N�X�ɑI��������s��\��
    TotemoNagaiListBox.Clear
    TotemoNagaiListBox.AddItem TmpProcedureName
    
    TmpProcedureName = ������؂�(TmpProcedureName, "(", 1)
    TmpProcedureName = Replace(TmpProcedureName, "�@", "")
    TmpProcedureName = Replace(TmpProcedureName, "��", "")
    
    TmpCode = �w��v���V�[�W���̃R�[�h�擾(TmpProcedureName, PbAllInfoList)

    If IsEmpty(TmpCode) Then Exit Sub
    
    '�R�[�h�\��
    txtListBox.Text = ""
    Dim ProcedureCode
    
    For I = 1 To UBound(TmpCode, 1)
        ProcedureCode = ProcedureCode & TmpCode(I) & Chr(10)
'        CodeListBox.AddItem TmpCode(I)
    Next I
    
    txtListBox.Text = ProcedureCode
    

End Sub
Private Sub SiyosakiListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    '���X�g�{�b�N�X�őI���_�u���N���b�N������VBE�N�����ăR�[�h��\������B
        
    Dim TmpProcedureName As String
    
    TmpProcedureName = SiyosakiListBox.List(SiyosakiListBox.ListIndex)
    
    On Error GoTo ErrorEscape
    Application.Goto Reference:=TmpProcedureName

ErrorEscape:

End Sub
Private Sub SiyosakiListBox_Click()

    Dim I% '�����グ�p(Integer�^)
    Dim TmpCode
    Dim TmpProcedureName As String
    
    TmpProcedureName = SiyosakiListBox.List(SiyosakiListBox.ListIndex) '����������������������������������������
            
    '�I�����X�g�S�̕\�����X�g�{�b�N�X�ɑI��������s��\��
    TotemoNagaiListBox.Clear
    TotemoNagaiListBox.AddItem TmpProcedureName

    TmpCode = �w��v���V�[�W���̃R�[�h�擾(TmpProcedureName, PbAllInfoList)
    
    If IsEmpty(TmpCode) Then Exit Sub

    '�R�[�h�\��
    txtListBox.Text = ""
    Dim ProcedureCode As String
    
    For I = 1 To UBound(TmpCode, 1)
        ProcedureCode = ProcedureCode & TmpCode(I) & Chr(10)
'        CodeListBox.AddItem TmpCode(I)
    Next I
    
    txtListBox.Text = ProcedureCode

End Sub
