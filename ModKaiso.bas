Attribute VB_Name = "ModKaiso"
Function Kaiso()
    '�K�w�t�H�[���N��
    Kaiso = "�K�w��"
    Call frmKaiso.Show
    
End Function
Function VBProject���X�g�擾()
    '�N������VBProject�����X�g�����Ď擾����B
    '�擾����VBProject1��1�̓I�u�W�F�N�g�`���B

    Dim VBProjectList '���o��
    Dim VBProjectCount As Byte
    Dim I% '�����グ�p(Integer�^)
    Dim Dummy1
    
    VBProjectCount = ActiveWorkbook.VBProject.VBE.VBProjects.Count 'VBProject�̌��v�Z
    
    ReDim VBProjectList(1 To VBProjectCount)
    
    For I = 1 To VBProjectCount
        Set Dummy1 = ActiveWorkbook.VBProject.VBE.VBProjects.Item(I)
        Set VBProjectList(I) = Dummy1
        
    Next I
    
    '�o��
    VBProject���X�g�擾 = VBProjectList
    
End Function
Function �񃍃b�N��VBProject���X�g�擾()
    '�񃍃b�N��VBProject�����X�g�����Ď擾����B
    '�擾����񃍃b�N��VBProject1��1�̓I�u�W�F�N�g�`���B
    
    Dim VBProjectList
    Dim UnLockVBProjectList '���o��
    Dim VBProjectCount As Byte
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Dummy1, Dummy2
    
    VBProjectList = VBProject���X�g�擾
    VBProjectCount = UBound(VBProjectList, 1)
        
    K = 0 '�����グ������
    ReDim UnLockVBProjectList(1 To 1)
    
    For I = 1 To VBProjectCount
    
        Set Dummy1 = VBProjectList(I)
        Dummy2 = Dummy1.Protection
            
        If Dummy2 = 1 Then
            '���b�N����Ă���
        ElseIf Dummy2 = 0 Then
            '���b�N����Ă��Ȃ�
            K = K + 1
            ReDim Preserve UnLockVBProjectList(1 To K)
            Set UnLockVBProjectList(K) = Dummy1
            
        End If
        
    Next I
    
    '�o��
    �񃍃b�N��VBProject���X�g�擾 = UnLockVBProjectList
    
End Function
Function ���W���[���ꗗ�擾(objVBProject As Object, TmpProcedureList)
    '�w��VBProject�̃��W���[���ꗗ���擾����B
    '�擾���郂�W���[��1��1�̓I�u�W�F�N�g�`���B
    
    Dim I%, J%, K% '�����グ�p(Integer�^)
    
    Dim TmpVBProject As Object
    Set TmpVBProject = objVBProject.VBComponents
    
    Dim ModuleCount As Integer
    ModuleCount = TmpVBProject.Count
    
    Dim ModuleList '���o��
    
    ReDim ModuleList(1 To ModuleCount, 1 To 2)
    'ModuleList(:,1)'���W���[��(�I�u�W�F�N�g�`��)
    'ModuleList(:,2)'�v���V�[�W���̃��X�g
    
    Dim TmpModuleName As String, TmpProcedureNameList
    
    For I = 1 To ModuleCount
        Set ModuleList(I, 1) = TmpVBProject(I)
        TmpModuleName = TmpVBProject(I).Name
        
        K = 0
        ReDim TmpProcedureNameList(1 To 1)
        For J = 1 To UBound(TmpProcedureList, 1)
            If TmpModuleName = TmpProcedureList(J, 1) Then
                K = K + 1
                ReDim Preserve TmpProcedureNameList(1 To K)
                TmpProcedureNameList(K) = TmpProcedureList(J, 2)
            End If
        Next J
        
        If K = 0 Then
            '���W���[���Ƀv���V�[�W�����Ȃ��ꍇ
            TmpProcedureNameList = Empty
        End If
        
        ModuleList(I, 2) = TmpProcedureNameList
        
    Next I
    
    ���W���[���ꗗ�擾 = ModuleList
    
End Function
Function �v���V�[�W���ꗗ�擾(objVBProject As Object)
    
    Dim I%, J%, K%, M%, N%, K2% '�����グ�p(Integer�^)
    Dim Dummy1, Dummy2, Dummy3
    
    Dim TmpVBProject As Object
    Set TmpVBProject = objVBProject.VBComponents
    
    Dim ModuleKosuu As Integer
    ModuleKosuu = TmpVBProject.Count

    Dim Gyosuu As Integer
    Dim TmpModule As Object
    Dim ProcedureCount As Integer
    Dim CodeStartList, CodeEndList
    ReDim CodeStartList(1 To 50000)
    ReDim CodeEndList(1 To 50000)
    
    Dim Output '�擾����v���V�[�W���̐����s���Ȃ̂łƂ肠������������i�[�ł���悤�ɂ���(��)
    ReDim Output(1 To 50000, 1 To 3)
    '1�F���W���[����
    '2�F�v���V�[�W����
    '3�F�R�[�h(��������1�����z��)
    Dim AllCodeList
    
    Dim TmpStartIti%, TmpEndIti%, CodeNagasa%, TmpCode As String
    Dim FirstCodeStartIti
    
    Dim ProcedureName As String
    
    K = 0 '�����グ������
    For I = 1 To ModuleKosuu
        Set TmpModule = TmpVBProject(I)
        Gyosuu = TmpModule.CodeModule.CountOfLines '���W���[�����s��
        
        If Gyosuu = 0 Then GoTo ForEscape
        
        ReDim AllCodeList(1 To Gyosuu)
        For J = 1 To Gyosuu
            AllCodeList(J) = TmpModule.CodeModule.ProcofLine(J, 0)
        Next J
                
        ProcedureName = ""
        K2 = 0 '���W���[�����ł̐����グ�̏�����
        For J = 1 To Gyosuu
            Dummy2 = TmpModule.CodeModule.ProcofLine(J, 0)
            If ProcedureName <> TmpModule.CodeModule.ProcofLine(J, 0) Then
                K = K + 1
                K2 = K2 + 1
                ProcedureName = TmpModule.CodeModule.ProcofLine(J, 0)
                
                Output(K, 1) = TmpModule.Name '���W���[����
                Output(K, 2) = ProcedureName '�v���V�[�W����
                
                On Error Resume Next '�N���X���W���[���̎���CodeModule.ProcStartLine�Ōv�Z�ł��Ȃ��H�H
                CodeStartList(K) = TmpModule.CodeModule.ProcStartLine(ProcedureName, 0)  '�J�n�s
                On Error GoTo 0
                If IsEmpty(CodeStartList(K)) Then
                    CodeStartList(K) = J
                End If

                If K2 > 1 Then '���W���[������2�ڈڍs�̃v���V�[�W���̂�
                    CodeEndList(K - 1) = CodeStartList(K) - 1 '�I���s(1�O�Ɏ擾�����v���V�[�W���̏I���s)
                    
                    '�R�[�h�̍ŏI�s�������������_�ŃR�[�h�̎擾
                    TmpStartIti = CodeStartList(K - 1)
                    TmpEndIti = CodeEndList(K - 1)
                    CodeNagasa = TmpEndIti - TmpStartIti + 1
                    TmpCode = TmpModule.CodeModule.Lines(TmpStartIti, CodeNagasa)
                    Dummy3 = ���s���ꂽ����������s�ŕ����Ĕz��ɂ���(TmpCode)
                    Output(K - 1, 3) = �R�[�h�̐擪�󔒂����O����(Dummy3)
                                        
                End If
                
            End If
        Next J
        
        If K2 <> 0 Then 'K2<>0�Ƃ������Ƃ̓��W���[�����Ƀv���V�[�W�������݂���Ƃ������ƁB
            '���W���[�����Ō�̃v���V�[�W���̏I���s�����W���[���̍s��
            CodeEndList(K) = Gyosuu
            
            '�R�[�h�̍ŏI�s�������������_�ŃR�[�h�̎擾
            TmpStartIti = CodeStartList(K)
            TmpEndIti = CodeEndList(K)
            CodeNagasa = TmpEndIti - TmpStartIti + 1
            TmpCode = TmpModule.CodeModule.Lines(TmpStartIti, CodeNagasa)
            Dummy3 = ���s���ꂽ����������s�ŕ����Ĕz��ɂ���(TmpCode)
            Output(K, 3) = �R�[�h�̐擪�󔒂����O����(Dummy3)
            
        End If

ForEscape:
        
    Next I
        
    ProcedureCount = K  '�擾�����v���V�[�W���̌�
    
   '�擾�����v���V�[�W���̌����̔z��ɂ���B
    Dim Output2
    ReDim Output2(1 To ProcedureCount, 1 To UBound(Output, 2))
    
    For I = 1 To ProcedureCount
        For J = 1 To UBound(Output, 2)
            Output2(I, J) = Output(I, J)
        Next J
    Next I
    
    �v���V�[�W���ꗗ�擾 = Output2
    
End Function
Function ���s���ꂽ����������s�ŕ����Ĕz��ɂ���(Mojiretu)
    Dim Hairetu
    Hairetu = Split(Mojiretu, Chr(10))
    Hairetu = Application.Transpose(Hairetu)
    Hairetu = Application.Transpose(Hairetu)
    
    Dim I%
    For I = 1 To UBound(Hairetu, 1)
        Hairetu(I) = Replace(Hairetu(I), Chr(13), "")
    Next I
    
    ���s���ꂽ����������s�ŕ����Ĕz��ɂ��� = Hairetu

End Function
Function �R�[�h�̐擪�󔒂����O����(CodeHairetu)
    Dim I%
    Dim KuhakuKosu%
    Dim CodeNagasa%
    CodeNagasa = UBound(CodeHairetu, 1)
    
    Dim TmpItiGyo As String
    
    For I = 1 To CodeNagasa
        TmpItiGyo = CodeHairetu(I)
        TmpItiGyo = Replace(TmpItiGyo, Chr(13), "") '���s������
        TmpItiGyo = Replace(TmpItiGyo, Chr(10), "") '���s������
        TmpItiGyo = Replace(TmpItiGyo, " ", "") '�󔒂�����
        If TmpItiGyo <> "" Then
            KuhakuKosu = I - 1
            Exit For
        End If
    Next I
    
    Dim RealCodeNagasa
    RealCodeNagasa = CodeNagasa - KuhakuKosu
    
    Dim Output
    ReDim Output(1 To RealCodeNagasa)
    
    For I = 1 To RealCodeNagasa
        Output(I) = CodeHairetu(I + KuhakuKosu)
    Next I
    
    '�o��
    �R�[�h�̐擪�󔒂����O���� = Output
    
End Function
Function �v���V�[�W�����̎g�p�v���V�[�W���̃��X�g�擾(InputCode, ProcedureNameList, KensakuProcedureNameList, ProcedureOfCode As String)
    '20210428�C��
    Dim I%, J%, K% '�����グ�p(Integer�^)
    Dim CodeNagasa%
    CodeNagasa = UBound(InputCode, 1)

    Dim ProcedureKosu%
    ProcedureKosu = UBound(ProcedureNameList, 1)
    
    Dim SiyoProcedureList
    
    Dim HikakuCodeItigyo As String, HikakuProcedureName As String
    Dim ProcedureAruNaraTrue As Boolean
    
    Dim KensakuCode '�����p�ɕ�����ϊ������R�[�h
    KensakuCode = �����p�R�[�h������ϊ�(InputCode)
    
    
    K = 0
    For I = 1 To ProcedureKosu
        HikakuProcedureName = KensakuProcedureNameList(I)

        For J = 1 To CodeNagasa
            HikakuCodeItigyo = KensakuCode(J)
            If HikakuProcedureName <> StrConv(ProcedureOfCode, vbLowerCase) Then '20210428�C��
                '�R�[�h���g�̃v���V�[�W���͑ΏۂƂ��Ȃ�
                ProcedureAruNaraTrue = �R�[�h��s���Ƀv���V�[�W�������邩����(HikakuCodeItigyo, HikakuProcedureName)
                
                If ProcedureAruNaraTrue = True Then
                    K = K + 1
                    If K = 1 Then
                        ReDim SiyoProcedureList(1 To K)
                    Else
                        ReDim Preserve SiyoProcedureList(1 To K)
                    End If
                    
                    SiyoProcedureList(K) = ProcedureNameList(I)
                    
                    Exit For '�R�[�h���Ɍ��������̂ł��̃v���V�[�W���̌����͏I��
                End If
            End If
        Next J
    Next I
    
    '�o��
    �v���V�[�W�����̎g�p�v���V�[�W���̃��X�g�擾 = SiyoProcedureList

End Function
Function �����p�R�[�h������ϊ�(InputCode)
    '�����p�ɃR�[�h�������ϊ�����
    
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim TmpCodeItiretu As String
    
    N = UBound(InputCode, 1)
    
    Dim Output
    Output = InputCode
    
    Dim FirstMoji As String
    
    For I = 1 To N
        TmpCodeItiretu = InputCode(I)
        
         '�擪�̋󔒂����O����B
        FirstMoji = Mid(TmpCodeItiretu, 1, 1)
        Do While FirstMoji = " "
            TmpCodeItiretu = Mid(TmpCodeItiretu, 2)
            FirstMoji = Mid(TmpCodeItiretu, 1, 1)
        Loop
        
        '�������ɂ���
        TmpCodeItiretu = StrConv(TmpCodeItiretu, vbLowerCase)
        
        ' �X�y�[�X�����������p�ɒu��������
        TmpCodeItiretu = Replace(TmpCodeItiretu, " ", "@")
        
        ' ���s������
        TmpCodeItiretu = Replace(TmpCodeItiretu, Chr(13), "")
        
        Output(I) = TmpCodeItiretu
        
    Next I
    
    �����p�R�[�h������ϊ� = Output

End Function
Function �R�[�h��s���Ƀv���V�[�W�������邩����(CodeItigyo As String, ProcedureName As String) As Boolean

'    Dim KensakuProcedureName
    Dim FirstMoji As String
    Dim ProcedureAruNaraTrue As Boolean
    Dim MojiIti As Integer, MojiLastIti As Integer
    Dim HitotumaeMoji As String, HitotuAtoMoji As String
    
    '�����p�ɃR�[�h���߂��Ⴍ����ϊ�����B'������������������������������������������������������
'    CodeItigyo = CodeItigyo
    
'    '�擪�̋󔒂����O����B
'    FirstMoji = Mid(CODEITIGYO, 1, 1)
'    Do While FirstMoji = " "
'        CODEITIGYO = Mid(CODEITIGYO, 2, Len(CODEITIGYO))
'        FirstMoji = Mid(CODEITIGYO, 1, 1)
'    Loop
'
'    '�������ɂ���
'    CODEITIGYO = StrConv(CODEITIGYO, vbLowerCase)
'
'    ' �X�y�[�X�����������p�ɒu��������
'    CODEITIGYO = Replace(CODEITIGYO, " ", "@")
'
'    ' ���s������
'    CODEITIGYO = Replace(CODEITIGYO, Chr(13), "")
'
    '�����v���V�[�W���̕ϊ�'������������������������������������������������������
'    KensakuProcedureName = ProcedureName 'StrConv(ProcedureName, vbLowerCase)
        
        
    '����'������������������������������������������������������
    ProcedureAruNaraTrue = False '���菉����
    
    If CodeItigyo = "" Or Mid(CodeItigyo, 1, 1) = "'" Then
        '�����@�F�\�����󔒁A�܂��͐擪������"'"�Ŗ����\���E�E�E�ł͂Ȃ��B
    ElseIf Mid(CodeItigyo, 1, 3) = "dim" Or _
            Mid(CodeItigyo, 1, 5) = "redim" Then
        '�����A�F������`�E�E�E�ł͂Ȃ��B
    ElseIf Mid(CodeItigyo, 1, 9) = "function@" Or _
            Mid(CodeItigyo, 1, 4) = "sub@" Or _
            Mid(CodeItigyo, 1, 8) = "private@" Then
        '�����B�F�v���V�[�W���̖`���E�E�E�ł͂Ȃ�
    ElseIf Mid(CodeItigyo, 1, 11) = "endfunction" Or _
            Mid(CodeItigyo, 1, 6) = "endsub" Then
        '�����C�F�v���V�[�W���̍Ō�E�E�E�ł͂Ȃ�
        
    ElseIf Len(CodeItigyo) < Len(ProcedureName) Then
        '�����D�F�R�[�h�̒������v���V�[�W���̒����ȉ�

    Else
    
        MojiIti = InStr(CodeItigyo, ProcedureName)
        MojiLastIti = MojiIti + Len(ProcedureName) - 1
       
        If MojiIti <> 0 Then
            '�����D�F�����v���V�[�W���̖��O�̕����񂪁A�R�[�h���ɑ��݂���B
            
            HitotumaeMoji = "" '������
            HitotuAtoMoji = "" '������
            
            If MojiIti > 1 Then
                '�����v���V�[�W�������镶�����1�O�̕���
                HitotumaeMoji = Mid(CodeItigyo, MojiIti - 1, 1)
            End If
            
            If MojiLastIti < Len(CodeItigyo) Then
                '�����v���V�[�W�������镶�����1��̕���
                HitotuAtoMoji = Mid(CodeItigyo, MojiLastIti + 1, 1)
            End If
            
            If HitotumaeMoji = "" Then
                '�O��ɕ������Ȃ��E�E�ECall�̕t���Ă��Ȃ��v���V�[�W��(Call�͕t���悤�ˁI)
                If HitotuAtoMoji = "" Or HitotuAtoMoji = "'" Then
                    '���ɕ��������A�������̓R�����g
                    '��FSubSample
                    '��FSubSample'�R�����g
                    ProcedureAruNaraTrue = True
                ElseIf HitotuAtoMoji = "(" Then
                    '��FSubSample(Input)
                    ProcedureAruNaraTrue = True
                End If
                
            ElseIf HitotumaeMoji = "@" Then
                '1�O����" "
                If HitotuAtoMoji = "" Or HitotuAtoMoji = "'" Then
                    '���ɕ��������A�������̓R�����g
                    '��FDummy = FunctionSample
                    '��FCall SubSample
                    '��FDummy = FunctionSample'�R�����g
                    '��FCall SubSample'�R�����g
                    ProcedureAruNaraTrue = True
                ElseIf HitotuAtoMoji = "(" Then
                    '��FDummy = FunctionSample(Input�j
                    '��FCall SubSample(Input�j
                    ProcedureAruNaraTrue = True
                ElseIf HitotuAtoMoji = "," Or HitotuAtoMoji = ")" Then
                    '�v���V�[�W���̈����Ŏg�p���Ă���B
                    '��FDummy = FunctionSample1(Input1, FunctionSample2, Input2)
                    '��FDummy = FunctionSample1(Input, FunctionSample2)
                    ProcedureAruNaraTrue = True
                ElseIf HitotuAtoMoji = "@" Then
                    '�O�オ�󔒁B�u�������Ɩ������ǁv
                    '��FIf FunctionSample = Hikaku Then
                    ProcedureAruNaraTrue = True
                End If
                            
            ElseIf HitotumaeMoji = "(" Then
                '��O��"("
                If HitotuAtoMoji = ")" Or HitotuAtoMoji = "," Then
                    '�v���V�[�W���̈����Ő擪
                    '��FDummy = FunctionSample1(FunctionSample2, Input)
                    ProcedureAruNaraTrue = True
                End If

            End If
        End If
    End If
    
    '�o��
    �R�[�h��s���Ƀv���V�[�W�������邩���� = ProcedureAruNaraTrue

End Function
Function �v���V�[�W���̎g�p��̃v���V�[�W���̃��X�g�擾(ProcedureList, SiyoProcedureListList)
    Dim I%, J%, K%, M%, N%, I2% '�����グ�p(Integer�^)
    Dim ProcedureKosu As Integer
    Dim TmpProcedureName As String, TmpList
    Dim TmpKakunouList
    Dim Output
    ProcedureKosu = UBound(ProcedureList, 1)
        
    ReDim Output(1 To ProcedureKosu)
    
    For I = 1 To ProcedureKosu
        TmpProcedureName = ProcedureList(I)
                
        TmpKakunouList = Empty '�i�[����z���������
        K = 0 '�����グ������
        For J = 1 To ProcedureKosu
            If I <> J Then '���g�͒��ׂȂ�
                TmpList = SiyoProcedureListList(J)
            
                If IsArray(TmpList) Then
                    For I2 = 1 To UBound(TmpList, 1)
                        If TmpList(I2) = TmpProcedureName Then
                            K = K + 1
                            If K = 1 Then
                                ReDim TmpKakunouList(1 To K)
                            Else
                                ReDim Preserve TmpKakunouList(1 To K)
                            End If
                            
                            TmpKakunouList(K) = ProcedureList(J)
                            
                            Exit For
                            
                        End If
                    Next I2
                End If
            End If
        Next J
        Output(I) = TmpKakunouList
    Next I
    
    �v���V�[�W���̎g�p��̃v���V�[�W���̃��X�g�擾 = Output
        
End Function
Sub �������R�s�[()
    
    Dim ModuleKosu As Integer
    Dim I As Integer
    Dim ProcedureKosu As Integer
    Dim ProcedureKosuAdd As Integer
    Dim strDummy As String
    Dim strDummy1 As String
    Dim IntDummy1 As Integer
    Dim IntDummy2 As Integer
    Dim HairetuDummy As Variant
    
    Dim ProcedureNameItiran
    
    '�T�C�Y����
    Application.WindowState = xlMaximized
    Dim TakasaWariai
    
    'TakasaWariai = 0.9
    'Me.Zoom = Me.Zoom * ((Application.Width * TakasaWariai) / Me.Width)
    'Me.Height = (Application.Height * TakasaWariai)
    'Me.Width = (Application.Width * TakasaWariai)
    'Stop
    
    
    '�A�h�C���Ɩ{�u�b�N��VBE�ԍ��擾
    Dim strFileName As String
    strFileName = ActiveWorkbook.FullName
    Dim VBEFileName As Variant
    Dim VBEFileName2 As Variant
    Dim VBECount As Integer
    
    VBECount = ActiveWorkbook.VBProject.VBE.VBProjects.Count
    ReDim VBEFileName(1 To VBECount)
    ReDim VBEFileName2(1 To VBECount)
    For I = 1 To VBECount
        VBEFileName(I) = ActiveWorkbook.VBProject.VBE.VBProjects.Item(I).FileName
        VBEFileName2(I) = F_FileName2(VBEFileName(I))
    Next I
    
    Dim ThisNum As Integer '�{�u�b�N��VBE�ԍ�
    Dim AddinNum As Integer '�A�h�C����VBE�ԍ�
    For I = 1 To VBECount
        If VBEFileName2(I) = F_FileName2(ActiveWorkbook.FullName) Then
            ThisNum = I
        ElseIf VBEFileName2(I) = "FukamiAddIns2" Then '����������������������������������������
            AddinNum = I
        End If
    Next I
        
    BookFileName = ActiveWorkbook.Name
    AddinFileName = "FukamiAddIns2.xla" '����������������������������������������
    
    Dim VBNum As Integer
    VBNum = AddinNum '����������������������������������������������\������VBProject�̔ԍ�
    
    '���擾
    ModuleItiran = F_���W���[���ꗗ�擾(VBNum)
    ModuleKosu = UBound(ModuleItiran, 1)
    ProcedureItiranAdd = F_�v���V�[�W�������ʒu�擾(AddinNum) '����������������������������������������
    ProcedureItiran = F_�v���V�[�W�������ʒu�擾(VBNum) '����������������������������������������
    ProcedureKosu = UBound(ProcedureItiran, 1)
    ProcedureKosuAdd = UBound(ProcedureItiranAdd, 1)
    ReDim ProcedureSetumeiItiran(1 To ProcedureKosu)
    ReDim ProcedureSetumeiItiranAdd(1 To ProcedureKosuAdd)
    ReDim ProcedureKobunItiran(1 To ProcedureKosu)
    ReDim ProcedureShiyoProcedureItiran(1 To ProcedureKosu)
    ReDim ProcedureShiyoSakiItiran(1 To ProcedureKosu)
    ReDim ProcedureNameItiran(1 To ProcedureKosu)
    ReDim ProcedureNameItiranAdd(1 To ProcedureKosuAdd) '�A�h�C���̃v���V�[�W����
    
    For I = 1 To ProcedureKosu
        ProcedureNameItiran(I) = ProcedureItiran(I, 4)
    Next I
    For I = 1 To ProcedureKosuAdd '20180111�C��
        ProcedureNameItiranAdd(I) = ProcedureItiranAdd(I, 4)
    Next I
    
    'Stop
    
    For I = 1 To ProcedureKosu
        strDummy1 = ProcedureItiran(I, 2)
        IntDummy1 = ProcedureItiran(I, 7)
        IntDummy2 = ProcedureItiran(I, 8)
        ProcedureSetumeiItiran(I) = _
            F_�v���V�[�W���\���擾(strDummy1, _
            IntDummy1, IntDummy2, VBNum)
    
        strDummy1 = ProcedureNameItiran(I) '�A�h�C���̃v���V�[�W����
        HairetuDummy = ProcedureSetumeiItiran(I) '�����Ώۃv���V�[�W���̍\��
        ProcedureShiyoProcedureItiran(I) = _
            F_�v���V�[�W�����g�p����v���V�[�W���擾(ProcedureNameItiranAdd, _
            HairetuDummy, strDummy1)
            
    Next I
    
    
    For I = 1 To ProcedureKosuAdd
        strDummy1 = ProcedureItiranAdd(I, 2)
        IntDummy1 = ProcedureItiranAdd(I, 7)
        IntDummy2 = ProcedureItiranAdd(I, 8)
        ProcedureSetumeiItiranAdd(I) = _
            F_�v���V�[�W���\���擾(strDummy1, _
            IntDummy1, IntDummy2, AddinNum)
    Next I
    
    
    For I = 1 To ProcedureKosu
        strDummy1 = ProcedureNameItiran(I)
        ProcedureShiyoSakiItiran(I) = _
            F_�v���V�[�W���g�p��v���V�[�W���擾(ProcedureShiyoProcedureItiran, _
            ProcedureNameItiran, strDummy1)
       
    Next I
    
    
    Dim N As Integer
    N = UBound(ModuleItiran, 1)
    
    With ModuleListBox
    
        For I = 1 To N
            .AddItem ModuleItiran(I)
        Next I
        
    End With
    
    '�L���v�V�����Ƀo�[�W�����\��'20180111�ǉ�
    Dim Version As String
    Version = F_Version()
'    Me.Caption = "�A�h�C���K�w�\��" & " " & Version

End Sub
Function ���d�z������ɂ܂Ƃ߂�(TajuHairetu)
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    N = UBound(TajuHairetu, 1)
    
    Dim TmpHairetu
    Dim Output
    K = 0
    ReDim Output(1 To 1)
    
    For I = 1 To N
        TmpHairetu = TajuHairetu(I)
        M = UBound(TmpHairetu, 1)
        
        For J = 1 To M
            K = K + 1
            ReDim Preserve Output(1 To K)
            Output(K) = TmpHairetu(J)
        Next J
    Next I
    
    '�o��
    ���d�z������ɂ܂Ƃ߂� = Output

End Function
Function �w��v���V�[�W���̃R�[�h�擾(ProcedureName As String, AllInfoList)
    Dim I% '�����グ�p(Integer�^)
    Dim Output
    
    For I = 1 To UBound(AllInfoList, 1)
        If ProcedureName = AllInfoList(I, 3) Then
            Output = AllInfoList(I, 4)
            Exit For
        End If
    Next I
    
    �w��v���V�[�W���̃R�[�h�擾 = Output
    
End Function
Function �w��v���V�[�W���̎g�p��擾(ProcedureName, ProcedureNameList, SiyosakiProcedureList)
    Dim I, K, J, M, N
    Dim Output
    
    For I = 1 To UBound(ProcedureList, 1)
        If ProcedureName = SiyosakiProcedureList(I) Then
            Stop
'            Output = PbProcedureCodeList(i)
            Exit For
        End If
    Next I
    
    �w��v���V�[�W���̎g�p��擾 = Output
    
End Function
Function �S�����ЂƂ܂Ƃ߂ɂ���(VBProjectFileNameList, ProcedureList, SiyosakiProcedureList)

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim ProcedureKosu As Integer
    Dim Output
'    Output(*,1)�FVBProject���O
'    Output(*,2)�F���W���[����
'    Output(*,3)�F�v���V�[�W����
'    Output(*,4)�F�v���V�[�W���̃R�[�h
'    Output(*,5)�F�v���V�[�W���̎g�p�惊�X�g
    
    Dim TmpKakunoHairetu
    ReDim TmpKakunoHairetu(1 To 50000, 1 To 5)
    
    Dim VBProjectKosu As Byte
    VBProjectKosu = UBound(VBProjectFileNameList, 1)
    
    Dim TmpVBProjectFileName As String
    Dim TmpProcedureInfo, TmpSiyosakiProcedureList
    
    K = 0
    For I = 1 To VBProjectKosu
        TmpVBProjectFileName = VBProjectFileNameList(I)
        TmpProcedureInfo = ProcedureList(I)
        TmpSiyosakiProcedureList = SiyosakiProcedureList(I)
    
            
        For J = 1 To UBound(TmpProcedureInfo, 1)
            K = K + 1
            TmpKakunoHairetu(K, 1) = TmpVBProjectFileName
            TmpKakunoHairetu(K, 2) = TmpProcedureInfo(J, 1)
            TmpKakunoHairetu(K, 3) = TmpProcedureInfo(J, 2)
            TmpKakunoHairetu(K, 4) = TmpProcedureInfo(J, 3)
            TmpKakunoHairetu(K, 5) = TmpSiyosakiProcedureList(J)
        Next J
    Next I
    
    '�K�v���������o��
    ProcedureKosu = K
    ReDim Output(1 To ProcedureKosu, 1 To 5)
    
    For I = 1 To ProcedureKosu
        For J = 1 To UBound(Output, 2)
            Output(I, J) = TmpKakunoHairetu(I, J)
        Next J
    Next I
    
    '�o��
    �S�����ЂƂ܂Ƃ߂ɂ��� = Output
    

End Function
Function �v���V�[�W���̊K�w�\���擾(ProcedureName As String, SiyoSakiItiran, ProcedureItiran)

    Dim ProcedureBango As Integer
    
    ProcedureBango = �v���V�[�W���̔ԍ��擾(ProcedureName, ProcedureItiran)
    
    Dim I%, K%, J%, M%, N%
    Dim I1%, I2%, I3%, I4%, I5%
    Dim TmpBango%, TmpName As String
    Dim SiyosakiList1
    Dim SiyosakiList2
    Dim SiyosakiList3
    Dim SiyosakiList4
    Dim SiyosakiList5
    Dim SiyosakiList6
    SiyosakiList1 = SiyoSakiItiran(ProcedureBango)
        
    K = 1
    Dim Output
    ReDim Output(1 To 1)
    Output(1) = ProcedureName
    
    If IsEmpty(SiyosakiList1) Then
        '�������Ȃ�
    Else
        '��1�K�w
        For I1 = 1 To UBound(SiyosakiList1, 1)
            TmpName = SiyosakiList1(I1)
            K = K + 1
            ReDim Preserve Output(1 To K)
            Output(K) = "��" & TmpName
            
            TmpBango = �v���V�[�W���̔ԍ��擾(TmpName, ProcedureItiran)
            If IsEmpty(TmpBango) Then
                SiyosakiList2 = Empty
            Else
                SiyosakiList2 = SiyoSakiItiran(TmpBango)
            End If
            
            If IsEmpty(SiyosakiList2) Then
                '�������Ȃ�
            Else
                '��2�K�w
                For I2 = 1 To UBound(SiyosakiList2, 1)
                    TmpName = SiyosakiList2(I2)
                    K = K + 1
                    ReDim Preserve Output(1 To K)
                    Output(K) = "�@��" & TmpName
                    TmpBango = �v���V�[�W���̔ԍ��擾(TmpName, ProcedureItiran)
                    If IsEmpty(TmpBango) Then
                        SiyosakiList3 = Empty
                    Else
                        SiyosakiList3 = SiyoSakiItiran(TmpBango)
                    End If
                    
                    If IsEmpty(SiyosakiList3) Then
                        '�������Ȃ�
                    Else
                        '��3�K�w
                        For I3 = 1 To UBound(SiyosakiList3, 1)
                            TmpName = SiyosakiList3(I3)
                            K = K + 1
                            ReDim Preserve Output(1 To K)
                            Output(K) = "�@�@��" & TmpName
                            TmpBango = �v���V�[�W���̔ԍ��擾(TmpName, ProcedureItiran)
                            If IsEmpty(TmpBango) Then
                                SiyosakiList4 = Empty
                            Else
                                SiyosakiList4 = SiyoSakiItiran(TmpBango)
                            End If
                            
                            If IsEmpty(SiyosakiList4) Then
                                '�������Ȃ�
                            Else
                                '��4�K�w
                                For I4 = 1 To UBound(SiyosakiList4, 1)
                                    TmpName = SiyosakiList4(I4)
                                    K = K + 1
                                    ReDim Preserve Output(1 To K)
                                    Output(K) = "�@�@�@��" & TmpName
                                    TmpBango = �v���V�[�W���̔ԍ��擾(TmpName, ProcedureItiran)
                                    If IsEmpty(TmpBango) Then
                                        SiyosakiList5 = Empty
                                    Else
                                        SiyosakiList5 = SiyoSakiItiran(TmpBango)
                                    End If
                                    If IsEmpty(SiyosakiList5) Then
                                        '�������Ȃ�
                                    Else
                                        '��5�K�w
                                        For I5 = 1 To UBound(SiyosakiList5, 1)
                                            TmpName = SiyosakiList5(I5)
                                            K = K + 1
                                            ReDim Preserve Output(1 To K)
                                            Output(K) = "�@�@�@�@��" & TmpName
'                                            SiyosakiList5 = SIYOSAKIITIRAN(TmpBango)
                                            
'                                            If SiyosakiList5(1) = "" Then
'                                                Exit For
'                                            Else

'                                            �����܂�
                                                
                                        Next I5
                                    End If
                                Next I4
                            End If
                        Next I3
                    End If
                Next I2
            End If
        Next I1
    End If
    
    �v���V�[�W���̊K�w�\���擾 = Output
    
    
End Function
Function �v���V�[�W���̔ԍ��擾(ProcedureName As String, ProcedureNameList) As Integer
    Dim I% '�����グ�p(Integer�^)
    Dim Output
    For I = 1 To UBound(ProcedureNameList, 1)
        If ProcedureName = ProcedureNameList(I) Then
            Output = I
            Exit For
        End If
    Next I
    
    �v���V�[�W���̔ԍ��擾 = Output
    
End Function
Function ������؂�(Mojiretu, KugiriMoji As String, OutputMojiretuBango As Byte) As String
    '��������w�蕶���ŋ�؂����Ƃ��́A�w��ԍ��̕������Ԃ��B
    
    Dim KugiriHairetu
    KugiriHairetu = Split(Mojiretu, KugiriMoji)
    KugiriHairetu = Application.Transpose(KugiriHairetu)
    KugiriHairetu = Application.Transpose(KugiriHairetu)
    
    Dim Output As String
    Output = KugiriHairetu(OutputMojiretuBango)
    
    ������؂� = Output
    
End Function
Function �K�w���X�g�̕t���֌W�v�Z(KaisoList)
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    N = UBound(KaisoList, 1)
    
    Dim KaisoNumList
    ReDim KaisoNumList(1 To N)
    
    Dim TmpOkikaeMojiretu As String
    Dim TmpList
    Dim KakuteiNaraTrue As Boolean
    
    For I = 1 To N
        TmpList = KaisoList(I)
        KakuteiNaraTrue = False
        K = -1
        TmpOkikaeMojiretu = "��"
        Do While KakuteiNaraTrue = False
            K = K + 1
            If K = 0 Then
                If Mid(TmpList, 1, 1) <> "��" And Mid(TmpList, 1, 1) <> "�@" Then
                    KakuteiNaraTrue = True
                    Exit Do
                End If
            ElseIf Mid(TmpList, 1, K) = TmpOkikaeMojiretu Then
                KakuteiNaraTrue = True
                Exit Do
            Else
                TmpOkikaeMojiretu = "�@" & TmpOkikaeMojiretu
            End If
            
        Loop
        
        KaisoNumList(I) = K + 1
    Next I
    
    Dim HuzokuKosuList
    ReDim HuzokuKosuList(1 To N)
    Dim TmpHuzokuKosu As Integer, TmpKaisoNum As Integer
    
    
    For I = 1 To N
        TmpKaisoNum = KaisoNumList(I)
        TmpHuzokuKosu = 0
        
        For J = I + 1 To N
            If KaisoNumList(J) = TmpKaisoNum + 1 Then
                TmpHuzokuKosu = TmpHuzokuKosu + 1
            ElseIf KaisoNumList(J) <= TmpKaisoNum - 1 Then
                Exit For
            End If
        Next J
        
        HuzokuKosuList(I) = TmpHuzokuKosu
    Next I
    
    '�o��
    Dim Output
    ReDim Output(1 To 2)
    Output(1) = KaisoNumList
    Output(2) = HuzokuKosuList
    
    �K�w���X�g�̕t���֌W�v�Z = Output
    
End Function
Function �K�w���X�g���w��K�w�܂ł̃��X�g�擾(KaisoList, ByVal SiteiKaisoNum)
    Dim KaisoNumList
    Dim HuzokuKosuList
    Dim Dummy1
    
    Dummy1 = �K�w���X�g�̕t���֌W�v�Z(KaisoList)
    KaisoNumList = Dummy1(1)
    HuzokuKosuList = Dummy1(2)
    
    Dim Output
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    N = UBound(KaisoList, 1)
    
    Dim MaxKaisoNum%
    MaxKaisoNum = WorksheetFunction.Max(KaisoNumList)
    
    If SiteiKaisoNum = 0 Then '���ׂĕ\������ɐݒ肷��B
        SiteiKaisoNum = MaxKaisoNum
    End If
    
    K = 0
    ReDim Output(1 To 1)
    For I = 1 To N
        If KaisoNumList(I) <= SiteiKaisoNum Then
            K = K + 1
            ReDim Preserve Output(1 To K)
            Output(K) = KaisoList(I) & "(" & HuzokuKosuList(I) & ")"
        End If
    Next I
    
    �K�w���X�g���w��K�w�܂ł̃��X�g�擾 = Output
    
End Function
