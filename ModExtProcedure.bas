Attribute VB_Name = "ModExtProcedure"
Option Explicit
'�O���Q�ƃv���V�[�W���̎擾�p���W���[��
'frmExtRef�ƘA�g���Ă���

Function Kaiso()
    '�O���Q�ƃv���V�[�W���ꗗ�\���t�H�[���N��
    Kaiso = "�K�w�\��"
    Call frmKaiso.Show
    
End Function

Function �t�H�[���pVBProject�쐬()
    
    Dim I%, J%, II%, K%, M%, N% '�����グ�p(Integer�^)
    Dim OutputVBProjectList() As classVBProject
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    Dim VBProjectList As VBProjects, TmpVBProject As VBProject
    Dim TmpModule As VBComponent, TmpProcedureNameList, TmpCodeDict As Object
    Dim TmpProcedureName$
    Dim Dummy
    
    Set VBProjectList = ActiveWorkbook.VBProject.VBE.VBProjects
    ReDim OutputVBProjectList(1 To VBProjectList.Count)
    For I = 1 To VBProjectList.Count
        Set TmpVBProject = VBProjectList(I)
        Set TmpClassVBProject = New classVBProject
        TmpClassVBProject.MyName = TmpVBProject.Name
        TmpClassVBProject.MyBookName = Dir(TmpVBProject.FileName)
        
        For J = 1 To TmpVBProject.VBComponents.Count
'            If I = 2 And J = 25 Then Stop
            Set TmpClassModule = New classModule
            Set TmpModule = TmpVBProject.VBComponents(J)
            
            TmpClassModule.Name = TmpModule.Name
            TmpClassModule.VBProjectName = TmpClassVBProject.Name
            TmpClassModule.ModuleType = ���W���[����ޔ���(TmpModule.Type)
            TmpClassModule.BookName = TmpClassVBProject.BookName
            
'            TmpProcedureNameList = ���W���[���̃v���V�[�W�����ꗗ�擾(TmpModule)
            Set TmpCodeDict = ���W���[���̃R�[�h�ꗗ�擾(TmpModule)
            
            If TmpCodeDict Is Nothing Then
                TmpProcedureNameList = Empty
            Else
                TmpProcedureNameList = TmpCodeDict.Keys
                TmpProcedureNameList = Application.Transpose(Application.Transpose(TmpProcedureNameList))
            End If
            
            
            If IsEmpty(TmpProcedureNameList) = False Then
                For II = 1 To UBound(TmpProcedureNameList)
                    Set TmpClassProcedure = New ClassProcedure
                    TmpProcedureName = TmpProcedureNameList(II)
                    TmpClassProcedure.Name = TmpProcedureName
                    TmpClassProcedure.Code = TmpCodeDict(TmpProcedureName)
                    Dummy = �R�[�h����v���V�[�W���̃^�C�v�Ǝg�p�͈͎擾(TmpClassProcedure.Code, TmpProcedureName)
                    TmpClassProcedure.RangeOfUse = Dummy(1)
                    TmpClassProcedure.ProcedureType = Dummy(2)
                    Set TmpClassProcedure.KensakuCode = �R�[�h�������p�ɕύX(TmpCodeDict(TmpProcedureName))
                    TmpClassProcedure.VBProjectName = TmpClassVBProject.Name
                    TmpClassProcedure.ModuleName = TmpClassModule.Name
                    TmpClassProcedure.BookName = TmpClassModule.BookName
                    TmpClassModule.AddProcedure TmpClassProcedure
                Next II
            End If
            
            TmpClassVBProject.AddModule TmpClassModule
            
        Next J
        
        Set OutputVBProjectList(I) = TmpClassVBProject
        
    Next I
    
    �t�H�[���pVBProject�쐬 = OutputVBProjectList
    
End Function

Private Function ���W���[����ޔ���(ModuleType%)
'http://officetanaka.net/excel/vba/vbe/04.htm

    Dim Output$
    Select Case ModuleType
    Case 1
        Output = "�W�����W���[��"
    Case 2
        Output = "�N���X���W���[��"
    Case 3
        Output = "���[�U�[�t�H�[��"
    Case 11
        Output = "ActiveX �f�U�C�i"
    Case 100
        Output = "Document ���W���[��"
    Case Else
        MsgBox ("���W���[����ނ�����ł��܂���")
        Stop
    End Select
    
    ���W���[����ޔ��� = Output
    
End Function

Function ���W���[���̃v���V�[�W�����ꗗ�擾(InputModule As VBComponent)
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim TmpStr$
    Dim Output
    ReDim Output(1 To 1)
    With InputModule.CodeModule
        K = 0
        For I = 1 To .CountOfLines
            If TmpStr <> .ProcOfLine(I, 0) Then
                TmpStr = .ProcOfLine(I, 0)
                K = K + 1
                ReDim Preserve Output(1 To K)
                Output(K) = TmpStr
            End If
        Next I
    End With
    
    If K = 0 Then '���W���[�����Ƀv���V�[�W�����Ȃ��ꍇ
        Output = Empty
    End If
    
    ���W���[���̃v���V�[�W�����ꗗ�擾 = Output
        
End Function

Function ���W���[���̃R�[�h�ꗗ�擾(InputModule As VBComponent)
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim ProcedureList
    ProcedureList = ���W���[���̃v���V�[�W�����ꗗ�擾(InputModule)
    Dim Output As Object
    Dim TmpProcedureName$, TmpStart&, TmpEnd&, TmpCode$
    Dim Hantei As Boolean
    Dim Dummy
    Dim TmpProcedureType$
    If IsEmpty(ProcedureList) Then
        '�v���V�[�W������
        Set Output = Nothing
    Else
        '�v���V�[�W���L��
        N = UBound(ProcedureList, 1)
        Set Output = CreateObject("Scripting.Dictionary")
        For I = 1 To N
            TmpProcedureName = ProcedureList(I)
            
            Hantei = �v���V�[�W�����v���p�e�B������(InputModule, TmpProcedureName)
                        
            If Hantei = False Then '�v���V�[�W�����v���p�e�B�łȂ�
                TmpCode = �R�[�h�̎擾�ŋ���(InputModule, TmpProcedureName)
                Output.Add TmpProcedureName, TmpCode
            Else
                Dummy = �R�[�h�̎擾�ŋ��Ńv���p�e�B��p(InputModule, TmpProcedureName)
                For J = 1 To UBound(Dummy, 1)
                    TmpProcedureName = Dummy(J, 1)
                    TmpCode = Dummy(J, 2)
                    Output.Add TmpProcedureName, TmpCode
                Next J
            End If
        Next I
    End If
    
    Set ���W���[���̃R�[�h�ꗗ�擾 = Output

End Function

Function �R�[�h�������p�ɕύX(InputCode) As Object
    
    Dim CodeList, TmpStr$, TmpRowStr$
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    CodeList = Split(InputCode, vbLf)
    CodeList = Application.Transpose(CodeList)
    CodeList = Application.Transpose(CodeList)
    N = UBound(CodeList, 1)
    
    Dim BunkatuStrList, HenkanStr$, TmpBunkatu
    BunkatuStrList = Array(" ", ":", ",", """", "(", ")")
    BunkatuStrList = Application.Transpose(Application.Transpose(BunkatuStrList))
    HenkanStr = Chr(13)
    
    Dim BunkatuDict As Object
    Set BunkatuDict = CreateObject("Scripting.Dictionary")
    Dim Output As Object
    Set Output = CreateObject("Scripting.Dictionary")
    
    For I = 1 To N
        TmpStr = CodeList(I)
        TmpRowStr = TmpStr
        TmpStr = Trim(TmpStr) '���E�̋󔒏���
        TmpStr = StrConv(TmpStr, vbUpperCase) '�������ɕϊ�
'        TmpStr = StrConv(TmpStr, vbNarrow) '���p�ɕϊ�
        If InStr(1, TmpStr, "'") > 0 Then
            TmpStr = Split(TmpStr, "'")(0) '�R�����g�̏���
        End If
        TmpStr = Replace(TmpStr, Chr(13), "") '���s������
        TmpStr = �R�[�h��s�������p�ɕϊ�(TmpStr)
        
        If TmpStr <> "" Then
            '�w�蕶���ŕ�������
            For J = 1 To UBound(BunkatuStrList, 1)
                TmpStr = Replace(TmpStr, BunkatuStrList(J), HenkanStr)
            Next J
            TmpBunkatu = Split(TmpStr, HenkanStr)
            
            For J = 0 To UBound(TmpBunkatu)
                If �v���V�[�W�������p�����񂩂ǂ�������(TmpRowStr, TmpBunkatu(J)) Then
                    If BunkatuDict.Exists(TmpBunkatu(J)) = False Then
                        BunkatuDict.Add TmpBunkatu(J), ""
                    End If
                Else
'                    Stop
                End If
            Next J
        End If
    Next I
    
    Set Output = BunkatuDict
    Set �R�[�h�������p�ɕύX = Output

End Function

Function �v���V�[�W�������p�����񂩂ǂ�������(RowStr$, Str)

    Dim Hantei As Boolean
    Dim HanteiStr$
    HanteiStr = Replace(RowStr, """", "!")
    
    If Str = "+" Or Str = "=" Or Str = "-" Or Str = "/" Or Str = "" Then
        Hantei = False
    ElseIf InStr(1, RowStr, """" & Str & """") > 0 Then '�������镶���񂪁u"�v�ŋ��܂ꂽ�����łȂ�
        Hantei = False
    ElseIf Str = "SUB" Or Str = "FUNCTION" Or Str = "END" Or Str = "EXIT" Or Str = "DIM" Or Str = "BYVAL" Or Str = "AS" Or Str = "RANGE" Or Str = "CALL" Then '�\���
        Hantei = False
    ElseIf Str = "ON" Or Str = "ERROR" Or Str = "NEXT" Or Str = "SET" Or Str = "RESUME" Or Str = "OR" Or Str = "ELSEIF" Then '�\���
        Hantei = False
    ElseIf IsNumeric(Mid(Str, 1, 1)) Then '1�����ڂ������łȂ�
        Hantei = False
    Else
        Hantei = True
    End If
    
    �v���V�[�W�������p�����񂩂ǂ������� = Hantei
    
End Function

Private Sub Test�R�[�h��s�������p�ɕϊ�()
    
    Dim Str$
    Str = "A" & """" & """" & """" & """" & "B"
    Call �R�[�h��s�������p�ɕϊ�(Str)
    
End Sub

Function �R�[�h��s�������p�ɕϊ�(ByVal RowStr$)
'�u"�v�ŋ��܂ꂽ���������������
    RowStr = Replace(RowStr, """" & """", "")
    Dim TmpSplit
    Dim Output$
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    If InStr(1, RowStr, """") > 0 Then
        TmpSplit = Split(RowStr, """")
        
        For I = 0 To UBound(TmpSplit, 1) '��Ԗڂ́u"�v�ŋ��܂ꂽ������ł���
            If I Mod 2 = 0 Then
                Output = Output & TmpSplit(I) & " "
            End If
        Next I
        
    Else
        Output = RowStr
    End If
        
    �R�[�h��s�������p�ɕϊ� = Output

    
    
End Function

Function �S�v���V�[�W���ꗗ�쐬(VBProjectList)
    
    Dim I&, J&, II&, K&, M&, N& '�����グ�p(Long�^)
    Dim ProcedureCount&
    '�v���V�[�W���̌����v�Z
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    
    ProcedureCount = 0
    For I = 1 To UBound(VBProjectList, 1)
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            ProcedureCount = ProcedureCount + TmpClassModule.Procedures.Count
        Next J
    Next
    
    Dim Output
    ReDim Output(1 To ProcedureCount, 1 To 6)
    '1:VBProject��
    '2:Module��
    '3:Procedure��
    '4:VBProject�̔ԍ�
    '5:Module�̔ԍ�
    '6:Procedure�̔ԍ�
    
    K = 0
    For I = 1 To UBound(VBProjectList, 1)
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            For II = 1 To TmpClassModule.Procedures.Count
                K = K + 1
                Set TmpClassProcedure = TmpClassModule.Procedures(II)
                Output(K, 1) = TmpClassVBProject.Name
                Output(K, 2) = TmpClassModule.Name
                Output(K, 3) = TmpClassProcedure.Name
                Output(K, 4) = I
                Output(K, 5) = J
                Output(K, 6) = II
            Next II
        Next J
    Next
    
    �S�v���V�[�W���ꗗ�쐬 = Output
    
End Function

Sub �v���V�[�W�����̎g�p�v���V�[�W���擾(VBProjectList() As classVBProject, AllProcedureList)
    
    Dim I&, J&, II&, JJ&, III&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(AllProcedureList, 1)
    '�v���V�[�W���̌����v�Z
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    Dim TmpVBProjectNum%, TmpModuleNum%, TmpProcedureNum%
    Dim TmpKensakuCode As Object
    Dim TmpVBProjectName$, TmpModuleName$, TmpProcedureName$
    Dim TmpSiyosakiList As Object
    Dim TmpSiyoProcedure As ClassProcedure
    Dim TmpSiyoProcedureList() As ClassProcedure
    Dim NaibuSansyoNaraTrue As Boolean
    Dim TmpHantei As Boolean
    
    For I = 1 To UBound(VBProjectList, 1) '�eVBProject�ɂ����Ă�
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count '�e���W���[���ɂ����Ă�
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            For II = 1 To TmpClassModule.Procedures.Count '�e�v���V�[�W���ɂ����Ă�
                Set TmpClassProcedure = TmpClassModule.Procedures(II)
                Set TmpKensakuCode = TmpClassProcedure.KensakuCode
                K = 0
                ReDim TmpSiyoProcedureList(1 To 1)
                For JJ = 1 To N
                    TmpVBProjectName = AllProcedureList(JJ, 1)
                    TmpModuleName = AllProcedureList(JJ, 2)
                    TmpProcedureName = AllProcedureList(JJ, 3)
                    
                    If TmpProcedureName <> TmpClassProcedure.Name Then '�������g�̃v���V�[�W���͌�������Ȃ�
                        TmpVBProjectName = StrConv(TmpVBProjectName, vbUpperCase) '�����p�ɑ啶���ɕϊ�
                        TmpModuleName = StrConv(TmpModuleName, vbUpperCase) '�����p�ɑ啶���ɕϊ�
                        TmpProcedureName = StrConv(TmpProcedureName, vbUpperCase) '�����p�ɑ啶���ɕϊ�
                        
                        If TmpKensakuCode.Exists(TmpVBProjectName & "." & TmpModuleName & "." & TmpProcedureName) Or _
                           TmpKensakuCode.Exists(TmpModuleName & "." & TmpProcedureName) Or _
                           TmpKensakuCode.Exists(TmpProcedureName) Then

                            TmpVBProjectNum = AllProcedureList(JJ, 4)
                            TmpModuleNum = AllProcedureList(JJ, 5)
                            TmpProcedureNum = AllProcedureList(JJ, 6)
                            Set TmpSiyoProcedure = VBProjectList(TmpVBProjectNum).Modules(TmpModuleNum).Procedures(TmpProcedureNum)
                            
                            TmpHantei = True
                            If TmpSiyoProcedure.RangeOfUse = "Private" Then
                                If TmpSiyoProcedure.ModuleName = TmpClassProcedure.ModuleName And _
                                   TmpSiyoProcedure.VBProjectName = TmpClassProcedure.VBProjectName Then
                                    '�g�p�v���V�[�W����Private�œ������W���[���AVBProject���ɂ���
                                    TmpHantei = True
                                Else
                                    TmpHantei = False
                                End If
                            Else
                                TmpHantei = True
                            End If
                            
                            If TmpHantei = True Then
'                                TmpClassProcedure.AddUseProcedure TmpSiyoProcedure
                                K = K + 1
                                ReDim Preserve TmpSiyoProcedureList(1 To K)
                                Set TmpSiyoProcedureList(K) = TmpSiyoProcedure
                            End If
                            
'                            Debug.Assert TmpSiyoProcedure.Name <> "OutputText"
                            
                        End If
                    End If
                Next JJ
                
                '�O���Q�Ƃ��Ă��邪�A�����ł��������O�ŎQ�Ƃ��Ă���Ƃ��͏��O����
                If K = 0 Then
                    '�g�p�v���V�[�W���Ȃ��E�E�E�������Ȃ�
                ElseIf K = 1 Then
                    '�g�p�v���V�[�W��1�E�E�E���̂܂܎g�p��Ŋi�[
                    TmpClassProcedure.AddUseProcedure TmpSiyoProcedureList(1)
                Else '�g�p�v���V�[�W����2�ȏ�
                    For JJ = 1 To K
                        Set TmpSiyoProcedure = TmpSiyoProcedureList(JJ)
                        TmpVBProjectName = TmpSiyoProcedure.VBProjectName
                        TmpProcedureName = TmpSiyoProcedure.Name
                        
                        If TmpVBProjectName = TmpClassVBProject.Name Then
                            '�����Q��(�g�p�v���V�[�W����VBProject�������g��VBProject���ƈ�v���Ă���)
                            TmpClassProcedure.AddUseProcedure TmpSiyoProcedure
                        Else
                            '�O���Q��(�g�p�v���V�[�W����VBProject�������g��VBProject���ƈ�v���Ă��Ȃ�)
                            
                            '�����Q�Ƃ����łɂ��Ă��邩����
                            NaibuSansyoNaraTrue = False
                            For III = 1 To K
                                If JJ <> III Then
                                    If TmpProcedureName = TmpSiyoProcedureList(III).Name And _
                                       TmpClassVBProject.Name = TmpSiyoProcedureList(III).VBProjectName Then
                                        '�����Q�ƍς�
                                        NaibuSansyoNaraTrue = True
                                        Exit For
                                    End If
                                End If
                            Next III
                            
                            If NaibuSansyoNaraTrue = False Then
                                TmpClassProcedure.AddUseProcedure TmpSiyoProcedure
                            End If
                        End If
                    Next JJ
                End If
            Next II
        Next J
    Next

End Sub

Function �O���Q�ƃv���V�[�W���A�z�z��쐬(VBProjectList() As classVBProject)
    
    Dim I&, J&, II&, K&, M&, N& '�����グ�p(Long�^)
    '�v���V�[�W���̌����v�Z
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    
    Dim TmpVBProjectName$, TmpModuleName$, TmpProcedureName$
    Dim TmpCode$
    
    Dim TmpVBProject
    
    Dim TmpExtProcedureDict As Object
    N = UBound(VBProjectList, 1)
    ReDim Output(1 To N)
    For I = 1 To N
        Set TmpExtProcedureDict = CreateObject("Scripting.Dictionary")
        TmpVBProjectName = VBProjectList(I).Name
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            For II = 1 To TmpClassModule.Procedures.Count
                Set TmpClassProcedure = TmpClassModule.Procedures(II)
                Call �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾�A�z�z��p(TmpVBProjectName, TmpClassProcedure, TmpExtProcedureDict)
            Next II
        Next J
        
        Set Output(I) = TmpExtProcedureDict
        
    Next I
        
    �O���Q�ƃv���V�[�W���A�z�z��쐬 = Output
    
End Function



Sub �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾�A�z�z��p(VBProjectName$, ClassProcedure As ClassProcedure, ExtProcedureDict As Object)
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim TmpUseProcedure As ClassProcedure
    Dim TmpUseProcedure2 As ClassProcedure
    Dim GaibuSansyoNaraTrue As Boolean
    
    If ClassProcedure.UseProcedure.Count = 0 Then
        '�g�p���Ă���v���V�[�W�������̏ꍇ�������Ȃ�
    Else
        For I = 1 To ClassProcedure.UseProcedure.Count
            Set TmpUseProcedure = ClassProcedure.UseProcedure(I)
            
            '�ċA(�g�p�v���V�[�W�����̊O���Q�Ƃ�T��)
            Call �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾�A�z�z��p(VBProjectName, TmpUseProcedure, ExtProcedureDict)
            
            If TmpUseProcedure.VBProjectName <> VBProjectName Then 'VBProject�����قȂ�ΊO���Q��
                
                '���Ɏ�����VBProject���ɓ������O�̃v���V�[�W�������݂���΁A�O���Q�ƂłȂ�
                GaibuSansyoNaraTrue = True
                For J = 1 To ClassProcedure.UseProcedure.Count
                    Set TmpUseProcedure2 = ClassProcedure.UseProcedure(J)
                    If TmpUseProcedure2.VBProjectName = VBProjectName And TmpUseProcedure2.Name = TmpUseProcedure.Name Then
                        GaibuSansyoNaraTrue = False
                        Exit For
                    End If
                Next J
                
                If GaibuSansyoNaraTrue = True And ExtProcedureDict.Exists(TmpUseProcedure.Name) = False Then
                    ExtProcedureDict.Add TmpUseProcedure.Name, TmpUseProcedure.Code
                End If
            End If
        Next I
    End If

End Sub

Function �O���Q�ƃv���V�[�W�����X�g�쐬(VBProjectList() As classVBProject)
    
    Dim I&, J&, II&, K&, M&, N& '�����グ�p(Long�^)
    '�v���V�[�W���̌����v�Z
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    
    Dim TmpVBProjectName$, TmpModuleName$, TmpProcedureName$
    Dim TmpCode$
    
    Dim TmpVBProject
    
    Dim TmpExtProcedureList() As ClassProcedure
    N = UBound(VBProjectList, 1)
    ReDim Output(1 To N)
    For I = 1 To N
        ReDim TmpExtProcedureList(1 To 1)
        TmpVBProjectName = VBProjectList(I).Name
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            For II = 1 To TmpClassModule.Procedures.Count
                Set TmpClassProcedure = TmpClassModule.Procedures(II)
                Call �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾(TmpVBProjectName, TmpClassProcedure, TmpExtProcedureList, 0)
            Next II
        Next J
        
        Output(I) = TmpExtProcedureList
    Next I
        
    �O���Q�ƃv���V�[�W�����X�g�쐬 = Output
    
End Function

Sub �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾(VBProjectName$, ClassProcedure As ClassProcedure, ExtProcedureList() As ClassProcedure, ByVal Depth&)
    
    '�ċA�֐��̐[���i���[�v�j�����ȏ㒴���Ȃ��悤�ɂ���B
    Depth = Depth + 1
    If Depth > 10 Then
        Debug.Print "�O���Q�ƃv���V�[�W���T���ŁA�K�萔�̊K�w�𒴂��܂����B"
        Debug.Print ClassProcedure.Name
        Exit Sub
    End If
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim TmpUseProcedure As ClassProcedure
    Dim TmpUseProcedure2 As ClassProcedure
    Dim GaibuSansyoNaraTrue As Boolean
    Dim TmpHantei As Boolean
    
    If ClassProcedure.UseProcedure.Count = 0 Then
        '�g�p���Ă���v���V�[�W�������̏ꍇ�������Ȃ�
    Else
        For I = 1 To ClassProcedure.UseProcedure.Count
            Set TmpUseProcedure = ClassProcedure.UseProcedure(I)
            
            '�ċA(�g�p�v���V�[�W�����̊O���Q�Ƃ�T��)
            Call �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾(VBProjectName, TmpUseProcedure, ExtProcedureList, Depth)
            
            If TmpUseProcedure.VBProjectName <> VBProjectName Then 'VBProject�����قȂ�ΊO���Q��
                
                '���Ɏ�����VBProject���ɓ������O�̃v���V�[�W�������݂���΁A�O���Q�ƂłȂ�
                GaibuSansyoNaraTrue = True
                For J = 1 To ClassProcedure.UseProcedure.Count
                    Set TmpUseProcedure2 = ClassProcedure.UseProcedure(J)
                    If TmpUseProcedure2.VBProjectName = VBProjectName And TmpUseProcedure2.Name = TmpUseProcedure.Name Then
                        GaibuSansyoNaraTrue = False
                        Exit For
                    End If
                Next J
                
                TmpHantei = False
                
                If Not ExtProcedureList(1) Is Nothing Then
                    For J = 1 To UBound(ExtProcedureList, 1)
                        If ExtProcedureList(J).Name = TmpUseProcedure.Name Then
                            '���Ɏ擾�ς�
                            TmpHantei = True
                            Exit For
                        End If
                    Next J
                End If
                
                If GaibuSansyoNaraTrue = True And TmpHantei = False Then
                
                    If Not ExtProcedureList(1) Is Nothing Then
                        ReDim Preserve ExtProcedureList(1 To UBound(ExtProcedureList, 1) + 1)
                    End If
                    
                    Set ExtProcedureList(UBound(ExtProcedureList, 1)) = TmpUseProcedure
                End If
            End If
        Next I
    End If

End Sub

Private Function �R�[�h����v���V�[�W���̃^�C�v�Ǝg�p�͈͎擾(InputCode, ProcedureName$)
    
    Dim ProcedureName2$
    '�v���V�[�W�����v���p�e�B�̏ꍇ�̑Ή�
    If InStr(1, ProcedureName, ")") > 0 Then
        ProcedureName2 = Split(ProcedureName, ")")(1)
    Else
        ProcedureName2 = ProcedureName
    End If
    
    Dim HeadStr$
    HeadStr = Split(InputCode, ProcedureName2)(0)
    
    Dim ProcedureType$, RangeOfUse$
    If InStr(1, HeadStr, "Sub") > 0 Then
        ProcedureType = "Sub"
    ElseIf InStr(1, HeadStr, "Function") > 0 Then
        ProcedureType = "Function"
    ElseIf InStr(1, HeadStr, "Property Get") > 0 Then
        ProcedureType = "Property Get"
    ElseIf InStr(1, HeadStr, "Property Let") > 0 Then
        ProcedureType = "Property Let"
    ElseIf InStr(1, HeadStr, "Property Set") > 0 Then
        ProcedureType = "Property Set"
    Else
        MsgBox ("�v���V�[�W���̃^�C�v������ł��܂���")
        Stop
    End If
    
    If InStr(1, HeadStr, "Public") > 0 Then
        RangeOfUse = "Public"
    ElseIf InStr(1, HeadStr, "Private") > 0 Then
        RangeOfUse = "Private"
    Else
        RangeOfUse = "Public"
    End If
        
    Dim Output(1 To 2)
    Output(1) = RangeOfUse
    Output(2) = ProcedureType
    
    �R�[�h����v���V�[�W���̃^�C�v�Ǝg�p�͈͎擾 = Output

End Function

Private Function �R�[�h�̎擾�C��(InputModule As VBComponent, CodeStart&, CodeCount&)

    '�ʏ�擾
    Dim TmpCode
    TmpCode = InputModule.CodeModule.Lines(CodeStart, CodeCount)
    Dim LastStr$, TmpSplit, TmpSplit2
    TmpSplit = Split(TmpCode, vbLf)
    LastStr = TmpSplit(UBound(TmpSplit))

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output$

    '�R�[�h�̃X�^�[�g�ʒu����ŏI�s��T������悤�ɂ���B
    For I = 2 To UBound(TmpSplit) + 100
        TmpCode = InputModule.CodeModule.Lines(CodeStart, I)
        TmpSplit2 = Split(TmpCode, vbLf)
        LastStr = TmpSplit2(UBound(TmpSplit2))
        
        LastStr = Trim(LastStr) '�擪�̃X�y�[�X������
        If InStr(1, LastStr, "'") > 0 Then
            LastStr = Split(LastStr, "'")(0) '�R�����g������
        End If
        
        If InStr(1, LastStr, "End Function") > 0 _
            Or InStr(1, LastStr, "End Sub") > 0 _
            Or InStr(1, LastStr, "End Property") > 0 Then
            Output = TmpCode
'            Debug.Print LastStr
            Exit For
        End If
    Next I

    If Output = "" Then
        '����ł��ŏI�s��������Ȃ������ꍇ
        Output = InputModule.CodeModule.Lines(CodeStart, CodeCount)
        Debug.Print Output '�m�F�p
        Stop
    End If

    �R�[�h�̎擾�C�� = Output

End Function


Private Function �R�[�h�̎擾�ŋ���(InputModule As VBComponent, ProcedureName$)
    

    
    Dim Output$
    Dim TmpStart&, TmpCount&, TmpProcKind%
    
    '�v���V�[�W���̊J�n�ʒu�擾
    '�Q�l�Fhttps://docs.microsoft.com/ja-jp/office/vba/api/access.module.procbodyline
    TmpStart = -1
    '�v���V�[�W����Sub/Function�v���V�[�W�����AProperty Get/Let/Set�v���V�[�W�����܂��s���Ȃ̂ŁA�肠���莟��T��B
    On Error Resume Next
    With InputModule.CodeModule
        TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Proc) 'Sub/Function�v���V�[�W��
        TmpProcKind = vbext_pk_Proc
        If TmpStart = -1 Then
            TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Get) 'Property Get�v���V�[�W��
            TmpProcKind = vbext_pk_Get
            If TmpStart = -1 Then
                TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Let) 'Property Let�v���V�[�W��
                TmpProcKind = vbext_pk_Let
                If TmpStart = -1 Then
                    TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Set) 'Property Set�v���V�[�W��
                    TmpProcKind = vbext_pk_Set
                End If
            End If
        End If
        TmpCount = .ProcCountLines(ProcedureName, TmpProcKind)
        
        Output = �R�[�h�̎擾�C��(InputModule, TmpStart, TmpCount)
'        Output = .Lines(TmpStart, TmpCount)
    End With
    On Error GoTo 0
    
    �R�[�h�̎擾�ŋ��� = Output

End Function

Function �v���V�[�W�����v���p�e�B������(InputModule As VBComponent, ProcedureName$) As Boolean

    Dim Output As Boolean
    Dim TmpStart&, TmpCount&, TmpProcKind%
    
    '�v���V�[�W���̊J�n�ʒu�擾
    '�Q�l�Fhttps://docs.microsoft.com/ja-jp/office/vba/api/access.module.procbodyline
    TmpStart = -1
    '�v���V�[�W����Sub/Function�v���V�[�W�����AProperty Get/Let/Set�v���V�[�W�����܂��s���Ȃ̂ŁA�肠���莟��T��B
    On Error Resume Next
    With InputModule.CodeModule
        TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Proc) 'Sub/Function�v���V�[�W��
        TmpProcKind = vbext_pk_Proc
        '�����ŃG���[�łȂ�������Sub�܂���Function�v���V�[�W��
        If TmpStart = -1 Then
            'TmpStart�̒l���擾�ł��Ă��Ȃ��̂ŁAProperty
            Output = True
        Else
            Output = False
        End If
    End With
    On Error GoTo 0
    
    �v���V�[�W�����v���p�e�B������ = Output

End Function

Private Function �R�[�h�̎擾�ŋ��Ńv���p�e�B��p(InputModule As VBComponent, ProcedureName$)
    
    
    Dim Output
    Dim TmpStart&, TmpCount&, TmpProcKind%
    Dim HanteiGet As Boolean, HanteiLet As Boolean, HanteiSet As Boolean
    
    '�v���V�[�W���̊J�n�ʒu�擾
    '�Q�l�Fhttps://docs.microsoft.com/ja-jp/office/vba/api/access.module.procbodyline
    
    '�܂��v���V�[�W����Property Get/Let/Set�ǂ�ɂȂ邩����
    On Error Resume Next
    With InputModule.CodeModule
        TmpStart = -1
        TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Get) 'Property Get�v���V�[�W��
        If TmpStart <> -1 Then
            HanteiGet = True
        Else
            HanteiGet = False
        End If
        
        TmpStart = -1
        TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Let) 'Property Let�v���V�[�W��
        If TmpStart <> -1 Then
            HanteiLet = True
        Else
            HanteiLet = False
        End If
        
        TmpStart = -1
        TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Set) 'Property Set�v���V�[�W��
        If TmpStart <> -1 Then
            HanteiSet = True
        Else
            HanteiSet = False
        End If
        
    End With
    On Error GoTo 0
    
    'Property Get/Let/Set�ʁX�ŃR�[�h���擾
    Dim CodeCount& '�o�͂���R�[�h�̌�
    CodeCount = Abs(HanteiGet + HanteiLet + HanteiSet)
    ReDim Output(1 To CodeCount, 1 To 2) '1:�v���V�[�W����,2:�R�[�h
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim TmpProcedureName$
    Dim TmpCode$
    K = 0
    
    If HanteiGet Then
        K = K + 1
        TmpStart = InputModule.CodeModule.ProcBodyLine(ProcedureName, vbext_pk_Get)
        TmpCount = InputModule.CodeModule.ProcCountLines(ProcedureName, vbext_pk_Get)
        Output(K, 1) = "(Get)" & ProcedureName
        Output(K, 2) = �R�[�h�̎擾�C��(InputModule, TmpStart, TmpCount)
    End If
    If HanteiLet Then
        K = K + 1
        TmpStart = InputModule.CodeModule.ProcBodyLine(ProcedureName, vbext_pk_Let)
        TmpCount = InputModule.CodeModule.ProcCountLines(ProcedureName, vbext_pk_Let)
        Output(K, 1) = "(Let)" & ProcedureName
        Output(K, 2) = �R�[�h�̎擾�C��(InputModule, TmpStart, TmpCount)
    End If
    If HanteiSet Then
        K = K + 1
        TmpStart = InputModule.CodeModule.ProcBodyLine(ProcedureName, vbext_pk_Set)
        TmpCount = InputModule.CodeModule.ProcCountLines(ProcedureName, vbext_pk_Set)
        Output(K, 1) = "(Set)" & ProcedureName
        Output(K, 2) = �R�[�h�̎擾�C��(InputModule, TmpStart, TmpCount)
    End If
    
    
    �R�[�h�̎擾�ŋ��Ńv���p�e�B��p = Output

End Function

