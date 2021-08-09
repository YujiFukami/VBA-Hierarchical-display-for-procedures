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
        
        For J = 1 To TmpVBProject.VBComponents.Count
            Set TmpClassModule = New classModule
            Set TmpModule = TmpVBProject.VBComponents(J)
            
            TmpClassModule.Name = TmpModule.Name
            TmpClassModule.VBProjectName = TmpClassVBProject.Name
            TmpClassModule.ModuleType = ���W���[����ޔ���(TmpModule.Type)
            
            TmpProcedureNameList = ���W���[���̃v���V�[�W�����ꗗ�擾(TmpModule)
            Set TmpCodeDict = ���W���[���̃R�[�h�ꗗ�擾(TmpModule)
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
    If IsEmpty(ProcedureList) Then
        '�v���V�[�W������
        Set Output = Nothing
    Else
        '�v���V�[�W���L��
        N = UBound(ProcedureList, 1)
        Set Output = CreateObject("Scripting.Dictionary")
        For I = 1 To N
            TmpProcedureName = ProcedureList(I)
            With InputModule.CodeModule
                On Error Resume Next
                TmpStart = 0
                TmpEnd = 0
                TmpStart = .ProcBodyLine(TmpProcedureName, 0)
                TmpEnd = .ProcCountLines(TmpProcedureName, 0)
                      
                If TmpStart = 0 Then '�N���X���W���[���̃R�[�h�擾�p
                    TmpStart = .ProcBodyLine(TmpProcedureName, vbext_pk_Get)
                    TmpEnd = .ProcCountLines(TmpProcedureName, vbext_pk_Let)
                    If TmpEnd = 0 Then
                        TmpEnd = .ProcCountLines(TmpProcedureName, vbext_pk_Get)
                    End If
                End If
                
                On Error GoTo 0
                
                TmpCode = .Lines(TmpStart, TmpEnd)
            End With
            
            Output.Add TmpProcedureName, TmpCode
        Next I
    End If
    
    Set ���W���[���̃R�[�h�ꗗ�擾 = Output

End Function

Function �R�[�h�������p�ɕύX(InputCode) As Object
    
    Dim CodeList, TmpStr$
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    CodeList = Split(InputCode, vbLf)
    CodeList = Application.Transpose(CodeList)
    CodeList = Application.Transpose(CodeList)
    N = UBound(CodeList, 1)
    
    Dim BunkatuStrList, HenkanStr$, TmpBunkatu
    BunkatuStrList = Array(" ", ":", "_", ",", """", "(", ")")
    BunkatuStrList = Application.Transpose(Application.Transpose(BunkatuStrList))
    HenkanStr = Chr(13)
    
    Dim BunkatuDict As Object
    Set BunkatuDict = CreateObject("Scripting.Dictionary")
    Dim Output As Object
    Set Output = CreateObject("Scripting.Dictionary")
    
    For I = 1 To N
        TmpStr = CodeList(I)
        TmpStr = Trim(TmpStr) '���E�̋󔒏���
        TmpStr = StrConv(TmpStr, vbUpperCase) '�������ɕϊ�
'        TmpStr = StrConv(TmpStr, vbNarrow) '���p�ɕϊ�
        If InStr(1, TmpStr, "'") > 0 Then
            TmpStr = Split(TmpStr, "'")(0) '�R�����g�̏���
        End If
        TmpStr = Replace(TmpStr, Chr(13), "") '���s������
        
        
        If TmpStr <> "" Then
            '�w�蕶���ŕ�������
            For J = 1 To UBound(BunkatuStrList, 1)
                TmpStr = Replace(TmpStr, BunkatuStrList(J), HenkanStr)
            Next J
            TmpBunkatu = Split(TmpStr, HenkanStr)
            
            For J = 0 To UBound(TmpBunkatu)
                If BunkatuDict.Exists(TmpBunkatu(J)) = False Then
                    BunkatuDict.Add TmpBunkatu(J), ""
                End If
            Next J
        End If
    Next I
    
    Set Output = BunkatuDict
    Set �R�[�h�������p�ɕύX = Output

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
                Call �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾(TmpVBProjectName, TmpClassProcedure, TmpExtProcedureList)
            Next II
        Next J
        
        Output(I) = TmpExtProcedureList
    Next I
        
    �O���Q�ƃv���V�[�W�����X�g�쐬 = Output
    
End Function

Sub �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾(VBProjectName$, ClassProcedure As ClassProcedure, ExtProcedureList() As ClassProcedure)
    
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
            Call �v���V�[�W�����̊O���Q�ƃv���V�[�W���擾(VBProjectName, TmpUseProcedure, ExtProcedureList)
            
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
    
    Dim HeadStr$
    HeadStr = Split(InputCode, ProcedureName)(0)
    
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
