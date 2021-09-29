Attribute VB_Name = "ModExportProcedure"
Option Explicit

'�v���V�[�W���P�̂����W���[���Ƃ��ďo�͂���

'�g�p���W���[��
'ModExtProcedure

Private PriVBProjectList() As classVBProject
Private PriAllProcedureList


Sub Test_�w�薼�̃v���V�[�W�������W���[���ŏo��()

    Dim InputProcedureName$
    InputProcedureName = "FukamiAddins3.ModArray.TestSortArray2D"
    InputProcedureName = "FukamiAddins3.ModExportProcedure.ExportProcedure"
    
    Call ExportProcedure(InputProcedureName, ActiveWorkbook.Path)
    
End Sub

Function GetProcedureAllCode(InputProcedureName$)
'�w��v���V�[�W�����g�p���Ă���v���V�[�W�����S���擾���āA�R�[�h���擾����
'20210916

'����
'InputProcedureName �E�E�E�v���V�[�W���̖��O�iVBProject.Module.Procedure�̃t�����œ��́j
'��FFukamiAddins3.ModExtProcedure.Kaiso
   
    Dim TmpProcedure As classProcedure
    Set TmpProcedure = �w�薼�̃v���V�[�W�����擾(InputProcedureName)
    
    Dim ProcedureNameDict
    Set ProcedureNameDict = �v���V�[�W���̎g�p�S�v���V�[�W�����擾(TmpProcedure)
    
    '���W���[���Ƃ��ďo��
    Dim ModuleFileName$
    ModuleFileName = "Mod" & TmpProcedure.Name
    If ���{����܂ނ�����(ModuleFileName) Then
        ModuleFileName = Mid(ModuleFileName, 1, 15) & "_"     '���{����܂ޏꍇ���W���[�����̒����ɂ͌��E������
    End If
    
    Dim ProcedureNameList, CodeList
    If ProcedureNameDict.Count > 0 Then
        ProcedureNameList = Application.Transpose(Application.Transpose(ProcedureNameDict.Keys))
        CodeList = Application.Transpose(Application.Transpose(ProcedureNameDict.Items))
    End If
    
    '���W���[���錾���̎擾
    Dim SengenList
    SengenList = ���W���[���̐錾�����擾(TmpProcedure, ProcedureNameList)
    
    '�擾�����v���V�[�W�����A�R�[�h�A�錾�����Ȃ���
    Dim FixProcedureName$, FixCode$, FixSengen$
    Dim TmpProcedureName$, TmpCode$, TmpSengen$
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    If IsEmpty(ProcedureNameList) Then
        N = 0
    Else
        N = UBound(ProcedureNameList, 1)
    End If
    
    Dim StrProcedureNameList '�R�[�h�ŕ\���p�̃v���V�[�W����
    ReDim StrProcedureNameList(1 To N + 1, 1 To 2)
    
    For I = 0 To N
    
        If I = 0 Then
            TmpProcedureName = TmpProcedure.VBProjectName & "." & TmpProcedure.ModuleName & "." & TmpProcedure.Name
            TmpCode = TmpProcedure.Code
        Else
            TmpProcedureName = ProcedureNameList(I)
            TmpCode = CodeList(I)
            TmpCode = �R�[�h���v���C�x�[�g�ɕϊ�(TmpCode)
        End If
        
        StrProcedureNameList(I + 1, 1) = FixProcedureName & "'" & Split(TmpProcedureName, ".")(2)
        StrProcedureNameList(I + 1, 2) = "���ꏊ�F" & Split(TmpProcedureName, ".")(0) & "." & Split(TmpProcedureName, ".")(1)
        
        FixCode = FixCode & TmpCode & vbLf & vbLf
        
    Next I
    
    FixProcedureName = MakeAligmentedArray(StrProcedureNameList, "�E�E�E")
    
    For I = 1 To UBound(SengenList)
        TmpSengen = SengenList(I)
        FixSengen = FixSengen & "'------------------------------" & vbLf
        FixSengen = FixSengen & TmpSengen & vbLf
    Next I
    FixSengen = FixSengen & "'------------------------------" & vbLf
    
    '�e�L�X�g�ŏo��
    Dim OutputStr$
    OutputStr = "Attribute VB_Name = " & """" & ModuleFileName & """" & vbLf
    OutputStr = OutputStr & "Option Explicit" & vbLf
    OutputStr = OutputStr & vbLf
    OutputStr = OutputStr & FixProcedureName & vbLf
    OutputStr = OutputStr & vbLf
    OutputStr = OutputStr & FixSengen & vbLf
    OutputStr = OutputStr & vbLf
    OutputStr = OutputStr & FixCode & vbLf
    
    Dim OutputList
    OutputList = Split(OutputStr, vbLf)
    OutputList = Application.Transpose(OutputList)
    
    '�o��
    GetProcedureAllCode = OutputStr
    
End Function


Sub ExportProcedure(InputProcedureName$, Optional FolderPath$)
'�w��v���V�[�W�����g�p���Ă���v���V�[�W�����S���擾���āA1�̃��W���[���Ƃ��ăG�N�X�|�[�g����
'20210915

'����
'InputProcedureName                      �E�E�E�v���V�[�W���̖��O�iVBProject.Module.Procedure�̃t�����œ��́j
'��FFukamiAddins3.ModExtProcedure.Kaiso
'[FolderPath]                            �E�E�E�o�͐�̃t�H���_�B�ȗ��Ȃ炱�̃u�b�N�̃t�H���_�p�X

    If FolderPath = "" Then
        FolderPath = ThisWorkbook.Path
    End If
    
    Dim TmpProcedure As classProcedure
    Set TmpProcedure = �w�薼�̃v���V�[�W�����擾(InputProcedureName)
    
    Dim ProcedureNameDict
    Set ProcedureNameDict = �v���V�[�W���̎g�p�S�v���V�[�W�����擾(TmpProcedure)
    
    '���W���[���Ƃ��ďo��
    Dim ModuleFileName$
    ModuleFileName = "Mod" & TmpProcedure.Name
    If ���{����܂ނ�����(ModuleFileName) Then
        ModuleFileName = Mid(ModuleFileName, 1, 15) & "_"     '���{����܂ޏꍇ���W���[�����̒����ɂ͌��E������
    End If
    
    Dim ProcedureNameList, CodeList
    If ProcedureNameDict.Count > 0 Then
        ProcedureNameList = Application.Transpose(Application.Transpose(ProcedureNameDict.Keys))
        CodeList = Application.Transpose(Application.Transpose(ProcedureNameDict.Items))
    End If
    
    '���W���[���錾���̎擾
    Dim SengenList
    SengenList = ���W���[���̐錾�����擾(TmpProcedure, ProcedureNameList)
    
    '�擾�����v���V�[�W�����A�R�[�h�A�錾�����Ȃ���
    Dim FixProcedureName$, FixCode$, FixSengen$
    Dim TmpProcedureName$, TmpCode$, TmpSengen$
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    If IsEmpty(ProcedureNameList) Then
        N = 0
    Else
        N = UBound(ProcedureNameList, 1)
    End If
    
    Dim StrProcedureNameList '�R�[�h�ŕ\���p�̃v���V�[�W����
    ReDim StrProcedureNameList(1 To N + 1, 1 To 2)
    
    For I = 0 To N
    
        If I = 0 Then
            TmpProcedureName = TmpProcedure.VBProjectName & "." & TmpProcedure.ModuleName & "." & TmpProcedure.Name
            TmpCode = TmpProcedure.Code
            TmpCode = �R�[�h���p�u���b�N�ɕϊ�(TmpCode)
        Else
            TmpProcedureName = ProcedureNameList(I)
            TmpCode = CodeList(I)
            TmpCode = �R�[�h���v���C�x�[�g�ɕϊ�(TmpCode)
        End If
        
        StrProcedureNameList(I + 1, 1) = FixProcedureName & "'" & Split(TmpProcedureName, ".")(2)
        StrProcedureNameList(I + 1, 2) = "���ꏊ�F" & Split(TmpProcedureName, ".")(0) & "." & Split(TmpProcedureName, ".")(1)
        
        FixCode = FixCode & TmpCode & vbLf & vbLf
        
    Next I
    
    FixProcedureName = MakeAligmentedArray(StrProcedureNameList, "�E�E�E")
    
    For I = 1 To UBound(SengenList)
        TmpSengen = SengenList(I)
        FixSengen = FixSengen & "'------------------------------" & vbLf
        FixSengen = FixSengen & TmpSengen & vbLf
    Next I
    FixSengen = FixSengen & "'------------------------------" & vbLf
    
    '�e�L�X�g�ŏo��
    Dim OutputStr$
    OutputStr = "Attribute VB_Name = " & """" & ModuleFileName & """" & vbLf
    OutputStr = OutputStr & "Option Explicit" & vbLf
    OutputStr = OutputStr & vbLf
    OutputStr = OutputStr & FixProcedureName & vbLf
    OutputStr = OutputStr & vbLf
    OutputStr = OutputStr & FixSengen & vbLf
    OutputStr = OutputStr & vbLf
    OutputStr = OutputStr & FixCode & vbLf
    
    Dim OutputList
    OutputList = Split(OutputStr, vbLf)
    OutputList = Application.Transpose(OutputList)
    
    Dim TmpRowStr$
    '��s���ɉ��s���܂�ł������������
    For I = 1 To UBound(OutputList, 1)
        TmpRowStr = OutputList(I, 1)
        
        TmpRowStr = Replace(TmpRowStr, vbLf, "")
        TmpRowStr = Replace(TmpRowStr, vbCrLf, "")
        TmpRowStr = Replace(TmpRowStr, Chr(13), "")
        TmpRowStr = Replace(TmpRowStr, Chr(10), "")
        OutputList(I, 1) = TmpRowStr
        
    Next I
    
    Call OutputText(FolderPath, ModuleFileName & ".bas", OutputList, "")
            
    Debug.Print "�u" & ModuleFileName & ".bas" & "�v���o��", " �o�͐恨" & FolderPath
    
End Sub

Private Sub ������()

    If IsEmpty(PriAllProcedureList) Then
        PriVBProjectList = �t�H�[���pVBProject�쐬
        PriAllProcedureList = �S�v���V�[�W���ꗗ�쐬(PriVBProjectList)
        Call �v���V�[�W�����̎g�p�v���V�[�W���擾(PriVBProjectList, PriAllProcedureList)
    End If
    
End Sub

Private Function �w�薼�̃v���V�[�W�����擾(InputProcedureName$) As classProcedure

'����
'InputProcedureName�E�E�E�v���V�[�W���̖��O�iVBProject.Module.Procedure�̃t�����œ��́j
'��FFukamiAddins3.ModExtProcedure.Kaiso

    Dim VBProjectName$, ModuleName$, ProcedureName$
    VBProjectName = Split(InputProcedureName, ".")(0)
    ModuleName = Split(InputProcedureName, ".")(1)
    ProcedureName = Split(InputProcedureName, ".")(2)
    
    Call ������
    
    Dim AllProcedureList, VBProjectList() As classVBProject
    AllProcedureList = PriAllProcedureList
    VBProjectList = PriVBProjectList
    
    Dim Output As classProcedure
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim Num1&, Num2&, Num3&
    For I = 1 To UBound(AllProcedureList)
        If AllProcedureList(I, 1) = VBProjectName And _
           AllProcedureList(I, 2) = ModuleName And _
           AllProcedureList(I, 3) = ProcedureName Then
            
            Num1 = AllProcedureList(I, 4)
            Num2 = AllProcedureList(I, 5)
            Num3 = AllProcedureList(I, 6)
            Exit For
            
        End If
    Next I
    
    Set Output = VBProjectList(Num1).Modules(Num2).Procedures(Num3)
    
    Set �w�薼�̃v���V�[�W�����擾 = Output
    
End Function

Private Function �v���V�[�W���̎g�p�S�v���V�[�W�����擾(InputProcedure As classProcedure)
    
    Dim ProcedureNameDict As Object
    Set ProcedureNameDict = CreateObject("Scripting.Dictionary")
    
    Call �ċA�^�g�p�v���V�[�W���擾(InputProcedure, ProcedureNameDict, 1, False)
    
    Set �v���V�[�W���̎g�p�S�v���V�[�W�����擾 = ProcedureNameDict
    
End Function

Private Sub �ċA�^�g�p�v���V�[�W���擾(ByVal InputProcedure As classProcedure, ByRef ProcedureNameDict As Object, ByVal Depth&, _
                                       Optional Kakunin As Boolean = True)
    
'    Debug.Print "�K�w�[��", Depth
    
    If Depth = 1 Then
        If Kakunin Then Debug.Print InputProcedure.Name & "(" & InputProcedure.ModuleName & ")"
    End If
    
    If Depth > 10 Then
        Debug.Print "�w��[���ȏ�̊K�w�ɒB���܂���"
        Stop
        Exit Sub
    End If
    
    If InputProcedure.UseProcedure.Count = 0 Then
        Exit Sub
    End If
    
    Dim TmpUseProcedure As classProcedure
    Dim TmpProcedureName$
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    For I = 1 To InputProcedure.UseProcedure.Count
        Set TmpUseProcedure = InputProcedure.UseProcedure(I)
        With TmpUseProcedure
            TmpProcedureName = .VBProjectName & "." & .ModuleName & "." & .Name
        End With
        
        If ProcedureNameDict.Exists(TmpProcedureName) = False Then
            '�o�^�ς݂łȂ�
            If Kakunin Then Debug.Print String(Depth - 1, Chr(9)) & "��" & TmpUseProcedure.Name & "(" & TmpUseProcedure.ModuleName & ")"
            ProcedureNameDict.Add TmpProcedureName, TmpUseProcedure.Code
            Call �ċA�^�g�p�v���V�[�W���擾(TmpUseProcedure, ProcedureNameDict, Depth + 1, Kakunin)
        End If
    Next I
    
End Sub

Private Function �R�[�h���v���C�x�[�g�ɕϊ�(InputCode$)
    
'    Stop
    Dim Output$
    If StrConv(Mid(InputCode, 1, 3), vbUpperCase) = "SUB" Then
        Output = "Private " & InputCode
    ElseIf StrConv(Mid(InputCode, 1, 8), vbUpperCase) = "FUNCTION" Then
        Output = "Private " & InputCode
    ElseIf StrConv(Mid(InputCode, 1, 6), vbUpperCase) = "PUBLIC" Then
        Output = Mid(InputCode, 7)
        Output = "Private" & Output
    Else
        Output = InputCode
    End If
    
    �R�[�h���v���C�x�[�g�ɕϊ� = Output
    
End Function

Private Function �R�[�h���p�u���b�N�ɕϊ�(InputCode$)
    
'    Stop
    Dim Output$
    If StrConv(Mid(InputCode, 1, 3), vbUpperCase) = "SUB" Then
        Output = "Public " & InputCode
    ElseIf StrConv(Mid(InputCode, 1, 8), vbUpperCase) = "FUNCTION" Then
        Output = "Public " & InputCode
    ElseIf StrConv(Mid(InputCode, 1, 7), vbUpperCase) = "PRIVATE" Then
        Output = Mid(InputCode, 8)
        Output = "Public" & Output
    Else
        Output = InputCode
    End If
    
    �R�[�h���p�u���b�N�ɕϊ� = Output
    
End Function

Private Function ���W���[���̐錾�����擾(TopProcedure As classProcedure, UseProcedureNameList)

    Call ������
    Dim AllProcedureList, VBProjectList() As classVBProject
    AllProcedureList = PriAllProcedureList
    VBProjectList = PriVBProjectList
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    If IsEmpty(UseProcedureNameList) Then
        N = 1
    Else
        N = UBound(UseProcedureNameList, 1) + 1
    End If
    
    Dim ProcedureNameList
    ReDim ProcedureNameList(1 To N)
    For I = 1 To N
        If I = 1 Then
            ProcedureNameList(1) = TopProcedure.VBProjectName & "." & TopProcedure.ModuleName & "." & TopProcedure.Name
        Else
            ProcedureNameList(I) = UseProcedureNameList(I - 1)
        End If
    Next I
    
    'VBProject�� & ���W���[���� �ŏd��������
    Dim ModuleNameDict As Object
    Set ModuleNameDict = CreateObject("Scripting.Dictionary")
    Dim TmpModuleName$
    For I = 1 To N
        TmpModuleName = ProcedureNameList(I)
        TmpModuleName = Split(TmpModuleName, ".")(0) & "." & Split(TmpModuleName, ".")(1)
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

Private Function ���{����܂ނ�����(InputStr$)
    
    Dim Hantei As Boolean
    Dim TmpStr$, TmpASC&
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    For I = 1 To Len(InputStr)
        TmpStr = Mid(InputStr, I, 1)
        TmpASC = Asc(TmpStr)
        If Asc(0) <= TmpASC And TmpASC <= Asc("z") Then
            Hantei = False
        Else
            Hantei = True
            Exit For
        End If
    Next I
        
    ���{����܂ނ����� = Hantei
    
End Function
