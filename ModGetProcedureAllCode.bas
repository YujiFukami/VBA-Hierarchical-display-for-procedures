Attribute VB_Name = "ModGetProcedureAllCode"
Option Explicit

'GetProcedureAllCode                         �E�E�E���ꏊ�FFukamiAddins3.ModExportProcedure
'�w�薼�̃v���V�[�W�����擾                  �E�E�E���ꏊ�FFukamiAddins3.ModExportProcedure
'������                                      �E�E�E���ꏊ�FFukamiAddins3.ModExportProcedure
'�t�H�[���pVBProject�쐬                     �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'���W���[����ޔ���                          �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'���W���[���̃R�[�h�ꗗ�擾                  �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'���W���[���̃v���V�[�W�����ꗗ�擾          �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'�R�[�h�̎擾�ŋ���                          �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'�R�[�h�̎擾�C��                            �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'�R�[�h��s�������p�ɕϊ�                    �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'�v���V�[�W�����v���p�e�B������              �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'�R�[�h�̎擾�ŋ��Ńv���p�e�B��p            �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'�R�[�h�������p�ɕύX                        �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'�v���V�[�W�������p�����񂩂ǂ�������        �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'�R�[�h����v���V�[�W���̃^�C�v�Ǝg�p�͈͎擾�E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'���W���[���̖`�����擾                      �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'�S�v���V�[�W���ꗗ�쐬                      �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'�v���V�[�W�����̎g�p�v���V�[�W���擾        �E�E�E���ꏊ�FFukamiAddins3.ModExtProcedure   
'�v���V�[�W���̎g�p�S�v���V�[�W�����擾      �E�E�E���ꏊ�FFukamiAddins3.ModExportProcedure
'�ċA�^�g�p�v���V�[�W���擾                  �E�E�E���ꏊ�FFukamiAddins3.ModExportProcedure
'�R�[�h���v���C�x�[�g�ɕϊ�                  �E�E�E���ꏊ�FFukamiAddins3.ModExportProcedure
'���W���[���̐錾�����擾                    �E�E�E���ꏊ�FFukamiAddins3.ModExportProcedure
'���{����܂ނ�����                          �E�E�E���ꏊ�FFukamiAddins3.ModExportProcedure
'MakeAligmentedArray                         �E�E�E���ꏊ�FFukamiAddins3.ModExportProcedure

'------------------------------


'�v���V�[�W���P�̂����W���[���Ƃ��ďo�͂���

'�g�p���W���[��
'ModExtProcedure

Private PriVBProjectList() As classVBProject
Private PriAllProcedureList


'------------------------------

'�O���Q�ƃv���V�[�W���̎擾�p���W���[��
'frmExtRef�ƘA�g���Ă���

'------------------------------


Public Function GetProcedureAllCode(InputProcedureName$)
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

Private Sub ������()

    If IsEmpty(PriAllProcedureList) Then
        PriVBProjectList = �t�H�[���pVBProject�쐬
        PriAllProcedureList = �S�v���V�[�W���ꗗ�쐬(PriVBProjectList)
        Call �v���V�[�W�����̎g�p�v���V�[�W���擾(PriVBProjectList, PriAllProcedureList)
    End If
    
End Sub

Private Function �t�H�[���pVBProject�쐬()
    
    Dim I%, J%, II%, K%, M%, N% '�����グ�p(Integer�^)
    Dim OutputVBProjectList() As classVBProject
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As classProcedure
    Dim VBProjectList As VBProjects, TmpVBProject As VBProject
    Dim TmpModule As VBComponent, TmpProcedureNameList, TmpCodeDict As Object
    Dim TmpProcedureName$
    Dim TmpSengenStr$, TmpFirstProcedureName$
    Dim Dummy
    
    Set VBProjectList = ActiveWorkbook.VBProject.VBE.VBProjects
    ReDim OutputVBProjectList(1 To VBProjectList.Count)
    For I = 1 To VBProjectList.Count
        Set TmpVBProject = VBProjectList(I)
        Set TmpClassVBProject = New classVBProject
        TmpClassVBProject.MyName = TmpVBProject.Name
        
        On Error Resume Next '�p�X��������܂���ւ̑Ώ�
        TmpClassVBProject.MyBookName = Dir(TmpVBProject.FileName)
        On Error GoTo 0
        If TmpClassVBProject.BookName = "" Then '�p�X��������܂���ւ̑Ώ�
            TmpClassVBProject.MyBookName = TmpVBProject.Name & Format(I, "00") '���Ȃ��悤�ɃI���W�i���ԍ���ł�
        End If
        
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
                TmpFirstProcedureName = ""
            Else
                TmpProcedureNameList = TmpCodeDict.Keys
                TmpProcedureNameList = Application.Transpose(Application.Transpose(TmpProcedureNameList))
                TmpFirstProcedureName = TmpProcedureNameList(1)
            End If
            
            TmpSengenStr = ���W���[���̖`�����擾(TmpModule, TmpFirstProcedureName)
            TmpClassModule.Sengen = TmpSengenStr
            
            If IsEmpty(TmpProcedureNameList) = False Then
                For II = 1 To UBound(TmpProcedureNameList)
                    Set TmpClassProcedure = New classProcedure
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

Private Function ���W���[���̃R�[�h�ꗗ�擾(InputModule As VBComponent)
    
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

Private Function ���W���[���̃v���V�[�W�����ꗗ�擾(InputModule As VBComponent)
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim TmpStr$
    Dim Output
    ReDim Output(1 To 1)
    With InputModule.CodeModule
        K = 0
        For I = 1 To .CountOfLines
            If TmpStr <> .ProcofLine(I, 0) Then
                TmpStr = .ProcofLine(I, 0)
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
        
        LastStr = �R�[�h��s�������p�ɕϊ�(LastStr)
        
        If Mid(LastStr, 1, Len("End Function")) = "End Function" Or _
           Mid(LastStr, 1, Len("End Sub")) = "End Sub" Or _
           Mid(LastStr, 1, Len("End Property")) = "End Property" Then
            Output = TmpCode
            Exit For
        End If
        
'        If InStr(1, LastStr, "End Function") > 0 _
'            Or InStr(1, LastStr, "End Sub") > 0 _
'            Or InStr(1, LastStr, "End Property") > 0 Then
'            Output = TmpCode
''            Debug.Print LastStr
'            Exit For
'        End If
        
        
    Next I

    If Output = "" Then
        '����ł��ŏI�s��������Ȃ������ꍇ
        Output = InputModule.CodeModule.Lines(CodeStart, CodeCount)
        Debug.Print Output '�m�F�p
        Stop
    End If

    �R�[�h�̎擾�C�� = Output

End Function

Private Function �R�[�h��s�������p�ɕϊ�(ByVal RowStr$)
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

Private Function �v���V�[�W�����v���p�e�B������(InputModule As VBComponent, ProcedureName$) As Boolean

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

Private Function �R�[�h�������p�ɕύX(InputCode) As Object
    
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

Private Function �v���V�[�W�������p�����񂩂ǂ�������(RowStr$, Str)

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

Private Function ���W���[���̖`�����擾(InputModule As VBComponent, FirstProcedureName$)
    
    Dim Output$
    Dim CodeCount&
    Dim TmpStart&
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    
    If FirstProcedureName <> "" Then
'        Stop
        '�ŏ��̃v���V�[�W���̊J�n�ʒu�擾���āA�J�n�s����v���V�[�W���J�n�ʒu�̎�O�܂ł��擾����
        '�Q�l�Fhttps://docs.microsoft.com/ja-jp/office/vba/api/access.module.procbodyline
        TmpStart = -1
        '�v���V�[�W����Sub/Function�v���V�[�W�����AProperty Get/Let/Set�v���V�[�W�����܂��s���Ȃ̂ŁA�肠���莟��T��B
        On Error Resume Next
        With InputModule.CodeModule
            TmpStart = .ProcBodyLine(FirstProcedureName, vbext_pk_Proc) 'Sub/Function�v���V�[�W��
            If TmpStart = -1 Then
                TmpStart = .ProcBodyLine(FirstProcedureName, vbext_pk_Get) 'Property Get�v���V�[�W��
                If TmpStart = -1 Then
                    TmpStart = .ProcBodyLine(FirstProcedureName, vbext_pk_Let) 'Property Let�v���V�[�W��
                    If TmpStart = -1 Then
                        TmpStart = .ProcBodyLine(FirstProcedureName, vbext_pk_Set) 'Property Set�v���V�[�W��
                    End If
                End If
            End If
            
            Output = .Lines(1, TmpStart - 1)
        End With
'        Stop
        On Error GoTo 0
    Else
'        Stop
        '�v���V�[�W�����Ȃ��ꍇ
        CodeCount = InputModule.CodeModule.CountOfLines
        If CodeCount = 0 Then
            Output = ""
        Else
            Output = InputModule.CodeModule.Lines(1, CodeCount)
        End If
    End If
    
    ���W���[���̖`�����擾 = Output
    
End Function

Private Function �S�v���V�[�W���ꗗ�쐬(VBProjectList)
    
    Dim I&, J&, II&, K&, M&, N& '�����グ�p(Long�^)
    Dim ProcedureCount&
    '�v���V�[�W���̌����v�Z
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As classProcedure
    
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

Private Sub �v���V�[�W�����̎g�p�v���V�[�W���擾(VBProjectList() As classVBProject, AllProcedureList)
    
    Dim I&, J&, II&, JJ&, III&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(AllProcedureList, 1)
    '�v���V�[�W���̌����v�Z
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As classProcedure
    Dim TmpVBProjectNum%, TmpModuleNum%, TmpProcedureNum%
    Dim TmpKensakuCode As Object
    Dim TmpVBProjectName$, TmpModuleName$, TmpProcedureName$
    Dim TmpSiyosakiList As Object
    Dim TmpSiyoProcedure As classProcedure
    Dim TmpSiyoProcedureList() As classProcedure
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
                
                
'                If TmpClassProcedure.Name = "MakeDictFromArrayWithItem" Then Stop
                
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
                        TmpModuleName = TmpSiyoProcedure.ModuleName
                        TmpProcedureName = TmpSiyoProcedure.Name
                        
                        If TmpVBProjectName = TmpClassVBProject.Name And TmpModuleName = TmpClassModule.Name Then
                            '�����Q��(�g�p�v���V�[�W����VBProject���ƃ��W���[���������g�ƈ�v���Ă���)
                            TmpClassProcedure.AddUseProcedure TmpSiyoProcedure
                        ElseIf TmpVBProjectName = TmpClassVBProject.Name And TmpModuleName <> TmpClassModule.Name Then
                            '����VBProject���Q�Ƃ����A�OModule����Q�Ƃ��Ă���B
                            
                            '�������W���[�����ŎQ�Ƃ����łɂ��Ă��邩����
                            NaibuSansyoNaraTrue = False
                            For III = 1 To K
                                If JJ <> III Then
                                    If TmpProcedureName = TmpSiyoProcedureList(III).Name And _
                                       TmpClassModule.Name = TmpSiyoProcedureList(III).ModuleName Then
                                        '�����Q�ƍς�
                                        NaibuSansyoNaraTrue = True
                                        Exit For
                                    End If
                                End If
                            Next III
                            
                            If NaibuSansyoNaraTrue = False Then
                                TmpClassProcedure.AddUseProcedure TmpSiyoProcedure
                            End If
                            
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


