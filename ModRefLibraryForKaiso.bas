Attribute VB_Name = "ModRefLibraryForKaiso"
Option Explicit

'RefLibraryForKaiso                                   �E�E�E���ꏊ�FFukamiAddins3.ModRefLibraryForKaiso
'VBIDE�Q��                                            �E�E�E���ꏊ�FFukamiAddins3.ModRefLibraryForKaiso
'SetRefLibraryGuid                                    �E�E�E���ꏊ�FFukamiAddins3.ModRefLibrary        
'GetLibNameFromGuid                                   �E�E�E���ꏊ�FFukamiAddins3.ModRefLibrary        
'GetRefLibrary                                        �E�E�E���ꏊ�FFukamiAddins3.ModRefLibrary        
'VBA�v���W�F�N�g�ւ̃A�N�Z�X���ݒ�x�����b�Z�[�W�\���E�E�E���ꏊ�FFukamiAddins3.ModRefLibrary        
'ExtractColArray                                      �E�E�E���ꏊ�FFukamiAddins3.ModArray             
'CheckArray2D                                         �E�E�E���ꏊ�FFukamiAddins3.ModArray             
'CheckArray2DStart1                                   �E�E�E���ꏊ�FFukamiAddins3.ModArray             
'MakeDictFromArray1D                                  �E�E�E���ꏊ�FFukamiAddins3.ModDictionary        
'CheckArray1D                                         �E�E�E���ꏊ�FFukamiAddins3.ModDictionary        
'CheckArray1DStart1                                   �E�E�E���ꏊ�FFukamiAddins3.ModDictionary        
'MSForms�Q��                                          �E�E�E���ꏊ�FFukamiAddins3.ModRefLibraryForKaiso
'MSComctlLib�Q��                                      �E�E�E���ꏊ�FFukamiAddins3.ModRefLibraryForKaiso

'------------------------------


'�K�w���t�H�[���p�̃��C�u���������Q�ƃv���O����

'------------------------------

'20210908
'�u�I�������C�u�����Q�Ɖ����v�ǉ�

'------------------------------


'�z��̏����֌W�̃v���V�[�W��

'------------------------------


'�A�z�z��֘A���W���[��
'------------------------------


Public Sub RefLibraryForKaiso()
    '�K�w���t�H�[���p�K�v���C�u�����Q��
    Call VBIDE�Q��
    Call MSForms�Q��
    Call MSComctlLib�Q��

End Sub

Private Sub VBIDE�Q��()
    
    Dim LibGuid$, LibMajor&, LibMinor&
    LibGuid = "{0002E157-0000-0000-C000-000000000046}"
    LibMajor = 5
    LibMinor = 3
    
    Call SetRefLibraryGuid(LibGuid, LibMajor, LibMinor)
    
End Sub

Private Function SetRefLibraryGuid(LibGuid$, LibMajor&, LibMinor&, Optional TargetBook As Workbook, Optional ShowAlert As Boolean = True)
'�w��Guid,Major,Minor�̃��C�u�������Q�Ƃ��Č��ʂ�Ԃ�
'�C�~�f�B�G�C�g�E�B���h�E�Ɍ��ʂ�\������
'20210928

'����
'LibGuid     �E�E�E�Q�Ƃ��郉�C�u������Guid�iString�^�j
'LibMajor    �E�E�E�Q�Ƃ��郉�C�u������Major�iLong�^�j
'LibMinor    �E�E�E�Q�Ƃ��郉�C�u������Minor�iLong�^�j
'[TargetBook]�E�E�E�Q�ƑΏۂ̃u�b�N�iWorkbook�^�j
'[ShowAlert] �E�E�EVBA�v���W�F�N�g�ւ̃A�N�Z�X���x���̕\�����邩�ǂ����iBoolean�^�j
                                       
    '�����`�F�b�N
    If TargetBook Is Nothing Then
        Set TargetBook = ActiveWorkbook
    End If
    
    '����
    Dim AddCheck As Boolean
    Dim LibName$
    
    On Error Resume Next
    Call TargetBook.VBProject.References.AddFromGuid(LibGuid, LibMajor, LibMinor)
    
    Select Case Err.Number
        Case 1004
            If ShowAlert = True Then
                Call VBA�v���W�F�N�g�ւ̃A�N�Z�X���ݒ�x�����b�Z�[�W�\��
            Else
                Debug.Print "���C�u�����Q�Ƃ̏������ł��܂���ł���"
            End If
            
            AddCheck = False
        Case 32813 '���ɎQ�ƒ�
            LibName = GetLibNameFromGuid(LibGuid, TargetBook) '���C�u�������擾
            Debug.Print "���C�u�������u" & LibName & "�v"
            Debug.Print "Guid�u" & LibGuid & "�v�͊��ɎQ�ƒ��ł��B"
            Debug.Print ""
            '�������Ȃ�
            AddCheck = True
        Case -2147319779
            
            Debug.Print "Guid�u" & LibGuid & "�v�͎Q�Ƃł��܂���ł����B"
            Debug.Print ""
            AddCheck = False
            
        Case Else '�Q�ƂŒǉ�����
            LibName = GetLibNameFromGuid(LibGuid, TargetBook) '���C�u�������擾
            Debug.Print "���C�u�������u" & LibName & "�v"
            Debug.Print "Guid�u" & LibGuid & "�v���Q�Ƃ��܂����B"
            Debug.Print ""
            AddCheck = True
    End Select
    
'        Debug.Print Err.Number '�m�F�p
    On Error GoTo 0
    
    '�o��
    SetRefLibraryGuid = AddCheck
    
End Function

Private Function GetLibNameFromGuid(LibGuid$, Optional TargetBook As Workbook)
'���C�u������Guid���烉�C�u���������擾����
'20210928
   
'����
'LibGuid     �E�E�E�Q�Ƃ��郉�C�u������Guid�iString�^�j
'[TargetBook]�E�E�E�Ώۂ̃��[�N�u�b�N(Workbook�I�u�W�F�N�g)�i�f�t�H���g��ThisWorkbook�j
    
    '�����`�F�b�N
    If TargetBook Is Nothing Then
        Set TargetBook = ActiveWorkbook
    End If
    
    '�Q�ƒ��̃��C�u�������X�g���擾
    Dim LibraryList
    LibraryList = GetRefLibrary(TargetBook)
    
    'Guid�̃��X�g�ƁA���O�̃��X�g���擾
    Dim GuidList, NameList
    GuidList = ExtractColArray(LibraryList, 5)
    NameList = ExtractColArray(LibraryList, 3)
    
    '�t���p�X��Key,���O��Item�ɘA�z�z��쐬
    Dim GuidDict As Object
    Set GuidDict = MakeDictFromArray1D(GuidList, NameList)
    
    '���C�u���������擾
    Dim Output$
    If GuidDict.Exists(LibGuid) = True Then
        Output = GuidDict(LibGuid)
    Else
        Debug.Print "�u" & LibGuid & "�v�̖��O�͕�����܂���ł���"
        Output = ""
    End If
    
    '�o��
    GetLibNameFromGuid = Output

End Function

Private Function GetRefLibrary(Optional ByVal TargetBook As Workbook)
'���ݎQ�ƒ��̃��C�u�����̈ꗗ��񎟌��z��Ŏ擾����
'20210928

'VBA�v���W�F�N�g�ւ̃A�N�Z�X�������Ă�������
    
'����
'[TargetBook]�E�E�E�Ώۂ̃��[�N�u�b�N(Workbook�I�u�W�F�N�g)�i�f�t�H���g��ThisWorkbook�j
    
    '�����`�F�b�N
    If TargetBook Is Nothing Then
        Set TargetBook = ThisWorkbook
    End If
    
    Dim OutputStr$, LibName$, LibDes$, LibPath$, TmpStatus$, LibGuid$, LibMajor&, LibMinor&
    
    Dim TmpRef
    OutputStr = ""
    
    On Error GoTo ErrorEscape:
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = TargetBook.VBProject.References.Count '�Q�ƃ��C�u�����̌��擾
    Dim Output
    ReDim Output(1 To N, 1 To 7) '1:�Q�Ə�,2:���O�i�ȗ��j,3:���O,4:�t���p�X,5:Guid,6:Major,7:Minor
    
    K = 0
    For Each TmpRef In TargetBook.VBProject.References
        
        If TmpRef.IsBroken = False Then '�Q�ƒ�
            TmpStatus = "�Q�ƒ�"
            LibName = TmpRef.Name '���C�u�����̖��O�i�ȗ��j
            LibDes = TmpRef.Description '���C�u�����̖��O
            LibPath = TmpRef.FullPath '���C�u�����̃t���p�X
            LibGuid = TmpRef.GUID
            LibMajor = TmpRef.Major
            LibMinor = TmpRef.Minor
            
            K = K + 1
            Output(K, 1) = TmpStatus
            Output(K, 2) = LibName
            Output(K, 3) = LibDes
            Output(K, 4) = LibPath
            Output(K, 5) = LibGuid
            Output(K, 6) = LibMajor
            Output(K, 7) = LibMinor
        Else
            TmpStatus = "�Q�ƕs��"
            LibGuid = TmpRef.GUID
            LibMajor = TmpRef.Major
            LibMinor = TmpRef.Minor
        
            K = K + 1
            Output(K, 1) = TmpStatus
            Output(K, 2) = ""
            Output(K, 3) = ""
            Output(K, 4) = ""
            Output(K, 5) = LibGuid
            Output(K, 6) = LibMajor
            Output(K, 7) = LibMinor
        End If
        
    Next
    
    GetRefLibrary = Output
    Exit Function
    
ErrorEscape:
    If Err.Number = 1004 Then
        Call VBA�v���W�F�N�g�ւ̃A�N�Z�X���ݒ�x�����b�Z�[�W�\��
    End If
    
End Function

Private Sub VBA�v���W�F�N�g�ւ̃A�N�Z�X���ݒ�x�����b�Z�[�W�\��()
    
    Dim MsgAns As Integer
    
    MsgAns = vbNo
    
    Do While MsgAns = vbNo
        MsgAns = MsgBox("VBA�v���W�F�N�g�ւ̃A�N�Z�X���̐ݒ�����Ă��������B" & vbLf & _
                "���ݒ���@��" & vbLf & _
                "�u�^�u�F�t�@�C���v" & vbLf & "��" & vbLf & _
                "�u�I�v�V�����v" & vbLf & "��" & vbLf & _
                "�u�Z�L�����e�B�Z���^�[�v" & vbLf & "��" & vbLf & _
                "�u�Z�L�����e�B�[�Z���^�[�̐ݒ�v" & vbLf & "��" & vbLf & _
                "�u�}�N���̐ݒ�v" & vbLf & "��" & vbLf & _
                "�uVBA�v���W�F�N�g�I�u�W�F�N�g���f���ւ̃A�N�Z�X��M������v�Ƀ`�F�b�N", vbYesNo)
    Loop
    
    End

End Sub

Private Function ExtractColArray(Array2D, TargetCol&)
'�񎟌��z��̎w�����ꎟ���z��Œ��o����
'20210917

'����
'Array2D  �E�E�E�񎟌��z��
'TargetCol�E�E�E���o����Ώۂ̗�ԍ�


    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��
 
    If TargetCol < 1 Then
        MsgBox ("���o�����ԍ���1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf TargetCol > N Then
        MsgBox ("���o�����ԍ��͌��̓񎟌��z��̍s��" & M & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If
    
    '����
    Dim Output
    ReDim Output(1 To N)
    
    For I = 1 To N
        Output(I) = Array2D(I, TargetCol)
    Next I
    
    '�o��
    ExtractColArray = Output
    
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName$ = "�z��")
'���͔z��2�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy2%, Dummy3%
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "��2�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Function MakeDictFromArray1D(KeyArray1D, ItemArray1D)
'�z�񂩂�A�z�z����쐬����
'�e�z��̗v�f�̊J�n�ԍ���1�Ƃ��邱��
'20210806�쐬

'KeyArray1D   �FKey���������ꎟ���z��
'ItemArray1D  �FItem���������ꎟ���z��

    '�����`�F�b�N
    Call CheckArray1D(KeyArray1D, "KeyArray1D") '2�����z�񂩃`�F�b�N
    Call CheckArray1DStart1(KeyArray1D, "KeyArray1D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    Call CheckArray1D(ItemArray1D, "ItemArray1D") '1�����z�񂩃`�F�b�N
    Call CheckArray1DStart1(ItemArray1D, "ItemArray1D") '�v�f�̊J�n�ԍ���1���`�F�b�N
    If UBound(KeyArray1D, 1) <> UBound(ItemArray1D, 1) Then
        MsgBox ("�uKeyArray1D�v�ƁuItemArray1D�v�̏c�v�f������v�����Ă�������")
        Stop
        End
    End If
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(KeyArray1D, 1)
    
    Dim Output As Object
    Set Output = CreateObject("Scripting.Dictionary")
    
    Dim TmpKey$
    
    For I = 1 To N
        TmpKey = KeyArray1D(I)
        If Output.Exists(TmpKey) = False Then
            Output.Add TmpKey, ItemArray1D(I)
        End If
    Next I
    
    Set MakeDictFromArray1D = Output
        
End Function

Private Sub CheckArray1D(InputArray, Optional HairetuName$ = "�z��")
'���͔z��1�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "��1�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName$ = "�z��")
'����1�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub MSForms�Q��()
    
    Dim LibGuid$, LibMajor&, LibMinor&
    LibGuid = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}"
    LibMajor = 2
    LibMinor = 0
    
    Call SetRefLibraryGuid(LibGuid, LibMajor, LibMinor)
    
End Sub

Private Sub MSComctlLib�Q��()
    
    Dim LibGuid$, LibMajor&, LibMinor&
    LibGuid = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}"
    LibMajor = 2
    LibMinor = 2
    
    Call SetRefLibraryGuid(LibGuid, LibMajor, LibMinor)
    
End Sub


