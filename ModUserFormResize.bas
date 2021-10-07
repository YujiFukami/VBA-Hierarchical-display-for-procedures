Attribute VB_Name = "ModUserFormResize"
Option Explicit

'// Win32API�p�萔
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000
'// Win32API�Q�Ɛ錾
'// 64bit��
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
'// 32bit��
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#End If

Private PriIniWidth#           '���[�U�[�t�H�[���̃��T�C�Y�O�̕�
Private PriIniHeight#          '���[�U�[�t�H�[���̃��T�C�Y�O�̍���
Private PriResizeCount&        '���[�U�[�t�H�[���̃��T�C�Y��
Private PriFontSizeRateList#() '�e�R���g���[���̃t�H���g�T�C�Y�ύX�p�̔䗦���i�[

Public Sub SetFormEnableResize()
'�Q�l�Fhttps://vbabeginner.net/change-form-size-minimize-and-maximize/
'���[�U�[�t�H�[���̃��T�C�Y���\�ɂ���
'���[�U�[�t�H�[���̃C�x���g(UserForm_Activate)�Ŏ��s����
'����Activate�C�x���g�ɓ\��t���ăR�����g����
'   Call SetFormEnableResize

'20211007

#If VBA7 And Win64 Then
    Dim hwnd As LongPtr  '�E�C���h�E�n���h��
    Dim style As LongPtr '�E�C���h�E�X�^�C��
#Else
    Dim hwnd As Long  '�E�C���h�E�n���h��
    Dim style As Long '�E�C���h�E�X�^�C��
#End If

    '�E�C���h�E�n���h���擾
    hwnd = GetActiveWindow()
    
    '�E�C���h�E�̃X�^�C�����擾
    style = GetWindowLong(hwnd, GWL_STYLE)
    
    '�E�C���h�E�̃X�^�C���ɃE�C���h�E�T�C�Y�ρ{�ŏ��{�^���{�ő�{�^����ǉ�
    style = style Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
 
    '�E�C���h�E�̃X�^�C�����Đݒ�
    Call SetWindowLong(hwnd, GWL_STYLE, style)
    
End Sub

Public Sub InitializeFormResize(TargetForm As Object)
'���[�U�[�t�H�[���̃��T�C�Y�p�̏����ݒ�
'���[�U�[�t�H�[���̃C�x���g(UserForm_Initialize)�Ŏ��s����B
'����Initialize�C�x���g�ɓ\��t���ăR�����g����
'   Call InitializeFormResize(Me)

'20211007

'����
'TargetForm�E�E�E�ΏۂƂ��郆�[�U�[�t�H�[��/�I�u�W�F�N�g�^

    PriIniHeight = TargetForm.Height '������Ԃ̃��[�U�[�t�H�[���̍����擾
    PriIniWidth = TargetForm.Width   '������Ԃ̃��[�U�[�t�H�[���̕��擾
    PriResizeCount = 0               '���T�C�Y�̉񐔏�����
    
End Sub

Public Sub ResizeForm(TargetForm As Object, Optional FontSizeResize As Boolean = True)
'���[�U�[�t�H�[���̃R���g���[�������T�C�Y����
'���[�U�[�t�H�[���̃C�x���g(UserForm_Resize)�Ŏ��s����
'����Resize�C�x���g�ɓ\��t���ăR�����g����
'   Call ResizeForm(Me)

'20211007

'����
'TargetForm      �E�E�E�ΏۂƂ��郆�[�U�[�t�H�[��/�I�u�W�F�N�g�^
'[FontSizeResize]�E�E�E�t�H���g�T�C�Y��ύX���邩�ǂ���/Boolean�^/�f�t�H���g�ł̓T�C�Y�ύX����

    PriResizeCount = PriResizeCount + 1 '���T�C�Y�̉�+1
    
    Dim TmpControl As MSForms.Control                '���[�U�[�t�H�[�����̊e�R���g���[��
    Dim NowFormHeight#, NowFormWidth#                '�T�C�Y�ύX��̃��[�U�[�t�H�[���̃T�C�Y
    Dim HeightRate#, WidthRate#                      '�T�C�Y�ύX�ɂ��T�C�Y�̔䗦�ω�
    Dim Top1#, Left1#, Height1#, Width1#, FontSize1# '�ύX�O�̊e�T�C�Y
    Dim Top2#, Left2#, Height2#, Width2#, FontSize2# '�ύX��̊e�T�C�Y
    
    NowFormHeight = TargetForm.Height         '���T�C�Y��̃��[�U�[�t�H�[���̍����擾
    NowFormWidth = TargetForm.Width           '���T�C�Y��̃��[�U�[�t�H�[���̕��擾
    HeightRate = NowFormHeight / PriIniHeight '���T�C�Y�O��ł̍����䗦
    WidthRate = NowFormWidth / PriIniWidth    '���T�C�Y�O��ł̕��䗦
    
    Dim K&
    If PriResizeCount = 1 Then '�R���g���[���̐������t�H���g�T�C�Y�̔䗦�̏�����Ԃ�ۑ����Ă���
        
        ReDim PriFontSizeRateList(1 To TargetForm.Controls.Count)
        
        K = 0
        For Each TmpControl In TargetForm.Controls '�e�R���g���[���̃t�H���g�T�C�Y/(����+��)���擾
            K = K + 1
            
            FontSize1 = 0
            On Error Resume Next '�R���g���[���ɂ���Ă̓t�H���g���Ȃ��ꍇ������̂ł��̍ۂ̃G���[���
            FontSize1 = TmpControl.FontSize
            If FontSize1 <> 0 Then
                PriFontSizeRateList(K) = FontSize1 / (TmpControl.Height + TmpControl.Width)
            Else
                FontSize1 = TmpControl.Font.Size '�c���[�r���[�⃊�X�g�r���[�͂��̃v���p�e�B�ݒ�
                If FontSize1 <> 0 Then
                    PriFontSizeRateList(K) = FontSize1 / (TmpControl.Height + TmpControl.Width)
                End If
            End If
            On Error GoTo 0
        Next
        
    End If
    
    K = 0
    For Each TmpControl In TargetForm.Controls
        K = K + 1
        With TmpControl '�R���g���[���̃��T�C�Y�O�̈ʒu�A�T�C�Y�擾
            Top1 = .Top
            Left1 = .Left
            Height1 = .Height
            Width1 = .Width
'            FontSize1 = .FontSize
        End With
        
        '�R���g���[���̃��T�C�Y��̈ʒu�A�T�C�Y�v�Z
        Top2 = Top1 * HeightRate
        Left2 = Left1 * WidthRate
        Height2 = Height1 * HeightRate
        Width2 = Width1 * WidthRate
        
        '�R���g���[���̃��T�C�Y��̃t�H���g�T�C�Y�v�Z
        FontSize2 = (Height2 + Width2) * PriFontSizeRateList(K) '�t�H���g�T�C�Y�͍����ƕ��ɑ΂���䗦�Őݒ�

        With TmpControl '�R���g���[���̃��T�C�Y��̈ʒu�A�T�C�Y�A�t�H���g�T�C�Y�ݒ�
            .Top = Top2
            .Left = Left2
            .Height = Height2
            .Width = Width2
            
            If FontSizeResize = True Then
                On Error Resume Next '�R���g���[���ɂ���Ă̓t�H���g���Ȃ��ꍇ������̂ł��̍ۂ̃G���[���
                .FontSize = FontSize2
                .Font.Size = FontSize2
                On Error GoTo 0
            End If
        End With
        
    Next
    
    '���̃��T�C�Y�̍ۂ̂��߂ɁA���݂̃��[�U�[�t�H�[���̍����A��������Ă���
    PriIniHeight = NowFormHeight
    PriIniWidth = NowFormWidth
    
End Sub
