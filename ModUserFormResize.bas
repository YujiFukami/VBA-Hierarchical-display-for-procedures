Attribute VB_Name = "ModUserFormResize"
Option Explicit

'// Win32API用定数
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_THICKFRAME = &H40000
'// Win32API参照宣言
'// 64bit版
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
'// 32bit版
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#End If

Private PriIniWidth#           'ユーザーフォームのリサイズ前の幅
Private PriIniHeight#          'ユーザーフォームのリサイズ前の高さ
Private PriResizeCount&        'ユーザーフォームのリサイズ回数
Private PriFontSizeRateList#() '各コントロールのフォントサイズ変更用の比率を格納

Public Sub SetFormEnableResize()
'参考：https://vbabeginner.net/change-form-size-minimize-and-maximize/
'ユーザーフォームのリサイズを可能にする
'ユーザーフォームのイベント(UserForm_Activate)で実行する
'↓をActivateイベントに貼り付けてコメント解除
'   Call SetFormEnableResize

'20211007

#If VBA7 And Win64 Then
    Dim hwnd As LongPtr  'ウインドウハンドル
    Dim style As LongPtr 'ウインドウスタイル
#Else
    Dim hwnd As Long  'ウインドウハンドル
    Dim style As Long 'ウインドウスタイル
#End If

    'ウインドウハンドル取得
    hwnd = GetActiveWindow()
    
    'ウインドウのスタイルを取得
    style = GetWindowLong(hwnd, GWL_STYLE)
    
    'ウインドウのスタイルにウインドウサイズ可変＋最小ボタン＋最大ボタンを追加
    style = style Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
 
    'ウインドウのスタイルを再設定
    Call SetWindowLong(hwnd, GWL_STYLE, style)
    
End Sub

Public Sub InitializeFormResize(TargetForm As Object)
'ユーザーフォームのリサイズ用の初期設定
'ユーザーフォームのイベント(UserForm_Initialize)で実行する。
'↓をInitializeイベントに貼り付けてコメント解除
'   Call InitializeFormResize(Me)

'20211007

'引数
'TargetForm・・・対象とするユーザーフォーム/オブジェクト型

    PriIniHeight = TargetForm.Height '初期状態のユーザーフォームの高さ取得
    PriIniWidth = TargetForm.Width   '初期状態のユーザーフォームの幅取得
    PriResizeCount = 0               'リサイズの回数初期化
    
End Sub

Public Sub ResizeForm(TargetForm As Object, Optional FontSizeResize As Boolean = True)
'ユーザーフォームのコントロールをリサイズする
'ユーザーフォームのイベント(UserForm_Resize)で実行する
'↓をResizeイベントに貼り付けてコメント解除
'   Call ResizeForm(Me)

'20211007

'引数
'TargetForm      ・・・対象とするユーザーフォーム/オブジェクト型
'[FontSizeResize]・・・フォントサイズを変更するかどうか/Boolean型/デフォルトではサイズ変更する

    PriResizeCount = PriResizeCount + 1 'リサイズの回数+1
    
    Dim TmpControl As MSForms.Control                'ユーザーフォーム内の各コントロール
    Dim NowFormHeight#, NowFormWidth#                'サイズ変更後のユーザーフォームのサイズ
    Dim HeightRate#, WidthRate#                      'サイズ変更によるサイズの比率変化
    Dim Top1#, Left1#, Height1#, Width1#, FontSize1# '変更前の各サイズ
    Dim Top2#, Left2#, Height2#, Width2#, FontSize2# '変更後の各サイズ
    
    NowFormHeight = TargetForm.Height         'リサイズ後のユーザーフォームの高さ取得
    NowFormWidth = TargetForm.Width           'リサイズ後のユーザーフォームの幅取得
    HeightRate = NowFormHeight / PriIniHeight 'リサイズ前後での高さ比率
    WidthRate = NowFormWidth / PriIniWidth    'リサイズ前後での幅比率
    
    Dim K&
    If PriResizeCount = 1 Then 'コントロールの数だけフォントサイズの比率の初期状態を保存しておく
        
        ReDim PriFontSizeRateList(1 To TargetForm.Controls.Count)
        
        K = 0
        For Each TmpControl In TargetForm.Controls '各コントロールのフォントサイズ/(高さ+幅)を取得
            K = K + 1
            
            FontSize1 = 0
            On Error Resume Next 'コントロールによってはフォントがない場合もあるのでその際のエラー回避
            FontSize1 = TmpControl.FontSize
            If FontSize1 <> 0 Then
                PriFontSizeRateList(K) = FontSize1 / (TmpControl.Height + TmpControl.Width)
            Else
                FontSize1 = TmpControl.Font.Size 'ツリービューやリストビューはこのプロパティ設定
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
        With TmpControl 'コントロールのリサイズ前の位置、サイズ取得
            Top1 = .Top
            Left1 = .Left
            Height1 = .Height
            Width1 = .Width
'            FontSize1 = .FontSize
        End With
        
        'コントロールのリサイズ後の位置、サイズ計算
        Top2 = Top1 * HeightRate
        Left2 = Left1 * WidthRate
        Height2 = Height1 * HeightRate
        Width2 = Width1 * WidthRate
        
        'コントロールのリサイズ後のフォントサイズ計算
        FontSize2 = (Height2 + Width2) * PriFontSizeRateList(K) 'フォントサイズは高さと幅に対する比率で設定

        With TmpControl 'コントロールのリサイズ後の位置、サイズ、フォントサイズ設定
            .Top = Top2
            .Left = Left2
            .Height = Height2
            .Width = Width2
            
            If FontSizeResize = True Then
                On Error Resume Next 'コントロールによってはフォントがない場合もあるのでその際のエラー回避
                .FontSize = FontSize2
                .Font.Size = FontSize2
                On Error GoTo 0
            End If
        End With
        
    Next
    
    '次のリサイズの際のために、現在のユーザーフォームの高さ、幅を取っておく
    PriIniHeight = NowFormHeight
    PriIniWidth = NowFormWidth
    
End Sub
