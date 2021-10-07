VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKaiso 
   Caption         =   "階層化表示フォーム"
   ClientHeight    =   9072
   ClientLeft      =   36
   ClientTop       =   408
   ClientWidth     =   15396
   OleObjectBlob   =   "frmKaiso.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmKaiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'色付け用の列挙型(モジュール)
Private Enum ModuleColor
    Module_ = rgbBlue
    Document_ = rgbGreen
    Class_ = rgbLightPink
    Form_ = rgbRed
    ActiveX_ = rgbLightGray
End Enum

'色付け用の列挙型(プロシージャ)
Private Enum ProcedureColor
    SubColor = rgbOrange
    FunctionColor = rgbGreen
    GetColor = rgbBlue
    SetColor = rgbRed
    LetColor = rgbPink
    
End Enum

'起動中VBProjectの全情報
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

'各リストビューのヘッダーの初期幅
Private PriListViewColumnWidthModuleList&(1 To 2)
Private PriListViewColumnWidthProcedureList&(1 To 6)
Private PriListViewColumnWidthUseProcedureList&(1 To 6)
Private PriListViewColumnWidthCodeList&(1 To 2)

'ユーザーフォームの初期サイズ
Private PriIniFormHeight#
Private PriIniFormWidth#

Private Sub Cmb文字サイズ_Change()
    
    '各リスト表示のコントロールのフォントサイズ変更
    Dim SelectedFontSize&
    SelectedFontSize = Me.Cmb文字サイズ.Text
    
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
    
    Call コードの使用先含め全部コピー

End Sub

Private Sub CmdCodeCopy_Click()
    
    Call コードのコピー
    
End Sub

Private Sub CmdExt_Click()
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim ExtProcedureList
    ExtProcedureList = 外部参照プロシージャリスト作成(PriVBProjectList)
    
    If Me.ListVBProject.ListIndex < 0 Then
        Exit Sub
    End If
    
    PriExtProcedureList = ExtProcedureList(Me.ListVBProject.ListIndex + 1)
    
    PriProcedureList = PriExtProcedureList
    
    Call プロシージャリストビュー初期化
    Call 使用プロシージャリストビュー初期化
    
    If Not PriExtProcedureList(1) Is Nothing Then
        
        For I = 1 To UBound(PriExtProcedureList, 1)
                        
            With Me.ListViewProcedure.ListItems.Add
                .Text = PriExtProcedureList(I).Name 'プロシージャ名
                .SubItems(1) = PriExtProcedureList(I).UseProcedure.Count '使用プロシージャ個数
                .SubItems(2) = PriExtProcedureList(I).VBProjectName 'VBProject名
                .SubItems(3) = PriExtProcedureList(I).ModuleName 'モジュール名
                .SubItems(4) = PriExtProcedureList(I).RangeOfUse 'プロシージャ使用可能範囲
                .SubItems(5) = PriExtProcedureList(I).ProcedureType 'プロシージャ種類
            End With
            
            Me.ListViewProcedure.ListItems(I).ForeColor = プロシージャ種類での色取得(PriExtProcedureList(I).ProcedureType)
        
        Next I
        
        Me.CmdExtCodeCopy.Enabled = True
    Else
        MsgBox ("外部参照しているプロシージャは見つかりませんでした")
    End If

End Sub

Private Sub 外部参照プロシージャコードコピー()
    
    If IsEmpty(PriExtProcedureList) Then Exit Sub
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim TmpCode$, TmpProcedureName$
    Dim TmpProcedureDict As Object
    Set TmpProcedureDict = CreateObject("Scripting.Dictionary")
    Dim TmpClassProcedure As classProcedure
    
    Dim SengenList
    SengenList = モジュールの宣言文を取得(PriExtProcedureList)
    
    Dim ProcedureItiran$
    ProcedureItiran = プロシージャ一覧を作成(PriExtProcedureList)
    
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
    MsgBox ("外部参照プロシージャの全コードをクリップボードにコピーしました。" & vbLf & _
            "プロシージャ個数：" & UBound(PriExtProcedureList, 1) & vbLf & _
            "全コード長：" & UBound(Split(TmpCode, vbLf), 1) & vbLf & _
            "文字数：" & Len(TmpCode))

End Sub

Private Sub CmdExtCodeCopy_Click()

    Call 外部参照プロシージャコードコピー

End Sub

Private Sub CmdSwitch_Click()
    
    If Me.CmdSwitch.Caption = "階層表示" Then
        '階層表示モードに切り替え
        Me.CmdSwitch.Caption = "コード表示"
'        Me.ListViewCode.Visible = False
'        Me.ListViewCode.Height = 300
        Me.ListViewCode.Top = Me.TreeProcedure.Top + Me.TreeProcedure.Height + 1
        Me.ListViewCode.Height = Me.Height - Me.ListViewCode.Top - 32
        Me.TreeProcedure.Visible = True
        If Not PriShowProcedure Is Nothing Then
            Call プロシージャコード表示(PriShowProcedure)
            Call ツリービューにプロシージャの階層表示(PriShowProcedure)
        End If
    Else
        'コード表示モードに切り替え
        Me.CmdSwitch.Caption = "階層表示"
'        Me.ListViewCode.Visible = True
        Me.ListViewCode.Top = Me.CmdSwitch.Top + Me.CmdSwitch.Height + 1
        Me.ListViewCode.Height = Me.Height - Me.ListViewCode.Top - 32
        Me.TreeProcedure.Visible = False
        If Not PriShowProcedure Is Nothing Then
            Call プロシージャコード表示(PriShowProcedure)
        End If
    End If

End Sub


Private Sub listVBProject_Click()

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
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
                        .Text = PriModuleList(J).Name 'モジュール名
                        .SubItems(1) = PriModuleList(J).Procedures.Count 'プロシージャの個数
                        .SubItems(2) = PriModuleList(J).ModuleType 'モジュール種類
                        
                    End With
                    
                    Me.ListViewModule.ListItems(J).ForeColor = モジュール種類での色取得(PriModuleList(J).ModuleType)
                Next J
                
                Exit For
                
        End Select
    Next I

End Sub

Private Sub listViewModule_Click()
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
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
                            .Text = PriProcedureList(J).Name 'プロシージャ名
                            .SubItems(1) = PriProcedureList(J).UseProcedure.Count '使用プロシージャ個数
                            .SubItems(2) = PriProcedureList(J).VBProjectName 'VBProject名
                            .SubItems(3) = PriProcedureList(J).ModuleName 'モジュール名
                            .SubItems(4) = PriProcedureList(J).RangeOfUse 'プロシージャ使用可能範囲
                            .SubItems(5) = PriProcedureList(J).ProcedureType 'プロシージャ種類
                        End With
                        
                        Me.ListViewProcedure.ListItems(J).ForeColor = プロシージャ種類での色取得(PriProcedureList(J).ProcedureType)
                    
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
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    
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
                If Me.CmdSwitch.Caption = "階層表示" Then
                    Call プロシージャコード表示(PriShowProcedure)
                Else
                    Call ツリービューにプロシージャの階層表示(PriShowProcedure)
                    Call プロシージャコード表示(PriShowProcedure)
                End If
                
                If PriProcedureList(I).UseProcedure.Count <> 0 Then
                    ReDim PriUseProcedureList(1 To PriProcedureList(I).UseProcedure.Count)
                    
                    For J = 1 To UBound(PriUseProcedureList, 1)
                        
                        Set PriUseProcedureList(J) = PriProcedureList(I).UseProcedure(J)
                        
                        With Me.ListViewUseProcedure.ListItems.Add
                            .Text = PriUseProcedureList(J).Name 'プロシージャ名
                            .SubItems(1) = PriUseProcedureList(J).UseProcedure.Count '使用プロシージャ個数
                            .SubItems(2) = PriUseProcedureList(J).VBProjectName 'VBProject名
                            .SubItems(3) = PriUseProcedureList(J).ModuleName 'モジュール名
                            .SubItems(4) = PriUseProcedureList(J).RangeOfUse 'プロシージャ使用可能範囲
                            .SubItems(5) = PriUseProcedureList(J).ProcedureType 'プロシージャ種類
                        End With
                        
                        Me.ListViewUseProcedure.ListItems(J).ForeColor = プロシージャ種類での色取得(PriUseProcedureList(J).ProcedureType)
                    
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

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    
    On Error GoTo ErrorEscape
    If UBound(PriProcedureList, 1) <= 0 Then
        Exit Sub
    End If
    
    For I = 1 To UBound(PriProcedureList, 1)
        Select Case Me.ListViewProcedure.SelectedItem
            Case PriProcedureList(I).Name
                
                Set PriShowProcedure = PriProcedureList(I)
                
                Call 指定プロシージャVBE画面表示(PriShowProcedure)
                                
        End Select
    Next I
    
    
ErrorEscape:
    Debug.Print Err.Number
    On Error GoTo 0

End Sub

Private Sub ListViewUseProcedure_Click()

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    
    Me.txtVBProject.Text = ""
    Me.txtModule.Text = ""
    
    If UBound(PriUseProcedureList, 1) <= 0 Then
        Exit Sub
    End If
    
    For I = 1 To UBound(PriUseProcedureList, 1)
        Select Case Me.ListViewUseProcedure.SelectedItem
            Case PriUseProcedureList(I).Name
                
                Set PriShowProcedure = PriUseProcedureList(I)
                If Me.CmdSwitch.Caption = "階層表示" Then
                    Call プロシージャコード表示(PriShowProcedure)
                Else
                    Call ツリービューにプロシージャの階層表示(PriShowProcedure)
                    Call プロシージャコード表示(PriShowProcedure)
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

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
       
    If UBound(PriProcedureList, 1) <= 0 Then
        Exit Sub
    End If
    
    For I = 1 To UBound(PriUseProcedureList, 1)
        Select Case Me.ListViewUseProcedure.SelectedItem
            Case PriUseProcedureList(I).Name
                
                Set PriShowProcedure = PriUseProcedureList(I)
                
                Call 指定プロシージャVBE画面表示(PriShowProcedure)
                                
        End Select
    Next I

End Sub

Private Sub TreeProcedure_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim TmpProcedure As classProcedure
    Set TmpProcedure = PriTreeProcedureList(Node.Index)
    
    Call プロシージャコード表示(TmpProcedure)
    
End Sub


Private Sub UserForm_Activate()
    Call SetFormEnableResize
End Sub

Private Sub UserForm_Initialize()

    PriVBProjectList = フォーム用VBProject作成
    
    Dim AllProcedureList
    AllProcedureList = 全プロシージャ一覧作成(PriVBProjectList)
    Call プロシージャ内の使用プロシージャ取得(PriVBProjectList, AllProcedureList)
    
    'フォーム用パブリック変数設定
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    N = UBound(PriVBProjectList, 1)
    ReDim PbVBProjectNameList(1 To N)
    For I = 1 To N
        PbVBProjectNameList(I) = PriVBProjectList(I).Name
    Next I
        
    'フォーム設定
    Me.ListVBProject.List = PbVBProjectNameList
    
    With Me.ListViewModule 'モジュールのリストビューのタブ設定
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add , "モジュール名", "モジュール名"
        .ColumnHeaders.Add , "個数", "個数"
        .ColumnHeaders.Add , "種類", "種類"
        .ColumnHeaders(2).Width = 16
        
        '各ヘッダーの幅をリサイズ用にとっておく
        PriListViewColumnWidthModuleList(1) = .ColumnHeaders(1).Width
        PriListViewColumnWidthModuleList(2) = .ColumnHeaders(2).Width
        
    End With
    
    With Me.ListViewProcedure 'プロシージャのリストビューのタブ設定
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add , "プロシージャ名", "プロシージャ名"
        .ColumnHeaders.Add , "個数", "個数"
        .ColumnHeaders.Add , "VBProject", "VBProject"
        .ColumnHeaders.Add , "モジュール", "モジュール"
        .ColumnHeaders.Add , "範囲", "範囲"
        .ColumnHeaders.Add , "種類", "種類"
        .ColumnHeaders(2).Width = 16
        .ColumnHeaders(3).Width = 20
        .ColumnHeaders(4).Width = 35
        .ColumnHeaders(5).Width = 25
        .ColumnHeaders(6).Width = 25
    
        '各ヘッダーの幅をリサイズ用にとっておく
        PriListViewColumnWidthProcedureList(1) = .ColumnHeaders(1).Width
        PriListViewColumnWidthProcedureList(2) = .ColumnHeaders(2).Width
        PriListViewColumnWidthProcedureList(3) = .ColumnHeaders(3).Width
        PriListViewColumnWidthProcedureList(4) = .ColumnHeaders(4).Width
        PriListViewColumnWidthProcedureList(5) = .ColumnHeaders(5).Width
        PriListViewColumnWidthProcedureList(6) = .ColumnHeaders(6).Width
    
    End With
   
    With Me.ListViewUseProcedure '使用プロシージャのリストビューのタブ設定
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add , "プロシージャ名", "プロシージャ名"
        .ColumnHeaders.Add , "個数", "個数"
        .ColumnHeaders.Add , "VBProject", "VBProject"
        .ColumnHeaders.Add , "モジュール", "モジュール"
        .ColumnHeaders.Add , "範囲", "範囲"
        .ColumnHeaders.Add , "種類", "種類"
        .ColumnHeaders(2).Width = 16
        .ColumnHeaders(3).Width = 20
        .ColumnHeaders(4).Width = 35
        .ColumnHeaders(5).Width = 25
        .ColumnHeaders(6).Width = 25

        '各ヘッダーの幅をリサイズ用にとっておく
        PriListViewColumnWidthUseProcedureList(1) = .ColumnHeaders(1).Width
        PriListViewColumnWidthUseProcedureList(2) = .ColumnHeaders(2).Width
        PriListViewColumnWidthUseProcedureList(3) = .ColumnHeaders(3).Width
        PriListViewColumnWidthUseProcedureList(4) = .ColumnHeaders(4).Width
        PriListViewColumnWidthUseProcedureList(5) = .ColumnHeaders(5).Width
        PriListViewColumnWidthUseProcedureList(6) = .ColumnHeaders(6).Width

    End With

    With Me.ListViewCode 'コードのリストビューのタブ設定
        .View = lvwReport
        .LabelEdit = lvwManual
        .HideSelection = False
        .AllowColumnReorder = True
        .FullRowSelect = True
        .Gridlines = True
        .ColumnHeaders.Add , "行", "行"
        .ColumnHeaders.Add , "コード", "コード"
        .ColumnHeaders(1).Width = 16
        .ColumnHeaders(2).Width = 1000
        
        '各ヘッダーの幅をリサイズ用にとっておく
        PriListViewColumnWidthCodeList(1) = .ColumnHeaders(1).Width
        PriListViewColumnWidthCodeList(2) = .ColumnHeaders(2).Width
    
    End With
    
    Me.ListViewCode.Visible = True
    Me.TreeProcedure.Visible = False
    Me.CmdExtCodeCopy.Enabled = False
    
    With Me.Cmb文字サイズ '文字サイズのコンボボックスの設定
        .List = Array(6, 7, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72)
        .Text = 8
        .ListWidth = 35
        .ColumnWidths = 35
    End With
    
    Call InitializeFormResize(Me)
    
    'リサイズ調整用にユーザーフォームのサイズを取っておく
    PriIniFormHeight = Me.Height
    PriIniFormWidth = Me.Width

End Sub


Private Sub listUseProcedure_Click()

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
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


Private Sub Cmd検索_Click()
    Call コード検索実行(Me.txtKensaku.Text)
End Sub

Private Sub Cmd検索_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call コード検索実行(Me.txtKensaku.Text)
    End If
End Sub

Private Sub txtKensaku_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call コード検索実行(Me.txtKensaku.Text)
    End If
End Sub

Sub コード検索実行(SearchStr$)

    Dim TmpVBProject As classVBProject
    Dim TmpModule As classModule
    Dim TmpProcedure As classProcedure
    
    ReDim PriSearchProcedureList(1 To 1)
    
    Dim I%, J%, II%, K%, M%, N% '数え上げ用(Integer型)
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
        '検索結果あり
        
        'リストビュー初期化
        For I = 1 To Me.ListViewProcedure.ListItems.Count
            Me.ListViewProcedure.ListItems.Remove (1)
        Next I
        For I = 1 To Me.ListViewUseProcedure.ListItems.Count
            Me.ListViewUseProcedure.ListItems.Remove (1)
        Next I
        
        For I = 1 To UBound(PriSearchProcedureList, 1)
            
            With Me.ListViewProcedure.ListItems.Add
                .Text = PriSearchProcedureList(I).Name 'プロシージャ名
                .SubItems(1) = PriSearchProcedureList(I).UseProcedure.Count '使用プロシージャ個数
                .SubItems(2) = PriSearchProcedureList(I).VBProjectName 'VBProject名
                .SubItems(3) = PriSearchProcedureList(I).ModuleName 'モジュール名
                .SubItems(4) = PriSearchProcedureList(I).RangeOfUse 'プロシージャ使用可能範囲
                .SubItems(5) = PriSearchProcedureList(I).ProcedureType 'プロシージャ種類
            End With
            
            Me.ListViewProcedure.ListItems(I).ForeColor = プロシージャ種類での色取得(PriSearchProcedureList(I).ProcedureType)
        
        Next I
        
        PriProcedureList = PriSearchProcedureList
        
        Me.Cmd検索.Caption = "検索" & "(" & UBound(PriSearchProcedureList, 1) & ")"
        
    Else
        MsgBox ("「" & SearchStr & "」" & "検索で見つかりませんでした")
    End If
              
End Sub

Private Function モジュール種類での色取得(ModuleType$)
    
    Dim Output&
    Select Case ModuleType
    Case "標準モジュール"
        Output = ModuleColor.Module_
    Case "クラスモジュール"
        Output = ModuleColor.Class_
    Case "ユーザーフォーム"
        Output = ModuleColor.Form_
    Case "ActiveX デザイナ"
        Output = ModuleColor.ActiveX_
    Case "Document モジュール"
        Output = ModuleColor.Document_
    End Select
        
    モジュール種類での色取得 = Output
    
End Function

Private Function プロシージャ種類での色取得(ProcedureType$)
    
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
        
    プロシージャ種類での色取得 = Output
    
End Function

Private Sub プロシージャコード表示(ShowProcedure As classProcedure)
    
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim TmpCode
    
    '初期化
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

Private Sub コードのコピー()

    If Not PriShowProcedure Is Nothing Then
        Call ClipboardCopy(PriShowProcedure.Code, False)
        
        MsgBox ("「" & PriShowProcedure.Name & "」" & vbLf & _
               "のコードをクリップボードにコピーしました" & vbLf & _
               "コード長：" & UBound(Split(PriShowProcedure.Code, vbLf)) & vbLf & _
               "文字数：" & Len(PriShowProcedure.Code))
        
    End If

End Sub

Private Sub コードの使用先含め全部コピー()
    
    Dim TmpProcedureName$
    Dim CopyCode$
    
    If Not PriShowProcedure Is Nothing Then
        TmpProcedureName = PriShowProcedure.VBProjectName & "." & PriShowProcedure.ModuleName & "." & PriShowProcedure.Name
        CopyCode = GetProcedureAllCode(TmpProcedureName)
        
        Call ClipboardCopy(CopyCode, False)
        
        MsgBox ("「" & PriShowProcedure.Name & "」" & vbLf & _
               "の使用先含めた全コードをクリップボードにコピーしました" & vbLf & _
               "コード長：" & UBound(Split(CopyCode, vbLf)) & vbLf & _
               "文字数：" & Len(CopyCode))
        
    End If

End Sub

Private Sub ツリービューにプロシージャの階層表示(ShowProcedure As classProcedure)
    
    '初期化
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    For I = Me.TreeProcedure.Nodes.Count To 1 Step -1
        Me.TreeProcedure.Nodes.Remove (I)
    Next I
    ReDim PriTreeProcedureList(1 To 1)
    Set PriTreeProcedureList(1) = ShowProcedure
    
    With Me.TreeProcedure
        .Nodes.Add Key:=ShowProcedure.Name, Text:=ShowProcedure.Name & "(" & ShowProcedure.UseProcedure.Count & ")"
        .Nodes(1).Expanded = True
        .Nodes(1).ForeColor = プロシージャ種類での色取得(ShowProcedure.ProcedureType)
        
    End With
    
    Me.txtVBProject.Text = ShowProcedure.VBProjectName
    Me.txtModule.Text = ShowProcedure.ModuleName
    
    Call 再帰型ツリービューにプロシージャの階層表示(ShowProcedure, ShowProcedure.Name, 0)
    
End Sub

Private Sub 再帰型ツリービューにプロシージャの階層表示(ShowProcedure As classProcedure, ParentKey$, ByVal Depth&)
        
    '再帰関数の深さ（ループ）が一定以上超えないようにする。
    Depth = Depth + 1
    Debug.Print String(Depth, " ") & "└" & ShowProcedure.Name
    If Depth > 15 Then
'        Debug.Print "規定数の階層を超えました", ShowProcedure.Name
        Exit Sub
    End If
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
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
                .Nodes(TmpKey).ForeColor = プロシージャ種類での色取得(TmpProcedure.ProcedureType)
                
                Call 再帰型ツリービューにプロシージャの階層表示(TmpProcedure, TmpKey, Depth)
                .Nodes(TmpKey).Expanded = True
                
                
            Next I
        End If
    End With
    
End Sub

Private Sub モジュールリストビュー初期化()
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    For I = 1 To Me.ListViewModule.ListItems.Count
        Me.ListViewModule.ListItems.Remove (1)
    Next I

End Sub

Private Sub プロシージャリストビュー初期化()
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    For I = 1 To Me.ListViewProcedure.ListItems.Count
        Me.ListViewProcedure.ListItems.Remove (1)
    Next I

End Sub

Private Sub 使用プロシージャリストビュー初期化()
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    For I = 1 To Me.ListViewUseProcedure.ListItems.Count
        Me.ListViewUseProcedure.ListItems.Remove (1)
    Next I

End Sub

Private Sub コードプロシージャリストビュー初期化()
        Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    For I = 1 To Me.ListViewCode.ListItems.Count
        Me.ListViewCode.ListItems.Remove (1)
    Next I

End Sub

Private Sub 指定プロシージャVBE画面表示(ShowProcedure As classProcedure)
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
'入力テキストをクリップボードに格納
'配列ならば列方向をTabわけ、行方向を改行する。
'20210719作成
    
    '入力した引数が配列か、配列の場合は1次元配列か、2次元配列か判定
    Dim HairetuHantei%
    Dim Jigen1%, Jigen2%
    If IsArray(InputClipText) = False Then
        '入力引数が配列でない
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
    
    'クリップボードに格納用のテキスト変数を作成
    Dim Output$
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    
    If HairetuHantei = 0 Then '配列でない場合
        Output = InputClipText
    ElseIf HairetuHantei = 1 Then '1次元配列の場合
    
        If LBound(InputClipText, 1) <> 1 Then '最初の要素番号が1出ない場合は最初の要素番号を1にする
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
    ElseIf HairetuHantei = 2 Then '2次元配列の場合
        
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
    
    
    'クリップボードに格納'参考 https://www.ka-net.org/blog/?p=7537
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = Output
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With

    '格納したテキスト変数をメッセージ表示
    If MessageIrunaraTrue Then
        MsgBox ("「" & Output & "」" & vbLf & _
                "をクリップボードにコピーしました。")
    End If
    
End Sub

Private Function モジュールの宣言文を取得(UseProcedureList() As classProcedure)

    Dim AllProcedureList, VBProjectList() As classVBProject
    AllProcedureList = 全プロシージャ一覧作成(PriVBProjectList)
    VBProjectList = PriVBProjectList
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(UseProcedureList, 1)
        
    'VBProject名 & モジュール名 で重複を消去
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
    
    '宣言文を取得
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
    
    'Option Explicitを消去する
    For I = 1 To N
        SengenList(I) = Replace(SengenList(I), "Option Explicit", "")
    Next I
    
    モジュールの宣言文を取得 = SengenList
    
End Function

Private Function プロシージャ一覧を作成(classProcedureList() As classProcedure)

    Dim TmpClassProcedure As classProcedure
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(classProcedureList, 1)
    Dim StrProcedureNameList
    ReDim StrProcedureNameList(1 To N, 1 To 2)
    
    For I = 1 To N
        Set TmpClassProcedure = classProcedureList(I)
        
        StrProcedureNameList(I, 1) = "'" & TmpClassProcedure.Name
        StrProcedureNameList(I, 2) = "元場所：" & TmpClassProcedure.VBProjectName & "." & TmpClassProcedure.ModuleName
        
    Next I
    
    Dim OutputStr$
    OutputStr = MakeAligmentedArray(StrProcedureNameList, "・・・")
    
    プロシージャ一覧を作成 = OutputStr

End Function


'MakeAligmentedArray・・・元場所：FukamiAddins3.ModAlignmentArray
'------------------------------
'文字列配列を整列させて1つの文字列として出力する
'------------------------------

Private Function MakeAligmentedArray(ByVal StrArray, Optional SikiriMoji$ = "：")
    '20210916
    '文字列配列を整列させて1つの文字列として出力する
    
    Dim I&, J&, K&, M&, N&                     '数え上げ用(Long型)
    Dim TateMin&, TateMax&, YokoMin&, YokoMax& '配列の縦横インデックス最大最小
    Dim WithTableArray                         'テーブル付配列…イミディエイトウィンドウに表示する際にインデックス番号を表示したテーブルを追加した配列
    Dim NagasaList, MaxNagasaList              '各文字の文字列長さを格納、各列での文字列長さの最大値を格納
    Dim NagasaOnajiList                        '" "（半角スペース）を文字列に追加して各列で文字列長さを同じにした文字列を格納
    Dim OutputStr                              '文字列を格納
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '入力引数の処理
    Dim Jigen2%
    On Error Resume Next
    Jigen2 = UBound(StrArray, 2)
    On Error GoTo 0
    If Jigen2 = 0 Then '1次元配列は2次元配列にする
        StrArray = Application.Transpose(StrArray)
    End If
    
    TateMin = LBound(StrArray, 1) '配列の縦番号（インデックス）の最小
    TateMax = UBound(StrArray, 1) '配列の縦番号（インデックス）の最大
    YokoMin = LBound(StrArray, 2) '配列の横番号（インデックス）の最小
    YokoMax = UBound(StrArray, 2) '配列の横番号（インデックス）の最大
    
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '各列の幅を同じに整えるために文字列長さとその各列の最大値を計算する。
    N = UBound(StrArray, 1) '「StrArray」の縦インデックス数（行数）
    M = UBound(StrArray, 2) '「StrArray」の横インデックス数（列数）
    ReDim NagasaList(1 To N, 1 To M)
    ReDim MaxNagasaList(1 To M)
    
    Dim TmpStr$
    For J = 1 To M
        For I = 1 To N
        
'            If J > 1 And HyoujiMaxNagasa <> 0 Then
'                '最大表示長さが指定されている場合。
'                '1列目のテーブルはそのままにする。
'                TmpStr = StrArray(I, J)
'                StrArray(I, J) = 文字列を指定バイト数文字数に省略(TmpStr, HyoujiMaxNagasa)
'            End If
            
            NagasaList(I, J) = LenB(StrConv(StrArray(I, J), vbFromUnicode)) '全角と半角を区別して長さを計算する。
            MaxNagasaList(J) = WorksheetFunction.Max(MaxNagasaList(J), NagasaList(I, J))
            
        Next I
    Next J
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '" "(半角スペース)を追加して文字列長さを同じにする。
    ReDim NagasaOnajiList(1 To N, 1 To M)
    Dim TmpMaxNagasa&
    
    For J = 1 To M
        TmpMaxNagasa = MaxNagasaList(J) 'その列の最大文字列長さ
            For I = 1 To N
            'Rept…指定文字列を指定個数連続してつなげた文字列を出力する。
            '（最大文字数-文字数）の分" "（半角スペース）を後ろにくっつける。
            NagasaOnajiList(I, J) = StrArray(I, J) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(I, J))
       
        Next I
    Next J
    
    '※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '文字列を作成
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
    
    ''※※※※※※※※※※※※※※※※※※※※※※※※※※※
    '出力
    MakeAligmentedArray = OutputStr
    
End Function

Private Sub UserForm_Resize()
    Call ResizeForm(Me)
    
    With Me 'VBProjectのリストのサイズだけへんちくりんなる問題を強制解決
        .ListVBProject.Height = .Label1.Top - .ListVBProject.Top
        .ListVBProject.Width = .ListViewModule.Width
    End With
    
    Call リストビューのヘッダー幅調整(Me.ListViewModule, PriListViewColumnWidthModuleList)
    Call リストビューのヘッダー幅調整(Me.ListViewProcedure, PriListViewColumnWidthProcedureList)
    Call リストビューのヘッダー幅調整(Me.ListViewUseProcedure, PriListViewColumnWidthUseProcedureList)
    Call リストビューのヘッダー幅調整(Me.ListViewCode, PriListViewColumnWidthCodeList)
    
End Sub


Private Sub リストビューのヘッダー幅調整(ListView As ListView, HeaderWidthList)
    
    Dim NowFormWidth&
    NowFormWidth = Me.Width
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    For I = 1 To UBound(HeaderWidthList, 1)
        ListView.ColumnHeaders(I).Width = HeaderWidthList(I) * NowFormWidth / PriIniFormWidth
    Next I
    
End Sub
