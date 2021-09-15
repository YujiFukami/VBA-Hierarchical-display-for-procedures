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
    
    Me.CmdExtCodeCopy.Enabled = True
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
    TmpCode = ""
    Dim TmpClassProcedure As classProcedure
    
    For I = 1 To UBound(PriExtProcedureList)
        Set TmpClassProcedure = PriProcedureList(I)
        TmpProcedureName = TmpClassProcedure.Name
        
        If TmpProcedureDict.Exists(TmpProcedureName) = False Then
            TmpProcedureDict.Add TmpProcedureName, ""
            TmpCode = TmpCode & vbLf & vbLf & TmpClassProcedure.Code
        End If
    Next I
    
    Call ClipboardCopy(TmpCode)
    MsgBox ("外部参照プロシージャの全コードをクリップボードにコピーしました。")

End Sub

Private Sub CmdExtCodeCopy_Click()

    Call 外部参照プロシージャコードコピー

End Sub

Private Sub CmdSwitch_Click()
    
    If Me.CmdSwitch.Caption = "コード表示" Then
        '階層表示モードに切り替え
        Me.CmdSwitch.Caption = "階層表示"
'        Me.ListViewCode.Visible = False
        Me.ListViewCode.Height = 210
        Me.ListViewCode.Top = 240
        Me.TreeProcedure.Visible = True
        If Not PriShowProcedure Is Nothing Then
            Call プロシージャコード表示(PriShowProcedure)
            Call ツリービューにプロシージャの階層表示(PriShowProcedure)
        End If
    Else
        'コード表示モードに切り替え
        Me.CmdSwitch.Caption = "コード表示"
'        Me.ListViewCode.Visible = True
        Me.ListViewCode.Height = 408
        Me.ListViewCode.Top = 39
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
    
    If UBound(PriProcedureList, 1) <= 0 Then
        Exit Sub
    End If
    
    For I = 1 To UBound(PriProcedureList, 1)
        Select Case Me.ListViewProcedure.SelectedItem
            Case PriProcedureList(I).Name
                
                Set PriShowProcedure = PriProcedureList(I)
                If Me.CmdSwitch.Caption = "コード表示" Then
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
                If Me.CmdSwitch.Caption = "コード表示" Then
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
        .ColumnHeaders(2).Width = 500
    End With
    
    Me.ListViewCode.Visible = True
    Me.TreeProcedure.Visible = False
    Me.CmdExtCodeCopy.Enabled = False
    
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
               "のコードをクリップボードにコピーしました")
        
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
    Application.GoTo Reference:=ReferenceStr
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

