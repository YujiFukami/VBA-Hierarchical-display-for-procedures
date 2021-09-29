Attribute VB_Name = "ModRefLibraryForKaiso"
Option Explicit

'RefLibraryForKaiso                                   ・・・元場所：FukamiAddins3.ModRefLibraryForKaiso
'VBIDE参照                                            ・・・元場所：FukamiAddins3.ModRefLibraryForKaiso
'SetRefLibraryGuid                                    ・・・元場所：FukamiAddins3.ModRefLibrary        
'GetLibNameFromGuid                                   ・・・元場所：FukamiAddins3.ModRefLibrary        
'GetRefLibrary                                        ・・・元場所：FukamiAddins3.ModRefLibrary        
'VBAプロジェクトへのアクセス許可設定警告メッセージ表示・・・元場所：FukamiAddins3.ModRefLibrary        
'ExtractColArray                                      ・・・元場所：FukamiAddins3.ModArray             
'CheckArray2D                                         ・・・元場所：FukamiAddins3.ModArray             
'CheckArray2DStart1                                   ・・・元場所：FukamiAddins3.ModArray             
'MakeDictFromArray1D                                  ・・・元場所：FukamiAddins3.ModDictionary        
'CheckArray1D                                         ・・・元場所：FukamiAddins3.ModDictionary        
'CheckArray1DStart1                                   ・・・元場所：FukamiAddins3.ModDictionary        
'MSForms参照                                          ・・・元場所：FukamiAddins3.ModRefLibraryForKaiso
'MSComctlLib参照                                      ・・・元場所：FukamiAddins3.ModRefLibraryForKaiso

'------------------------------


'階層化フォーム用のライブラリ自動参照プログラム

'------------------------------

'20210908
'「終了時ライブラリ参照解除」追加

'------------------------------


'配列の処理関係のプロシージャ

'------------------------------


'連想配列関連モジュール
'------------------------------


Public Sub RefLibraryForKaiso()
    '階層化フォーム用必要ライブラリ参照
    Call VBIDE参照
    Call MSForms参照
    Call MSComctlLib参照

End Sub

Private Sub VBIDE参照()
    
    Dim LibGuid$, LibMajor&, LibMinor&
    LibGuid = "{0002E157-0000-0000-C000-000000000046}"
    LibMajor = 5
    LibMinor = 3
    
    Call SetRefLibraryGuid(LibGuid, LibMajor, LibMinor)
    
End Sub

Private Function SetRefLibraryGuid(LibGuid$, LibMajor&, LibMinor&, Optional TargetBook As Workbook, Optional ShowAlert As Boolean = True)
'指定Guid,Major,Minorのライブラリを参照して結果を返す
'イミディエイトウィンドウに結果を表示する
'20210928

'引数
'LibGuid     ・・・参照するライブラリのGuid（String型）
'LibMajor    ・・・参照するライブラリのMajor（Long型）
'LibMinor    ・・・参照するライブラリのMinor（Long型）
'[TargetBook]・・・参照対象のブック（Workbook型）
'[ShowAlert] ・・・VBAプロジェクトへのアクセス許可警告の表示するかどうか（Boolean型）
                                       
    '引数チェック
    If TargetBook Is Nothing Then
        Set TargetBook = ActiveWorkbook
    End If
    
    '処理
    Dim AddCheck As Boolean
    Dim LibName$
    
    On Error Resume Next
    Call TargetBook.VBProject.References.AddFromGuid(LibGuid, LibMajor, LibMinor)
    
    Select Case Err.Number
        Case 1004
            If ShowAlert = True Then
                Call VBAプロジェクトへのアクセス許可設定警告メッセージ表示
            Else
                Debug.Print "ライブラリ参照の処理ができませんでした"
            End If
            
            AddCheck = False
        Case 32813 '既に参照中
            LibName = GetLibNameFromGuid(LibGuid, TargetBook) 'ライブラリ名取得
            Debug.Print "ライブラリ名「" & LibName & "」"
            Debug.Print "Guid「" & LibGuid & "」は既に参照中です。"
            Debug.Print ""
            '何もしない
            AddCheck = True
        Case -2147319779
            
            Debug.Print "Guid「" & LibGuid & "」は参照できませんでした。"
            Debug.Print ""
            AddCheck = False
            
        Case Else '参照で追加した
            LibName = GetLibNameFromGuid(LibGuid, TargetBook) 'ライブラリ名取得
            Debug.Print "ライブラリ名「" & LibName & "」"
            Debug.Print "Guid「" & LibGuid & "」を参照しました。"
            Debug.Print ""
            AddCheck = True
    End Select
    
'        Debug.Print Err.Number '確認用
    On Error GoTo 0
    
    '出力
    SetRefLibraryGuid = AddCheck
    
End Function

Private Function GetLibNameFromGuid(LibGuid$, Optional TargetBook As Workbook)
'ライブラリのGuidからライブラリ名を取得する
'20210928
   
'引数
'LibGuid     ・・・参照するライブラリのGuid（String型）
'[TargetBook]・・・対象のワークブック(Workbookオブジェクト)（デフォルトはThisWorkbook）
    
    '引数チェック
    If TargetBook Is Nothing Then
        Set TargetBook = ActiveWorkbook
    End If
    
    '参照中のライブラリリストを取得
    Dim LibraryList
    LibraryList = GetRefLibrary(TargetBook)
    
    'Guidのリストと、名前のリストを取得
    Dim GuidList, NameList
    GuidList = ExtractColArray(LibraryList, 5)
    NameList = ExtractColArray(LibraryList, 3)
    
    'フルパスをKey,名前をItemに連想配列作成
    Dim GuidDict As Object
    Set GuidDict = MakeDictFromArray1D(GuidList, NameList)
    
    'ライブラリ名を取得
    Dim Output$
    If GuidDict.Exists(LibGuid) = True Then
        Output = GuidDict(LibGuid)
    Else
        Debug.Print "「" & LibGuid & "」の名前は分かりませんでした"
        Output = ""
    End If
    
    '出力
    GetLibNameFromGuid = Output

End Function

Private Function GetRefLibrary(Optional ByVal TargetBook As Workbook)
'現在参照中のライブラリの一覧を二次元配列で取得する
'20210928

'VBAプロジェクトへのアクセスを許可してください
    
'引数
'[TargetBook]・・・対象のワークブック(Workbookオブジェクト)（デフォルトはThisWorkbook）
    
    '引数チェック
    If TargetBook Is Nothing Then
        Set TargetBook = ThisWorkbook
    End If
    
    Dim OutputStr$, LibName$, LibDes$, LibPath$, TmpStatus$, LibGuid$, LibMajor&, LibMinor&
    
    Dim TmpRef
    OutputStr = ""
    
    On Error GoTo ErrorEscape:
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = TargetBook.VBProject.References.Count '参照ライブラリの個数取得
    Dim Output
    ReDim Output(1 To N, 1 To 7) '1:参照状況,2:名前（省略）,3:名前,4:フルパス,5:Guid,6:Major,7:Minor
    
    K = 0
    For Each TmpRef In TargetBook.VBProject.References
        
        If TmpRef.IsBroken = False Then '参照中
            TmpStatus = "参照中"
            LibName = TmpRef.Name 'ライブラリの名前（省略）
            LibDes = TmpRef.Description 'ライブラリの名前
            LibPath = TmpRef.FullPath 'ライブラリのフルパス
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
            TmpStatus = "参照不可"
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
        Call VBAプロジェクトへのアクセス許可設定警告メッセージ表示
    End If
    
End Function

Private Sub VBAプロジェクトへのアクセス許可設定警告メッセージ表示()
    
    Dim MsgAns As Integer
    
    MsgAns = vbNo
    
    Do While MsgAns = vbNo
        MsgAns = MsgBox("VBAプロジェクトへのアクセス許可の設定をしてください。" & vbLf & _
                "＜設定方法＞" & vbLf & _
                "「タブ：ファイル」" & vbLf & "↓" & vbLf & _
                "「オプション」" & vbLf & "↓" & vbLf & _
                "「セキュリティセンター」" & vbLf & "↓" & vbLf & _
                "「セキュリティーセンターの設定」" & vbLf & "↓" & vbLf & _
                "「マクロの設定」" & vbLf & "↓" & vbLf & _
                "「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェック", vbYesNo)
    Loop
    
    End

End Sub

Private Function ExtractColArray(Array2D, TargetCol&)
'二次元配列の指定列を一次元配列で抽出する
'20210917

'引数
'Array2D  ・・・二次元配列
'TargetCol・・・抽出する対象の列番号


    '引数チェック
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(Array2D, 1) '行数
    M = UBound(Array2D, 2) '列数
 
    If TargetCol < 1 Then
        MsgBox ("抽出する列番号は1以上の値を入れてください")
        Stop
        End
    ElseIf TargetCol > N Then
        MsgBox ("抽出する列番号は元の二次元配列の行数" & M & "以下の値を入れてください")
        Stop
        End
    End If
    
    '処理
    Dim Output
    ReDim Output(1 To N)
    
    For I = 1 To N
        Output(I) = Array2D(I, TargetCol)
    Next I
    
    '出力
    ExtractColArray = Output
    
End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName$ = "配列")
'入力配列が2次元配列かどうかチェックする
'20210804

    Dim Dummy2%, Dummy3%
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "は2次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "配列")
'入力2次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Function MakeDictFromArray1D(KeyArray1D, ItemArray1D)
'配列から連想配列を作成する
'各配列の要素の開始番号は1とすること
'20210806作成

'KeyArray1D   ：Keyが入った一次元配列
'ItemArray1D  ：Itemが入った一次元配列

    '引数チェック
    Call CheckArray1D(KeyArray1D, "KeyArray1D") '2次元配列かチェック
    Call CheckArray1DStart1(KeyArray1D, "KeyArray1D") '要素の開始番号が1かチェック
    Call CheckArray1D(ItemArray1D, "ItemArray1D") '1次元配列かチェック
    Call CheckArray1DStart1(ItemArray1D, "ItemArray1D") '要素の開始番号が1かチェック
    If UBound(KeyArray1D, 1) <> UBound(ItemArray1D, 1) Then
        MsgBox ("「KeyArray1D」と「ItemArray1D」の縦要素数を一致させてください")
        Stop
        End
    End If
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
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

Private Sub CheckArray1D(InputArray, Optional HairetuName$ = "配列")
'入力配列が1次元配列かどうかチェックする
'20210804

    Dim Dummy%
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "は1次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName$ = "配列")
'入力1次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub MSForms参照()
    
    Dim LibGuid$, LibMajor&, LibMinor&
    LibGuid = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}"
    LibMajor = 2
    LibMinor = 0
    
    Call SetRefLibraryGuid(LibGuid, LibMajor, LibMinor)
    
End Sub

Private Sub MSComctlLib参照()
    
    Dim LibGuid$, LibMajor&, LibMinor&
    LibGuid = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}"
    LibMajor = 2
    LibMinor = 2
    
    Call SetRefLibraryGuid(LibGuid, LibMajor, LibMinor)
    
End Sub


