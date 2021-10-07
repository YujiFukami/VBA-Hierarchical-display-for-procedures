Attribute VB_Name = "ModGetProcedureAllCode"
Option Explicit

'GetProcedureAllCode                         ・・・元場所：FukamiAddins3.ModExportProcedure
'指定名のプロシージャを取得                  ・・・元場所：FukamiAddins3.ModExportProcedure
'初期化                                      ・・・元場所：FukamiAddins3.ModExportProcedure
'フォーム用VBProject作成                     ・・・元場所：FukamiAddins3.ModExtProcedure   
'モジュール種類判定                          ・・・元場所：FukamiAddins3.ModExtProcedure   
'モジュールのコード一覧取得                  ・・・元場所：FukamiAddins3.ModExtProcedure   
'モジュールのプロシージャ名一覧取得          ・・・元場所：FukamiAddins3.ModExtProcedure   
'コードの取得最強版                          ・・・元場所：FukamiAddins3.ModExtProcedure   
'コードの取得修正                            ・・・元場所：FukamiAddins3.ModExtProcedure   
'コード一行を検索用に変換                    ・・・元場所：FukamiAddins3.ModExtProcedure   
'プロシージャがプロパティか判定              ・・・元場所：FukamiAddins3.ModExtProcedure   
'コードの取得最強版プロパティ専用            ・・・元場所：FukamiAddins3.ModExtProcedure   
'コードを検索用に変更                        ・・・元場所：FukamiAddins3.ModExtProcedure   
'プロシージャ検索用文字列かどうか判定        ・・・元場所：FukamiAddins3.ModExtProcedure   
'コードからプロシージャのタイプと使用範囲取得・・・元場所：FukamiAddins3.ModExtProcedure   
'モジュールの冒頭文取得                      ・・・元場所：FukamiAddins3.ModExtProcedure   
'全プロシージャ一覧作成                      ・・・元場所：FukamiAddins3.ModExtProcedure   
'プロシージャ内の使用プロシージャ取得        ・・・元場所：FukamiAddins3.ModExtProcedure   
'プロシージャの使用全プロシージャを取得      ・・・元場所：FukamiAddins3.ModExportProcedure
'再帰型使用プロシージャ取得                  ・・・元場所：FukamiAddins3.ModExportProcedure
'コードをプライベートに変換                  ・・・元場所：FukamiAddins3.ModExportProcedure
'モジュールの宣言文を取得                    ・・・元場所：FukamiAddins3.ModExportProcedure
'日本語を含むか判定                          ・・・元場所：FukamiAddins3.ModExportProcedure
'MakeAligmentedArray                         ・・・元場所：FukamiAddins3.ModExportProcedure

'------------------------------


'プロシージャ単体をモジュールとして出力する

'使用モジュール
'ModExtProcedure

Private PriVBProjectList() As classVBProject
Private PriAllProcedureList


'------------------------------

'外部参照プロシージャの取得用モジュール
'frmExtRefと連携している

'------------------------------


Public Function GetProcedureAllCode(InputProcedureName$)
'指定プロシージャを使用しているプロシージャも全部取得して、コードを取得する
'20210916

'引数
'InputProcedureName ・・・プロシージャの名前（VBProject.Module.Procedureのフル名で入力）
'例：FukamiAddins3.ModExtProcedure.Kaiso
   
    Dim TmpProcedure As classProcedure
    Set TmpProcedure = 指定名のプロシージャを取得(InputProcedureName)
    
    Dim ProcedureNameDict
    Set ProcedureNameDict = プロシージャの使用全プロシージャを取得(TmpProcedure)
    
    'モジュールとして出力
    Dim ModuleFileName$
    ModuleFileName = "Mod" & TmpProcedure.Name
    If 日本語を含むか判定(ModuleFileName) Then
        ModuleFileName = Mid(ModuleFileName, 1, 15) & "_"     '日本語を含む場合モジュール名の長さには限界がある
    End If
    
    Dim ProcedureNameList, CodeList
    If ProcedureNameDict.Count > 0 Then
        ProcedureNameList = Application.Transpose(Application.Transpose(ProcedureNameDict.Keys))
        CodeList = Application.Transpose(Application.Transpose(ProcedureNameDict.Items))
    End If
    
    'モジュール宣言文の取得
    Dim SengenList
    SengenList = モジュールの宣言文を取得(TmpProcedure, ProcedureNameList)
    
    '取得したプロシージャ名、コード、宣言文をつなげる
    Dim FixProcedureName$, FixCode$, FixSengen$
    Dim TmpProcedureName$, TmpCode$, TmpSengen$
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    If IsEmpty(ProcedureNameList) Then
        N = 0
    Else
        N = UBound(ProcedureNameList, 1)
    End If
    
    Dim StrProcedureNameList 'コードで表示用のプロシージャ名
    ReDim StrProcedureNameList(1 To N + 1, 1 To 2)
    
    For I = 0 To N
    
        If I = 0 Then
            TmpProcedureName = TmpProcedure.VBProjectName & "." & TmpProcedure.ModuleName & "." & TmpProcedure.Name
            TmpCode = TmpProcedure.Code
        Else
            TmpProcedureName = ProcedureNameList(I)
            TmpCode = CodeList(I)
            TmpCode = コードをプライベートに変換(TmpCode)
        End If
        
        StrProcedureNameList(I + 1, 1) = FixProcedureName & "'" & Split(TmpProcedureName, ".")(2)
        StrProcedureNameList(I + 1, 2) = "元場所：" & Split(TmpProcedureName, ".")(0) & "." & Split(TmpProcedureName, ".")(1)
        
        FixCode = FixCode & TmpCode & vbLf & vbLf
        
    Next I
    
    FixProcedureName = MakeAligmentedArray(StrProcedureNameList, "・・・")
    
    For I = 1 To UBound(SengenList)
        TmpSengen = SengenList(I)
        FixSengen = FixSengen & "'------------------------------" & vbLf
        FixSengen = FixSengen & TmpSengen & vbLf
    Next I
    FixSengen = FixSengen & "'------------------------------" & vbLf
    
    'テキストで出力
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
    
    '出力
    GetProcedureAllCode = OutputStr
    
End Function

Private Function 指定名のプロシージャを取得(InputProcedureName$) As classProcedure

'引数
'InputProcedureName・・・プロシージャの名前（VBProject.Module.Procedureのフル名で入力）
'例：FukamiAddins3.ModExtProcedure.Kaiso

    Dim VBProjectName$, ModuleName$, ProcedureName$
    VBProjectName = Split(InputProcedureName, ".")(0)
    ModuleName = Split(InputProcedureName, ".")(1)
    ProcedureName = Split(InputProcedureName, ".")(2)
    
    Call 初期化
    
    Dim AllProcedureList, VBProjectList() As classVBProject
    AllProcedureList = PriAllProcedureList
    VBProjectList = PriVBProjectList
    
    Dim Output As classProcedure
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
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
    
    Set 指定名のプロシージャを取得 = Output
    
End Function

Private Sub 初期化()

    If IsEmpty(PriAllProcedureList) Then
        PriVBProjectList = フォーム用VBProject作成
        PriAllProcedureList = 全プロシージャ一覧作成(PriVBProjectList)
        Call プロシージャ内の使用プロシージャ取得(PriVBProjectList, PriAllProcedureList)
    End If
    
End Sub

Private Function フォーム用VBProject作成()
    
    Dim I%, J%, II%, K%, M%, N% '数え上げ用(Integer型)
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
        
        On Error Resume Next 'パスが見つかりませんへの対処
        TmpClassVBProject.MyBookName = Dir(TmpVBProject.FileName)
        On Error GoTo 0
        If TmpClassVBProject.BookName = "" Then 'パスが見つかりませんへの対処
            TmpClassVBProject.MyBookName = TmpVBProject.Name & Format(I, "00") '被らないようにオリジナル番号を打つ
        End If
        
        For J = 1 To TmpVBProject.VBComponents.Count
'            If I = 2 And J = 25 Then Stop
            Set TmpClassModule = New classModule
            Set TmpModule = TmpVBProject.VBComponents(J)
            
            TmpClassModule.Name = TmpModule.Name
            TmpClassModule.VBProjectName = TmpClassVBProject.Name
            TmpClassModule.ModuleType = モジュール種類判定(TmpModule.Type)
            TmpClassModule.BookName = TmpClassVBProject.BookName
            
'            TmpProcedureNameList = モジュールのプロシージャ名一覧取得(TmpModule)
            Set TmpCodeDict = モジュールのコード一覧取得(TmpModule)

            If TmpCodeDict Is Nothing Then
                TmpProcedureNameList = Empty
                TmpFirstProcedureName = ""
            Else
                TmpProcedureNameList = TmpCodeDict.Keys
                TmpProcedureNameList = Application.Transpose(Application.Transpose(TmpProcedureNameList))
                TmpFirstProcedureName = TmpProcedureNameList(1)
            End If
            
            TmpSengenStr = モジュールの冒頭文取得(TmpModule, TmpFirstProcedureName)
            TmpClassModule.Sengen = TmpSengenStr
            
            If IsEmpty(TmpProcedureNameList) = False Then
                For II = 1 To UBound(TmpProcedureNameList)
                    Set TmpClassProcedure = New classProcedure
                    TmpProcedureName = TmpProcedureNameList(II)
                    TmpClassProcedure.Name = TmpProcedureName
                    TmpClassProcedure.Code = TmpCodeDict(TmpProcedureName)
                    Dummy = コードからプロシージャのタイプと使用範囲取得(TmpClassProcedure.Code, TmpProcedureName)
                    TmpClassProcedure.RangeOfUse = Dummy(1)
                    TmpClassProcedure.ProcedureType = Dummy(2)
                    Set TmpClassProcedure.KensakuCode = コードを検索用に変更(TmpCodeDict(TmpProcedureName))
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
    
    フォーム用VBProject作成 = OutputVBProjectList
    
End Function

Private Function モジュール種類判定(ModuleType%)
'http://officetanaka.net/excel/vba/vbe/04.htm

    Dim Output$
    Select Case ModuleType
    Case 1
        Output = "標準モジュール"
    Case 2
        Output = "クラスモジュール"
    Case 3
        Output = "ユーザーフォーム"
    Case 11
        Output = "ActiveX デザイナ"
    Case 100
        Output = "Document モジュール"
    Case Else
        MsgBox ("モジュール種類が判定できません")
        Stop
    End Select
    
    モジュール種類判定 = Output
    
End Function

Private Function モジュールのコード一覧取得(InputModule As VBComponent)
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim ProcedureList
    ProcedureList = モジュールのプロシージャ名一覧取得(InputModule)
    Dim Output As Object
    Dim TmpProcedureName$, TmpStart&, TmpEnd&, TmpCode$
    Dim Hantei As Boolean
    Dim Dummy
    Dim TmpProcedureType$
    If IsEmpty(ProcedureList) Then
        'プロシージャ無し
        Set Output = Nothing
    Else
        'プロシージャ有り
        N = UBound(ProcedureList, 1)
        Set Output = CreateObject("Scripting.Dictionary")
        For I = 1 To N
            TmpProcedureName = ProcedureList(I)
            
            Hantei = プロシージャがプロパティか判定(InputModule, TmpProcedureName)
                        
            If Hantei = False Then 'プロシージャがプロパティでない
                TmpCode = コードの取得最強版(InputModule, TmpProcedureName)
                Output.Add TmpProcedureName, TmpCode
            Else
                Dummy = コードの取得最強版プロパティ専用(InputModule, TmpProcedureName)
                For J = 1 To UBound(Dummy, 1)
                    TmpProcedureName = Dummy(J, 1)
                    TmpCode = Dummy(J, 2)
                    Output.Add TmpProcedureName, TmpCode
                Next J
            End If
        Next I
    End If
    
    Set モジュールのコード一覧取得 = Output

End Function

Private Function モジュールのプロシージャ名一覧取得(InputModule As VBComponent)
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
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
    
    If K = 0 Then 'モジュール内にプロシージャがない場合
        Output = Empty
    End If
    
    モジュールのプロシージャ名一覧取得 = Output
        
End Function

Private Function コードの取得最強版(InputModule As VBComponent, ProcedureName$)

    Dim Output$
    Dim TmpStart&, TmpCount&, TmpProcKind%
    
    'プロシージャの開始位置取得
    '参考：https://docs.microsoft.com/ja-jp/office/vba/api/access.module.procbodyline
    TmpStart = -1
    'プロシージャがSub/Functionプロシージャか、Property Get/Let/Setプロシージャかまだ不明なので、手あたり次第探る。
    On Error Resume Next
    With InputModule.CodeModule
        TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Proc) 'Sub/Functionプロシージャ
        TmpProcKind = vbext_pk_Proc
        If TmpStart = -1 Then
            TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Get) 'Property Getプロシージャ
            TmpProcKind = vbext_pk_Get
            If TmpStart = -1 Then
                TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Let) 'Property Letプロシージャ
                TmpProcKind = vbext_pk_Let
                If TmpStart = -1 Then
                    TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Set) 'Property Setプロシージャ
                    TmpProcKind = vbext_pk_Set
                End If
            End If
        End If
        TmpCount = .ProcCountLines(ProcedureName, TmpProcKind)
        
        Output = コードの取得修正(InputModule, TmpStart, TmpCount)
'        Output = .Lines(TmpStart, TmpCount)
    End With
    On Error GoTo 0
    
    コードの取得最強版 = Output

End Function

Private Function コードの取得修正(InputModule As VBComponent, CodeStart&, CodeCount&)

    '通常取得
    Dim TmpCode
    TmpCode = InputModule.CodeModule.Lines(CodeStart, CodeCount)
    Dim LastStr$, TmpSplit, TmpSplit2
    TmpSplit = Split(TmpCode, vbLf)
    LastStr = TmpSplit(UBound(TmpSplit))

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Output$

    'コードのスタート位置から最終行を探索するようにする。
    For I = 2 To UBound(TmpSplit) + 100
        TmpCode = InputModule.CodeModule.Lines(CodeStart, I)
        TmpSplit2 = Split(TmpCode, vbLf)
        LastStr = TmpSplit2(UBound(TmpSplit2))
        
        LastStr = Trim(LastStr) '先頭のスペースを除去
        If InStr(1, LastStr, "'") > 0 Then
            LastStr = Split(LastStr, "'")(0) 'コメントを除去
        End If
        
        LastStr = コード一行を検索用に変換(LastStr)
        
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
        'それでも最終行が見つからなかった場合
        Output = InputModule.CodeModule.Lines(CodeStart, CodeCount)
        Debug.Print Output '確認用
        Stop
    End If

    コードの取得修正 = Output

End Function

Private Function コード一行を検索用に変換(ByVal RowStr$)
'「"」で挟まれた文字列を消去する
    RowStr = Replace(RowStr, """" & """", "")
    Dim TmpSplit
    Dim Output$
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    If InStr(1, RowStr, """") > 0 Then
        TmpSplit = Split(RowStr, """")
        
        For I = 0 To UBound(TmpSplit, 1) '奇数番目は「"」で挟まれた文字列である
            If I Mod 2 = 0 Then
                Output = Output & TmpSplit(I) & " "
            End If
        Next I
        
    Else
        Output = RowStr
    End If
        
    コード一行を検索用に変換 = Output
    
End Function

Private Function プロシージャがプロパティか判定(InputModule As VBComponent, ProcedureName$) As Boolean

    Dim Output As Boolean
    Dim TmpStart&, TmpCount&, TmpProcKind%
    
    'プロシージャの開始位置取得
    '参考：https://docs.microsoft.com/ja-jp/office/vba/api/access.module.procbodyline
    TmpStart = -1
    'プロシージャがSub/Functionプロシージャか、Property Get/Let/Setプロシージャかまだ不明なので、手あたり次第探る。
    On Error Resume Next
    With InputModule.CodeModule
        TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Proc) 'Sub/Functionプロシージャ
        TmpProcKind = vbext_pk_Proc
        'ここでエラーでなかったらSubまたはFunctionプロシージャ
        If TmpStart = -1 Then
            'TmpStartの値が取得できていないので、Property
            Output = True
        Else
            Output = False
        End If
    End With
    On Error GoTo 0
    
    プロシージャがプロパティか判定 = Output

End Function

Private Function コードの取得最強版プロパティ専用(InputModule As VBComponent, ProcedureName$)
    
    
    Dim Output
    Dim TmpStart&, TmpCount&, TmpProcKind%
    Dim HanteiGet As Boolean, HanteiLet As Boolean, HanteiSet As Boolean
    
    'プロシージャの開始位置取得
    '参考：https://docs.microsoft.com/ja-jp/office/vba/api/access.module.procbodyline
    
    'まずプロシージャがProperty Get/Let/Setどれになるか判定
    On Error Resume Next
    With InputModule.CodeModule
        TmpStart = -1
        TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Get) 'Property Getプロシージャ
        If TmpStart <> -1 Then
            HanteiGet = True
        Else
            HanteiGet = False
        End If
        
        TmpStart = -1
        TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Let) 'Property Letプロシージャ
        If TmpStart <> -1 Then
            HanteiLet = True
        Else
            HanteiLet = False
        End If
        
        TmpStart = -1
        TmpStart = .ProcBodyLine(ProcedureName, vbext_pk_Set) 'Property Setプロシージャ
        If TmpStart <> -1 Then
            HanteiSet = True
        Else
            HanteiSet = False
        End If
        
    End With
    On Error GoTo 0
    
    'Property Get/Let/Set別々でコードを取得
    Dim CodeCount& '出力するコードの個数
    CodeCount = Abs(HanteiGet + HanteiLet + HanteiSet)
    ReDim Output(1 To CodeCount, 1 To 2) '1:プロシージャ名,2:コード
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim TmpProcedureName$
    Dim TmpCode$
    K = 0
    
    If HanteiGet Then
        K = K + 1
        TmpStart = InputModule.CodeModule.ProcBodyLine(ProcedureName, vbext_pk_Get)
        TmpCount = InputModule.CodeModule.ProcCountLines(ProcedureName, vbext_pk_Get)
        Output(K, 1) = "(Get)" & ProcedureName
        Output(K, 2) = コードの取得修正(InputModule, TmpStart, TmpCount)
    End If
    If HanteiLet Then
        K = K + 1
        TmpStart = InputModule.CodeModule.ProcBodyLine(ProcedureName, vbext_pk_Let)
        TmpCount = InputModule.CodeModule.ProcCountLines(ProcedureName, vbext_pk_Let)
        Output(K, 1) = "(Let)" & ProcedureName
        Output(K, 2) = コードの取得修正(InputModule, TmpStart, TmpCount)
    End If
    If HanteiSet Then
        K = K + 1
        TmpStart = InputModule.CodeModule.ProcBodyLine(ProcedureName, vbext_pk_Set)
        TmpCount = InputModule.CodeModule.ProcCountLines(ProcedureName, vbext_pk_Set)
        Output(K, 1) = "(Set)" & ProcedureName
        Output(K, 2) = コードの取得修正(InputModule, TmpStart, TmpCount)
    End If
    
    
    コードの取得最強版プロパティ専用 = Output

End Function

Private Function コードを検索用に変更(InputCode) As Object
    
    Dim CodeList, TmpStr$, TmpRowStr$
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
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
        TmpStr = Trim(TmpStr) '左右の空白除去
        TmpStr = StrConv(TmpStr, vbUpperCase) '小文字に変換
'        TmpStr = StrConv(TmpStr, vbNarrow) '半角に変換
        If InStr(1, TmpStr, "'") > 0 Then
            TmpStr = Split(TmpStr, "'")(0) 'コメントの除去
        End If
        TmpStr = Replace(TmpStr, Chr(13), "") '改行を消去
        TmpStr = コード一行を検索用に変換(TmpStr)
        
        If TmpStr <> "" Then
            '指定文字で分割する
            For J = 1 To UBound(BunkatuStrList, 1)
                TmpStr = Replace(TmpStr, BunkatuStrList(J), HenkanStr)
            Next J
            TmpBunkatu = Split(TmpStr, HenkanStr)
            
            For J = 0 To UBound(TmpBunkatu)
                If プロシージャ検索用文字列かどうか判定(TmpRowStr, TmpBunkatu(J)) Then
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
    Set コードを検索用に変更 = Output

End Function

Private Function プロシージャ検索用文字列かどうか判定(RowStr$, Str)

    Dim Hantei As Boolean
    Dim HanteiStr$
    HanteiStr = Replace(RowStr, """", "!")
    
    If Str = "+" Or Str = "=" Or Str = "-" Or Str = "/" Or Str = "" Then
        Hantei = False
    ElseIf InStr(1, RowStr, """" & Str & """") > 0 Then '分割する文字列が「"」で挟まれた文字でない
        Hantei = False
    ElseIf Str = "SUB" Or Str = "FUNCTION" Or Str = "END" Or Str = "EXIT" Or Str = "DIM" Or Str = "BYVAL" Or Str = "AS" Or Str = "RANGE" Or Str = "CALL" Then '予約語
        Hantei = False
    ElseIf Str = "ON" Or Str = "ERROR" Or Str = "NEXT" Or Str = "SET" Or Str = "RESUME" Or Str = "OR" Or Str = "ELSEIF" Then '予約語
        Hantei = False
    ElseIf IsNumeric(Mid(Str, 1, 1)) Then '1文字目が数字でない
        Hantei = False
    Else
        Hantei = True
    End If
    
    プロシージャ検索用文字列かどうか判定 = Hantei
    
End Function

Private Function コードからプロシージャのタイプと使用範囲取得(InputCode, ProcedureName$)
    
    Dim ProcedureName2$
    'プロシージャがプロパティの場合の対応
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
        MsgBox ("プロシージャのタイプが判定できません")
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
    
    コードからプロシージャのタイプと使用範囲取得 = Output

End Function

Private Function モジュールの冒頭文取得(InputModule As VBComponent, FirstProcedureName$)
    
    Dim Output$
    Dim CodeCount&
    Dim TmpStart&
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    
    If FirstProcedureName <> "" Then
'        Stop
        '最初のプロシージャの開始位置取得して、開始行からプロシージャ開始位置の手前までを取得する
        '参考：https://docs.microsoft.com/ja-jp/office/vba/api/access.module.procbodyline
        TmpStart = -1
        'プロシージャがSub/Functionプロシージャか、Property Get/Let/Setプロシージャかまだ不明なので、手あたり次第探る。
        On Error Resume Next
        With InputModule.CodeModule
            TmpStart = .ProcBodyLine(FirstProcedureName, vbext_pk_Proc) 'Sub/Functionプロシージャ
            If TmpStart = -1 Then
                TmpStart = .ProcBodyLine(FirstProcedureName, vbext_pk_Get) 'Property Getプロシージャ
                If TmpStart = -1 Then
                    TmpStart = .ProcBodyLine(FirstProcedureName, vbext_pk_Let) 'Property Letプロシージャ
                    If TmpStart = -1 Then
                        TmpStart = .ProcBodyLine(FirstProcedureName, vbext_pk_Set) 'Property Setプロシージャ
                    End If
                End If
            End If
            
            Output = .Lines(1, TmpStart - 1)
        End With
'        Stop
        On Error GoTo 0
    Else
'        Stop
        'プロシージャがない場合
        CodeCount = InputModule.CodeModule.CountOfLines
        If CodeCount = 0 Then
            Output = ""
        Else
            Output = InputModule.CodeModule.Lines(1, CodeCount)
        End If
    End If
    
    モジュールの冒頭文取得 = Output
    
End Function

Private Function 全プロシージャ一覧作成(VBProjectList)
    
    Dim I&, J&, II&, K&, M&, N& '数え上げ用(Long型)
    Dim ProcedureCount&
    'プロシージャの個数を計算
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
    '1:VBProject名
    '2:Module名
    '3:Procedure名
    '4:VBProjectの番号
    '5:Moduleの番号
    '6:Procedureの番号
    
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
    
    全プロシージャ一覧作成 = Output
    
End Function

Private Sub プロシージャ内の使用プロシージャ取得(VBProjectList() As classVBProject, AllProcedureList)
    
    Dim I&, J&, II&, JJ&, III&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(AllProcedureList, 1)
    'プロシージャの個数を計算
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
    
    For I = 1 To UBound(VBProjectList, 1) '各VBProjectにおいての
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count '各モジュールにおいての
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            For II = 1 To TmpClassModule.Procedures.Count '各プロシージャにおいての
                Set TmpClassProcedure = TmpClassModule.Procedures(II)
                Set TmpKensakuCode = TmpClassProcedure.KensakuCode
                K = 0
                ReDim TmpSiyoProcedureList(1 To 1)
                For JJ = 1 To N
                    TmpVBProjectName = AllProcedureList(JJ, 1)
                    TmpModuleName = AllProcedureList(JJ, 2)
                    TmpProcedureName = AllProcedureList(JJ, 3)
                    
                    If TmpProcedureName <> TmpClassProcedure.Name Then '自分自身のプロシージャは検索から省く
                        TmpVBProjectName = StrConv(TmpVBProjectName, vbUpperCase) '検索用に大文字に変換
                        TmpModuleName = StrConv(TmpModuleName, vbUpperCase) '検索用に大文字に変換
                        TmpProcedureName = StrConv(TmpProcedureName, vbUpperCase) '検索用に大文字に変換
                        
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
                                    '使用プロシージャがPrivateで同じモジュール、VBProject内にいる
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
                
                '外部参照しているが、内部でも同じ名前で参照しているときは除外する
                If K = 0 Then
                    '使用プロシージャなし・・・何もしない
                ElseIf K = 1 Then
                    '使用プロシージャ1つ・・・そのまま使用先で格納
                    TmpClassProcedure.AddUseProcedure TmpSiyoProcedureList(1)
                Else '使用プロシージャが2つ以上
                    For JJ = 1 To K
                        Set TmpSiyoProcedure = TmpSiyoProcedureList(JJ)
                        TmpVBProjectName = TmpSiyoProcedure.VBProjectName
                        TmpModuleName = TmpSiyoProcedure.ModuleName
                        TmpProcedureName = TmpSiyoProcedure.Name
                        
                        If TmpVBProjectName = TmpClassVBProject.Name And TmpModuleName = TmpClassModule.Name Then
                            '内部参照(使用プロシージャのVBProject名とモジュール名が自身と一致している)
                            TmpClassProcedure.AddUseProcedure TmpSiyoProcedure
                        ElseIf TmpVBProjectName = TmpClassVBProject.Name And TmpModuleName <> TmpClassModule.Name Then
                            '同じVBProject内参照だが、外Moduleから参照している。
                            
                            '同じモジュール内で参照がすでにしてあるか判定
                            NaibuSansyoNaraTrue = False
                            For III = 1 To K
                                If JJ <> III Then
                                    If TmpProcedureName = TmpSiyoProcedureList(III).Name And _
                                       TmpClassModule.Name = TmpSiyoProcedureList(III).ModuleName Then
                                        '内部参照済み
                                        NaibuSansyoNaraTrue = True
                                        Exit For
                                    End If
                                End If
                            Next III
                            
                            If NaibuSansyoNaraTrue = False Then
                                TmpClassProcedure.AddUseProcedure TmpSiyoProcedure
                            End If
                            
                        Else
                            '外部参照(使用プロシージャのVBProject名が自身のVBProject名と一致していない)
                            
                            '内部参照がすでにしてあるか判定
                            NaibuSansyoNaraTrue = False
                            For III = 1 To K
                                If JJ <> III Then
                                    If TmpProcedureName = TmpSiyoProcedureList(III).Name And _
                                       TmpClassVBProject.Name = TmpSiyoProcedureList(III).VBProjectName Then
                                        '内部参照済み
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

Private Function プロシージャの使用全プロシージャを取得(InputProcedure As classProcedure)
    
    Dim ProcedureNameDict As Object
    Set ProcedureNameDict = CreateObject("Scripting.Dictionary")
    
    Call 再帰型使用プロシージャ取得(InputProcedure, ProcedureNameDict, 1, False)
    
    Set プロシージャの使用全プロシージャを取得 = ProcedureNameDict
    
End Function

Private Sub 再帰型使用プロシージャ取得(ByVal InputProcedure As classProcedure, ByRef ProcedureNameDict As Object, ByVal Depth&, _
                                       Optional Kakunin As Boolean = True)
    
'    Debug.Print "階層深さ", Depth
    
    If Depth = 1 Then
        If Kakunin Then Debug.Print InputProcedure.Name & "(" & InputProcedure.ModuleName & ")"
    End If
    
    If Depth > 10 Then
        Debug.Print "指定深さ以上の階層に達しました"
        Stop
        Exit Sub
    End If
    
    If InputProcedure.UseProcedure.Count = 0 Then
        Exit Sub
    End If
    
    Dim TmpUseProcedure As classProcedure
    Dim TmpProcedureName$
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    For I = 1 To InputProcedure.UseProcedure.Count
        Set TmpUseProcedure = InputProcedure.UseProcedure(I)
        With TmpUseProcedure
            TmpProcedureName = .VBProjectName & "." & .ModuleName & "." & .Name
        End With
        
        If ProcedureNameDict.Exists(TmpProcedureName) = False Then
            '登録済みでない
            If Kakunin Then Debug.Print String(Depth - 1, Chr(9)) & "└" & TmpUseProcedure.Name & "(" & TmpUseProcedure.ModuleName & ")"
            ProcedureNameDict.Add TmpProcedureName, TmpUseProcedure.Code
            Call 再帰型使用プロシージャ取得(TmpUseProcedure, ProcedureNameDict, Depth + 1, Kakunin)
        End If
    Next I
    
End Sub

Private Function コードをプライベートに変換(InputCode$)
    
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
    
    コードをプライベートに変換 = Output
    
End Function

Private Function モジュールの宣言文を取得(TopProcedure As classProcedure, UseProcedureNameList)

    Call 初期化
    Dim AllProcedureList, VBProjectList() As classVBProject
    AllProcedureList = PriAllProcedureList
    VBProjectList = PriVBProjectList
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
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
    
    'VBProject名 & モジュール名 で重複を消去
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

Private Function 日本語を含むか判定(InputStr$)
    
    Dim Hantei As Boolean
    Dim TmpStr$, TmpASC&
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
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
        
    日本語を含むか判定 = Hantei
    
End Function

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


