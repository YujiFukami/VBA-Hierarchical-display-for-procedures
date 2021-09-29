Attribute VB_Name = "ModExportProcedure"
Option Explicit

'プロシージャ単体をモジュールとして出力する

'使用モジュール
'ModExtProcedure

Private PriVBProjectList() As classVBProject
Private PriAllProcedureList


Sub Test_指定名のプロシージャをモジュールで出力()

    Dim InputProcedureName$
    InputProcedureName = "FukamiAddins3.ModArray.TestSortArray2D"
    InputProcedureName = "FukamiAddins3.ModExportProcedure.ExportProcedure"
    
    Call ExportProcedure(InputProcedureName, ActiveWorkbook.Path)
    
End Sub

Function GetProcedureAllCode(InputProcedureName$)
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


Sub ExportProcedure(InputProcedureName$, Optional FolderPath$)
'指定プロシージャを使用しているプロシージャも全部取得して、1つのモジュールとしてエクスポートする
'20210915

'引数
'InputProcedureName                      ・・・プロシージャの名前（VBProject.Module.Procedureのフル名で入力）
'例：FukamiAddins3.ModExtProcedure.Kaiso
'[FolderPath]                            ・・・出力先のフォルダ。省略ならこのブックのフォルダパス

    If FolderPath = "" Then
        FolderPath = ThisWorkbook.Path
    End If
    
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
            TmpCode = コードをパブリックに変換(TmpCode)
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
    
    Dim TmpRowStr$
    '一行内に改行を含んでいたら消去する
    For I = 1 To UBound(OutputList, 1)
        TmpRowStr = OutputList(I, 1)
        
        TmpRowStr = Replace(TmpRowStr, vbLf, "")
        TmpRowStr = Replace(TmpRowStr, vbCrLf, "")
        TmpRowStr = Replace(TmpRowStr, Chr(13), "")
        TmpRowStr = Replace(TmpRowStr, Chr(10), "")
        OutputList(I, 1) = TmpRowStr
        
    Next I
    
    Call OutputText(FolderPath, ModuleFileName & ".bas", OutputList, "")
            
    Debug.Print "「" & ModuleFileName & ".bas" & "」を出力", " 出力先→" & FolderPath
    
End Sub

Private Sub 初期化()

    If IsEmpty(PriAllProcedureList) Then
        PriVBProjectList = フォーム用VBProject作成
        PriAllProcedureList = 全プロシージャ一覧作成(PriVBProjectList)
        Call プロシージャ内の使用プロシージャ取得(PriVBProjectList, PriAllProcedureList)
    End If
    
End Sub

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

Private Function コードをパブリックに変換(InputCode$)
    
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
    
    コードをパブリックに変換 = Output
    
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
