Attribute VB_Name = "ModExtProcedure"
Option Explicit
'外部参照プロシージャの取得用モジュール
'frmExtRefと連携している

Function Kaiso()
    '外部参照プロシージャ一覧表示フォーム起動
    Kaiso = "階層表示"
    Call frmKaiso.Show
    
End Function

Function フォーム用VBProject作成()
    
    Dim I%, J%, II%, K%, M%, N% '数え上げ用(Integer型)
    Dim OutputVBProjectList() As classVBProject
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    Dim VBProjectList As VBProjects, TmpVBProject As VBProject
    Dim TmpModule As VBComponent, TmpProcedureNameList, TmpCodeDict As Object
    Dim TmpProcedureName$
    Dim Dummy
    
    Set VBProjectList = ActiveWorkbook.VBProject.VBE.VBProjects
    ReDim OutputVBProjectList(1 To VBProjectList.Count)
    For I = 1 To VBProjectList.Count
        Set TmpVBProject = VBProjectList(I)
        Set TmpClassVBProject = New classVBProject
        TmpClassVBProject.MyName = TmpVBProject.Name
        TmpClassVBProject.MyBookName = Dir(TmpVBProject.FileName)
        
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
            Else
                TmpProcedureNameList = TmpCodeDict.Keys
                TmpProcedureNameList = Application.Transpose(Application.Transpose(TmpProcedureNameList))
            End If
            
            
            If IsEmpty(TmpProcedureNameList) = False Then
                For II = 1 To UBound(TmpProcedureNameList)
                    Set TmpClassProcedure = New ClassProcedure
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

Function モジュールのプロシージャ名一覧取得(InputModule As VBComponent)
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim TmpStr$
    Dim Output
    ReDim Output(1 To 1)
    With InputModule.CodeModule
        K = 0
        For I = 1 To .CountOfLines
            If TmpStr <> .ProcOfLine(I, 0) Then
                TmpStr = .ProcOfLine(I, 0)
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

Function モジュールのコード一覧取得(InputModule As VBComponent)
    
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

Function コードを検索用に変更(InputCode) As Object
    
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

Function プロシージャ検索用文字列かどうか判定(RowStr$, Str)

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

Private Sub Testコード一行を検索用に変換()
    
    Dim Str$
    Str = "A" & """" & """" & """" & """" & "B"
    Call コード一行を検索用に変換(Str)
    
End Sub

Function コード一行を検索用に変換(ByVal RowStr$)
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

Function 全プロシージャ一覧作成(VBProjectList)
    
    Dim I&, J&, II&, K&, M&, N& '数え上げ用(Long型)
    Dim ProcedureCount&
    'プロシージャの個数を計算
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    
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

Sub プロシージャ内の使用プロシージャ取得(VBProjectList() As classVBProject, AllProcedureList)
    
    Dim I&, J&, II&, JJ&, III&, K&, M&, N& '数え上げ用(Long型)
    N = UBound(AllProcedureList, 1)
    'プロシージャの個数を計算
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    Dim TmpVBProjectNum%, TmpModuleNum%, TmpProcedureNum%
    Dim TmpKensakuCode As Object
    Dim TmpVBProjectName$, TmpModuleName$, TmpProcedureName$
    Dim TmpSiyosakiList As Object
    Dim TmpSiyoProcedure As ClassProcedure
    Dim TmpSiyoProcedureList() As ClassProcedure
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
                        TmpProcedureName = TmpSiyoProcedure.Name
                        
                        If TmpVBProjectName = TmpClassVBProject.Name Then
                            '内部参照(使用プロシージャのVBProject名が自身のVBProject名と一致している)
                            TmpClassProcedure.AddUseProcedure TmpSiyoProcedure
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

Function 外部参照プロシージャ連想配列作成(VBProjectList() As classVBProject)
    
    Dim I&, J&, II&, K&, M&, N& '数え上げ用(Long型)
    'プロシージャの個数を計算
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    
    Dim TmpVBProjectName$, TmpModuleName$, TmpProcedureName$
    Dim TmpCode$
    
    Dim TmpVBProject
    
    Dim TmpExtProcedureDict As Object
    N = UBound(VBProjectList, 1)
    ReDim Output(1 To N)
    For I = 1 To N
        Set TmpExtProcedureDict = CreateObject("Scripting.Dictionary")
        TmpVBProjectName = VBProjectList(I).Name
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            For II = 1 To TmpClassModule.Procedures.Count
                Set TmpClassProcedure = TmpClassModule.Procedures(II)
                Call プロシージャ内の外部参照プロシージャ取得連想配列用(TmpVBProjectName, TmpClassProcedure, TmpExtProcedureDict)
            Next II
        Next J
        
        Set Output(I) = TmpExtProcedureDict
        
    Next I
        
    外部参照プロシージャ連想配列作成 = Output
    
End Function



Sub プロシージャ内の外部参照プロシージャ取得連想配列用(VBProjectName$, ClassProcedure As ClassProcedure, ExtProcedureDict As Object)
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim TmpUseProcedure As ClassProcedure
    Dim TmpUseProcedure2 As ClassProcedure
    Dim GaibuSansyoNaraTrue As Boolean
    
    If ClassProcedure.UseProcedure.Count = 0 Then
        '使用しているプロシージャ無しの場合何もしない
    Else
        For I = 1 To ClassProcedure.UseProcedure.Count
            Set TmpUseProcedure = ClassProcedure.UseProcedure(I)
            
            '再帰(使用プロシージャ内の外部参照を探る)
            Call プロシージャ内の外部参照プロシージャ取得連想配列用(VBProjectName, TmpUseProcedure, ExtProcedureDict)
            
            If TmpUseProcedure.VBProjectName <> VBProjectName Then 'VBProject名が異なれば外部参照
                
                '既に自分のVBProject内に同じ名前のプロシージャが存在すれば、外部参照でない
                GaibuSansyoNaraTrue = True
                For J = 1 To ClassProcedure.UseProcedure.Count
                    Set TmpUseProcedure2 = ClassProcedure.UseProcedure(J)
                    If TmpUseProcedure2.VBProjectName = VBProjectName And TmpUseProcedure2.Name = TmpUseProcedure.Name Then
                        GaibuSansyoNaraTrue = False
                        Exit For
                    End If
                Next J
                
                If GaibuSansyoNaraTrue = True And ExtProcedureDict.Exists(TmpUseProcedure.Name) = False Then
                    ExtProcedureDict.Add TmpUseProcedure.Name, TmpUseProcedure.Code
                End If
            End If
        Next I
    End If

End Sub

Function 外部参照プロシージャリスト作成(VBProjectList() As classVBProject)
    
    Dim I&, J&, II&, K&, M&, N& '数え上げ用(Long型)
    'プロシージャの個数を計算
    Dim TmpClassVBProject As classVBProject
    Dim TmpClassModule As classModule
    Dim TmpClassProcedure As ClassProcedure
    
    Dim TmpVBProjectName$, TmpModuleName$, TmpProcedureName$
    Dim TmpCode$
    
    Dim TmpVBProject
    
    Dim TmpExtProcedureList() As ClassProcedure
    N = UBound(VBProjectList, 1)
    ReDim Output(1 To N)
    For I = 1 To N
        ReDim TmpExtProcedureList(1 To 1)
        TmpVBProjectName = VBProjectList(I).Name
        Set TmpClassVBProject = VBProjectList(I)
        For J = 1 To TmpClassVBProject.Modules.Count
            Set TmpClassModule = TmpClassVBProject.Modules(J)
            For II = 1 To TmpClassModule.Procedures.Count
                Set TmpClassProcedure = TmpClassModule.Procedures(II)
                Call プロシージャ内の外部参照プロシージャ取得(TmpVBProjectName, TmpClassProcedure, TmpExtProcedureList, 0)
            Next II
        Next J
        
        Output(I) = TmpExtProcedureList
    Next I
        
    外部参照プロシージャリスト作成 = Output
    
End Function

Sub プロシージャ内の外部参照プロシージャ取得(VBProjectName$, ClassProcedure As ClassProcedure, ExtProcedureList() As ClassProcedure, ByVal Depth&)
    
    '再帰関数の深さ（ループ）が一定以上超えないようにする。
    Depth = Depth + 1
    If Depth > 10 Then
        Debug.Print "外部参照プロシージャ探索で、規定数の階層を超えました。"
        Debug.Print ClassProcedure.Name
        Exit Sub
    End If
    
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim TmpUseProcedure As ClassProcedure
    Dim TmpUseProcedure2 As ClassProcedure
    Dim GaibuSansyoNaraTrue As Boolean
    Dim TmpHantei As Boolean
    
    If ClassProcedure.UseProcedure.Count = 0 Then
        '使用しているプロシージャ無しの場合何もしない
    Else
        For I = 1 To ClassProcedure.UseProcedure.Count
            Set TmpUseProcedure = ClassProcedure.UseProcedure(I)
            
            '再帰(使用プロシージャ内の外部参照を探る)
            Call プロシージャ内の外部参照プロシージャ取得(VBProjectName, TmpUseProcedure, ExtProcedureList, Depth)
            
            If TmpUseProcedure.VBProjectName <> VBProjectName Then 'VBProject名が異なれば外部参照
                
                '既に自分のVBProject内に同じ名前のプロシージャが存在すれば、外部参照でない
                GaibuSansyoNaraTrue = True
                For J = 1 To ClassProcedure.UseProcedure.Count
                    Set TmpUseProcedure2 = ClassProcedure.UseProcedure(J)
                    If TmpUseProcedure2.VBProjectName = VBProjectName And TmpUseProcedure2.Name = TmpUseProcedure.Name Then
                        GaibuSansyoNaraTrue = False
                        Exit For
                    End If
                Next J
                
                TmpHantei = False
                
                If Not ExtProcedureList(1) Is Nothing Then
                    For J = 1 To UBound(ExtProcedureList, 1)
                        If ExtProcedureList(J).Name = TmpUseProcedure.Name Then
                            '既に取得済み
                            TmpHantei = True
                            Exit For
                        End If
                    Next J
                End If
                
                If GaibuSansyoNaraTrue = True And TmpHantei = False Then
                
                    If Not ExtProcedureList(1) Is Nothing Then
                        ReDim Preserve ExtProcedureList(1 To UBound(ExtProcedureList, 1) + 1)
                    End If
                    
                    Set ExtProcedureList(UBound(ExtProcedureList, 1)) = TmpUseProcedure
                End If
            End If
        Next I
    End If

End Sub

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
        
        If InStr(1, LastStr, "End Function") > 0 _
            Or InStr(1, LastStr, "End Sub") > 0 _
            Or InStr(1, LastStr, "End Property") > 0 Then
            Output = TmpCode
'            Debug.Print LastStr
            Exit For
        End If
    Next I

    If Output = "" Then
        'それでも最終行が見つからなかった場合
        Output = InputModule.CodeModule.Lines(CodeStart, CodeCount)
        Debug.Print Output '確認用
        Stop
    End If

    コードの取得修正 = Output

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

Function プロシージャがプロパティか判定(InputModule As VBComponent, ProcedureName$) As Boolean

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

