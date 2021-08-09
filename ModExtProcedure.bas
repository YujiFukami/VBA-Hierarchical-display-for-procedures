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
        
        For J = 1 To TmpVBProject.VBComponents.Count
            Set TmpClassModule = New classModule
            Set TmpModule = TmpVBProject.VBComponents(J)
            
            TmpClassModule.Name = TmpModule.Name
            TmpClassModule.VBProjectName = TmpClassVBProject.Name
            TmpClassModule.ModuleType = モジュール種類判定(TmpModule.Type)
            
            TmpProcedureNameList = モジュールのプロシージャ名一覧取得(TmpModule)
            Set TmpCodeDict = モジュールのコード一覧取得(TmpModule)
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
    If IsEmpty(ProcedureList) Then
        'プロシージャ無し
        Set Output = Nothing
    Else
        'プロシージャ有り
        N = UBound(ProcedureList, 1)
        Set Output = CreateObject("Scripting.Dictionary")
        For I = 1 To N
            TmpProcedureName = ProcedureList(I)
            With InputModule.CodeModule
                On Error Resume Next
                TmpStart = 0
                TmpEnd = 0
                TmpStart = .ProcBodyLine(TmpProcedureName, 0)
                TmpEnd = .ProcCountLines(TmpProcedureName, 0)
                      
                If TmpStart = 0 Then 'クラスモジュールのコード取得用
                    TmpStart = .ProcBodyLine(TmpProcedureName, vbext_pk_Get)
                    TmpEnd = .ProcCountLines(TmpProcedureName, vbext_pk_Let)
                    If TmpEnd = 0 Then
                        TmpEnd = .ProcCountLines(TmpProcedureName, vbext_pk_Get)
                    End If
                End If
                
                On Error GoTo 0
                
                TmpCode = .Lines(TmpStart, TmpEnd)
            End With
            
            Output.Add TmpProcedureName, TmpCode
        Next I
    End If
    
    Set モジュールのコード一覧取得 = Output

End Function

Function コードを検索用に変更(InputCode) As Object
    
    Dim CodeList, TmpStr$
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    CodeList = Split(InputCode, vbLf)
    CodeList = Application.Transpose(CodeList)
    CodeList = Application.Transpose(CodeList)
    N = UBound(CodeList, 1)
    
    Dim BunkatuStrList, HenkanStr$, TmpBunkatu
    BunkatuStrList = Array(" ", ":", "_", ",", """", "(", ")")
    BunkatuStrList = Application.Transpose(Application.Transpose(BunkatuStrList))
    HenkanStr = Chr(13)
    
    Dim BunkatuDict As Object
    Set BunkatuDict = CreateObject("Scripting.Dictionary")
    Dim Output As Object
    Set Output = CreateObject("Scripting.Dictionary")
    
    For I = 1 To N
        TmpStr = CodeList(I)
        TmpStr = Trim(TmpStr) '左右の空白除去
        TmpStr = StrConv(TmpStr, vbUpperCase) '小文字に変換
'        TmpStr = StrConv(TmpStr, vbNarrow) '半角に変換
        If InStr(1, TmpStr, "'") > 0 Then
            TmpStr = Split(TmpStr, "'")(0) 'コメントの除去
        End If
        TmpStr = Replace(TmpStr, Chr(13), "") '改行を消去
        
        
        If TmpStr <> "" Then
            '指定文字で分割する
            For J = 1 To UBound(BunkatuStrList, 1)
                TmpStr = Replace(TmpStr, BunkatuStrList(J), HenkanStr)
            Next J
            TmpBunkatu = Split(TmpStr, HenkanStr)
            
            For J = 0 To UBound(TmpBunkatu)
                If BunkatuDict.Exists(TmpBunkatu(J)) = False Then
                    BunkatuDict.Add TmpBunkatu(J), ""
                End If
            Next J
        End If
    Next I
    
    Set Output = BunkatuDict
    Set コードを検索用に変更 = Output

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
                Call プロシージャ内の外部参照プロシージャ取得(TmpVBProjectName, TmpClassProcedure, TmpExtProcedureList)
            Next II
        Next J
        
        Output(I) = TmpExtProcedureList
    Next I
        
    外部参照プロシージャリスト作成 = Output
    
End Function

Sub プロシージャ内の外部参照プロシージャ取得(VBProjectName$, ClassProcedure As ClassProcedure, ExtProcedureList() As ClassProcedure)
    
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
            Call プロシージャ内の外部参照プロシージャ取得(VBProjectName, TmpUseProcedure, ExtProcedureList)
            
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
    
    Dim HeadStr$
    HeadStr = Split(InputCode, ProcedureName)(0)
    
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
