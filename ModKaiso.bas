Attribute VB_Name = "ModKaiso"
Function Kaiso()
    '階層フォーム起動
    Kaiso = "階層化"
    Call frmKaiso.Show
    
End Function
Function VBProjectリスト取得()
    '起動中のVBProjectをリスト化して取得する。
    '取得するVBProject1つ1つはオブジェクト形式。

    Dim VBProjectList '←出力
    Dim VBProjectCount As Byte
    Dim I% '数え上げ用(Integer型)
    Dim Dummy1
    
    VBProjectCount = ActiveWorkbook.VBProject.VBE.VBProjects.Count 'VBProjectの個数計算
    
    ReDim VBProjectList(1 To VBProjectCount)
    
    For I = 1 To VBProjectCount
        Set Dummy1 = ActiveWorkbook.VBProject.VBE.VBProjects.Item(I)
        Set VBProjectList(I) = Dummy1
        
    Next I
    
    '出力
    VBProjectリスト取得 = VBProjectList
    
End Function
Function 非ロックのVBProjectリスト取得()
    '非ロックのVBProjectをリスト化して取得する。
    '取得する非ロックのVBProject1つ1つはオブジェクト形式。
    
    Dim VBProjectList
    Dim UnLockVBProjectList '←出力
    Dim VBProjectCount As Byte
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Dummy1, Dummy2
    
    VBProjectList = VBProjectリスト取得
    VBProjectCount = UBound(VBProjectList, 1)
        
    K = 0 '数え上げ初期化
    ReDim UnLockVBProjectList(1 To 1)
    
    For I = 1 To VBProjectCount
    
        Set Dummy1 = VBProjectList(I)
        Dummy2 = Dummy1.Protection
            
        If Dummy2 = 1 Then
            'ロックされている
        ElseIf Dummy2 = 0 Then
            'ロックされていない
            K = K + 1
            ReDim Preserve UnLockVBProjectList(1 To K)
            Set UnLockVBProjectList(K) = Dummy1
            
        End If
        
    Next I
    
    '出力
    非ロックのVBProjectリスト取得 = UnLockVBProjectList
    
End Function
Function モジュール一覧取得(objVBProject As Object, TmpProcedureList)
    '指定VBProjectのモジュール一覧を取得する。
    '取得するモジュール1つ1つはオブジェクト形式。
    
    Dim I%, J%, K% '数え上げ用(Integer型)
    
    Dim TmpVBProject As Object
    Set TmpVBProject = objVBProject.VBComponents
    
    Dim ModuleCount As Integer
    ModuleCount = TmpVBProject.Count
    
    Dim ModuleList '←出力
    
    ReDim ModuleList(1 To ModuleCount, 1 To 2)
    'ModuleList(:,1)'モジュール(オブジェクト形式)
    'ModuleList(:,2)'プロシージャのリスト
    
    Dim TmpModuleName As String, TmpProcedureNameList
    
    For I = 1 To ModuleCount
        Set ModuleList(I, 1) = TmpVBProject(I)
        TmpModuleName = TmpVBProject(I).Name
        
        K = 0
        ReDim TmpProcedureNameList(1 To 1)
        For J = 1 To UBound(TmpProcedureList, 1)
            If TmpModuleName = TmpProcedureList(J, 1) Then
                K = K + 1
                ReDim Preserve TmpProcedureNameList(1 To K)
                TmpProcedureNameList(K) = TmpProcedureList(J, 2)
            End If
        Next J
        
        If K = 0 Then
            'モジュールにプロシージャがない場合
            TmpProcedureNameList = Empty
        End If
        
        ModuleList(I, 2) = TmpProcedureNameList
        
    Next I
    
    モジュール一覧取得 = ModuleList
    
End Function
Function プロシージャ一覧取得(objVBProject As Object)
    
    Dim I%, J%, K%, M%, N%, K2% '数え上げ用(Integer型)
    Dim Dummy1, Dummy2, Dummy3
    
    Dim TmpVBProject As Object
    Set TmpVBProject = objVBProject.VBComponents
    
    Dim ModuleKosuu As Integer
    ModuleKosuu = TmpVBProject.Count

    Dim Gyosuu As Integer
    Dim TmpModule As Object
    Dim ProcedureCount As Integer
    Dim CodeStartList, CodeEndList
    ReDim CodeStartList(1 To 50000)
    ReDim CodeEndList(1 To 50000)
    
    Dim Output '取得するプロシージャの数が不明なのでとりあえずたくさん格納できるようにする(笑)
    ReDim Output(1 To 50000, 1 To 3)
    '1：モジュール名
    '2：プロシージャ名
    '3：コード(長さ分の1次元配列)
    Dim AllCodeList
    
    Dim TmpStartIti%, TmpEndIti%, CodeNagasa%, TmpCode As String
    Dim FirstCodeStartIti
    
    Dim ProcedureName As String
    
    K = 0 '数え上げ初期化
    For I = 1 To ModuleKosuu
        Set TmpModule = TmpVBProject(I)
        Gyosuu = TmpModule.CodeModule.CountOfLines 'モジュール内行数
        
        If Gyosuu = 0 Then GoTo ForEscape
        
        ReDim AllCodeList(1 To Gyosuu)
        For J = 1 To Gyosuu
            AllCodeList(J) = TmpModule.CodeModule.ProcofLine(J, 0)
        Next J
                
        ProcedureName = ""
        K2 = 0 'モジュール内での数え上げの初期化
        For J = 1 To Gyosuu
            Dummy2 = TmpModule.CodeModule.ProcofLine(J, 0)
            If ProcedureName <> TmpModule.CodeModule.ProcofLine(J, 0) Then
                K = K + 1
                K2 = K2 + 1
                ProcedureName = TmpModule.CodeModule.ProcofLine(J, 0)
                
                Output(K, 1) = TmpModule.Name 'モジュール名
                Output(K, 2) = ProcedureName 'プロシージャ名
                
                On Error Resume Next 'クラスモジュールの時はCodeModule.ProcStartLineで計算できない？？
                CodeStartList(K) = TmpModule.CodeModule.ProcStartLine(ProcedureName, 0)  '開始行
                On Error GoTo 0
                If IsEmpty(CodeStartList(K)) Then
                    CodeStartList(K) = J
                End If

                If K2 > 1 Then 'モジュール内で2つ目移行のプロシージャのみ
                    CodeEndList(K - 1) = CodeStartList(K) - 1 '終了行(1つ前に取得したプロシージャの終了行)
                    
                    'コードの最終行が分かった時点でコードの取得
                    TmpStartIti = CodeStartList(K - 1)
                    TmpEndIti = CodeEndList(K - 1)
                    CodeNagasa = TmpEndIti - TmpStartIti + 1
                    TmpCode = TmpModule.CodeModule.Lines(TmpStartIti, CodeNagasa)
                    Dummy3 = 改行された文字列を改行で分けて配列にする(TmpCode)
                    Output(K - 1, 3) = コードの先頭空白を除外する(Dummy3)
                                        
                End If
                
            End If
        Next J
        
        If K2 <> 0 Then 'K2<>0ということはモジュール内にプロシージャが存在するということ。
            'モジュール内最後のプロシージャの終了行＝モジュールの行数
            CodeEndList(K) = Gyosuu
            
            'コードの最終行が分かった時点でコードの取得
            TmpStartIti = CodeStartList(K)
            TmpEndIti = CodeEndList(K)
            CodeNagasa = TmpEndIti - TmpStartIti + 1
            TmpCode = TmpModule.CodeModule.Lines(TmpStartIti, CodeNagasa)
            Dummy3 = 改行された文字列を改行で分けて配列にする(TmpCode)
            Output(K, 3) = コードの先頭空白を除外する(Dummy3)
            
        End If

ForEscape:
        
    Next I
        
    ProcedureCount = K  '取得したプロシージャの個数
    
   '取得したプロシージャの個数分の配列にする。
    Dim Output2
    ReDim Output2(1 To ProcedureCount, 1 To UBound(Output, 2))
    
    For I = 1 To ProcedureCount
        For J = 1 To UBound(Output, 2)
            Output2(I, J) = Output(I, J)
        Next J
    Next I
    
    プロシージャ一覧取得 = Output2
    
End Function
Function 改行された文字列を改行で分けて配列にする(Mojiretu)
    Dim Hairetu
    Hairetu = Split(Mojiretu, Chr(10))
    Hairetu = Application.Transpose(Hairetu)
    Hairetu = Application.Transpose(Hairetu)
    
    Dim I%
    For I = 1 To UBound(Hairetu, 1)
        Hairetu(I) = Replace(Hairetu(I), Chr(13), "")
    Next I
    
    改行された文字列を改行で分けて配列にする = Hairetu

End Function
Function コードの先頭空白を除外する(CodeHairetu)
    Dim I%
    Dim KuhakuKosu%
    Dim CodeNagasa%
    CodeNagasa = UBound(CodeHairetu, 1)
    
    Dim TmpItiGyo As String
    
    For I = 1 To CodeNagasa
        TmpItiGyo = CodeHairetu(I)
        TmpItiGyo = Replace(TmpItiGyo, Chr(13), "") '改行を消去
        TmpItiGyo = Replace(TmpItiGyo, Chr(10), "") '改行を消去
        TmpItiGyo = Replace(TmpItiGyo, " ", "") '空白を消去
        If TmpItiGyo <> "" Then
            KuhakuKosu = I - 1
            Exit For
        End If
    Next I
    
    Dim RealCodeNagasa
    RealCodeNagasa = CodeNagasa - KuhakuKosu
    
    Dim Output
    ReDim Output(1 To RealCodeNagasa)
    
    For I = 1 To RealCodeNagasa
        Output(I) = CodeHairetu(I + KuhakuKosu)
    Next I
    
    '出力
    コードの先頭空白を除外する = Output
    
End Function
Function プロシージャ内の使用プロシージャのリスト取得(InputCode, ProcedureNameList, KensakuProcedureNameList, ProcedureOfCode As String)
    '20210428修正
    Dim I%, J%, K% '数え上げ用(Integer型)
    Dim CodeNagasa%
    CodeNagasa = UBound(InputCode, 1)

    Dim ProcedureKosu%
    ProcedureKosu = UBound(ProcedureNameList, 1)
    
    Dim SiyoProcedureList
    
    Dim HikakuCodeItigyo As String, HikakuProcedureName As String
    Dim ProcedureAruNaraTrue As Boolean
    
    Dim KensakuCode '検索用に文字を変換したコード
    KensakuCode = 検索用コード文字列変換(InputCode)
    
    
    K = 0
    For I = 1 To ProcedureKosu
        HikakuProcedureName = KensakuProcedureNameList(I)

        For J = 1 To CodeNagasa
            HikakuCodeItigyo = KensakuCode(J)
            If HikakuProcedureName <> StrConv(ProcedureOfCode, vbLowerCase) Then '20210428修正
                'コード自身のプロシージャは対象としない
                ProcedureAruNaraTrue = コード一行内にプロシージャがあるか検索(HikakuCodeItigyo, HikakuProcedureName)
                
                If ProcedureAruNaraTrue = True Then
                    K = K + 1
                    If K = 1 Then
                        ReDim SiyoProcedureList(1 To K)
                    Else
                        ReDim Preserve SiyoProcedureList(1 To K)
                    End If
                    
                    SiyoProcedureList(K) = ProcedureNameList(I)
                    
                    Exit For 'コード内に見つかったのでそのプロシージャの検索は終了
                End If
            End If
        Next J
    Next I
    
    '出力
    プロシージャ内の使用プロシージャのリスト取得 = SiyoProcedureList

End Function
Function 検索用コード文字列変換(InputCode)
    '検索用にコード文字列を変換する
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim TmpCodeItiretu As String
    
    N = UBound(InputCode, 1)
    
    Dim Output
    Output = InputCode
    
    Dim FirstMoji As String
    
    For I = 1 To N
        TmpCodeItiretu = InputCode(I)
        
         '先頭の空白を除外する。
        FirstMoji = Mid(TmpCodeItiretu, 1, 1)
        Do While FirstMoji = " "
            TmpCodeItiretu = Mid(TmpCodeItiretu, 2)
            FirstMoji = Mid(TmpCodeItiretu, 1, 1)
        Loop
        
        '小文字にする
        TmpCodeItiretu = StrConv(TmpCodeItiretu, vbLowerCase)
        
        ' スペース部分を検索用に置き換える
        TmpCodeItiretu = Replace(TmpCodeItiretu, " ", "@")
        
        ' 改行を消去
        TmpCodeItiretu = Replace(TmpCodeItiretu, Chr(13), "")
        
        Output(I) = TmpCodeItiretu
        
    Next I
    
    検索用コード文字列変換 = Output

End Function
Function コード一行内にプロシージャがあるか検索(CodeItigyo As String, ProcedureName As String) As Boolean

'    Dim KensakuProcedureName
    Dim FirstMoji As String
    Dim ProcedureAruNaraTrue As Boolean
    Dim MojiIti As Integer, MojiLastIti As Integer
    Dim HitotumaeMoji As String, HitotuAtoMoji As String
    
    '検索用にコードをめちゃくちゃ変換する。'※※※※※※※※※※※※※※※※※※※※※※※※※※※
'    CodeItigyo = CodeItigyo
    
'    '先頭の空白を除外する。
'    FirstMoji = Mid(CODEITIGYO, 1, 1)
'    Do While FirstMoji = " "
'        CODEITIGYO = Mid(CODEITIGYO, 2, Len(CODEITIGYO))
'        FirstMoji = Mid(CODEITIGYO, 1, 1)
'    Loop
'
'    '小文字にする
'    CODEITIGYO = StrConv(CODEITIGYO, vbLowerCase)
'
'    ' スペース部分を検索用に置き換える
'    CODEITIGYO = Replace(CODEITIGYO, " ", "@")
'
'    ' 改行を消去
'    CODEITIGYO = Replace(CODEITIGYO, Chr(13), "")
'
    '検索プロシージャの変換'※※※※※※※※※※※※※※※※※※※※※※※※※※※
'    KensakuProcedureName = ProcedureName 'StrConv(ProcedureName, vbLowerCase)
        
        
    '検索'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    ProcedureAruNaraTrue = False '判定初期化
    
    If CodeItigyo = "" Or Mid(CodeItigyo, 1, 1) = "'" Then
        '条件①：構文が空白、または先頭文字が"'"で無効構文・・・ではない。
    ElseIf Mid(CodeItigyo, 1, 3) = "dim" Or _
            Mid(CodeItigyo, 1, 5) = "redim" Then
        '条件②：引数定義・・・ではない。
    ElseIf Mid(CodeItigyo, 1, 9) = "function@" Or _
            Mid(CodeItigyo, 1, 4) = "sub@" Or _
            Mid(CodeItigyo, 1, 8) = "private@" Then
        '条件③：プロシージャの冒頭・・・ではない
    ElseIf Mid(CodeItigyo, 1, 11) = "endfunction" Or _
            Mid(CodeItigyo, 1, 6) = "endsub" Then
        '条件④：プロシージャの最後・・・ではない
        
    ElseIf Len(CodeItigyo) < Len(ProcedureName) Then
        '条件⑤：コードの長さがプロシージャの長さ以下

    Else
    
        MojiIti = InStr(CodeItigyo, ProcedureName)
        MojiLastIti = MojiIti + Len(ProcedureName) - 1
       
        If MojiIti <> 0 Then
            '条件⑤：検索プロシージャの名前の文字列が、コード内に存在する。
            
            HitotumaeMoji = "" '初期化
            HitotuAtoMoji = "" '初期化
            
            If MojiIti > 1 Then
                '検索プロシージャがある文字列の1つ前の文字
                HitotumaeMoji = Mid(CodeItigyo, MojiIti - 1, 1)
            End If
            
            If MojiLastIti < Len(CodeItigyo) Then
                '検索プロシージャがある文字列の1つ後の文字
                HitotuAtoMoji = Mid(CodeItigyo, MojiLastIti + 1, 1)
            End If
            
            If HitotumaeMoji = "" Then
                '前後に文字がない・・・Callの付いていないプロシージャ(Callは付けようね！)
                If HitotuAtoMoji = "" Or HitotuAtoMoji = "'" Then
                    '後ろに文字無し、もしくはコメント
                    '例：SubSample
                    '例：SubSample'コメント
                    ProcedureAruNaraTrue = True
                ElseIf HitotuAtoMoji = "(" Then
                    '例：SubSample(Input)
                    ProcedureAruNaraTrue = True
                End If
                
            ElseIf HitotumaeMoji = "@" Then
                '1つ前が空白" "
                If HitotuAtoMoji = "" Or HitotuAtoMoji = "'" Then
                    '後ろに文字無し、もしくはコメント
                    '例：Dummy = FunctionSample
                    '例：Call SubSample
                    '例：Dummy = FunctionSample'コメント
                    '例：Call SubSample'コメント
                    ProcedureAruNaraTrue = True
                ElseIf HitotuAtoMoji = "(" Then
                    '例：Dummy = FunctionSample(Input）
                    '例：Call SubSample(Input）
                    ProcedureAruNaraTrue = True
                ElseIf HitotuAtoMoji = "," Or HitotuAtoMoji = ")" Then
                    'プロシージャの引数で使用している。
                    '例：Dummy = FunctionSample1(Input1, FunctionSample2, Input2)
                    '例：Dummy = FunctionSample1(Input, FunctionSample2)
                    ProcedureAruNaraTrue = True
                ElseIf HitotuAtoMoji = "@" Then
                    '前後が空白。「見たこと無いけど」
                    '例：If FunctionSample = Hikaku Then
                    ProcedureAruNaraTrue = True
                End If
                            
            ElseIf HitotumaeMoji = "(" Then
                '一つ前が"("
                If HitotuAtoMoji = ")" Or HitotuAtoMoji = "," Then
                    'プロシージャの引数で先頭
                    '例：Dummy = FunctionSample1(FunctionSample2, Input)
                    ProcedureAruNaraTrue = True
                End If

            End If
        End If
    End If
    
    '出力
    コード一行内にプロシージャがあるか検索 = ProcedureAruNaraTrue

End Function
Function プロシージャの使用先のプロシージャのリスト取得(ProcedureList, SiyoProcedureListList)
    Dim I%, J%, K%, M%, N%, I2% '数え上げ用(Integer型)
    Dim ProcedureKosu As Integer
    Dim TmpProcedureName As String, TmpList
    Dim TmpKakunouList
    Dim Output
    ProcedureKosu = UBound(ProcedureList, 1)
        
    ReDim Output(1 To ProcedureKosu)
    
    For I = 1 To ProcedureKosu
        TmpProcedureName = ProcedureList(I)
                
        TmpKakunouList = Empty '格納する配列を初期化
        K = 0 '数え上げ初期化
        For J = 1 To ProcedureKosu
            If I <> J Then '自身は調べない
                TmpList = SiyoProcedureListList(J)
            
                If IsArray(TmpList) Then
                    For I2 = 1 To UBound(TmpList, 1)
                        If TmpList(I2) = TmpProcedureName Then
                            K = K + 1
                            If K = 1 Then
                                ReDim TmpKakunouList(1 To K)
                            Else
                                ReDim Preserve TmpKakunouList(1 To K)
                            End If
                            
                            TmpKakunouList(K) = ProcedureList(J)
                            
                            Exit For
                            
                        End If
                    Next I2
                End If
            End If
        Next J
        Output(I) = TmpKakunouList
    Next I
    
    プロシージャの使用先のプロシージャのリスト取得 = Output
        
End Function
Sub 初期化コピー()
    
    Dim ModuleKosu As Integer
    Dim I As Integer
    Dim ProcedureKosu As Integer
    Dim ProcedureKosuAdd As Integer
    Dim strDummy As String
    Dim strDummy1 As String
    Dim IntDummy1 As Integer
    Dim IntDummy2 As Integer
    Dim HairetuDummy As Variant
    
    Dim ProcedureNameItiran
    
    'サイズ調整
    Application.WindowState = xlMaximized
    Dim TakasaWariai
    
    'TakasaWariai = 0.9
    'Me.Zoom = Me.Zoom * ((Application.Width * TakasaWariai) / Me.Width)
    'Me.Height = (Application.Height * TakasaWariai)
    'Me.Width = (Application.Width * TakasaWariai)
    'Stop
    
    
    'アドインと本ブックのVBE番号取得
    Dim strFileName As String
    strFileName = ActiveWorkbook.FullName
    Dim VBEFileName As Variant
    Dim VBEFileName2 As Variant
    Dim VBECount As Integer
    
    VBECount = ActiveWorkbook.VBProject.VBE.VBProjects.Count
    ReDim VBEFileName(1 To VBECount)
    ReDim VBEFileName2(1 To VBECount)
    For I = 1 To VBECount
        VBEFileName(I) = ActiveWorkbook.VBProject.VBE.VBProjects.Item(I).FileName
        VBEFileName2(I) = F_FileName2(VBEFileName(I))
    Next I
    
    Dim ThisNum As Integer '本ブックのVBE番号
    Dim AddinNum As Integer 'アドインのVBE番号
    For I = 1 To VBECount
        If VBEFileName2(I) = F_FileName2(ActiveWorkbook.FullName) Then
            ThisNum = I
        ElseIf VBEFileName2(I) = "FukamiAddIns2" Then '←←←←←←←←←←←←←←←←←←←←
            AddinNum = I
        End If
    Next I
        
    BookFileName = ActiveWorkbook.Name
    AddinFileName = "FukamiAddIns2.xla" '←←←←←←←←←←←←←←←←←←←←
    
    Dim VBNum As Integer
    VBNum = AddinNum '←←←←←←←←←←←←←←←←←←←←説明を表示するVBProjectの番号
    
    '情報取得
    ModuleItiran = F_モジュール一覧取得(VBNum)
    ModuleKosu = UBound(ModuleItiran, 1)
    ProcedureItiranAdd = F_プロシージャ説明位置取得(AddinNum) '←←←←←←←←←←←←←←←←←←←←
    ProcedureItiran = F_プロシージャ説明位置取得(VBNum) '←←←←←←←←←←←←←←←←←←←←
    ProcedureKosu = UBound(ProcedureItiran, 1)
    ProcedureKosuAdd = UBound(ProcedureItiranAdd, 1)
    ReDim ProcedureSetumeiItiran(1 To ProcedureKosu)
    ReDim ProcedureSetumeiItiranAdd(1 To ProcedureKosuAdd)
    ReDim ProcedureKobunItiran(1 To ProcedureKosu)
    ReDim ProcedureShiyoProcedureItiran(1 To ProcedureKosu)
    ReDim ProcedureShiyoSakiItiran(1 To ProcedureKosu)
    ReDim ProcedureNameItiran(1 To ProcedureKosu)
    ReDim ProcedureNameItiranAdd(1 To ProcedureKosuAdd) 'アドインのプロシージャ名
    
    For I = 1 To ProcedureKosu
        ProcedureNameItiran(I) = ProcedureItiran(I, 4)
    Next I
    For I = 1 To ProcedureKosuAdd '20180111修正
        ProcedureNameItiranAdd(I) = ProcedureItiranAdd(I, 4)
    Next I
    
    'Stop
    
    For I = 1 To ProcedureKosu
        strDummy1 = ProcedureItiran(I, 2)
        IntDummy1 = ProcedureItiran(I, 7)
        IntDummy2 = ProcedureItiran(I, 8)
        ProcedureSetumeiItiran(I) = _
            F_プロシージャ構文取得(strDummy1, _
            IntDummy1, IntDummy2, VBNum)
    
        strDummy1 = ProcedureNameItiran(I) 'アドインのプロシージャ名
        HairetuDummy = ProcedureSetumeiItiran(I) '検索対象プロシージャの構文
        ProcedureShiyoProcedureItiran(I) = _
            F_プロシージャ内使用自作プロシージャ取得(ProcedureNameItiranAdd, _
            HairetuDummy, strDummy1)
            
    Next I
    
    
    For I = 1 To ProcedureKosuAdd
        strDummy1 = ProcedureItiranAdd(I, 2)
        IntDummy1 = ProcedureItiranAdd(I, 7)
        IntDummy2 = ProcedureItiranAdd(I, 8)
        ProcedureSetumeiItiranAdd(I) = _
            F_プロシージャ構文取得(strDummy1, _
            IntDummy1, IntDummy2, AddinNum)
    Next I
    
    
    For I = 1 To ProcedureKosu
        strDummy1 = ProcedureNameItiran(I)
        ProcedureShiyoSakiItiran(I) = _
            F_プロシージャ使用先プロシージャ取得(ProcedureShiyoProcedureItiran, _
            ProcedureNameItiran, strDummy1)
       
    Next I
    
    
    Dim N As Integer
    N = UBound(ModuleItiran, 1)
    
    With ModuleListBox
    
        For I = 1 To N
            .AddItem ModuleItiran(I)
        Next I
        
    End With
    
    'キャプションにバージョン表示'20180111追加
    Dim Version As String
    Version = F_Version()
'    Me.Caption = "アドイン階層構造" & " " & Version

End Sub
Function 多重配列を一列にまとめる(TajuHairetu)
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    N = UBound(TajuHairetu, 1)
    
    Dim TmpHairetu
    Dim Output
    K = 0
    ReDim Output(1 To 1)
    
    For I = 1 To N
        TmpHairetu = TajuHairetu(I)
        M = UBound(TmpHairetu, 1)
        
        For J = 1 To M
            K = K + 1
            ReDim Preserve Output(1 To K)
            Output(K) = TmpHairetu(J)
        Next J
    Next I
    
    '出力
    多重配列を一列にまとめる = Output

End Function
Function 指定プロシージャのコード取得(ProcedureName As String, AllInfoList)
    Dim I% '数え上げ用(Integer型)
    Dim Output
    
    For I = 1 To UBound(AllInfoList, 1)
        If ProcedureName = AllInfoList(I, 3) Then
            Output = AllInfoList(I, 4)
            Exit For
        End If
    Next I
    
    指定プロシージャのコード取得 = Output
    
End Function
Function 指定プロシージャの使用先取得(ProcedureName, ProcedureNameList, SiyosakiProcedureList)
    Dim I, K, J, M, N
    Dim Output
    
    For I = 1 To UBound(ProcedureList, 1)
        If ProcedureName = SiyosakiProcedureList(I) Then
            Stop
'            Output = PbProcedureCodeList(i)
            Exit For
        End If
    Next I
    
    指定プロシージャの使用先取得 = Output
    
End Function
Function 全情報をひとまとめにする(VBProjectFileNameList, ProcedureList, SiyosakiProcedureList)

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim ProcedureKosu As Integer
    Dim Output
'    Output(*,1)：VBProject名前
'    Output(*,2)：モジュール名
'    Output(*,3)：プロシージャ名
'    Output(*,4)：プロシージャのコード
'    Output(*,5)：プロシージャの使用先リスト
    
    Dim TmpKakunoHairetu
    ReDim TmpKakunoHairetu(1 To 50000, 1 To 5)
    
    Dim VBProjectKosu As Byte
    VBProjectKosu = UBound(VBProjectFileNameList, 1)
    
    Dim TmpVBProjectFileName As String
    Dim TmpProcedureInfo, TmpSiyosakiProcedureList
    
    K = 0
    For I = 1 To VBProjectKosu
        TmpVBProjectFileName = VBProjectFileNameList(I)
        TmpProcedureInfo = ProcedureList(I)
        TmpSiyosakiProcedureList = SiyosakiProcedureList(I)
    
            
        For J = 1 To UBound(TmpProcedureInfo, 1)
            K = K + 1
            TmpKakunoHairetu(K, 1) = TmpVBProjectFileName
            TmpKakunoHairetu(K, 2) = TmpProcedureInfo(J, 1)
            TmpKakunoHairetu(K, 3) = TmpProcedureInfo(J, 2)
            TmpKakunoHairetu(K, 4) = TmpProcedureInfo(J, 3)
            TmpKakunoHairetu(K, 5) = TmpSiyosakiProcedureList(J)
        Next J
    Next I
    
    '必要部分抜き出し
    ProcedureKosu = K
    ReDim Output(1 To ProcedureKosu, 1 To 5)
    
    For I = 1 To ProcedureKosu
        For J = 1 To UBound(Output, 2)
            Output(I, J) = TmpKakunoHairetu(I, J)
        Next J
    Next I
    
    '出力
    全情報をひとまとめにする = Output
    

End Function
Function プロシージャの階層構造取得(ProcedureName As String, SiyoSakiItiran, ProcedureItiran)

    Dim ProcedureBango As Integer
    
    ProcedureBango = プロシージャの番号取得(ProcedureName, ProcedureItiran)
    
    Dim I%, K%, J%, M%, N%
    Dim I1%, I2%, I3%, I4%, I5%
    Dim TmpBango%, TmpName As String
    Dim SiyosakiList1
    Dim SiyosakiList2
    Dim SiyosakiList3
    Dim SiyosakiList4
    Dim SiyosakiList5
    Dim SiyosakiList6
    SiyosakiList1 = SiyoSakiItiran(ProcedureBango)
        
    K = 1
    Dim Output
    ReDim Output(1 To 1)
    Output(1) = ProcedureName
    
    If IsEmpty(SiyosakiList1) Then
        '何もしない
    Else
        '第1階層
        For I1 = 1 To UBound(SiyosakiList1, 1)
            TmpName = SiyosakiList1(I1)
            K = K + 1
            ReDim Preserve Output(1 To K)
            Output(K) = "┗" & TmpName
            
            TmpBango = プロシージャの番号取得(TmpName, ProcedureItiran)
            If IsEmpty(TmpBango) Then
                SiyosakiList2 = Empty
            Else
                SiyosakiList2 = SiyoSakiItiran(TmpBango)
            End If
            
            If IsEmpty(SiyosakiList2) Then
                '何もしない
            Else
                '第2階層
                For I2 = 1 To UBound(SiyosakiList2, 1)
                    TmpName = SiyosakiList2(I2)
                    K = K + 1
                    ReDim Preserve Output(1 To K)
                    Output(K) = "　┗" & TmpName
                    TmpBango = プロシージャの番号取得(TmpName, ProcedureItiran)
                    If IsEmpty(TmpBango) Then
                        SiyosakiList3 = Empty
                    Else
                        SiyosakiList3 = SiyoSakiItiran(TmpBango)
                    End If
                    
                    If IsEmpty(SiyosakiList3) Then
                        '何もしない
                    Else
                        '第3階層
                        For I3 = 1 To UBound(SiyosakiList3, 1)
                            TmpName = SiyosakiList3(I3)
                            K = K + 1
                            ReDim Preserve Output(1 To K)
                            Output(K) = "　　┗" & TmpName
                            TmpBango = プロシージャの番号取得(TmpName, ProcedureItiran)
                            If IsEmpty(TmpBango) Then
                                SiyosakiList4 = Empty
                            Else
                                SiyosakiList4 = SiyoSakiItiran(TmpBango)
                            End If
                            
                            If IsEmpty(SiyosakiList4) Then
                                '何もしない
                            Else
                                '第4階層
                                For I4 = 1 To UBound(SiyosakiList4, 1)
                                    TmpName = SiyosakiList4(I4)
                                    K = K + 1
                                    ReDim Preserve Output(1 To K)
                                    Output(K) = "　　　┗" & TmpName
                                    TmpBango = プロシージャの番号取得(TmpName, ProcedureItiran)
                                    If IsEmpty(TmpBango) Then
                                        SiyosakiList5 = Empty
                                    Else
                                        SiyosakiList5 = SiyoSakiItiran(TmpBango)
                                    End If
                                    If IsEmpty(SiyosakiList5) Then
                                        '何もしない
                                    Else
                                        '第5階層
                                        For I5 = 1 To UBound(SiyosakiList5, 1)
                                            TmpName = SiyosakiList5(I5)
                                            K = K + 1
                                            ReDim Preserve Output(1 To K)
                                            Output(K) = "　　　　┗" & TmpName
'                                            SiyosakiList5 = SIYOSAKIITIRAN(TmpBango)
                                            
'                                            If SiyosakiList5(1) = "" Then
'                                                Exit For
'                                            Else

'                                            ここまで
                                                
                                        Next I5
                                    End If
                                Next I4
                            End If
                        Next I3
                    End If
                Next I2
            End If
        Next I1
    End If
    
    プロシージャの階層構造取得 = Output
    
    
End Function
Function プロシージャの番号取得(ProcedureName As String, ProcedureNameList) As Integer
    Dim I% '数え上げ用(Integer型)
    Dim Output
    For I = 1 To UBound(ProcedureNameList, 1)
        If ProcedureName = ProcedureNameList(I) Then
            Output = I
            Exit For
        End If
    Next I
    
    プロシージャの番号取得 = Output
    
End Function
Function 文字区切り(Mojiretu, KugiriMoji As String, OutputMojiretuBango As Byte) As String
    '文字列を指定文字で区切ったときの、指定番号の文字列を返す。
    
    Dim KugiriHairetu
    KugiriHairetu = Split(Mojiretu, KugiriMoji)
    KugiriHairetu = Application.Transpose(KugiriHairetu)
    KugiriHairetu = Application.Transpose(KugiriHairetu)
    
    Dim Output As String
    Output = KugiriHairetu(OutputMojiretuBango)
    
    文字区切り = Output
    
End Function
Function 階層リストの付属関係計算(KaisoList)
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    N = UBound(KaisoList, 1)
    
    Dim KaisoNumList
    ReDim KaisoNumList(1 To N)
    
    Dim TmpOkikaeMojiretu As String
    Dim TmpList
    Dim KakuteiNaraTrue As Boolean
    
    For I = 1 To N
        TmpList = KaisoList(I)
        KakuteiNaraTrue = False
        K = -1
        TmpOkikaeMojiretu = "┗"
        Do While KakuteiNaraTrue = False
            K = K + 1
            If K = 0 Then
                If Mid(TmpList, 1, 1) <> "┗" And Mid(TmpList, 1, 1) <> "　" Then
                    KakuteiNaraTrue = True
                    Exit Do
                End If
            ElseIf Mid(TmpList, 1, K) = TmpOkikaeMojiretu Then
                KakuteiNaraTrue = True
                Exit Do
            Else
                TmpOkikaeMojiretu = "　" & TmpOkikaeMojiretu
            End If
            
        Loop
        
        KaisoNumList(I) = K + 1
    Next I
    
    Dim HuzokuKosuList
    ReDim HuzokuKosuList(1 To N)
    Dim TmpHuzokuKosu As Integer, TmpKaisoNum As Integer
    
    
    For I = 1 To N
        TmpKaisoNum = KaisoNumList(I)
        TmpHuzokuKosu = 0
        
        For J = I + 1 To N
            If KaisoNumList(J) = TmpKaisoNum + 1 Then
                TmpHuzokuKosu = TmpHuzokuKosu + 1
            ElseIf KaisoNumList(J) <= TmpKaisoNum - 1 Then
                Exit For
            End If
        Next J
        
        HuzokuKosuList(I) = TmpHuzokuKosu
    Next I
    
    '出力
    Dim Output
    ReDim Output(1 To 2)
    Output(1) = KaisoNumList
    Output(2) = HuzokuKosuList
    
    階層リストの付属関係計算 = Output
    
End Function
Function 階層リストを指定階層までのリスト取得(KaisoList, ByVal SiteiKaisoNum)
    Dim KaisoNumList
    Dim HuzokuKosuList
    Dim Dummy1
    
    Dummy1 = 階層リストの付属関係計算(KaisoList)
    KaisoNumList = Dummy1(1)
    HuzokuKosuList = Dummy1(2)
    
    Dim Output
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    N = UBound(KaisoList, 1)
    
    Dim MaxKaisoNum%
    MaxKaisoNum = WorksheetFunction.Max(KaisoNumList)
    
    If SiteiKaisoNum = 0 Then 'すべて表示するに設定する。
        SiteiKaisoNum = MaxKaisoNum
    End If
    
    K = 0
    ReDim Output(1 To 1)
    For I = 1 To N
        If KaisoNumList(I) <= SiteiKaisoNum Then
            K = K + 1
            ReDim Preserve Output(1 To K)
            Output(K) = KaisoList(I) & "(" & HuzokuKosuList(I) & ")"
        End If
    Next I
    
    階層リストを指定階層までのリスト取得 = Output
    
End Function
