VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmKaiso 
   Caption         =   "階層化表示フォーム"
   ClientHeight    =   9432
   ClientLeft      =   36
   ClientTop       =   408
   ClientWidth     =   15480
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
'起動中VBProjectの全情報
Public PbUnLockVBProjectList
Public PbUnLockVBProjectFileNameList
Public PbModuleList
Public PbProcedureList
Public PbProcedureNameList
Public PbProcedureCodeList
Public PbShiyoProcedureListList
Public PbShiyoSakiProcedureListList
Public PbKaisoList

Public PbAllProcedureNameList
Public PbAllInfoList

'VBProjectListBoxで選択したVBProjectの情報のみを格納
Public PbOutputModuleList
Public PbOutputProcedureList
Public PbOutputProcedureNameList
Public PbOutputProcedureCodeList
Public PbOutputShiyoProcedureListList
Public PbOutputShiyoSakiProcedureListList
Public PbOutputKaisoList

Public PbTmpKaisoList

'Private Sub CodeListBox_Click()
'
'    Dim TmpCodeItigyo
'    TmpCodeItigyo = CodeListBox.List(CodeListBox.ListIndex)
'
'    '選択リスト全体表示リストボックスに選択した一行を表示
'    TotemoNagaiListBox.Clear
'    TotemoNagaiListBox.AddItem TmpCodeItigyo
'
'End Sub

Private Sub KaisoHyouji_Change()
    Dim ListNo As Integer
    ListNo = KaisoHyouji.ListIndex
'    Stop
    
    KaisoListBox.List = 階層リストを指定階層までのリスト取得(PbTmpKaisoList, ListNo)
    
End Sub
Sub コード検索実行()
    Dim KensakuWord As String
    KensakuWord = txtKensaku.Value
    KensakuWord = StrConv(KensakuWord, vbLowerCase)
    
    If KensakuWord = "" Then
        Exit Sub
    End If
    
    Dim KensakuWordList
    KensakuWordList = Split(KensakuWord, " ")
    KensakuWordList = Application.Transpose(KensakuWordList)
    KensakuWordList = Application.Transpose(KensakuWordList)
    
    Dim I%, J%, I2%, N%, K%
    Dim KensakuKosu%
    KensakuKosu = UBound(KensakuWordList, 1)
    
    Dim TmpCodeListList
    Dim TmpCodeList
    TmpCodeListList = PbProcedureCodeList
    Dim TmpProcedureListList
    Dim TmpProcedureList
    TmpProcedureListList = PbProcedureList
    Dim TmpSiyosakiListList
    Dim TmpSiyosakiList
    TmpSiyosakiListList = PbShiyoSakiProcedureListList
    Dim TmpSiyoListList
    Dim TmpSiyoList
    TmpSiyoListList = PbShiyoProcedureListList
    Dim TmpKaisoListList
    Dim TmpKaisoList
    TmpKaisoListList = PbKaisoList
    
    Dim TmpCode
    Dim TmpItiGyo
    Dim TmpKensakuWord As String
    Dim TmpKensakuHanteiList
    
    Dim NaiNaraTrue As Boolean
    Dim AruNaraTrue As Boolean
    
    Dim KensakuProcedureList, KensakuCodeList, KensakuSiyosakiList, KensakuSiyoList, KensakuKaisoList
    K = 0
    ReDim KensakuProcedureList(1 To 1)
    ReDim KensakuCodeList(1 To 1)
    ReDim KensakuSiyosakiList(1 To 1)
    ReDim KensakuSiyoList(1 To 1)
    ReDim KensakuKaisoList(1 To 1)
    
    For I = 1 To UBound(TmpCodeListList, 1)
        TmpCodeList = TmpCodeListList(I)
        TmpSiyosakiList = TmpSiyosakiListList(I)
        TmpSiyoList = TmpSiyoListList(I)
        TmpKaisoList = TmpKaisoListList(I)
        
        For I2 = 1 To UBound(TmpCodeList, 1)
            TmpCode = TmpCodeList(I2)
            ReDim TmpKensakuHanteiList(1 To KensakuKosu)
            For J = 1 To KensakuKosu
                TmpKensakuWord = KensakuWordList(J)
                
                For Each TmpItiGyo In TmpCode
                    If InStr(1, TmpItiGyo, TmpKensakuWord) <> 0 Then
                        TmpKensakuHanteiList(J) = 1
                        Exit For
                    End If
                Next
            Next J
            
            If WorksheetFunction.Sum(TmpKensakuHanteiList) = KensakuKosu Then
                K = K + 1
                ReDim Preserve KensakuProcedureList(1 To K)
                ReDim Preserve KensakuCodeList(1 To K)
                ReDim Preserve KensakuSiyosakiList(1 To K)
                ReDim Preserve KensakuSiyoList(1 To K)
                ReDim Preserve KensakuKaisoList(1 To K)
                TmpProcedureList = TmpProcedureListList(I)
                KensakuProcedureList(K) = TmpProcedureList(I2, 2)
                KensakuCodeList(K) = TmpCode
                KensakuSiyosakiList(K) = TmpSiyosakiList(I2)
                KensakuSiyoList(K) = TmpSiyoList(I2)
                KensakuKaisoList(K) = TmpKaisoList(I2)
                
            End If
        Next I2
    Next I
    
    PbOutputProcedureNameList = KensakuProcedureList
    PbOutputProcedureCodeList = KensakuCodeList
    PbOutputShiyoProcedureListList = KensakuSiyoList
    PbOutputShiyoSakiProcedureListList = KensakuSiyosakiList
    PbOutputKaisoList = KensakuKaisoList
    
    ProcedureListBox.List = KensakuProcedureList
    
    
End Sub
Private Sub tgl検索_Click()
    Call コード検索実行
    tgl検索.Value = False
End Sub
Private Sub tgl検索_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call コード検索実行
    tgl検索.Value = False
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'サイズ調整
    Application.WindowState = xlMaximized
    Dim TakasaWariai As Double
    TakasaWariai = 0.9
    
    Me.Zoom = 100 * (Application.Height) / Me.Height
    Me.Height = (Application.Height * TakasaWariai)
    Me.Width = (Application.Width * TakasaWariai)
    Me.Zoom = 100 * (Application.Height) / Me.Height
    Me.Height = (Application.Height * TakasaWariai)
    Me.Width = (Application.Width * TakasaWariai)
    Me.Top = 20
    Me.Left = 20

End Sub
Private Sub UserForm_Initialize()

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Dummy1, Dummy2
    Dim TmpFileName As String
    Dim TmpVBProject As Object, VBProjectCount As Byte
    Dim TmpModuleList, TmpProcedureList, TmpProcedureNameList, TmpProcedureCodeList
    Dim TmpProcedureKosu As Integer
    Dim TmpCode, TmpProcedureName As String
    Dim TmpSiyoProcedureList
    Dim TmpSiyoProcedureListList, TmpSiyosakiProcedureListList
    Dim AllSiyoProcedureListList
    Dim TmpKaiso, TmpKaisoList
    
    '起動中のVBProjectのリスト取得'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    PbUnLockVBProjectList = 非ロックのVBProjectリスト取得
    VBProjectCount = UBound(PbUnLockVBProjectList, 1)
    
    'VBProjectのファイル名のリストも作成しておく
    ReDim PbUnLockVBProjectFileNameList(1 To VBProjectCount)
    For I = 1 To VBProjectCount
        Set Dummy1 = PbUnLockVBProjectList(I)
        PbUnLockVBProjectFileNameList(I) = Dir(Dummy1.FileName) 'ファイル名抜出
    Next I

    '各モジュール、プロシージャの名前、コードを取得'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    ReDim PbModuleList(1 To VBProjectCount)
    ReDim PbProcedureList(1 To VBProjectCount)
    ReDim PbProcedureNameList(1 To VBProjectCount)
    ReDim PbProcedureCodeList(1 To VBProjectCount)
    
    For I = 1 To VBProjectCount
        Set TmpVBProject = PbUnLockVBProjectList(I)
        TmpProcedureList = プロシージャ一覧取得(TmpVBProject) '1列目モジュール、2列目プロシージャ名、3列目プロシージャコード
        TmpModuleList = モジュール一覧取得(TmpVBProject, TmpProcedureList) '1列目モジュール(オブジェクト形式)、2列目モジュール内のプロシージャリスト
        
        TmpProcedureKosu = UBound(TmpProcedureList, 1)
        
        ReDim TmpProcedureNameList(1 To TmpProcedureKosu)
        ReDim TmpProcedureCodeList(1 To TmpProcedureKosu)
           
        For J = 1 To TmpProcedureKosu
            TmpProcedureNameList(J) = TmpProcedureList(J, 2)
            TmpProcedureCodeList(J) = TmpProcedureList(J, 3)
        Next J
        
        'パブリック引数に格納
        PbModuleList(I) = TmpModuleList
        PbProcedureList(I) = TmpProcedureList
        PbProcedureNameList(I) = TmpProcedureNameList
        PbProcedureCodeList(I) = TmpProcedureCodeList
        
    Next I
        
    '全VBProjectのプロシージャ名一覧を作成
    PbAllProcedureNameList = 多重配列を一列にまとめる(PbProcedureNameList)
    
    
    Dim KensakuProcedureNameList
    ReDim KensakuProcedureNameList(1 To UBound(PbAllProcedureNameList, 1))
    For I = 1 To UBound(PbAllProcedureNameList, 1)
        KensakuProcedureNameList(I) = StrConv(PbAllProcedureNameList(I), vbLowerCase)
    Next I
    
    '各プロシージャの使用関係をコードを読み取って取得'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    ReDim PbShiyoProcedureListList(1 To VBProjectCount)
    ReDim PbShiyoSakiProcedureListList(1 To VBProjectCount)
    
    '使用プロシージャ取得
    For I = 1 To VBProjectCount
        TmpProcedureNameList = PbProcedureNameList(I)
        TmpProcedureCodeList = PbProcedureCodeList(I)
        TmpProcedureKosu = UBound(TmpProcedureNameList, 1)
        
        ReDim TmpSiyoProcedureListList(1 To TmpProcedureKosu)
        
        For J = 1 To TmpProcedureKosu
            TmpCode = TmpProcedureCodeList(J)
            TmpProcedureName = TmpProcedureNameList(J)
            TmpSiyoProcedureList = プロシージャ内の使用プロシージャのリスト取得(TmpCode, PbAllProcedureNameList, KensakuProcedureNameList, TmpProcedureName)
            TmpSiyoProcedureListList(J) = TmpSiyoProcedureList '使用先のリストをリストに格納
        Next J
        
        PbShiyoProcedureListList(I) = TmpSiyoProcedureListList
    
    Next I
    
    '全VBProjectの使用プロシージャリストを一列にまとめる。
    AllSiyoProcedureListList = 多重配列を一列にまとめる(PbShiyoProcedureListList)
    
    
    '使用先プロシージャ取得
    For I = 1 To VBProjectCount
        TmpProcedureNameList = PbProcedureNameList(I)
        TmpProcedureKosu = UBound(TmpProcedureNameList, 1)
        
        TmpSiyosakiProcedureListList = プロシージャの使用先のプロシージャのリスト取得(TmpProcedureNameList, AllSiyoProcedureListList)
        
        PbShiyoSakiProcedureListList(I) = TmpSiyosakiProcedureListList
    
    Next I

    '全情報をひとまとめにしておく
    PbAllInfoList = 全情報をひとまとめにする(PbUnLockVBProjectFileNameList, PbProcedureList, PbShiyoSakiProcedureListList)

    'プロシージャの階層リストを取得する
    ReDim PbKaisoList(1 To VBProjectCount)
    
    For I = 1 To VBProjectCount
        PbOutputShiyoProcedureListList = PbShiyoProcedureListList(I)
        PbOutputProcedureNameList = PbProcedureNameList(I)
        TmpProcedureKosu = UBound(PbOutputProcedureNameList, 1)
        
        ReDim TmpKaisoList(1 To TmpProcedureKosu)
        
        For J = 1 To TmpProcedureKosu
            TmpProcedureName = PbOutputProcedureNameList(J)
            TmpKaiso = プロシージャの階層構造取得(TmpProcedureName, AllSiyoProcedureListList, _
                                                    PbAllProcedureNameList)
            TmpKaisoList(J) = TmpKaiso
        Next J
        
        PbKaisoList(I) = TmpKaisoList
        
    Next I

    'ListBox「起動中VBProjecct」に起動中のVBProjectのリスト出力'※※※※※※※※※※※※※※※※※※※※※※※※※※※
    For I = 1 To UBound(PbUnLockVBProjectList, 1)
        VBProjectListBox.AddItem PbUnLockVBProjectFileNameList(I)
    Next I
    
    '階層表示のコンボボックスにアイテム追加
    With KaisoHyouji
        .AddItem "全部表示"
        .AddItem "第1階層まで"
        .AddItem "第2階層まで"
        .AddItem "第3階層まで"
        .AddItem "第4階層まで"
        .AddItem "第5階層まで"
    End With

    tgl検索.Value = False
    
End Sub
Private Sub VBProjectListBox_Click()

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim TmpModuleName As String
    Dim TmpProcedureList
    Dim TmpProcedureKosu As Integer
    
    ModuleListBox.Clear
    ProcedureListBox.Clear
    KaisoListBox.Clear
    ProcedureListBox.Clear
    SiyosakiListBox.Clear
    txtListBox.Text = ""
    
      
    For I = 1 To UBound(PbUnLockVBProjectFileNameList, 1)
        Select Case VBProjectListBox.List(VBProjectListBox.ListIndex)
            Case PbUnLockVBProjectFileNameList(I)
                
                '選択したVBProjectの情報を格納
                PbOutputModuleList = PbModuleList(I)
                PbOutputProcedureList = PbProcedureList(I)
                PbOutputProcedureNameList = PbProcedureNameList(I)
                PbOutputProcedureCodeList = PbProcedureCodeList(I)
                PbOutputShiyoProcedureListList = PbShiyoProcedureListList(I)
                PbOutputShiyoSakiProcedureListList = PbShiyoSakiProcedureListList(I)
                PbOutputKaisoList = PbKaisoList(I)
                                
                For J = 1 To UBound(PbOutputModuleList, 1)
                    TmpModuleName = PbOutputModuleList(J, 1).Name
                    TmpProcedureList = PbOutputModuleList(J, 2)
                    
                    If IsEmpty(TmpProcedureList) Then
                        TmpProcedureKosu = 0
                    Else
                        TmpProcedureKosu = UBound(TmpProcedureList, 1)
                    End If
                    
                    ModuleListBox.AddItem TmpModuleName & "(" & TmpProcedureKosu & ")" 'モジュール内のプロシージャの個数を後ろにつける
                
                Next J
                
                Exit For
                
        End Select
    Next I
    
End Sub
Private Sub ModuleListBox_Click()

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim TmpModuleBango As Integer
    Dim TmpModuleName As String
    Dim TmpProcedureList
    Dim TmpProcedureName As String
    Dim TmpKaisoList
    Dim TmpKaisoKosu As Integer
    Dim TmpListAddName As String
    
'    ModuleListBox.Clear
    ProcedureListBox.Clear
    KaisoListBox.Clear
    ProcedureListBox.Clear
    SiyosakiListBox.Clear
    txtListBox.Text = ""
        
    TmpModuleBango = ModuleListBox.ListIndex + 1
    TmpModuleName = PbOutputModuleList(TmpModuleBango, 1).Name
    TmpProcedureList = PbOutputModuleList(TmpModuleBango, 2)
    
    On Error GoTo ErrorEscape
    
    If IsEmpty(TmpProcedureList) = False Then
    
        For I = 1 To UBound(TmpProcedureList, 1)
            TmpProcedureName = TmpProcedureList(I)
            
            '各プロシージャでコード内で使用しているプロシージャの数を取得する。
            For J = 1 To UBound(PbOutputProcedureList, 1)
                If TmpProcedureName = PbOutputProcedureList(J, 2) Then
                    TmpKaisoList = PbOutputKaisoList(J)
                    TmpKaisoKosu = UBound(TmpKaisoList, 1) - 1
                    TmpListAddName = TmpProcedureName & "(" & TmpKaisoKosu & ")" '使用プロシージャの個数を後ろにつける
                    Exit For
                End If
            Next J
            
            ProcedureListBox.AddItem TmpListAddName

        Next I
    End If
 
ErrorEscape:

End Sub
Private Sub ProcedureListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'リストボックスで選択ダブルクリックしたらVBE起動してコードを表示する。
        
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    
    Dim TmpProcedureName As String

    TmpProcedureName = ProcedureListBox.List(ProcedureListBox.ListIndex)
    TmpProcedureName = 文字区切り(TmpProcedureName, "(", 1)
    
'    On Error GoTo ErrorEscape
    On Error Resume Next
    Application.Goto Reference:=TmpProcedureName
    On Error GoTo 0
'ErrorEscape:

End Sub
Private Sub ProcedureListBox_Click()
    
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    
    Dim TmpProcedureName As String
    Dim TmpSiyoProcedureList
    Dim TmpSiyosakiProcedureList
    Dim TmpKaisoList
    Dim TmpCode
    Dim ProcedureCode As String
    
    TmpProcedureName = ProcedureListBox.List(ProcedureListBox.ListIndex)
    TmpProcedureName = 文字区切り(TmpProcedureName, "(", 1)
    
    '選択リスト全体表示リストボックスに選択した一行を表示
    TotemoNagaiListBox.Clear
    TotemoNagaiListBox.AddItem ProcedureListBox.List(ProcedureListBox.ListIndex)
    
    For I = 1 To UBound(PbOutputProcedureNameList, 1)
        If TmpProcedureName = PbOutputProcedureNameList(I) Then
            
            'TxtLISTBOXにコードの表示
            txtListBox.Text = ""
    
            TmpCode = PbOutputProcedureCodeList(I)
            For J = 1 To UBound(TmpCode, 1)
                ProcedureCode = ProcedureCode & TmpCode(J) & Chr(10)
        '        CodeListBox.AddItem TmpCode(I)
            Next J
            txtListBox.Text = ProcedureCode
    
'            For J = 1 To UBound(TmpCode, 1)
'                CodeListBox.AddItem TmpCode(J)
'            Next J
            
            'KaisoListBoxに階層構造の表示
            KaisoListBox.Clear
            TmpKaisoList = PbOutputKaisoList(I)
                                                        
            For J = 1 To UBound(TmpKaisoList, 1)
                KaisoListBox.AddItem TmpKaisoList(J)
            Next J
            
            'SiyosakiListBoxに使用先プロシージャの表示
            SiyosakiListBox.Clear
            TmpSiyosakiProcedureList = PbOutputShiyoSakiProcedureListList(I)
            If IsEmpty(TmpSiyosakiProcedureList) Then
                '何もしない
            Else
                For J = 1 To UBound(TmpSiyosakiProcedureList, 1)
                    SiyosakiListBox.AddItem TmpSiyosakiProcedureList(J)
                Next J
            End If
            
            PbTmpKaisoList = TmpKaisoList
            
            Exit For
        End If
    Next I

End Sub
Private Sub KaisoListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'リストボックスで選択ダブルクリックしたらVBE起動してコードを表示する。
        
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    
    Dim TmpProcedureName As String
    
    TmpProcedureName = KaisoListBox.List(KaisoListBox.ListIndex)
    TmpProcedureName = 文字区切り(TmpProcedureName, "(", 1)
    TmpProcedureName = Replace(TmpProcedureName, "　", "")
    TmpProcedureName = Replace(TmpProcedureName, "┗", "")
    
    On Error GoTo ErrorEscape
    Application.Goto Reference:=TmpProcedureName

ErrorEscape:
End Sub
Private Sub KaisoListBox_Click()
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim TmpCode
    Dim TmpProcedureName As String
    
    TmpProcedureName = KaisoListBox.List(KaisoListBox.ListIndex)
    
    '選択リスト全体表示リストボックスに選択した一行を表示
    TotemoNagaiListBox.Clear
    TotemoNagaiListBox.AddItem TmpProcedureName
    
    TmpProcedureName = 文字区切り(TmpProcedureName, "(", 1)
    TmpProcedureName = Replace(TmpProcedureName, "　", "")
    TmpProcedureName = Replace(TmpProcedureName, "┗", "")
    
    TmpCode = 指定プロシージャのコード取得(TmpProcedureName, PbAllInfoList)

    If IsEmpty(TmpCode) Then Exit Sub
    
    'コード表示
    txtListBox.Text = ""
    Dim ProcedureCode
    
    For I = 1 To UBound(TmpCode, 1)
        ProcedureCode = ProcedureCode & TmpCode(I) & Chr(10)
'        CodeListBox.AddItem TmpCode(I)
    Next I
    
    txtListBox.Text = ProcedureCode
    

End Sub
Private Sub SiyosakiListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'リストボックスで選択ダブルクリックしたらVBE起動してコードを表示する。
        
    Dim TmpProcedureName As String
    
    TmpProcedureName = SiyosakiListBox.List(SiyosakiListBox.ListIndex)
    
    On Error GoTo ErrorEscape
    Application.Goto Reference:=TmpProcedureName

ErrorEscape:

End Sub
Private Sub SiyosakiListBox_Click()

    Dim I% '数え上げ用(Integer型)
    Dim TmpCode
    Dim TmpProcedureName As String
    
    TmpProcedureName = SiyosakiListBox.List(SiyosakiListBox.ListIndex) '←←←←←←←←←←←←←←←←←←←←
            
    '選択リスト全体表示リストボックスに選択した一行を表示
    TotemoNagaiListBox.Clear
    TotemoNagaiListBox.AddItem TmpProcedureName

    TmpCode = 指定プロシージャのコード取得(TmpProcedureName, PbAllInfoList)
    
    If IsEmpty(TmpCode) Then Exit Sub

    'コード表示
    txtListBox.Text = ""
    Dim ProcedureCode As String
    
    For I = 1 To UBound(TmpCode, 1)
        ProcedureCode = ProcedureCode & TmpCode(I) & Chr(10)
'        CodeListBox.AddItem TmpCode(I)
    Next I
    
    txtListBox.Text = ProcedureCode

End Sub
