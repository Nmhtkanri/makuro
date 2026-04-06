Attribute VB_Name = "c設定"

Option Explicit

' ============================================================
' 経費集計マクロ（設定シート対応版）
' ============================================================
' 【特徴】
' ・キーワードをコード内ではなく「設定」シートで管理
' ・キーワードの追加・変更時にVBAを触る必要なし
' ・処理結果のログを「集計ログ」シートに出力
' ============================================================

' === シート名 ===
Private Const SH_SUM As String = "集計"               ' 出力先
Private Const SH_SRC As String = "経費統合一覧表"     ' 取り込み元
Private Const SH_SETTING As String = "設定"           ' キーワード設定シート
Private Const SH_LOG As String = "集計ログ"           ' 処理ログ出力先

' === 出力列（集計シート）===
Private Const COL_EMP_NO As Long = 1   ' A: 社員番号
Private Const COL_NAME As Long = 2     ' B: 氏名
Private Const COL_TOTAL As Long = 3    ' C: 合計
Private Const COL_GK As Long = 4       ' D: 夜間当番手当
Private Const COL_RINK As Long = 5     ' E: RINK手当
Private Const COL_ALLOW2 As Long = 6   ' F: 手当2（D+E）
Private Const COL_BILL As Long = 7     ' G: 顧客請求分
Private Const COL_TRANS As Long = 8    ' H: 交通費
Private Const COL_NONTAX_TATEKAE As Long = 9  ' I: 非課税精算(立替金) ＜ NEW
Private Const COL_ETC As Long = 10     ' J: その他
Private Const COL_TW As Long = 11      ' K: テレワーク手当
Private Const COL_DATE As Long = 12    ' L: 請求日
' === 分類名（設定シートで使用）===
Private Const CAT_YAKAN As String = "夜間当番手当"
Private Const CAT_RINK As String = "RINK手当"
Private Const CAT_TRANS As String = "交通費"
Private Const CAT_TW As String = "テレワーク手当"
Private Const CAT_TRANS_NG As String = "交通費除外"
Private Const CAT_KOKYAKU_NG As String = "顧客請求除外"

' === キーワード格納用（モジュールレベル変数）===
Private kwYakan As Collection
Private kwRink As Collection
Private kwTrans As Collection
Private kwTW As Collection
Private kwTransNG As Collection
Private kwKokyakuNG As Collection

' ============================================================
' メイン処理
' ============================================================
Public Sub Run_経費集計_設定シート版()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' 1) 設定シートからキーワードを読み込む
    If Not LoadKeywordsFromSetting() Then
        MsgBox "設定シートの読み込みに失敗しました。" & vbCrLf & _
               "「設定」シートが存在するか確認してください。" & vbCrLf & _
               "初回は Setup_設定シート作成 を実行してください。", vbExclamation
        GoTo FinallyExit
    End If
    
    ' 2) ログシートを初期化
    InitLogSheet
    
    ' 3) データ集計
    Dim agg As Object, maxDate As Object, empList As Object, hitCount As Long
    Set agg = CreateObject("Scripting.Dictionary")
    Set maxDate = CreateObject("Scripting.Dictionary")
    Set empList = CreateObject("Scripting.Dictionary")
    
    hitCount = Collect_From_Source(agg, maxDate, empList)
    
    If hitCount = 0 Then
        MsgBox "取り込み件数が0でした。出力シートは変更していません。" & vbCrLf & _
               "・取り込み元シート名: " & SH_SRC & vbCrLf & _
               "・設定シートのキーワードをご確認ください。", vbExclamation
        GoTo FinallyExit
    End If
    
    ' 4) バックアップ作成後、出力
    BackupSheet SH_SUM
    Rewrite_Output agg, maxDate
    
    MsgBox "集計が完了しました。" & vbCrLf & _
           "処理件数: " & hitCount & "件" & vbCrLf & _
           "詳細は「集計ログ」シートをご確認ください。", vbInformation

FinallyExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    MsgBox "エラー: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume FinallyExit
End Sub

' ============================================================
' 設定シートからキーワードを読み込む
' ============================================================
Private Function LoadKeywordsFromSetting() As Boolean
    LoadKeywordsFromSetting = False
    
    If Not SheetExists(SH_SETTING) Then
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = Worksheets(SH_SETTING)
    
    ' Collectionを初期化
    Set kwYakan = New Collection
    Set kwRink = New Collection
    Set kwTrans = New Collection
    Set kwTW = New Collection
    Set kwTransNG = New Collection
    Set kwKokyakuNG = New Collection
    
    ' 設定シートの構成:
    ' A列: 分類名
    ' B列: キーワード
    
    Dim lastR As Long, r As Long
    Dim catName As String, keyword As String
    
    lastR = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lastR  ' 1行目はヘッダー
        catName = Trim$(CStr(ws.Cells(r, 1).value))
        keyword = Trim$(CStr(ws.Cells(r, 2).value))
        
        If keyword <> "" Then
            Select Case catName
                Case CAT_YAKAN
                    kwYakan.Add LCase$(keyword)
                Case CAT_RINK
                    kwRink.Add LCase$(keyword)
                Case CAT_TRANS
                    kwTrans.Add LCase$(keyword)
                Case CAT_TW
                    kwTW.Add LCase$(keyword)
                Case CAT_TRANS_NG
                    kwTransNG.Add LCase$(keyword)
                Case CAT_KOKYAKU_NG
                    kwKokyakuNG.Add LCase$(keyword)
            End Select
        End If
    Next r
    
    LoadKeywordsFromSetting = True
End Function

' ============================================================
' ログシートの初期化
' ============================================================
Private Sub InitLogSheet()
    Dim ws As Worksheet
    
    ' シートがなければ作成
    If Not SheetExists(SH_LOG) Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = SH_LOG
    End If
    
    Set ws = Worksheets(SH_LOG)
    ws.Cells.Clear
    
    ' ヘッダー作成
    ws.Cells(1, 1).value = "行番号"
    ws.Cells(1, 2).value = "社員番号"
    ws.Cells(1, 3).value = "氏名"
    ws.Cells(1, 4).value = "内訳"
    ws.Cells(1, 5).value = "金額"
    ws.Cells(1, 6).value = "判定結果"
    ws.Cells(1, 7).value = "マッチしたキーワード"
    ws.rows(1).Font.Bold = True
End Sub

' ============================================================
' ログに1行追加
' ============================================================
Private Sub AddLog(ByVal srcRow As Long, ByVal empNo As String, ByVal empNm As String, _
                   ByVal desc As String, ByVal amt As Double, _
                   ByVal result As String, ByVal matchedKw As String)
    Dim ws As Worksheet
    Set ws = Worksheets(SH_LOG)
    
    Dim NextR As Long
    NextR = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(NextR, 1).value = srcRow
    ws.Cells(NextR, 2).value = empNo
    ws.Cells(NextR, 3).value = empNm
    ws.Cells(NextR, 4).value = desc
    ws.Cells(NextR, 5).value = amt
    ws.Cells(NextR, 6).value = result
    ws.Cells(NextR, 7).value = matchedKw
End Sub

' ============================================================
' データ収集（メインロジック）
' ============================================================
Private Function Collect_From_Source(ByRef agg As Object, ByRef maxDate As Object, ByRef empList As Object) As Long
    Dim ws As Worksheet
    If Not SheetExists(SH_SRC) Then
        MsgBox "取り込み元シートが見つかりません: " & SH_SRC, vbExclamation
        Collect_From_Source = 0
        Exit Function
    End If
    Set ws = Worksheets(SH_SRC)
    
    ' 列を探す
    Dim cEmpNo As Long, cName As Long, cUch As Long, cTrans As Long
    Dim cAmt As Long, cFareAmt As Long, cBooked As Long, cEStaff As Long
    
    cEmpNo = FindCol(ws, 1, Array("社員番号", "従業員番号", "社員ID"))
    cName = FindCol(ws, 1, Array("氏名", "名前"))
    cUch = FindCol(ws, 1, Array("内訳", "摘要", "内容"))
    cTrans = FindCol(ws, 1, Array("交通機関", "経路", "移動手段"), True)
    cAmt = FindCol(ws, 1, Array("合計", "小計", "金額"))
    cFareAmt = FindCol(ws, 1, Array("金額(交通費)", "交通費金額", "交通費（円）", "交通費(円)"), True)
    cBooked = FindCol(ws, 1, Array("計上日", "計上", "計上日付"), True)
    cEStaff = FindCol(ws, 1, Array("顧客請求費", "顧客対応", "顧客当番", "夜間当番", "24時間準直当番", "深夜出動", "24時間準直当番手当", "糊客請求分"), True)
    If cEStaff = 0 Then cEStaff = 34 ' S列フォールバック
    
    If cEmpNo = 0 Or cName = 0 Or cUch = 0 Or cAmt = 0 Then
        MsgBox "必須列が見つからないため取り込み中止。" & vbCrLf & _
               "社員番号/氏名/内訳/合計（金額）の見出しをご確認ください。", vbExclamation
        Collect_From_Source = 0
        Exit Function
    End If
    
    Dim lastR As Long, r As Long, hits As Long
    lastR = ws.Cells(ws.rows.Count, cEmpNo).End(xlUp).Row
    If lastR < 2 Then
        Collect_From_Source = 0
        Exit Function
    End If
    
    For r = 2 To lastR
        Dim empNo As String, empNm As String, key As String
        Dim desc As String, trans As String
        Dim amt As Double, fa As Double, estAmt As Double
        Dim estFilled As Boolean
        Dim matchedKw As String, resultCat As String
        
        empNo = NormalizeId(ws.Cells(r, cEmpNo).value)
        empNm = NormalizeName(ws.Cells(r, cName).value)
        If empNo = "" And empNm = "" Then GoTo NextR
        
        key = BuildEmpKey(empNo, empNm)
        If Not empList.Exists(key) Then empList.Add key, Array(empNo, empNm)
        
        desc = NormalizeStr(ws.Cells(r, cUch).value)
        trans = IIf(cTrans > 0, NormalizeStr(ws.Cells(r, cTrans).value), "")
        estFilled = (cEStaff > 0 And Trim$(CStr(ws.Cells(r, cEStaff).value)) <> "")
        
        ' --- G列：顧客請求費（除外キーワードチェック）---
        If cEStaff > 0 And empNo <> "" Then
            estAmt = ParseAmount(ws.Cells(r, cEStaff).value)
            If estAmt <> 0 Then
                matchedKw = HitAnyCollection(desc, kwKokyakuNG)
                If matchedKw = "" Then
                    ' 除外キーワードに該当しない → G列に加算
                    If Not agg.Exists(key) Then agg.Add key, Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
                    Dim b0: b0 = agg(key)
                    b0(5) = b0(5) + estAmt
                    agg(key) = b0
                    hits = hits + 1
                    AddLog r, empNo, empNm, ws.Cells(r, cUch).value, estAmt, "G:顧客請求分", ""
                Else
                    AddLog r, empNo, empNm, ws.Cells(r, cUch).value, estAmt, "G:除外", matchedKw
                End If
            End If
        End If
        
        ' --- 金額取得 ---
        amt = ParseAmount(ws.Cells(r, cAmt).value)
        If amt = 0 And cFareAmt > 0 Then
            fa = ParseAmount(ws.Cells(r, cFareAmt).value)
            If fa > 0 Then amt = fa
        End If
        
        If amt <> 0 Then
            If Not agg.Exists(key) Then agg.Add key, Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
            Dim bucket: bucket = agg(key)
            
            ' キーワード判定（優先順位順）
            resultCat = ""
            matchedKw = ""
            
            ' 1. 夜間当番手当
            matchedKw = HitAnyCollection(desc, kwYakan)
            If matchedKw <> "" Then
                bucket(0) = bucket(0) + amt
                resultCat = "D:夜間当番手当"
                GoTo Decided
            End If
            
            ' 2. テレワーク手当
            matchedKw = HitAnyCollection(desc, kwTW)
            If matchedKw <> "" Then
                If Not estFilled Then bucket(4) = bucket(4) + amt
                resultCat = "J:テレワーク手当"
                GoTo Decided
            End If
            
            ' 3. RINK手当
            matchedKw = HitAnyCollection(desc, kwRink)
            If matchedKw <> "" Then
                bucket(1) = bucket(1) + amt
                resultCat = "E:RINK手当"
                GoTo Decided
            End If
            
            ' 4. 非課税精算(立替金)
            If cTrans > 0 And InStr(1, trans, "交通費", vbTextCompare) > 0 Then
                If Not estFilled Then bucket(6) = bucket(6) + amt
                resultCat = "I:非課税精算"
                GoTo Decided
            End If
            
            ' 5. 交通費（NGワードチェック含む）
            matchedKw = HitAnyCollection(desc, kwTrans)
            If matchedKw = "" Then matchedKw = HitAnyCollection(trans, kwTrans)
            If matchedKw <> "" Then
                Dim ngKw As String
                ngKw = HitAnyCollection(desc, kwTransNG)
                If ngKw = "" Then
                    If Not estFilled Then bucket(2) = bucket(2) + amt
                    resultCat = "H:交通費"
                Else
                    If Not estFilled Then bucket(3) = bucket(3) + amt
                    resultCat = "I:その他（交通費NG: " & ngKw & "）"
                    matchedKw = matchedKw & " → NG:" & ngKw
                End If
                GoTo Decided
            End If
            
            ' 6. その他（どのキーワードにもマッチしない場合は無条件で計上）
            bucket(3) = bucket(3) + amt
            resultCat = "I:その他"
            matchedKw = "(該当キーワードなし)"
            
Decided:
            agg(key) = bucket
            hits = hits + 1
            AddLog r, empNo, empNm, ws.Cells(r, cUch).value, amt, resultCat, matchedKw
        End If
        
        ' --- 計上日最大 ---
        If cBooked > 0 Then
            Dim d As Double: d = TryParseDate(ws.Cells(r, cBooked).value)
            If d > 0 Then DateMax maxDate, key, d
        End If
        
NextR:
    Next r
    
    Collect_From_Source = hits
End Function

' ============================================================
' 出力シート書き込み
' ============================================================
Private Sub Rewrite_Output(ByVal agg As Object, ByVal maxDate As Object)
    Dim ws As Worksheet
    If Not SheetExists(SH_SUM) Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = SH_SUM
    End If
    Set ws = Worksheets(SH_SUM)
    
    ' ヘッダー設定（C～Kのみ。A/Bは触らない）
    ws.Range(ws.Cells(1, 3), ws.Cells(1, 12)).ClearContents
    ws.Cells(1, 3).value = "合計"
    ws.Cells(1, 4).value = "夜間当番手当"
    ws.Cells(1, 5).value = "RINK手当"
    ws.Cells(1, 6).value = "手当2（夜間＋RINK）"
    ws.Cells(1, 7).value = "顧客請求分"
    ws.Cells(1, 8).value = "交通費"
    ws.Cells(1, 9).value = "非課税精算(立替金)"
    ws.Cells(1, 10).value = "その他(会議費・消耗品など)"
    ws.Cells(1, 11).value = "テレワーク手当"
    ws.Cells(1, 12).value = "請求日"
    ws.Columns(12).NumberFormatLocal = "yyyy/m/d"
    
    ' 社員番号単位に集約
    Dim byId As Object: Set byId = CreateObject("Scripting.Dictionary")
    Dim dateById As Object: Set dateById = CreateObject("Scripting.Dictionary")
    Dim k, id As String, arr
    
    For Each k In agg.Keys
        id = ParseEmpNo(CStr(k))
        If id <> "" Then
            arr = agg(k)
            If Not byId.Exists(id) Then byId.Add id, Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
            Dim cur: cur = byId(id)
            cur(0) = cur(0) + arr(0): cur(1) = cur(1) + arr(1): cur(2) = cur(2) + arr(2)
            cur(3) = cur(3) + arr(3): cur(4) = cur(4) + arr(4): cur(5) = cur(5) + arr(5)
            cur(6) = cur(6) + arr(6)
            byId(id) = cur
        End If
    Next k
    
    For Each k In maxDate.Keys
        id = ParseEmpNo(CStr(k))
        If id <> "" Then
            Dim dmax As Double: dmax = maxDate(k)
            If Not dateById.Exists(id) Then
                dateById.Add id, dmax
            Else
                If dmax > dateById(id) Then dateById(id) = dmax
            End If
        End If
    Next k
    
    ' 集計シートの既存の社員番号を取得してマッチング
    Dim lastR As Long
    lastR = ws.Cells(ws.rows.Count, COL_EMP_NO).End(xlUp).Row
    If lastR < 2 Then lastR = 1
    
    Dim Z As Long
    For Z = 2 To lastR
        Dim empId As String
        empId = NormalizeId(ws.Cells(Z, COL_EMP_NO).value)
        If empId <> "" And byId.Exists(empId) Then
            Dim vals: vals = byId(empId)
            ' bucket: 0=夜間当番, 1=RINK, 2=交通費, 3=その他, 4=テレワーク, 5=顧客請求
            ws.Cells(Z, COL_GK).value = vals(0)         ' D: 夜間当番手当
            ws.Cells(Z, COL_RINK).value = vals(1)       ' E: RINK手当
            ws.Cells(Z, COL_ALLOW2).value = Nz(vals(0)) + Nz(vals(1))  ' F: 手当2（D+E）
            ws.Cells(Z, COL_BILL).value = vals(5)       ' G: 顧客請求分
            ws.Cells(Z, COL_TRANS).value = vals(2)      ' H: 交通費
            ws.Cells(Z, COL_NONTAX_TATEKAE).value = vals(6)   ' I: 非課税精算(立替金) ＜ NEW
            ws.Cells(Z, COL_ETC).value = vals(3)        ' J: その他
            ws.Cells(Z, COL_TW).value = vals(4)         ' K: テレワーク手当
            
            ' C: 合計 = D + E + G + H + I + J + K
            ws.Cells(Z, COL_TOTAL).value = Nz(vals(0)) + Nz(vals(1)) + Nz(vals(5)) + _
                                            Nz(vals(2)) + Nz(vals(6)) + Nz(vals(3)) + Nz(vals(4))
            
            ' K: 計上日
            If dateById.Exists(empId) Then
                ws.Cells(Z, COL_DATE).value = CDate(dateById(empId))
            End If
            
            byId.Remove empId  ' 処理済みを除去
        End If
    Next Z
End Sub

' ============================================================
' Collection版のキーワードマッチ関数
' （設定シートからCollectionに読み込んだキーワードで判定）
' ============================================================
Private Function HitAnyCollection(ByVal s As String, ByRef kw As Collection) As String
    ' マッチしたキーワードを返す。マッチしなければ空文字列。
    HitAnyCollection = ""
    If kw Is Nothing Then Exit Function
    If kw.Count = 0 Then Exit Function
    
    Dim i As Long
    Dim keyword As String
    s = LCase$(s)
    
    For i = 1 To kw.Count
        keyword = kw(i)
        If InStr(1, s, keyword, vbTextCompare) > 0 Then
            HitAnyCollection = keyword
            Exit Function
        End If
    Next i
End Function

' ============================================================
' ユーティリティ関数群
' ============================================================

Private Function SheetExists(ByVal shtName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(shtName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

Private Sub BackupSheet(ByVal shtName As String)
    If Not SheetExists(shtName) Then Exit Sub
    Dim src As Worksheet: Set src = Worksheets(shtName)
    Dim nm As String: nm = shtName & "_backup_" & Format(Now, "yyyymmdd_HHMMSS")
    src.Copy After:=src
    ActiveSheet.Name = nm
End Sub

Private Function FindCol(ws As Worksheet, headerRow As Long, names As Variant, Optional allowMissing As Boolean = False) As Long
    Dim lastC As Long, c As Long, h As String, i As Long
    lastC = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        h = NormalizeStr(ws.Cells(headerRow, c).value)
        For i = LBound(names) To UBound(names)
            If h Like "*" & NormalizeStr(CStr(names(i))) & "*" Then
                FindCol = c: Exit Function
            End If
        Next i
    Next c
    FindCol = 0
End Function

Private Function NormalizeStr(ByVal s As String) As String
    s = Trim$(LCase$(Replace(Replace(Replace(Replace(CStr(s), vbCr, " "), vbLf, " "), vbTab, " "), "　", " ")))
    s = Replace(s, " ", "")
    NormalizeStr = s
End Function

Private Function NormalizeId(v) As String
    Dim s As String, i As Long, ch As String, out As String
    s = CStr(v)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    NormalizeId = out
End Function

Private Function NormalizeName(v) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, "　", "")
    s = Replace(s, " ", "")
    NormalizeName = LCase$(s)
End Function

Private Function BuildEmpKey(ByVal empNo As String, ByVal empNm As String) As String
    If empNo <> "" Then BuildEmpKey = "ID:" & empNo & "|NM:" & empNm Else BuildEmpKey = "NM:" & empNm
End Function

Private Function ParseEmpNo(ByVal key As String) As String
    Dim p As Long, q As Long, id As String
    p = InStr(1, key, "ID:")
    If p > 0 Then
        q = InStr(p + 3, key, "|NM:")
        If q > 0 Then
            id = Mid$(key, p + 3, q - (p + 3))
        Else
            id = Mid$(key, p + 3)
        End If
        ParseEmpNo = NormalizeId(id)
    Else
        ParseEmpNo = ""
    End If
End Function

Private Function ParseAmount(v) As Double
    Dim s As String: s = CStr(v)
    s = Replace(s, ",", "")
    s = Replace(s, "￥", "")
    s = Replace(s, "円", "")
    s = Replace(s, "(", "-")
    s = Replace(s, ")", "")
    s = Replace(s, "（", "-")
    s = Replace(s, "）", "")
    s = Trim$(s)
    If s = "" Then ParseAmount = 0: Exit Function
    If IsNumeric(s) Then
        ParseAmount = CDbl(s)
    Else
        ParseAmount = 0
    End If
End Function

Private Function TryParseDate(v) As Double
    On Error Resume Next
    If IsDate(v) Then
        TryParseDate = CDbl(CDate(v))
    Else
        TryParseDate = 0
    End If
    On Error GoTo 0
End Function

Private Sub DateMax(ByRef dict As Object, ByVal key As String, ByVal d As Double)
    If Not dict.Exists(key) Then
        dict.Add key, d
    Else
        If d > dict(key) Then dict(key) = d
    End If
End Sub

Private Function Nz(v) As Double
    If isEmpty(v) Or IsNull(v) Then Nz = 0 Else Nz = CDbl(v)
End Function

' ============================================================
' 設定シート作成（初回セットアップ用）
' ============================================================
Public Sub Setup_設定シート作成()
    If SheetExists(SH_SETTING) Then
        If MsgBox("「設定」シートは既に存在します。上書きしますか？", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        Application.DisplayAlerts = False
        Worksheets(SH_SETTING).Delete
        Application.DisplayAlerts = True
    End If
    
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = SH_SETTING
    
    ' ヘッダー
    ws.Cells(1, 1).value = "分類名"
    ws.Cells(1, 2).value = "キーワード"
    ws.rows(1).Font.Bold = True
    
    ' 初期データ（過去の修正をすべて反映済み）
    ' ※VBAの行継続文字の上限(25行)を超えないよう、AddRow関数で1行ずつ書き込み
    Dim r As Long: r = 2
    
    ' --- 夜間当番手当 ---
    r = AddRow(ws, r, CAT_YAKAN, "夜間当番")
    r = AddRow(ws, r, CAT_YAKAN, "24時間準直当番")
    r = AddRow(ws, r, CAT_YAKAN, "準直当番")
    r = AddRow(ws, r, CAT_YAKAN, "深夜出動")
    r = AddRow(ws, r, CAT_YAKAN, "顧客当番")
    r = AddRow(ws, r, CAT_YAKAN, "顧客対応当番")
    r = AddRow(ws, r, CAT_YAKAN, "オンコール")
    
    ' --- RINK手当 ---
    r = AddRow(ws, r, CAT_RINK, "RINK")
    
    ' --- テレワーク手当 ---
    r = AddRow(ws, r, CAT_TW, "テレワーク")
    r = AddRow(ws, r, CAT_TW, "在宅")
    
    ' --- 交通費 ---
    r = AddRow(ws, r, CAT_TRANS, "交通費")
    r = AddRow(ws, r, CAT_TRANS, "電車")
    r = AddRow(ws, r, CAT_TRANS, "バス")
    r = AddRow(ws, r, CAT_TRANS, "タクシー")
    r = AddRow(ws, r, CAT_TRANS, "地下鉄")
    r = AddRow(ws, r, CAT_TRANS, "鉄道")
    r = AddRow(ws, r, CAT_TRANS, "新幹線")
    r = AddRow(ws, r, CAT_TRANS, "モノレール")
    r = AddRow(ws, r, CAT_TRANS, "JR")
    r = AddRow(ws, r, CAT_TRANS, "私鉄")
    r = AddRow(ws, r, CAT_TRANS, "有料列車")
    r = AddRow(ws, r, CAT_TRANS, "飛行機")
    r = AddRow(ws, r, CAT_TRANS, "航空券")
    r = AddRow(ws, r, CAT_TRANS, "定期券")
    r = AddRow(ws, r, CAT_TRANS, "定期代")
    r = AddRow(ws, r, CAT_TRANS, "ガソリン")
    r = AddRow(ws, r, CAT_TRANS, "燃料")
    r = AddRow(ws, r, CAT_TRANS, "駐車場")
    r = AddRow(ws, r, CAT_TRANS, "パーキング")
    r = AddRow(ws, r, CAT_TRANS, "高速")
    r = AddRow(ws, r, CAT_TRANS, "ETC")
    r = AddRow(ws, r, CAT_TRANS, "車両")
    r = AddRow(ws, r, CAT_TRANS, "レンタカー")
    r = AddRow(ws, r, CAT_TRANS, "移動")
    r = AddRow(ws, r, CAT_TRANS, "旅費")
    r = AddRow(ws, r, CAT_TRANS, "出張")
    r = AddRow(ws, r, CAT_TRANS, "日当")
    r = AddRow(ws, r, CAT_TRANS, "宿泊")
    r = AddRow(ws, r, CAT_TRANS, "ホテル")
    r = AddRow(ws, r, CAT_TRANS, "北総鉄道北総線")
    
    ' --- 交通費除外（NGワード）---
    r = AddRow(ws, r, CAT_TRANS_NG, "会議")
    r = AddRow(ws, r, CAT_TRANS_NG, "交際")
    r = AddRow(ws, r, CAT_TRANS_NG, "接待")
    r = AddRow(ws, r, CAT_TRANS_NG, "飲食")
    r = AddRow(ws, r, CAT_TRANS_NG, "手土産")
    r = AddRow(ws, r, CAT_TRANS_NG, "福利厚生")
    r = AddRow(ws, r, CAT_TRANS_NG, "親睦")
    r = AddRow(ws, r, CAT_TRANS_NG, "定期健康診断")
    r = AddRow(ws, r, CAT_TRANS_NG, "健康診断")
    
    ' --- 顧客請求除外 ---
    r = AddRow(ws, r, CAT_KOKYAKU_NG, "夜間当番")
    r = AddRow(ws, r, CAT_KOKYAKU_NG, "RINK")
    r = AddRow(ws, r, CAT_KOKYAKU_NG, "顧客当番")
    r = AddRow(ws, r, CAT_KOKYAKU_NG, "顧客対応当番")
    r = AddRow(ws, r, CAT_KOKYAKU_NG, "オンコール")
    
    ' 列幅調整
    ws.Columns("A:B").AutoFit
    
    ' 分類名ごとに色分け
    Dim lastR As Long: lastR = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        Select Case ws.Cells(r, 1).value
            Case CAT_YAKAN
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(255, 230, 230)
            Case CAT_RINK
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(230, 255, 230)
            Case CAT_TW
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(230, 230, 255)
            Case CAT_TRANS
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(255, 255, 230)
            Case CAT_TRANS_NG
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(255, 220, 180)
            Case CAT_KOKYAKU_NG
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(220, 220, 220)
        End Select
    Next r
    
    MsgBox "「設定」シートを作成しました。" & vbCrLf & _
           "キーワードの追加・変更はこのシートで行ってください。", vbInformation
End Sub

' ============================================================
' 設定シートに1行追加するヘルパー関数
' （行番号を返すので、呼び出し側で r = AddRow(...) と書ける）
' ============================================================
Private Function AddRow(ws As Worksheet, ByVal r As Long, ByVal cat As String, ByVal kw As String) As Long
    ws.Cells(r, 1).value = cat
    ws.Cells(r, 2).value = kw
    AddRow = r + 1
End Function

