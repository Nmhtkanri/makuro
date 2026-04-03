Attribute VB_Name = "Module1"
' ===============================
' 勤務時間帯一覧：フル再実装（安定版 / 深夜回避なし / 3ルール休憩）
' ===============================
Option Explicit
' ▼参照元（パスリスト）シートと列
Private Const SHEET_PATHLIST As String = "パス一覧"
Private Const COL_EMP_ID As Long = 1 ' A
Private Const COL_EMP_NM As Long = 2 ' B
Private Const COL_PATH   As Long = 3 ' C

' ▼出力・ログ
Private Const OUTPUT_SHEET As String = "勤務時間帯一覧"
Private Const LOG_SHEET    As String = "抽出ログ"

' ========== エントリ ==========
Public Sub RunImport()
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wsPL As Worksheet
    If Not SheetExists(ThisWorkbook, SHEET_PATHLIST) Then
        MsgBox "シートがありません: " & SHEET_PATHLIST, vbExclamation
        GoTo FinallyExit
    End If
    Set wsPL = ThisWorkbook.Worksheets(SHEET_PATHLIST)

    Dim wsOut As Worksheet: Set wsOut = GetOrCreateSheet(OUTPUT_SHEET)
    Dim wsLog As Worksheet: Set wsLog = GetOrCreateSheet(LOG_SHEET)
    EnsureHeaders_Out wsOut
    ' ▼出力シート初期化（ヘッダー行を残して2行目以降を削除）
    With wsOut
        If .Rows.Count > 1 Then
            .Rows("2:" & .Rows.Count).ClearContents
        End If
    End With
    EnsureHeaders_Log wsLog

    Dim lastRow As Long: lastRow = wsPL.Cells(wsPL.Rows.Count, COL_PATH).End(xlUp).row

    Dim r As Long
    For r = 2 To lastRow ' 1行目は見出し
        Dim folderPath As String
        folderPath = Trim$(CStr(wsPL.Cells(r, COL_PATH).value))
        If Len(folderPath) = 0 Then GoTo NextRow
        If Right$(folderPath, 1) <> "\" And Right$(folderPath, 1) <> "/" Then folderPath = folderPath & "\"
        If Dir(folderPath, vbDirectory) = vbNullString Then
            LogMsg wsLog, "INFO", "フォルダなし: " & folderPath
            GoTo NextRow
        End If

        Dim empID As String, empNM As String
        empID = Trim$(CStr(wsPL.Cells(r, COL_EMP_ID).value))
        empNM = Trim$(CStr(wsPL.Cells(r, COL_EMP_NM).value))

        ' 勤務時間報告書*.xlsx を優先、無ければ *.xlsx
        Dim f As String
        f = Dir(folderPath & "勤務時間報告書*.xlsx")
        If Len(f) = 0 Then f = Dir(folderPath & "*.xlsx")
        If Len(f) = 0 Then
            LogMsg wsLog, "INFO", "xlsx不在: " & folderPath
            GoTo NextRow
        End If

        Do While Len(f) > 0
            ImportOne folderPath & f, empID, empNM, wsOut, wsLog
            f = Dir()
        Loop
NextRow:
    Next r

FinallyExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

EH:
    MsgBox "取込エラー: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume FinallyExit
    
   
End Sub

' ========== 個別ファイル取込 ==========
Private Sub ImportOne(ByVal filePath As String, ByVal empID As String, ByVal empNM As String, _
                      ByVal wsOut As Worksheet, ByVal wsLog As Worksheet)
    On Error GoTo ErrorHandler  ' ★変数名衝突を避けるため eH から変更
    Dim wb As Workbook: Set wb = Workbooks.Open(fileName:=filePath, ReadOnly:=True)

    Dim ws As Worksheet: Set ws = FindTargetSheet(wb)
    If ws Is Nothing Then
        LogMsg wsLog, "INFO", "シート不明: " & filePath
        GoTo Clean
    End If
    
    ' ★形式チェック
    If Not IsValidTimesheet(ws) Then
        LogMsg wsLog, "INFO", "スキップ（形式不一致）: " & filePath
        GoTo Clean
    End If
    
    ' ▼ 対象年月を「年月日」シートから取得
    Dim TARGET_YEAR  As Long
    Dim TARGET_MONTH As Long
    Dim wsYM As Worksheet
    On Error Resume Next
    Set wsYM = ThisWorkbook.Worksheets("年月日")
    On Error GoTo ErrorHandler  ' ★ここも変更
    
    If wsYM Is Nothing Then
        MsgBox "「年月日」シートが見つかりません。", vbCritical
        GoTo Clean
    Else
        TARGET_YEAR = Val(wsYM.Range("A1").value)
        TARGET_MONTH = Val(wsYM.Range("B1").value)
        If TARGET_YEAR = 0 Or TARGET_MONTH = 0 Then
            MsgBox "年月日シートのA1/B1に正しい年と月を入力してください。", vbExclamation
            GoTo Clean
        End If
    End If
    
    Dim r As Long
    For r = 13 To 43
        ' --- 日付（C列）---
        Dim vDate As Variant: vDate = ws.Cells(r, "C").value
        If Len(Trim$(CStr(vDate))) = 0 Then Exit For ' 以降終了

        Dim outDate As Date
        If IsDate(vDate) Then
            outDate = CDate(vDate)
        ElseIf IsNumeric(vDate) Then
            outDate = DateSerial(TARGET_YEAR, TARGET_MONTH, CLng(vDate))
        Else
            ' 文字列の場合も試みる
            On Error Resume Next
            outDate = DateSerial(TARGET_YEAR, TARGET_MONTH, CLng(vDate))
            On Error GoTo ErrorHandler  ' ★ここも変更
        End If

        ' --- 時刻（F:G / H:I または小数の時刻）---
        ' ★変数名を変更：eH → endH, eM → endM
        Dim sh As Long, sm As Long, endH As Long, endM As Long
        ReadTimePair ws, r, "F", "G", sh, sm   ' 開始
        ReadTimePair ws, r, "H", "I", endH, endM   ' 終了

        ' 空行でも行は出力するため、スキップしない
        Dim hasTime As Boolean
        hasTime = Not (sh = 0 And sm = 0 And endH = 0 And endM = 0)

        ' --- 休憩（Excelの時間＋分をそのまま使う）---
        Dim brH As Long, brM As Long
        brH = 0: brM = 0

        If Not IsEmpty(ws.Cells(r, "J").value) Then brH = Val(ws.Cells(r, "J").value)
        If Not IsEmpty(ws.Cells(r, "K").value) Then brM = Val(ws.Cells(r, "K").value)
        
        ' --- 勤務区分（E列：出勤／所休／法休／有給 などを文字列で）---
        Dim statusRaw As String
        statusRaw = Replace(Replace(Trim$(CStr(ws.Cells(r, "E").value)), vbCr, " "), vbLf, " ")

        ' --- テレワーク（R列そのまま＋確認用 "h:mm"）---
        Dim teleRaw As Variant: teleRaw = ws.Cells(r, "R").value
        Dim teleView As String
        teleView = CanonicalHHMM(teleRaw)  ' 例: 7:30 / ""（不正・空は空欄）

        ' --- 出力（既存列はそのまま）---
        Dim nr As Long: nr = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).row + 1
        wsOut.Cells(nr, 1).value = empID
        wsOut.Cells(nr, 2).value = empNM
        wsOut.Cells(nr, 3).value = outDate                 ' yyyy/m/d

        ' ▼開始(D)・終了(E) … 時刻がどっちもゼロなら空欄、それ以外は書く
        If (sh = 0 And sm = 0 And endH = 0 And endM = 0) Then
            wsOut.Cells(nr, 4).value = ""                  ' D: 開始
            wsOut.Cells(nr, 5).value = ""                  ' E: 終了
        Else
            wsOut.Cells(nr, 4).value = HMSerial(sh, sm)    ' [h]:mm
            wsOut.Cells(nr, 5).value = HMSerial(endH, endM)    ' [h]:mm  ★変数名変更
        End If

        ' 休憩開始に時間を入れて、終了は一旦空欄にする
        If brH = 0 And brM = 0 Then
            wsOut.Cells(nr, 6).value = ""
            wsOut.Cells(nr, 7).value = ""
        Else
            wsOut.Cells(nr, 6).value = HMSerial(brH, brM)  ' [h]:mm
            wsOut.Cells(nr, 7).value = ""                   ' 終了時刻は未設定
        End If

        ' テレワークは元値＋確認表示
        wsOut.Cells(nr, 8).NumberFormat = "@"
        wsOut.Cells(nr, 8).value = CStr(teleRaw)          ' 元値そのまま
        wsOut.Cells(nr, 9).NumberFormat = "@"
        wsOut.Cells(nr, 9).value = teleView                ' h:mm 文字列
        wsOut.Cells(nr, 10).NumberFormat = "@"
        wsOut.Cells(nr, 10).value = statusRaw
    Next r

Clean:
    On Error Resume Next
    wb.Close SaveChanges:=False
    Exit Sub
    
ErrorHandler:  ' ★エラーハンドラのラベル名も変更
    LogMsg wsLog, "ERROR", "取込失敗: " & filePath & " / " & Err.Number & ":" & Err.Description
    Resume Clean
End Sub

' ========== ターゲットシート検出 ==========
Private Function FindTargetSheet(ByVal wb As Workbook) As Worksheet
    Dim s As Worksheet
    For Each s In wb.Worksheets
        If Left$(s.Name, Len("勤務時間報告書")) = "勤務時間報告書" Then Set FindTargetSheet = s: Exit Function
    Next
    If wb.Worksheets.Count > 0 Then Set FindTargetSheet = wb.Worksheets(1) Else Set FindTargetSheet = Nothing
End Function

' ========== 表示形式/ログ ==========
Private Sub EnsureHeaders_Out(ByVal ws As Worksheet)
    ' 見出しは無いときだけ書く、表示形式は毎回強制
    If Len(CStr(ws.Cells(1, 1).value)) = 0 Then
        ws.Range("A1:J1").value = Array( _
            "従業員番号", "名前", "日付", "開始時刻", "終了時刻", "休憩開始", "休憩終了", _
            "テレワーク(元値)", "テレワーク(確認表示)", "勤務区分")
    End If
    With ws
        .Columns(3).NumberFormat = "yyyy/m/d"
        .Columns(4).NumberFormat = "[h]:mm"
        .Columns(5).NumberFormat = "[h]:mm"
        .Columns(6).NumberFormat = "[h]:mm"
        .Columns(7).NumberFormat = "[h]:mm"
        .Columns(8).NumberFormat = "@"
        .Columns(9).NumberFormat = "@"
        .Columns(10).NumberFormat = "@"
    End With
End Sub

Private Sub EnsureHeaders_Log(ByVal ws As Worksheet)
    If Len(CStr(ws.Cells(1, 1).value)) = 0 Then
        ws.Range("A1:C1").value = Array("時刻", "種別", "メッセージ")
    End If
    ws.Columns(1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
End Sub

Private Sub LogMsg(ByVal ws As Worksheet, ByVal kind As String, ByVal msg As String)
    Dim nr As Long: nr = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    ws.Cells(nr, 1).value = Now
    ws.Cells(nr, 2).value = kind
    ws.Cells(nr, 3).value = msg
End Sub

' ========== 入力の時刻を読む（小数/時分 自動判定） ==========
Private Sub ReadTimePair(ByVal ws As Worksheet, ByVal r As Long, _
                         ByVal colH As String, ByVal colM As String, _
                         ByRef outH As Long, ByRef outM As Long)
    Dim vH As Variant, vM As Variant
    vH = ws.Cells(r, colH).value
    vM = ws.Cells(r, colM).value

    If IsNumeric(vH) And Not IsNumeric(vM) Then
        ' vH が小数の時刻（0.0?数日分）・vM は空や文字列のケース
        Dim ser As Double: ser = CDbl(vH) * 24#
        outH = Fix(ser)
        outM = Round((ser - outH) * 60)
    ElseIf IsNumeric(vH) And IsNumeric(vM) And (CDbl(vH) < 5# And CDbl(vM) = 0) Then
        ' 両方小数のパターンは稀だが、vHのみ採用
        Dim ser2 As Double: ser2 = CDbl(vH) * 24#
        outH = Fix(ser2)
        outM = Round((ser2 - outH) * 60)
    Else
        ' 通常：時・分の別列
        outH = NzLng(vH)
        outM = NzLng(vM)
    End If
End Sub

' ========== 休憩（3ルール） ==========
Private Sub DecideBreakBySimpleRules(ByVal sh As Long, ByVal sm As Long, _
                                     ByVal EH As Long, ByVal eM As Long, _
                                     ByRef brS As Long, ByRef brE As Long)
    Dim dispEndH As Long: dispEndH = EH
    If EH < sh Then dispEndH = EH + 24 ' 翌日跨ぎ → +24 表示

    If sh = 15 And sm = 0 Then
        brS = 15: brE = 16          ' 15:00→16:00
    ElseIf dispEndH = 33 Then
        brS = 31: brE = 32          ' 終了33:00→31:00～32:00
    Else
        brS = 12: brE = 13          ' その他→12:00～13:00
    End If
End Sub

' ========== シート/書式ユーティリティ ==========
Private Function GetOrCreateSheet(ByVal nm As String) As Worksheet
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim w As Worksheet
    For Each w In wb.Worksheets
        If StrComp(w.Name, nm, vbTextCompare) = 0 Then Set GetOrCreateSheet = w: Exit Function
    Next
    On Error GoTo MK
    Set w = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    w.Name = nm: Set GetOrCreateSheet = w: Exit Function
MK:
    w.Name = nm & "_copy": Set GetOrCreateSheet = w
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim w As Worksheet
    For Each w In wb.Worksheets
        If StrComp(w.Name, nm, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next
    SheetExists = False
End Function

' ========== 時刻ヘルパ ==========
Private Function HMSerial(ByVal h As Long, ByVal m As Long) As Double
    HMSerial = (h / 24#) + (m / 1440#)  ' [h]:mm 表示向け
End Function

Private Function HSerial(ByVal h As Long) As Double
    HSerial = h / 24#
End Function

Private Function HHMMToSerial(ByVal s As String) As Double
    Dim p() As String, h As Long, m As Long
    If InStr(s, ":") = 0 Then HHMMToSerial = 0: Exit Function
    p = Split(s, ":")
    h = CLng(Trim$(p(0)))
    If UBound(p) >= 1 Then m = CLng(Trim$(p(1))) Else m = 0
    HHMMToSerial = (h / 24#) + (m / 1440#)
End Function

Private Function NzLng(ByVal v As Variant) As Long
    If IsNumeric(v) Then NzLng = CLng(v) Else NzLng = 0
End Function

' ========== 目視用 "h:mm" 文字列に正規化 ==========
Private Function CanonicalHHMM(ByVal v As Variant) As String
    Dim mins As Long
    mins = ToMinutesFromVariant(v)
    If mins <= 0 Then
        CanonicalHHMM = ""              ' 空や不正は空欄に
    Else
        CanonicalHHMM = CStr(mins \ 60) & ":" & Format$(mins Mod 60, "00")
    End If
End Function

' ========== Variant → 分（整数）に解釈 ==========
Private Function ToMinutesFromVariant(ByVal v As Variant) As Long
    If IsNumeric(v) Then
        ToMinutesFromVariant = CLng(Round(CDbl(v) * 24# * 60#)) ' シリアル/小数→分
        Exit Function
    End If
    Dim s As String: s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    If InStr(s, ":") > 0 Then
        Dim p() As String, h As Long, m As Long
        p = Split(s, ":")
        h = Val(p(0))
        If UBound(p) >= 1 Then m = Val(p(1)) Else m = 0
        ToMinutesFromVariant = h * 60& + m
    End If
End Function

' ========== 形式チェック関数 ==========
Private Function IsValidTimesheet(ByVal ws As Worksheet) As Boolean
    On Error GoTo NG
    Dim score As Long: score = 0

    ' 1) タイトルや表頭のサイン（どれか当たれば加点）
    If LCase$(CStr(ws.Range("C12").value)) Like "*日*" Then score = score + 1
    If LCase$(CStr(ws.Range("E12").value)) Like "*勤*" Then score = score + 1
    If LCase$(CStr(ws.Range("F12").value)) Like "*開始*" Then score = score + 1
    If LCase$(CStr(ws.Range("H12").value)) Like "*終了*" Then score = score + 1
    If LCase$(CStr(ws.Range("R12").value)) Like "*テレワ*" Then score = score + 1

    ' 2) 実データ行（13行目）の型感触
    Dim vC As Variant, vE As Variant, vF As Variant, vH As Variant, vR As Variant
    vC = ws.Cells(13, "C").value
    vE = ws.Cells(13, "E").value
    vF = ws.Cells(13, "F").value
    vH = ws.Cells(13, "H").value
    vR = ws.Cells(13, "R").value

    If IsDate(vC) Or IsNumeric(vC) Then score = score + 1
    If Len(Trim$(CStr(vE))) > 0 Then score = score + 1
    If (IsDate(vF) Or IsNumeric(vF)) Then score = score + 1
    If (IsDate(vH) Or IsNumeric(vH)) Then score = score + 1
    If (IsDate(vR) Or IsNumeric(vR) Or InStr(CStr(vR), ":") > 0) Then score = score + 1

    ' ★しきい値：合計 3 以上で「妥当」とする（4から3に緩和）
    IsValidTimesheet = (score >= 3)
    Exit Function
NG:
    IsValidTimesheet = False
End Function

' ▼ 連続重複（日付）の上側を消す
Public Sub RemoveConsecutiveDuplicateDates()
    On Error GoTo ErrH
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim id1 As Variant, id2 As Variant
    Dim nm1 As Variant, nm2 As Variant
    Dim d1 As Double, d2 As Double  ' 日付はシリアルで比較が安全

    Set ws = ThisWorkbook.Worksheets(OUTPUT_SHEET) ' = "勤務時間帯一覧"
    ' C列（日付）基準で最終行
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).row
    If lastRow < 3 Then Exit Sub

    ' ※ 並び順は今のまま（ソートしない）。連続しているペアだけを対象にする。
    r = 3
    Do While r <= lastRow
        id1 = ws.Cells(r - 1, 1).value
        id2 = ws.Cells(r, 1).value
        nm1 = ws.Cells(r - 1, 2).value
        nm2 = ws.Cells(r, 2).value

        ' 日付は空や文字も来るので Value2 を Double として扱う（IsDateならCDbl）
        d1 = 0: d2 = 0
        If IsDate(ws.Cells(r - 1, 3).value) Then d1 = CDbl(CDate(ws.Cells(r - 1, 3).value))
        If IsDate(ws.Cells(r, 3).value) Then d2 = CDbl(CDate(ws.Cells(r, 3).value))

        If (id1 = id2) And (nm1 = nm2) And (d1 <> 0) And (d1 = d2) Then
            ' 連続・同一日付 → 上の行（r-1）を削除、カウンタ巻き戻し
            ws.Rows(r - 1).Delete
            lastRow = lastRow - 1
            If r > 3 Then r = r - 1
        Else
            r = r + 1
        End If
    Loop
    Exit Sub
ErrH:
    MsgBox "RemoveConsecutiveDuplicateDates エラー: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub


