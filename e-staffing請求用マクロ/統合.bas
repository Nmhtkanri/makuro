Attribute VB_Name = "統合"
Option Explicit

' ==========================================================================
'  一括処理マクロ（統合版・修正版）
'  ファイルA（準備・取込）とファイルB（請求処理パイプライン）を統合
'  3つのファイルを選択後、全15ステップを自動実行します
'
'  【実行フロー】
'    Phase 1 (Step 1-3):  ファイル取込
'    Phase 2 (Step 4-6):  データ準備
'    Phase 3 (Step 7-15): 請求処理・ファイル出力
'
'  【修正内容】
'    - Step4_請求ヘッダ作成: 数式をINDIRECT+ROW関数に変更し、
'      コピー時に相対参照が効くように修正
' ==========================================================================

' ■モジュールレベル変数
Private p_wsMain As Worksheet          ' メインシート（勤務データ）
Private p_FolderPath As String         ' 出力先パス（設定シートC2セルの値）
Private p_DateStr As String            ' 日付スタンプ（mmdd）
Private p_TimeStr As String            ' 時間スタンプ（hhmm）
Private Const TOTAL_STEPS As Long = 15 ' 全ステップ数
Private p_CurrentStep As Long          ' 現在のステップ番号
Private p_LogPath As String            ' ログファイルパス
Private p_ErrorOccurred As Boolean     ' エラー発生フラグ（複数ポップアップ防止用）


' ==============================================================================
' 【メインエントリーポイント】このマクロをボタンに登録してください
' ==============================================================================
Sub 一括処理()
    On Error GoTo MainErrorHandler

    p_CurrentStep = 0
    p_ErrorOccurred = False  ' エラーフラグ初期化

    ' ─── ファイル選択（処理開始前に3つまとめて選択）───
    Dim fileTCH As Variant       ' webTC_data用 (TCH*.txt)
    Dim fileADVC As Variant      ' 立替金データ用 (ADVC*.csv)
    Dim fileTCnmht As Variant    ' e-staffing用 (TCnmht*.csv)

    fileTCH = Application.GetOpenFilename( _
        "TEXTファイル(*.txt),TCH*.txt", , _
        "【1/3】WebTimeCard勤怠データ(TCH*.txt)を選択してください")
    If VarType(fileTCH) = vbBoolean Then
        MsgBox "キャンセルされました。処理を中断します。", vbExclamation, "中断"
        Exit Sub
    End If

    fileADVC = Application.GetOpenFilename( _
        "CSVファイル(*.csv),ADVC*.csv", , _
        "【2/3】立替金データ(ADVC*.csv)を選択してください")
    If VarType(fileADVC) = vbBoolean Then
        MsgBox "キャンセルされました。処理を中断します。", vbExclamation, "中断"
        Exit Sub
    End If

    fileTCnmht = Application.GetOpenFilename( _
        "CSVファイル(*.csv),*.csv", , _
        "【3/3】e-Staffing契約データ(TCnmht*.csv)を選択してください")
    If VarType(fileTCnmht) = vbBoolean Then
        MsgBox "キャンセルされました。処理を中断します。", vbExclamation, "中断"
        Exit Sub
    End If

    ' 選択内容の確認
    Dim confirmMsg As String
    confirmMsg = "以下のファイルで処理を実行します。" & vbCrLf & vbCrLf & _
                 "1. webTC_data: " & Dir(CStr(fileTCH)) & vbCrLf & _
                 "2. 立替金データ: " & Dir(CStr(fileADVC)) & vbCrLf & _
                 "3. e-staffing: " & Dir(CStr(fileTCnmht)) & vbCrLf & vbCrLf & _
                 "実行しますか？"
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "確認") = vbNo Then
        Exit Sub
    End If

    ' ─── 初期設定 ───
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set p_wsMain = ThisWorkbook.Worksheets("勤務データ")

    p_FolderPath = ThisWorkbook.Worksheets("設定").Range("C2").Value
    If p_FolderPath = "" Then Err.Raise 1001, , "設定シートのC2セルに出力先フォルダパスが入力されていません。"
    If Right(p_FolderPath, 1) = "\" Then p_FolderPath = Left(p_FolderPath, Len(p_FolderPath) - 1)
    If Dir(p_FolderPath, vbDirectory) = "" Then Err.Raise 1002, , "指定された出力先フォルダが存在しません。" & vbCrLf & p_FolderPath

    p_DateStr = Format(Date, "mmdd")
    p_TimeStr = Format(Time, "hhmm")

    ' ログファイル初期化
    p_LogPath = p_FolderPath & "\一括処理_log_" & p_DateStr & p_TimeStr & ".txt"
    WriteLog "===== 一括処理 開始 ====="
    WriteLog "出力先: " & p_FolderPath
    WriteLog "ファイル1: " & CStr(fileTCH)
    WriteLog "ファイル2: " & CStr(fileADVC)
    WriteLog "ファイル3: " & CStr(fileTCnmht)


    ' === Phase 1: ファイル取込（Step 1-3）===

    Call UpdateStatus(1, "webTC_data ファイル取込")
    Call ImportFileToSheet(CStr(fileTCH), "webTC_data")

    Call UpdateStatus(2, "立替金データ ファイル取込")
    Call ImportFileToSheet(CStr(fileADVC), "立替金データ")

    Call UpdateStatus(3, "e-staffing データ ファイル取込")
    Call ImportFileToSheet(CStr(fileTCnmht), "e-staffing TCnmhtの最新情報")


    ' === Phase 2: データ準備（Step 4-6）===

    Call UpdateStatus(4, "名簿2作成")
    Call Phase2_名簿2作成

    Call UpdateStatus(5, "webTCデータの抽出")
    Call Phase2_webTCデータの抽出

    Call UpdateStatus(6, "スタッフコード更新")
    Call Phase2_スタッフコード


    ' === Phase 3: 請求処理パイプライン（Step 7-15）===

    Call UpdateStatus(7, "立替金集計")
    Call Step1_立替金集計

    Call UpdateStatus(8, "立替データ保存")
    Call Step2_立替データ保存

    Call UpdateStatus(9, "請求明細作成")
    Call Step3_請求明細作成

    Call UpdateStatus(10, "請求ヘッダ作成")
    Call Step4_請求ヘッダ作成

    Call UpdateStatus(11, "請求データとの照合")
    Call Step5_請求データとの照合

    Call UpdateStatus(12, "データバックアップ保存")
    Call Step6_データ保存

    Call UpdateStatus(13, "BH/BDファイル作成")
    Call Step7_提出データファイル作成

    Call UpdateStatus(14, "BTファイル作成")
    Call Step8_ExportTimecardData

    Call UpdateStatus(15, "ZIP作成・ファイル整理")
    Call Step9_Method


    ' ─── 終了処理 ───
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False

    WriteLog "===== 一括処理 正常完了 ====="

    MsgBox "全ての処理が完了しました。（全" & TOTAL_STEPS & "ステップ）" & vbCrLf & _
           "出力先: " & p_FolderPath & vbCrLf & _
           "ログ: " & p_LogPath, vbInformation, "完了"
    Exit Sub

MainErrorHandler:
    ' 既にエラーが発生している場合は、二重エラー防止のため即座に終了
    If p_ErrorOccurred Then
        Exit Sub
    End If

    ' エラーフラグを立てる（これ以降はエラーメッセージを表示しない）
    p_ErrorOccurred = True

    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    ' ログ出力（エラーが発生してもログは記録する）
    On Error Resume Next
    WriteLog "!!!!! エラー発生 !!!!!"
    WriteLog "Step: " & p_CurrentStep & "/" & TOTAL_STEPS
    WriteLog "Err.Number: " & errNum
    WriteLog "Err.Description: " & errDesc
    WriteLog "Err.Source: " & errSrc
    On Error GoTo 0

    ' アプリケーション設定を元に戻す
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    On Error GoTo 0

    ' エラーメッセージを表示（最初の1回のみ）
    MsgBox "ステップ " & p_CurrentStep & "/" & TOTAL_STEPS & " でエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "内容: " & errDesc & vbCrLf & vbCrLf & _
           "詳細はログを確認してください:" & vbCrLf & p_LogPath, vbCritical, "エラー中断"

    ' 処理を完全に中断
    Exit Sub
End Sub


' ==============================================================================
' ユーティリティ
' ==============================================================================

Private Sub UpdateStatus(stepNum As Long, stepName As String)
    p_CurrentStep = stepNum
    Application.StatusBar = "処理中：" & stepNum & "/" & TOTAL_STEPS & " " & stepName
    WriteLog "========== Step " & stepNum & "/" & TOTAL_STEPS & " " & stepName & " =========="
End Sub

Private Sub WriteLog(msg As String)
    ' イミディエイトウィンドウ＋ログファイルに出力
    Dim logLine As String
    logLine = Format(Now, "yyyy/mm/dd hh:nn:ss") & " | " & msg

    ' イミディエイトウィンドウ（常に出力）
    Debug.Print logLine

    ' ファイル出力（パスが設定済みの場合）
    If p_LogPath <> "" Then
        Dim f As Integer
        f = FreeFile
        Open p_LogPath For Append As #f
        Print #f, logLine
        Close #f
    End If
End Sub

Private Function CellInfo(ws As Worksheet, r As Long, c As Long) As String
    ' セルの値と型をログ用に文字列化する
    Dim v As Variant
    On Error Resume Next
    v = ws.Cells(r, c).Value
    If Err.Number <> 0 Then
        CellInfo = "ERROR(読取不可)"
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    CellInfo = "Val=[" & CStr(Nz(v, "NULL")) & "] Type=" & TypeName(v)
End Function

Private Function Nz(v As Variant, Optional defaultVal As String = "") As String
    ' Null/Emptyを安全に文字列化
    If IsNull(v) Or IsEmpty(v) Then
        Nz = defaultVal
    Else
        On Error Resume Next
        Nz = CStr(v)
        If Err.Number <> 0 Then Nz = "(変換不可)"
        On Error GoTo 0
    End If
End Function

Private Sub ImportFileToSheet(filePath As String, targetSheetName As String)
    ' ファイルを開いて全セルをコピーし、対象シートのA1に貼り付けて閉じる
    Dim wsTarget As Worksheet
    Dim wbSource As Workbook

    Set wsTarget = ThisWorkbook.Worksheets(targetSheetName)

    Set wbSource = Workbooks.Open(filePath)
    wbSource.ActiveSheet.Cells.Copy

    ThisWorkbook.Activate
    wsTarget.Select
    wsTarget.Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

    wbSource.Close SaveChanges:=False
End Sub


' ==============================================================================
' Phase 2: データ準備（ファイルAのマクロを統合）
' ==============================================================================

Private Sub Phase2_名簿2作成()
    ' 名簿シートからデータを「名簿 (2)」にコピー（列順を入替）
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim MR1 As Long, i As Long

    Set ws1 = Worksheets("名簿")
    Set ws2 = Worksheets("名簿 (2)")

    MR1 = ws1.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To MR1
        ws2.Cells(i, 1).Value = ws1.Cells(i, 2).Value
        ws2.Cells(i, 2).Value = ws1.Cells(i, 1).Value
        ws2.Cells(i, 3).Value = ws1.Cells(i, 3).Value
    Next
End Sub

Private Sub Phase2_webTCデータの抽出()
    ' webTC_dataシートから勤務データシートへデータを集計・展開
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim HC(1000) As Long
    Dim n1 As Long, n2 As Long
    Dim MR1 As Long, MR2 As Long, MR11 As Long
    Dim hc_cnt As Long, hc_cnt_max As Long
    Dim i As Long
    Dim JOB_C As Variant, S_name As Variant
    Dim hantei As Double

    Set ws1 = p_wsMain              ' 勤務データ（出力先）
    Set ws2 = Worksheets("webTC_data")
    Set ws3 = Worksheets("名簿")

    MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    MR2 = ws2.Cells(Rows.Count, 2).End(xlUp).Row

    ' 画面クリア
    If MR1 >= 2 Then ws1.Range("A2:H" & MR1).Clear

    hc_cnt = 1
    For i = 1 To MR2
        If ws2.Cells(i, 1).Value = "H" Then
            HC(hc_cnt) = i
            hc_cnt = hc_cnt + 1
        End If
    Next

    hc_cnt_max = hc_cnt
    HC(hc_cnt_max) = MR2

    hc_cnt = 1

    ' 【修正】hc_cnt_max は「Hレコード数+1」のためループは -1 まで
    '         On Error GoTo errlab_extract を廃止し、エラーは MainErrorHandler へ伝播させる
    For i = 1 To hc_cnt_max - 1
        JOB_C = ws2.Cells(HC(i), 5).Value
        S_name = ws2.Cells(HC(i), 6).Value

        Application.StatusBar = "Step 5/" & TOTAL_STEPS & " webTCデータ抽出中：" & _
                                S_name & " " & i & "/" & (hc_cnt_max - 1) & "件"

        ' 名簿参照用の数式セット
        ws1.Cells(hc_cnt + 1, 1).Value = "=+SUMIF(名簿!A:A,C" & Format(hc_cnt + 1) & ",名簿!B:B)"

        ws1.Cells(hc_cnt + 1, 2).Value = JOB_C
        ws1.Cells(hc_cnt + 1, 3).Value = S_name

        n1 = HC(hc_cnt) + 1
        n2 = HC(hc_cnt + 1) - 1

        ws1.Cells(hc_cnt + 1, 4).Value = Application.WorksheetFunction.Sum(ws2.Range("J" & n1 & ":J" & n2)) * 24
        ws1.Cells(hc_cnt + 1, 4).NumberFormatLocal = "0.00"

        ws1.Cells(hc_cnt + 1, 5).Value = Application.WorksheetFunction.Sum(ws2.Range("M" & n1 & ":M" & n2)) * 24
        ws1.Cells(hc_cnt + 1, 5).NumberFormatLocal = "0.00"

        ' 【修正】全期間（n1:n2）のJ列合計がゼロのとき未承認と判定（旧: 後半 n3:n2 のみチェック）
        '         旧ロジックでは勤務日が1日のみのスタッフが常に「未承認」になるバグがあった
        hantei = Application.WorksheetFunction.Sum(ws2.Range("J" & n1 & ":J" & n2)) * 24
        If hantei = 0 Then
            ws1.Cells(hc_cnt + 1, 8).Value = "未承認"
        End If

        ws1.Cells(hc_cnt + 1, 6).Value = Application.WorksheetFunction.Sum(ws2.Range("O" & n1 & ":O" & n2)) * 24
        ws1.Cells(hc_cnt + 1, 6).NumberFormatLocal = "0.00"

        ws1.Cells(hc_cnt + 1, 7).Value = Application.WorksheetFunction.Sum(ws2.Range("N" & n1 & ":N" & n2)) * 24
        ws1.Cells(hc_cnt + 1, 7).NumberFormatLocal = "0.00"

        hc_cnt = hc_cnt + 1
    Next

    ' ループ正常終了後の後処理（旧: errlab_extract ラベルへの Error Jump で実行していた）
    ws1.Cells(1, 8).Value = ws2.Cells(2, 2).Value

    MR11 = ws1.Cells(ws1.Rows.Count, 2).End(xlUp).Row
    If MR11 >= 2 Then
        ' 数式を計算してから値に変換（手動計算モードのため明示的に計算）
        ws1.Calculate
        ws1.Range("A2:A" & MR11).Value = ws1.Range("A2:A" & MR11).Value
    End If
End Sub

Private Sub Phase2_スタッフコード()
    ' e-staffingの最新情報シートと照合してスタッフコードを更新
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim MR1 As Long, MR2 As Long
    Dim i As Long, j As Long

    Set ws1 = p_wsMain              ' 勤務データ
    Set ws2 = Worksheets("e-staffing TCnmhtの最新情報")

    MR1 = ws1.Cells(Rows.Count, 1).End(xlUp).Row
    MR2 = ws2.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To MR1
        Application.StatusBar = "Step 6/" & TOTAL_STEPS & " スタッフコード更新中：" & i & "/" & MR1
        For j = 2 To MR2
            If ws1.Cells(i, 1).Value = ws2.Cells(j, 22).Value Then
                ws1.Cells(i, 3).Value = ws2.Cells(j, 21).Value
            End If
        Next
    Next
End Sub


' ==============================================================================
' Phase 3: 請求処理パイプライン（ファイルBのマクロを統合）
' ==============================================================================

Private Sub Step1_立替金集計()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet

    Dim targetSheetName As String
    targetSheetName = "立替金集計"

    WriteLog "[S1] 開始"

    On Error Resume Next
    Set ws2 = Worksheets(targetSheetName)
    On Error GoTo 0
    If ws2 Is Nothing Then
        Set ws2 = Worksheets.Add
        ws2.Name = targetSheetName
        WriteLog "[S1] 立替金集計シート新規作成"
    End If

    Set ws1 = Worksheets("立替金データ")
    WriteLog "[S1] 立替金データシート取得OK"

    ' 立替金データの先頭セルをダンプ（データ構造確認用）
    WriteLog "[S1] 立替金データ確認: A1=" & Nz(ws1.Cells(1, 1).Value) & " B1=" & Nz(ws1.Cells(1, 2).Value) & _
             " C1=" & Nz(ws1.Cells(1, 3).Value) & " D1=" & Nz(ws1.Cells(1, 4).Value) & " E1=" & Nz(ws1.Cells(1, 5).Value)
    WriteLog "[S1] 立替金データ確認: A2=" & Nz(ws1.Cells(2, 1).Value) & " B2=" & Nz(ws1.Cells(2, 2).Value) & _
             " D2=" & Nz(ws1.Cells(2, 4).Value) & " F2=" & Nz(ws1.Cells(2, 6).Value) & " O2(15列)=" & Nz(ws1.Cells(2, 15).Value)

    Dim MR0 As Long, MR1 As Long, MR2 As Long
    Dim i As Long, j As Long

    MR0 = ws2.Cells(Rows.Count, 2).End(xlUp).Row
    If MR0 > 1 Then ws2.Range("A2:F" & MR0).Clear
    WriteLog "[S1] 集計シートクリア完了 MR0=" & MR0

    MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    WriteLog "[S1] MR1(立替金データ最終行)=" & MR1

    ' 【修正】Dictionary を使い全重複（非連続含む）を正しく除去して書き込む
    '         旧コードは隣接行の連続重複しか削除できず、入力順次第で結果が変わるバグがあった
    Dim dictUniq As Object
    Set dictUniq = CreateObject("Scripting.Dictionary")
    Dim writeRow As Long
    writeRow = 2
    For i = 1 To MR1
        Dim empKey As String
        empKey = CStr(ws1.Cells(i, 4).Value)
        If empKey <> "" And Not dictUniq.Exists(empKey) Then
            dictUniq(empKey) = True
            ws2.Cells(writeRow, 2).Value = ws1.Cells(i, 4).Value  ' 社員番号
            ws2.Cells(writeRow, 3).Value = ws1.Cells(i, 5).Value  ' 氏名
            writeRow = writeRow + 1
        End If
    Next i
    Set dictUniq = Nothing
    WriteLog "[S1] 社員番号・名前コピー＆重複削除完了（一意件数=" & (writeRow - 2) & "件）"

    MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    MR2 = ws2.Cells(Rows.Count, 2).End(xlUp).Row
    WriteLog "[S1] 集計ループ開始 MR1=" & MR1 & " MR2=" & MR2

    For i = 2 To MR2
        WriteLog "[S1] i=" & i & " ws2.B=" & Nz(ws2.Cells(i, 2).Value) & " ws2.C=" & Nz(ws2.Cells(i, 3).Value)

        WriteLog "[S1] SumIf実行前 (i=" & i & ")"
        ws2.Cells(i, 4).Value = Application.WorksheetFunction.SumIf(ws1.Columns(4), ws2.Cells(i, 2), ws1.Columns(15))
        ws2.Cells(i, 4).NumberFormatLocal = "##,##"
        WriteLog "[S1] SumIf結果 D=" & Nz(ws2.Cells(i, 4).Value)

        Dim tempVal As Double
        tempVal = 0
        WriteLog "[S1] 顧客対応当番ループ開始 (j=1 to " & MR1 & ")"
        For j = 1 To MR1
            Dim instrVal As Variant
            instrVal = ws1.Cells(j, 12).Value
            If VarType(instrVal) = vbString Then
                If InStr(CStr(instrVal), "顧客対応当番") > 0 Then
                    If ws1.Cells(j, 4).Value = ws2.Cells(i, 2).Value Then
                        Dim addVal As Variant
                        addVal = ws1.Cells(j, 15).Value
                        WriteLog "[S1] 当番加算 j=" & j & " val=" & Nz(addVal) & " Type=" & TypeName(addVal)
                        tempVal = tempVal + CDbl(addVal)
                    End If
                End If
            End If
        Next
        ws2.Cells(i, 5).Value = tempVal
        ws2.Cells(i, 5).NumberFormatLocal = "##,##"
        WriteLog "[S1] 当番手当=" & tempVal

        ws2.Cells(i, 6).Value = ws2.Cells(i, 4).Value - ws2.Cells(i, 5).Value
        ws2.Cells(i, 6).NumberFormatLocal = "##,##"
        WriteLog "[S1] 交通費=" & Nz(ws2.Cells(i, 6).Value)

        On Error Resume Next
        ws2.Cells(i, 1).Value = Application.WorksheetFunction.SumIf(Worksheets("名簿").Columns(1), ws2.Cells(i, 3), Worksheets("名簿").Columns(2))
        If Err.Number <> 0 Then WriteLog "[S1] 名簿SumIf失敗 i=" & i & " Err=" & Err.Description
        On Error GoTo 0
    Next
    WriteLog "[S1] 集計ループ完了"

    WriteLog "[S1] 月表示セット前: ws1(2,6)=" & Nz(ws1.Cells(2, 6).Value) & " Type=" & TypeName(ws1.Cells(2, 6).Value)
    If MR1 > 0 Then
        ws2.Cells(1, 7).Value = Month(ws1.Cells(2, 6).Value) & "月"
    End If
    WriteLog "[S1] 完了"
End Sub

Private Sub Step2_立替データ保存()
    Dim ws1 As Worksheet
    Dim ws3 As Worksheet

    Set ws1 = Worksheets("立替金集計")
    Set ws3 = p_wsMain

    Dim MR1 As Long, MR2 As Long
    Dim i As Long, j As Long

    MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    MR2 = ws3.Cells(Rows.Count, 2).End(xlUp).Row

    ws3.Range("I2:K" & MR2).ClearContents

    For i = 2 To MR2
        For j = 2 To MR1
            If ws3.Cells(i, 1).Value = ws1.Cells(j, 1).Value Then
                ws3.Cells(i, 9).Value = ws1.Cells(j, 5).Value
                ws3.Cells(i, 9).NumberFormatLocal = "##,##"
                ws3.Cells(i, 10).Value = ws1.Cells(j, 6).Value
                ws3.Cells(i, 10).NumberFormatLocal = "##,##"
                Exit For
            End If
        Next j
    Next i
End Sub

Private Sub Step3_請求明細作成()
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim ws1 As Worksheet
    Set ws1 = p_wsMain
    Set ws2 = Worksheets("請求明細")
    Set ws3 = Worksheets("e-staffing TCnmhtの最新情報")

    Dim MR1 As Long, MR2 As Long, MR3 As Long
    Dim i As Long, j As Long

    WriteLog "[Step3] 開始"

    ' 【修正】Application.StandardFont / StandardFontSize の変更を削除
    '         請求処理マクロがExcel全体のフォント設定を永続的に変更するのは副作用であり不適切

    MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    MR2 = ws2.Cells(Rows.Count, 7).End(xlUp).Row
    WriteLog "[Step3] MR1(勤務データ最終行)=" & MR1 & "  MR2(請求明細最終行)=" & MR2

    ' --- ブロック1: 勤務データ L-O列に数式セット ---
    WriteLog "[Step3-1] 勤務データ L-O列 数式セット開始"
    ws1.Range("L2").Value = "=+SUMIF('e-staffing TCnmhtの最新情報'!M:M,B2,'e-staffing TCnmhtの最新情報'!BC:BC)"
    ws1.Range("M2").Value = "=+SUMIFS(請求内訳!H:H,請求内訳!C:C,A2,請求内訳!J:J,""基本料金"")"
    ws1.Range("N2").Value = "=+(L2=1)*(ROUNDDOWN(E2*2,0)/2+ROUNDDOWN(F2*2,0)/2)+(L2=3)*MAX(0,D2-M2)"
    ws1.Range("O2").Value = "=+(L2=1)*(D2-ROUNDDOWN(E2*2,0)/2-ROUNDDOWN(f2*2,0)/2)+(L2=3)*MIN(D2,M2)"
    WriteLog "[Step3-1] 数式セット完了"

    If MR1 >= 2 Then
        WriteLog "[Step3-1] 数式展開＆計算（L2:O" & MR1 & "）"
        ws1.Range("L2:O2").Copy
        ws1.Range("L2:O" & MR1).PasteSpecial xlPasteAll
        ws1.Calculate
        ws1.Range("L2:O" & MR1).Copy
        ws1.Range("L2:O" & MR1).PasteSpecial xlPasteValues
        WriteLog "[Step3-1] 値貼り付け完了"
    End If

    ' --- ブロック2: RoundDown補正 ---
    WriteLog "[Step3-2] RoundDown補正ループ開始（i=2 to " & MR1 & "）"
    For i = 2 To MR1
        If ws1.Cells(i, 12).Value = 1 Then
            Dim rdSrc As Variant
            rdSrc = ws1.Cells(i, 5).Value
            WriteLog "[Step3-2] i=" & i & " L列=" & Nz(ws1.Cells(i, 12).Value) & " E列=" & Nz(rdSrc) & " Type=" & TypeName(rdSrc)
            ws1.Cells(i, 14).Value = WorksheetFunction.RoundDown(rdSrc * 2, 0) / 2
        End If
    Next
    WriteLog "[Step3-2] RoundDown補正完了"

    ' --- ブロック3: 請求明細クリア＆氏名コピー ---
    WriteLog "[Step3-3] 請求明細クリア＆コピー開始"
    If MR2 >= 7 Then ws2.Range("A7:AT" & MR2 + 100).ClearContents

    ws1.Range("C2:C" & MR1).Copy
    ws2.Range("G7").PasteSpecial xlPasteAll
    ws1.Range("B2:B" & MR1).Copy
    ws2.Range("F7").PasteSpecial xlPasteAll
    ws2.Range("H7").PasteSpecial xlPasteAll
    WriteLog "[Step3-3] 氏名コピー完了"

    MR2 = ws2.Cells(Rows.Count, 7).End(xlUp).Row
    MR3 = ws3.Cells(Rows.Count, 1).End(xlUp).Row
    WriteLog "[Step3-3] 更新後 MR2=" & MR2 & "  MR3(e-staffing最終行)=" & MR3

    ' --- ブロック4: 契約番号書き込み（Match）---
    WriteLog "[Step3-4] 契約番号Matchループ開始（i=7 to " & MR2 & "）"
    For i = 7 To MR2
        Dim jc As Variant
        Dim ret As Variant
        jc = ws2.Cells(i, 6).Value
        On Error Resume Next
        ret = WorksheetFunction.Match(jc, ws3.Range("M1:M" & MR3), 0)
        If Err.Number = 0 Then
            ws2.Cells(i, 2).Value = ws3.Cells(ret, 1).Value
        Else
            WriteLog "[Step3-4] Match失敗 i=" & i & " jc=" & Nz(jc) & " Err=" & Err.Description
        End If
        On Error GoTo 0
    Next
    WriteLog "[Step3-4] 契約番号書き込み完了"

    ' --- ブロック5: VLOOKUP＆ual-nmht絞り込み ---
    WriteLog "[Step3-5] VLOOKUP＆ual-nmht絞り込み開始"
    Application.Calculation = xlCalculationAutomatic
    ws2.Range("C7").NumberFormatLocal = "G/標準"
    ws2.Range("C7").Value = "=VLOOKUP(B7,'e-staffing TCnmhtの最新情報'!A:E,5,0)"

    If MR2 > 7 Then
        ws2.Range("C7").Copy
        ws2.Range("C8:C" & MR2).PasteSpecial xlPasteFormulas
    End If
    ws2.Range("C7:C" & MR2).Copy
    ws2.Range("C7:C" & MR2).PasteSpecial xlPasteValues

    Dim lRow As Long
    lRow = ws2.Cells(Rows.Count, 3).End(xlUp).Row
    WriteLog "[Step3-5] 絞り込み前 lRow=" & lRow
    Dim delCount As Long
    delCount = 0
    For j = lRow To 7 Step -1
        Dim cellVal As Variant
        cellVal = ws2.Cells(j, 3).Value
        ' セルがエラー値（#N/A等）の場合、または"ual-nmht"でない場合は削除
        If IsError(cellVal) Then
            WriteLog "[Step3-5] 削除(エラー値) j=" & j & " CellError"
            ws2.Rows(j).Delete
            delCount = delCount + 1
        ElseIf CStr(cellVal) <> "ual-nmht" Then
            ws2.Rows(j).Delete
            delCount = delCount + 1
        End If
    Next j
    WriteLog "[Step3-5] 絞り込み完了 削除行数=" & delCount

    Application.Calculation = xlCalculationManual

    ' --- ブロック6: データ書込み（BHM～日付系）---
    MR2 = ws2.Cells(Rows.Count, 3).End(xlUp).Row
    WriteLog "[Step3-6] データ書込み開始 MR2=" & MR2

    ' K1, H1 の値を事前ログ出力（Format呼出しで型不一致になりやすい箇所）
    Dim valK1 As Variant, valH1 As Variant
    valK1 = ws1.Range("K1").Value
    valH1 = ws1.Range("H1").Value
    WriteLog "[Step3-6] K1=" & Nz(valK1) & " Type=" & TypeName(valK1)
    WriteLog "[Step3-6] H1=" & Nz(valH1) & " Type=" & TypeName(valH1)

    For i = 7 To MR2
        WriteLog "[Step3-6] 行" & i & " F列=" & Nz(ws2.Cells(i, 6).Value)

        ws2.Cells(i, 3).Value = "BHM" & Format(valK1, "yyyymmdd") & ws2.Cells(i, 6)
        ws2.Cells(i, 4).Value = 1
        ws2.Cells(i, 9).NumberFormatLocal = "@"
        ws2.Cells(i, 9).Value = Format(valK1, "yyyy") & "/" & Format(valK1, "mm")
        ws2.Cells(i, 10).Value = Format(valH1, "yyyy/mm/dd")
        ws2.Cells(i, 11).Value = Format(valK1, "yyyy/mm/dd")
    Next
    WriteLog "[Step3-6] データ書込み完了"

    ' --- ブロック7: 数式セット（大量）---
    WriteLog "[Step3-7] 請求計算数式セット開始"
    ws2.Cells(7, 5).Value = "=+VLOOKUP(G7,名簿!A:B,2,0)"
    ws2.Cells(7, 12).Value = "=+SUMIF('e-staffing TCnmhtの最新情報'!A:A,B7,'e-staffing TCnmhtの最新情報'!BB:BB)"
    ws2.Cells(7, 13).Value = "=+SUMIF('e-staffing TCnmhtの最新情報'!A:A,B7,'e-staffing TCnmhtの最新情報'!BC:BC)"
    ws2.Cells(7, 14).Value = "=+ROUNDDOWN(SUMIF(勤務データ!B:B,H7,勤務データ!o:o),0)"
    ws2.Cells(7, 15).Value = "=+ROUNDDOWN((SUMIF(勤務データ!B:B,H7,勤務データ!O:O)-N7+.001)*2,0)*30"
    ws2.Cells(7, 16).Value = "=+IF(M7=1,L7*(N7+O7/60),IF(M7=3,L7,""error!""))"

    ws2.Cells(7, 17).Value = "=+ROUNDDOWN(SUMIF(勤務データ!$B:$B,請求明細!F7,勤務データ!$N:$N),0)"
    ws2.Cells(7, 18).Value = "=+ROUNDDOWN((SUMIF(勤務データ!$B:$B,請求明細!F7,勤務データ!$N:$N)-Q7+0.00001)*2,0)*30"
    ws2.Cells(7, 19).Value = "=+rounddown((SUMIFS(請求内訳!K:K,請求内訳!C:C,E7,請求内訳!J:J,""超過単金"")+SUMIFS(請求内訳!K:K,請求内訳!C:C,E7,請求内訳!J:J,""時間外単金""))*(Q7+R7/60),0)"

    ws2.Cells(7, 20).Value = "=+ROUNDDOWN(SUMIF(勤務データ!$B:$B,請求明細!F7,勤務データ!$F:$F),0)"
    ws2.Cells(7, 21).Value = "=+ROUNDDOWN((SUMIF(勤務データ!$B:$B,請求明細!F7,勤務データ!$F:$F)-T7)*2,0)*30"
    ws2.Cells(7, 22).Value = "=+rounddown((SUMIFS(請求内訳!K:K,請求内訳!C:C,E7,請求内訳!J:J,""法定休日単金""))*(T7+U7/60),0)"

    ws2.Cells(7, 23).Value = "=+Q7+T7"
    ws2.Cells(7, 24).Value = "=+U7+R7"
    ws2.Cells(7, 25).Value = "=+S7+V7"

    ws2.Cells(7, 26).Value = "=+ROUNDDOWN(SUMIF(勤務データ!$B:$B,請求明細!F7,勤務データ!$G:$G),0)"
    ws2.Cells(7, 27).Value = "=+ROUNDDOWN((SUMIF(勤務データ!$B:$B,請求明細!F7,勤務データ!$G:$G)-AA7)*2,0)*30"
    ws2.Cells(7, 28).Value = "=+rounddown((SUMIFS(請求内訳!K:K,請求内訳!C:C,E7,請求内訳!J:J,""深夜割増単金""))*(Z7+AA7/60),0)"

    ws2.Cells(7, 32).Value = "=+IF(M7=3,SUMIF(勤務データ!$B:$B,請求明細!F7,勤務データ!$D:$D)-SUMIFS(請求内訳!I:I,請求内訳!C:C,E7,請求内訳!J:J,""基本料金""),0)"
    ws2.Cells(7, 29).Value = "=+IF(AF7<0,ROUNDDOWN(ABS(AF7),0),0)"
    ws2.Cells(7, 30).Value = "=+ABS(ROUNDdown(IF(AF7<0,(AF7+AC7))*2,0)*30)"
    ws2.Cells(7, 31).Value = "=+rounddown((SUMIFS(請求内訳!K:K,請求内訳!C:C,E7,請求内訳!J:J,""控除単金""))*(AC7+AD7/60),0)"
    ws2.Cells(7, 34).Value = "=+AE7+AG7"

    ws2.Cells(7, 40).Value = "=+SUMIF(勤務データ!$B:$B,請求明細!F7,勤務データ!$I:$I)+SUMIF(勤務データ!$B:$B,請求明細!F7,勤務データ!$J:$J)+SUMIF(勤務データ!$B:$B,請求明細!F7,勤務データ!$k:$K)"
    WriteLog "[Step3-7] 数式セット完了"

    ' --- ブロック8: 数式展開＆計算 ---
    WriteLog "[Step3-8] 数式展開開始"
    If MR2 > 7 Then
        ws2.Range("E7").Copy ws2.Range("E8:E" & MR2)
        ws2.Range("L7:AN7").Copy ws2.Range("L8:AN" & MR2)
    End If

    WriteLog "[Step3-8] Calculate実行"
    ws2.Calculate

    ws2.Range("B7:AQ7").Copy
    ws2.Range("B8:AQ" & MR2).PasteSpecial xlPasteFormats
    WriteLog "[Step3-8] 書式コピー完了"

    ' --- ブロック9: 月額制修正 ---
    WriteLog "[Step3-9] 月額制修正ループ開始"
    Dim k As Long
    For k = 7 To MR2
        If ws2.Cells(k, 13).Value = 3 Then
            WriteLog "[Step3-9] 月額制修正 行=" & k
            ws2.Cells(k, 20).Value = 0
            ws2.Cells(k, 21).Value = 0
            ws2.Cells(k, 26).Value = 0
            ws2.Cells(k, 27).Value = 0
        End If
    Next
    ws2.Calculate
    WriteLog "[Step3-9] 月額制修正完了"

    ' --- ブロック10: 値貼り付け＆時間繰上 ---
    WriteLog "[Step3-10] 値貼り付け開始"
    ws2.Range("B7:AQ" & MR2).Copy
    ws2.Range("B7:AQ" & MR2).PasteSpecial xlValues
    ws2.Range("W7:X" & MR2).Copy
    ws2.Range("W7:X" & MR2).PasteSpecial xlValues
    WriteLog "[Step3-10] 値貼り付け完了"

    WriteLog "[Step3-10] 時間繰上チェック開始"
    For k = 7 To MR2
        Dim xVal As Variant
        xVal = ws2.Cells(k, 24).Value
        If IsNumeric(xVal) Then
            If CDbl(xVal) > 30 Then
                WriteLog "[Step3-10] 繰上 行=" & k & " X列=" & Nz(xVal)
                ws2.Cells(k, 23).Value = ws2.Cells(k, 23).Value + 1
                ws2.Cells(k, 24).Value = 0
            End If
        Else
            WriteLog "[Step3-10] 警告: X列が非数値 行=" & k & " Val=" & Nz(xVal) & " Type=" & TypeName(xVal)
        End If
    Next
    WriteLog "[Step3-10] 時間繰上完了"

    ' --- ブロック11: 請求金額合計 ---
    WriteLog "[Step3-11] 請求金額合計開始"
    ws2.Cells(7, 39).NumberFormatLocal = "0_ "
    ws2.Cells(7, 39).Font.Size = 11
    ws2.Cells(7, 39).Value = "=+P7+Y7+AB7-AH7+AI7+AL7"
    If MR2 > 7 Then
        ws2.Range("AM7").Copy ws2.Range("AM8:AM" & MR2)
    End If
    ws2.Calculate
    ws2.Range("AM7:AM" & MR2).Copy
    ws2.Range("AM7:AM" & MR2).PasteSpecial xlPasteValues
    WriteLog "[Step3-11] 請求金額合計完了"

    ' --- ブロック12: 未使用欄クリア ---
    WriteLog "[Step3-12] 未使用欄クリア"
    ws2.Range("E7:E" & MR2).Value = "nmht"
    ' 【修正】AF列を単独でクリアする行を削除（次行の AF:AG クリアで十分）
    ws2.Range("AF7:AG" & MR2).Value = 0
    ws2.Range("AI7:AL" & MR2).Value = 0
    ws2.Range("AP7:AQ" & MR2).Value = 0
    WriteLog "[Step3] 完了"
End Sub

Private Sub Step4_請求ヘッダ作成()
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Set ws2 = Worksheets("請求明細")
    Set ws3 = Worksheets("請求ヘッダテンプレート (請求コメントあり)")

    Dim MR2 As Long, MR3 As Long
    Dim i As Long

    WriteLog "[Step4] 開始"

    MR3 = ws3.Cells(Rows.Count, 7).End(xlUp).Row
    WriteLog "[Step4] 請求ヘッダテンプレート既存最終行=" & MR3

    If MR3 >= 7 Then
        ws3.Range("B7:BG" & MR3).Delete
        WriteLog "[Step4] 既存データ削除完了"
    End If

    ws3.Cells.NumberFormatLocal = "G/標準"

    ' ★★★ 修正箇所：INDIRECT関数とROW関数を使用して相対参照を実現 ★★★
    WriteLog "[Step4] 数式セット開始（INDIRECT+ROW版）"

    ws3.Range("B7").Value = "請H_2"
    ws3.Range("C7").Value = 1
    ws3.Range("D7").Value = "=+INDIRECT(""請求明細!C"" & ROW())"
    ws3.Range("E7").Value = "=+INDIRECT(""請求明細!I"" & ROW())"
    ws3.Range("F7").Value = "ual-nmht"
    ws3.Range("G7").Value = "=+INDIRECT(""請求明細!C"" & ROW())"
    ws3.Range("H7").Value = "=""【"" & VLOOKUP(INDIRECT(""請求明細!F"" & ROW()),'e-staffing TCnmhtの最新情報'!U:V,2,0) &""】""& VLOOKUP(INDIRECT(""請求明細!F"" & ROW()),'e-staffing TCnmhtの最新情報'!U:CC,61,0)"
    ws3.Range("I7").Value = "=+INDIRECT(""請求明細!J"" & ROW())"
    ws3.Range("J7").Value = "=+INDIRECT(""請求明細!K"" & ROW())"
    ws3.Range("K7").Value = "=+INDIRECT(""請求明細!K"" & ROW())"
    ws3.Range("L7").Value = "=+EOMONTH(K7,1)"
    ws3.Range("L7").NumberFormatLocal = "yyyy/mm/dd"
    ws3.Range("M7").Value = "=+INDIRECT(""請求明細!AM"" & ROW())"
    ws3.Range("P7").Value = "=+rounddown(M7*0.1,0)"
    ws3.Range("Q7").Value = "=+rounddown(M7*1.1,0)"
    ws3.Range("R7").Value = "=+INDIRECT(""請求明細!AN"" & ROW())"
    ws3.Range("S7").Value = "=+Q7+R7"
    ws3.Range("T7").Value = "りそな銀行"
    ws3.Range("U7").Value = "九段支店"
    ws3.Range("V7").Value = 1
    ws3.Range("W7").Value = "1416351"
    ws3.Range("AR7").Value = "管理部"
    ws3.Range("AS7").Value = "03-6222-9420"
    ws3.Range("AT7").Value = 1
    ws3.Range("BG7").Value = 0

    WriteLog "[Step4] 数式セット完了"

    MR2 = ws2.Cells(Rows.Count, 7).End(xlUp).Row
    WriteLog "[Step4] 請求明細の最終行=" & MR2

    ' 7行目の数式を8行目以降にコピー
    If MR2 > 7 Then
        WriteLog "[Step4] 数式コピー開始（B7:BG7 → B8:BG" & MR2 & "）"
        ws3.Range("B7:BG7").Copy
        ws3.Range("B8:BG" & MR2).PasteSpecial xlPasteAll
        WriteLog "[Step4] 数式コピー完了"
    End If

    ' 数式を計算して値に変換
    WriteLog "[Step4] Calculate実行"
    ws3.Calculate

    WriteLog "[Step4] 値貼り付け開始"
    ws3.Range("B7:BG" & MR2).Copy
    ws3.Range("B7:BG" & MR2).PasteSpecial xlPasteValues
    WriteLog "[Step4] 値貼り付け完了"

    ' テキスト置換処理
    WriteLog "[Step4] テキスト置換開始"
    For i = 7 To MR2
        ws3.Cells(i, 8).Value = Replace(ws3.Cells(i, 8).Value, "ＮＵＬシステム購買室", "ユニアデックス株式会社")
        ws3.Cells(i, 8).Value = Replace(ws3.Cells(i, 8).Value, "。", "")
        ws3.Cells(i, 8).Value = Replace(ws3.Cells(i, 8).Value, "．", "")
    Next
    WriteLog "[Step4] テキスト置換完了"

    WriteLog "[Step4] 完了（出力行数=" & (MR2 - 6) & "人分）"
End Sub

Private Sub Step5_請求データとの照合()
    Dim ws11 As Worksheet
    Dim ws12 As Worksheet
    Dim ws13 As Worksheet
    Dim wb2 As Workbook
    Dim ws21 As Worksheet

    Set ws11 = p_wsMain
    Set ws12 = Worksheets("請求内訳")
    Set ws13 = Worksheets("名簿")

    Dim OpenFileName As String
    OpenFileName = ws11.Range("M1").Value

    ' ファイルがなければスキップ
    If OpenFileName = "" Or Dir(OpenFileName) = "" Then
        Exit Sub
    End If

    Set wb2 = Workbooks.Open(OpenFileName)
    Set ws21 = wb2.Worksheets("総合計請求書")

    Dim MR11 As Long, MR21 As Long, MR12 As Long, MR13 As Long
    Dim K_Month As Variant, K_Col As Long, MC12 As Long
    Dim i As Long, j As Long

    MR11 = ws11.Cells(Rows.Count, 6).End(xlUp).Row
    If MR11 >= 2 Then ws11.Rows("2:" & MR11).Delete

    MR21 = ws21.Cells(Rows.Count, 1).End(xlUp).Row
    ws21.Range("A1:I" & MR21).Copy ws11.Range("B1")

    K_Month = ws11.Range("K1").Value
    MC12 = ws12.Cells(1, Columns.Count).End(xlToLeft).Column

    For i = 1 To MC12
        If ws12.Cells(1, i).Value = K_Month Then
            K_Col = i + 1
            Exit For
        End If
    Next

    ' 【修正】対象月が請求内訳に存在しない場合、K_Col=0 のままでセルアクセスするとエラーになる
    '         wb2 を閉じてから明確なエラーメッセージで処理を中断する
    If K_Col = 0 Then
        wb2.Close SaveChanges:=False
        Err.Raise 1003, , "請求内訳シートの1行目に対象月「" & K_Month & "」が見つかりません。" & vbCrLf & _
                          "「勤務データ」シートのK1セルと請求内訳の列ヘッダを確認してください。"
    End If

    MR11 = ws11.Cells(Rows.Count, 6).End(xlUp).Row
    MR12 = ws12.Cells(Rows.Count, 1).End(xlUp).Row
    MR13 = ws13.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To MR11
        For j = 2 To MR13
            If ws11.Cells(i, 4).Value = ws13.Cells(j, 1).Value Then
                ws11.Cells(i, 1).Value = ws13.Cells(j, 2).Value
            End If
        Next
    Next

    For i = 2 To MR11
        For j = 3 To MR12
            If ws11.Cells(i, 1).Value = ws12.Cells(j, 3).Value Then
                If ws12.Cells(j, 10).Value = "合計" Then
                    ws11.Cells(i, 11).Value = ws12.Cells(j, K_Col).Value
                    ws11.Cells(i, 11).NumberFormatLocal = "##,#"
                    ws11.Cells(i, 12).Value = "=J" & i & "-K" & i
                    ws11.Cells(i, 12).NumberFormatLocal = "##,#"
                End If
            End If
        Next
    Next

    ws11.Cells(MR11 + 1, 11).Value = "=sum(K2:K" & MR11 & ")"
    ws11.Cells(MR11 + 1, 11).NumberFormatLocal = "##,#"
    ws11.Cells(MR11 + 1, 12).Value = "=J" & MR11 + 1 & "-K" & MR11 + 1
    ws11.Cells(MR11 + 1, 12).NumberFormatLocal = "##,#"

    ws11.Range("A1:L" & MR11 + 1).Borders.LineStyle = xlContinuous

    wb2.Close SaveChanges:=False
End Sub

Private Sub Step6_データ保存()
    Dim ws1 As Worksheet
    Set ws1 = p_wsMain

    Dim S_name As String
    If IsDate(ws1.Cells(1, 8).Value) Then
        S_name = Month(ws1.Cells(1, 8).Value) & "月"
    Else
        S_name = "Backup"
    End If

    Dim ws As Worksheet, flag As Boolean
    flag = False
    For Each ws In Worksheets
        If ws.Name = S_name Then flag = True
    Next ws

    Dim MR1 As Long
    MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row

    If flag = True Then
        ws1.Range("A1:J" & MR1).Copy Worksheets(S_name).Range("A1")
    Else
        ws1.Copy After:=ws1
        ActiveSheet.Name = S_name
        p_wsMain.Activate
    End If
End Sub

Private Sub Step7_提出データファイル作成()
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim wb2 As Workbook
    Dim wb3 As Workbook

    Set ws2 = Worksheets("請求明細")
    Set ws3 = Worksheets("請求ヘッダテンプレート (請求コメントあり)")

    Dim Hd_file As String, M_file As String

    Hd_file = p_FolderPath & "\" & "BH" & p_DateStr & p_TimeStr & ".txt"
    M_file = p_FolderPath & "\" & "BD" & p_DateStr & p_TimeStr & ".txt"

    WriteLog "[Step7] 出力開始 BHファイル=" & Hd_file
    WriteLog "[Step7] 出力開始 BDファイル=" & M_file

    ' 明細出力
    ' 【修正】1回目のSaveAs（A列・1-6行削除前）を削除（直後に上書きされるため無意味）
    ws2.Copy
    Set wb2 = ActiveWorkbook
    wb2.Sheets(1).Cells.NumberFormatLocal = "@"
    wb2.Sheets(1).Columns(1).Delete
    wb2.Sheets(1).Rows("1:6").Delete
    wb2.SaveAs fileName:=M_file, FileFormat:=xlText
    wb2.Close SaveChanges:=False
    Call RemovePureBlankLines(M_file)
    WriteLog "[Step7] BDファイル出力完了"

    ' ヘッダ出力
    ' 【修正】1回目のSaveAs（A列・1-6行削除前）を削除（直後に上書きされるため無意味）
    ws3.Copy
    Set wb3 = ActiveWorkbook
    wb3.Sheets(1).Columns(1).Delete
    wb3.Sheets(1).Rows("1:6").Delete
    wb3.SaveAs fileName:=Hd_file, FileFormat:=xlText
    wb3.Close SaveChanges:=False
    WriteLog "[Step7] BHファイル出力完了"
End Sub



Private Sub RemovePureBlankLines(filePath As String)
    On Error GoTo ExitPoint
    Dim f As Integer
    Dim txt As String
    Dim lines() As String
    Dim outText As String
    Dim i As Long

    If Dir(filePath) = "" Then Exit Sub

    f = FreeFile
    Open filePath For Input As #f
    txt = Input$(LOF(f), #f)
    Close #f
    f = 0

    lines = Split(txt, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If Len(lines(i)) > 0 Then
            If outText <> "" Then outText = outText & vbCrLf
            outText = outText & lines(i)
        End If
    Next i

    f = FreeFile
    Open filePath For Output As #f
    Print #f, outText;
    Close #f
    f = 0

ExitPoint:
    On Error Resume Next
    If f > 0 Then Close #f
End Sub

Private Sub Step8_ExportTimecardData()
    Dim wsEStaffing As Worksheet
    Dim wsWebTC As Worksheet
    Dim fileName As String
    Dim outputPath As String
    Dim fileNum As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim contractNo As String
    Dim contractNos As Collection
    Dim outputLines As Collection

    Set wsEStaffing = ThisWorkbook.Worksheets("e-staffing TCnmhtの最新情報")
    Set wsWebTC = ThisWorkbook.Worksheets("webTC_data")

    Set contractNos = New Collection
    lastRow = wsWebTC.Cells(wsWebTC.Rows.Count, 1).End(xlUp).Row

    For i = 1 To lastRow
        If wsWebTC.Cells(i, 1).Value = "H" Then
            contractNo = wsWebTC.Cells(i, 2).Value
            On Error Resume Next
            contractNos.Add contractNo, CStr(contractNo)
            On Error GoTo 0
        End If
    Next i

    If contractNos.Count = 0 Then Err.Raise 2001, , "webTC_dataに処理対象の契約Noが見つかりません。"

    Set outputLines = New Collection
    Dim contractItem As Variant
    Dim outputLine As String
    Dim tcCounter As Long
    tcCounter = 1

    For Each contractItem In contractNos
        outputLine = ProcessContract_Internal(CStr(contractItem), wsEStaffing, wsWebTC, tcCounter)
        If outputLine <> "" Then
            outputLines.Add outputLine
        End If
    Next contractItem

    If outputLines.Count = 0 Then Err.Raise 2002, , "BTファイルの出力データがありません。"

    fileName = "BT" & p_DateStr & p_TimeStr & ".txt"
    outputPath = p_FolderPath & "\" & fileName

    fileNum = FreeFile
    Open outputPath For Output As #fileNum
    Dim line As Variant
    For Each line In outputLines
        Print #fileNum, line
    Next line
    Close #fileNum

    ' --- 「勤怠 (就業場所区分あり)」シートへの反映 ---
    Dim wsKintai As Worksheet
    On Error Resume Next
    Set wsKintai = ThisWorkbook.Worksheets("勤怠 (就業場所区分あり)")
    On Error GoTo 0

    If Not wsKintai Is Nothing Then
        ' 既存データクリア（B7以降）
        Dim lastRowKintai As Long
        lastRowKintai = wsKintai.Cells(wsKintai.Rows.Count, 2).End(xlUp).Row
        If lastRowKintai >= 7 Then
            wsKintai.Range("B7:QS" & lastRowKintai).ClearContents
        End If

        ' BTファイルと同じデータをB7から書き込み
        Dim rowIdx As Long
        rowIdx = 7
        Dim lineItem As Variant
        For Each lineItem In outputLines
            Dim parts() As String
            parts = Split(CStr(lineItem), vbTab)
            Dim colIdx As Long
            For colIdx = 0 To UBound(parts)
                wsKintai.Cells(rowIdx, 2 + colIdx).Value = parts(colIdx)
            Next colIdx
            rowIdx = rowIdx + 1
        Next lineItem
    End If
End Sub

Private Function ProcessContract_Internal(contractNo As String, wsEStaffing As Worksheet, wsWebTC As Worksheet, tcNumber As Long) As String
    Dim fields() As String
    ReDim fields(1 To 500)
    Dim i As Long
    For i = 1 To 500
        fields(i) = ""
    Next i

    fields(2) = contractNo

    Dim contractRow As Long
    contractRow = FindContractRow_Internal(wsEStaffing, contractNo)
    If contractRow = 0 Then Err.Raise 2003, , "契約No " & contractNo & " がe-staffingシートに見つかりません。"

    Dim cStart As Variant, cEnd As Variant, cBreak As Variant
    cStart = wsEStaffing.Cells(contractRow, 46).Value
    cEnd = wsEStaffing.Cells(contractRow, 47).Value
    cBreak = wsEStaffing.Cells(contractRow, 49).Value

    If IsEmpty(cStart) Or IsEmpty(cEnd) Or IsEmpty(cBreak) Then Err.Raise 2004, , contractNo & ": 勤務時間データが空白です。"

    Dim startTimeValue As Date, endTimeValue As Date
    startTimeValue = IIf(IsDate(cStart), CDate(cStart), CDbl(cStart))
    endTimeValue = IIf(IsDate(cEnd), CDate(cEnd), CDbl(cEnd))

    Dim contractWorkMin As Long
    contractWorkMin = DateDiff("n", startTimeValue, endTimeValue) - CLng(cBreak)

    Dim jobCode As String
    Dim dataRows As Collection
    Set dataRows = New Collection
    Dim lastRow As Long
    lastRow = wsWebTC.Cells(wsWebTC.Rows.Count, 1).End(xlUp).Row
    Dim j As Long, foundH As Boolean
    foundH = False

    ' 祝日表シートから祝日辞書を構築（カテゴリ空白行の自動補完に使用）
    Dim holidayDict As Object
    Set holidayDict = BuildHolidayDict()

    For j = 1 To lastRow
        If wsWebTC.Cells(j, 1).Value = "H" And wsWebTC.Cells(j, 2).Value = contractNo Then
            jobCode = wsWebTC.Cells(j, 5).Value
            foundH = True
        ElseIf wsWebTC.Cells(j, 1).Value = "D" And foundH Then
            Dim rowCategory As String
            rowCategory = CStr(wsWebTC.Cells(j, 3).Value)

            ' カテゴリ空白の場合、祝日表と照合して自動補完
            If rowCategory = "" Then
                Dim rowDateVal As Variant
                rowDateVal = wsWebTC.Cells(j, 2).Value
                If IsDate(rowDateVal) Then
                    If holidayDict.Exists(Format(CDate(rowDateVal), "yyyy/mm/dd")) Then
                        rowCategory = "2"  ' 祝日表に存在する日 → 休日扱い
                    End If
                End If
            End If

            If rowCategory <> "" Then
                Dim dataRow As Object
                Set dataRow = CreateObject("Scripting.Dictionary")
                dataRow("Row") = j
                dataRow("Date") = wsWebTC.Cells(j, 2).Value
                dataRow("Category") = rowCategory
                dataRow("StartTime") = wsWebTC.Cells(j, 4).Value
                dataRow("EndTime") = wsWebTC.Cells(j, 5).Value
                dataRow("BreakTime") = wsWebTC.Cells(j, 6).Value
                dataRows.Add dataRow
            End If
        ElseIf wsWebTC.Cells(j, 1).Value = "H" And wsWebTC.Cells(j, 2).Value <> contractNo And foundH Then
            Exit For
        End If
    Next j

    If dataRows.Count = 0 Then Err.Raise 2005, , contractNo & ": webTC_dataにデータがありません。"

    fields(8) = jobCode

    ' --- 請求書コード: BH/BDファイルと同じコード体系で生成 ---
    fields(3) = "BHM" & Format(p_wsMain.Range("K1").Value, "yyyymmdd") & jobCode

    ' --- 請求書明細コード: 契約No（BD明細との紐づけ） ---
    fields(4) = contractNo

    ' --- タイムカード番号 ---
    fields(5) = CStr(tcNumber)

    ' --- 企業ID: e-staffingシートから取得 ---
    fields(6) = CStr(wsEStaffing.Cells(contractRow, 5).Value)

    ' --- スタッフコード: e-staffingシートから取得 ---
    fields(7) = CStr(wsEStaffing.Cells(contractRow, 21).Value)

    Dim workDays As Long, absentDays As Long, holidayDays As Long
    Dim totalWorkMin As Long, totalContractMin As Long, totalOverMin As Long

    Dim targetDate As Date
    targetDate = CDate(dataRows(1)("Date"))
    fields(9) = Year(targetDate) & "/" & Format(Month(targetDate), "00")
    fields(10) = Format(DateSerial(Year(targetDate), Month(targetDate) + 1, 0), "yyyy/mm/dd")


    Dim daysInMonth As Long
    Dim d As Long
    Dim initBaseCol As Long
    daysInMonth = 31

    ' 勤怠取込仕様に合わせ、日付枠は31日分を固定で初期化（カテゴリ既定=休日2）
    For d = 1 To daysInMonth
        initBaseCol = 22 + (d - 1) * 14
        fields(initBaseCol) = Format(DateSerial(Year(targetDate), Month(targetDate), d), "yyyy/mm/dd")
        fields(initBaseCol + 1) = "2"
    Next d
    Dim dayData As Variant
    For Each dayData In dataRows

        Dim dayOfMonth As Long
        dayOfMonth = Day(CDate(dayData("Date")))
        If dayOfMonth < 1 Or dayOfMonth > 31 Then GoTo NextDay

        Dim category As String
        category = CStr(dayData("Category"))

        If category = "1" Then workDays = workDays + 1
        If category = "4" Then absentDays = absentDays + 1
        If category = "2" Then holidayDays = holidayDays + 1

        Dim baseCol As Long
        baseCol = 22 + (dayOfMonth - 1) * 14
        fields(baseCol) = Format(CDate(dayData("Date")), "yyyy/mm/dd")
        fields(baseCol + 1) = category

        If category = "1" Then
            Dim sVal As Variant, eVal As Variant, bVal As Variant
            sVal = dayData("StartTime")
            eVal = dayData("EndTime")
            bVal = dayData("BreakTime")

            If IsEmpty(sVal) Or IsEmpty(eVal) Or IsEmpty(bVal) Then GoTo NextDay

            Dim sTime As Date, eTime As Date, bTime As Date
            sTime = IIf(IsNumeric(sVal), CDbl(sVal), CDate(sVal))
            eTime = IIf(IsNumeric(eVal), CDbl(eVal), CDate(eVal))
            bTime = IIf(IsNumeric(bVal), CDbl(bVal), CDate(bVal))

            fields(baseCol + 2) = Hour(sTime)
            fields(baseCol + 3) = Minute(sTime)
            fields(baseCol + 4) = Hour(eTime)
            fields(baseCol + 5) = Minute(eTime)
            fields(baseCol + 6) = Hour(bTime)
            fields(baseCol + 7) = Minute(bTime)

            Dim actualWorkMin As Long, contractDayMin As Long, overDayMin As Long
            actualWorkMin = DateDiff("n", sTime, eTime) - (Hour(bTime) * 60 + Minute(bTime))

            If actualWorkMin <= contractWorkMin Then
                contractDayMin = actualWorkMin
                overDayMin = 0
            Else
                contractDayMin = contractWorkMin
                overDayMin = actualWorkMin - contractWorkMin
            End If

            fields(baseCol + 10) = contractDayMin \ 60
            fields(baseCol + 11) = contractDayMin Mod 60
            fields(baseCol + 12) = overDayMin \ 60
            fields(baseCol + 13) = overDayMin Mod 60

            totalWorkMin = totalWorkMin + actualWorkMin
            totalContractMin = totalContractMin + contractDayMin
            totalOverMin = totalOverMin + overDayMin
        End If
NextDay:
    Next dayData

    fields(11) = workDays
    fields(12) = absentDays
    fields(13) = holidayDays
    fields(14) = totalWorkMin \ 60
    fields(15) = totalWorkMin Mod 60
    fields(16) = totalContractMin \ 60
    fields(17) = totalContractMin Mod 60
    fields(18) = totalOverMin \ 60
    fields(19) = totalOverMin Mod 60

    Dim maxCol As Long
    maxCol = 22 + 31 * 14 - 1
    Dim output As String
    output = ""
    For i = 2 To maxCol
        If i > 2 Then output = output & vbTab
        output = output & fields(i)
    Next i

    ProcessContract_Internal = output
End Function

Private Function FindContractRow_Internal(ws As Worksheet, contractNo As String) As Long
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = contractNo Then
            FindContractRow_Internal = i
            Exit Function
        End If
    Next i
    FindContractRow_Internal = 0
End Function

' ------------------------------------------------------------------------------
' 祝日表シートのB列（日付）を読み込み、"yyyy/mm/dd" をキーとする辞書を返す
' シートが存在しない場合は空の辞書を返す（エラーにしない）
' ------------------------------------------------------------------------------
Private Function BuildHolidayDict() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("祝日表")
    On Error GoTo 0

    If ws Is Nothing Then
        Set BuildHolidayDict = dict
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    Dim i As Long
    For i = 1 To lastRow
        Dim cellVal As Variant
        cellVal = ws.Cells(i, 2).Value
        If IsDate(cellVal) Then
            Dim key As String
            key = Format(CDate(cellVal), "yyyy/mm/dd")
            If Not dict.Exists(key) Then
                dict.Add key, CStr(ws.Cells(i, 1).Value)  ' 祝日名も格納
            End If
        End If
    Next i

    Set BuildHolidayDict = dict
End Function

Private Sub Step9_Method()
    Dim objFSO As Object
    Dim objShell As Object
    Dim objStream As Object
    Dim objCompress As Object
    Dim strCompressFilePath As String
    Dim objSourceFilePaths As Collection
    Dim objValue As Variant
    Dim nCounter As Integer

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("Shell.Application")
    Set objSourceFilePaths = New Collection

    strCompressFilePath = p_FolderPath & "\nmht_" & p_DateStr & p_TimeStr & ".zip"

    objSourceFilePaths.Add p_FolderPath & "\" & "BH" & p_DateStr & p_TimeStr & ".txt"
    objSourceFilePaths.Add p_FolderPath & "\" & "BD" & p_DateStr & p_TimeStr & ".txt"
    objSourceFilePaths.Add p_FolderPath & "\" & "BT" & p_DateStr & p_TimeStr & ".txt"

    If objFSO.FileExists(strCompressFilePath) Then
        objFSO.DeleteFile strCompressFilePath
    End If

    Set objStream = objFSO.CreateTextFile(strCompressFilePath, True)
    objStream.Write "PK" & Chr(5) & Chr(6) & String(18, 0)
    objStream.Close
    Set objStream = Nothing

    Set objCompress = objShell.Namespace(objFSO.GetAbsolutePathName(strCompressFilePath))

    ' 【修正】ZIP格納の待機ループにタイムアウト（120秒）を追加
    '         Shell.Application の CopyHere は環境によって完了通知が遅いため余裕を持たせる
    '         タイムアウトなしでは、格納失敗時に永久ループとなるバグがあった
    '         Application.Wait は Excel をブロックし Shell.Application の動作を妨げるため使用しない
    '         Timer は深夜0時にリセットされるため Now/DateDiff("s",...) を使用する
    Const ZIP_TIMEOUT_SEC As Long = 120
    Dim dtStart As Date
    For Each objValue In objSourceFilePaths
        If objFSO.FileExists(objValue) Then
            objCompress.CopyHere objValue
            nCounter = nCounter + 1
            dtStart = Now
            Do While objCompress.Items().Count < nCounter
                DoEvents
                If DateDiff("s", dtStart, Now) > ZIP_TIMEOUT_SEC Then
                    Err.Raise 9001, , "ZIPファイルへの追加がタイムアウトしました（" & ZIP_TIMEOUT_SEC & "秒超）。" & vbCrLf & _
                                      "対象ファイル: " & CStr(objValue) & vbCrLf & _
                                      "セキュリティソフトや権限の問題がないか確認してください。"
                End If
            Loop
        End If
    Next

    Set objCompress = Nothing
    Set objShell = Nothing
    Set objFSO = Nothing

    On Error Resume Next
    Kill p_FolderPath & "\" & "BH" & p_DateStr & p_TimeStr & ".txt"
    Kill p_FolderPath & "\" & "BD" & p_DateStr & p_TimeStr & ".txt"
    Kill p_FolderPath & "\" & "BT" & p_DateStr & p_TimeStr & ".txt"
    On Error GoTo 0
End Sub








