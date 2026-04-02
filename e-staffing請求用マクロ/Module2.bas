Attribute VB_Name = "Module2"
Option Explicit

' ■モジュールレベル変数（全工程で共通して使用する値）
Private p_wsMain As Worksheet      ' メインシート（勤務データ）
Private p_FolderPath As String     ' 出力先パス（P1セルの値）
Private p_DateStr As String        ' 共通の日付文字列（mmdd）
Private p_TimeStr As String        ' 共通の時間文字列（hhmm）

' ==============================================================================
' 【メイン処理】このマクロをボタンに登録してください
' ==============================================================================
Sub 一括自動実行()
    On Error GoTo MainErrorHandler
    
    ' --- 初期設定 ---
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual ' 処理高速化のため手動計算へ
    
    ' シートの特定
    Set p_wsMain = ThisWorkbook.Worksheets("勤務データ")
    
    ' パスの取得とチェック
    p_FolderPath = p_wsMain.Range("P1").Value
    If p_FolderPath = "" Then Err.Raise 1001, , "P1セルに出力先フォルダパスが入力されていません。"
    If Right(p_FolderPath, 1) = "\" Then p_FolderPath = Left(p_FolderPath, Len(p_FolderPath) - 1)
    If Dir(p_FolderPath, vbDirectory) = "" Then Err.Raise 1002, , "指定された出力先フォルダが存在しません。" & vbCrLf & p_FolderPath
    
    ' タイムスタンプの固定（全ファイルで統一）
    p_DateStr = Format(Date, "mmdd")
    p_TimeStr = Format(Time, "hhmm")
    
    
    ' --- 実行フロー ---
    
    ' 1. 立替金集計
    Application.StatusBar = "処理中：1/9 立替金集計"
    Call Step1_立替金集計
    
    ' 2. 立替データ保存（これを先にやらないと請求計算が合わない）
    Application.StatusBar = "処理中：2/9 立替データ保存"
    Call Step2_立替データ保存
    
    ' 3. 請求明細作成
    Application.StatusBar = "処理中：3/9 請求明細作成"
    Call Step3_請求明細作成
    
    ' 4. 請求ヘッダ作成
    Application.StatusBar = "処理中：4/9 請求ヘッダ作成"
    Call Step4_請求ヘッダ作成
    
    ' 5. 請求データとの照合（ファイルがなければスキップ）
    Application.StatusBar = "処理中：5/9 請求データとの照合"
    Call Step5_請求データとの照合
    
    ' 6. データ保存（バックアップ）
    Application.StatusBar = "処理中：6/9 データバックアップ保存"
    Call Step6_データ保存
    
    ' 7. 提出データファイル作成 (BH, BD)
    Application.StatusBar = "処理中：7/9 BH/BDファイル作成"
    Call Step7_提出データファイル作成
    
    ' 8. BTファイル作成 (コード②ベース)
    Application.StatusBar = "処理中：8/9 BTファイル作成"
    Call Step8_ExportTimecardData
    
    ' 9. ZIP作成・ファイル整理
    Application.StatusBar = "処理中：9/9 ZIP作成"
    Call Step9_Method
    
    
    ' --- 終了処理 ---
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    MsgBox "全ての処理が完了しました。" & vbCrLf & "出力先: " & p_FolderPath, vbInformation, "完了"
    Exit Sub

MainErrorHandler:
    ' エラー時の共通処理
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    MsgBox "エラーが発生したため処理を中断します。" & vbCrLf & _
           "内容: " & Err.Description, vbCritical, "エラー中断"
End Sub


' ==============================================================================
' 以下、各ステップのサブプロシージャ
' ==============================================================================

Private Sub Step1_立替金集計()
    Dim ws1 As Worksheet '立替金データ
    Dim ws2 As Worksheet '集計先
    
    Dim targetSheetName As String
    targetSheetName = "立替金集計"
    
    ' シートが存在するか確認、なければ作成
    On Error Resume Next
    Set ws2 = Worksheets(targetSheetName)
    On Error GoTo 0
    If ws2 Is Nothing Then
        Set ws2 = Worksheets.Add
        ws2.Name = targetSheetName
    End If
    
    Set ws1 = Worksheets("立替金データ")
    
    Dim MR0 As Long, MR1 As Long, MR2 As Long
    Dim i As Long, j As Long
    
    ' 初期化
    MR0 = ws2.Cells(Rows.Count, 2).End(xlUp).Row
    If MR0 > 1 Then ws2.Range("A2:F" & MR0).Clear
    
    ' 社員番号、スタッフ番号、名前をコピー
    MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    ws2.Range("B2:C" & MR1).Value = ws1.Range("D1:E" & MR1).Value
    
    ' 重複を削除
    ws2.Activate
    With ws2.Range("B2")
        For i = .CurrentRegion.Rows.Count To 1 Step -1
            If .offset(i, 0).Value = .offset(i - 1, 0).Value Then .offset(i, 0).EntireRow.Delete
        Next i
    End With
    p_wsMain.Activate ' メインに戻る
    
    ' 立替金の集計
    MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    MR2 = ws2.Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = 2 To MR2
        ' 立替金合計
        ws2.Cells(i, 4).Value = Application.WorksheetFunction.SumIf(ws1.Columns(4), ws2.Cells(i, 2), ws1.Columns(15))
        ws2.Cells(i, 4).NumberFormatLocal = "##,##"
        
        ' 顧客対応当番手当の積算
        Dim tempVal As Double
        tempVal = 0
        For j = 1 To MR1
            If InStr(ws1.Cells(j, 12).Value, "顧客対応当番") > 0 Then
                If ws1.Cells(j, 4).Value = ws2.Cells(i, 2).Value Then
                    tempVal = tempVal + ws1.Cells(j, 15).Value
                End If
            End If
        Next
        ws2.Cells(i, 5).Value = tempVal
        ws2.Cells(i, 5).NumberFormatLocal = "##,##"
        
        ' 交通費
        ws2.Cells(i, 6).Value = ws2.Cells(i, 4).Value - ws2.Cells(i, 5).Value
        ws2.Cells(i, 6).NumberFormatLocal = "##,##"
        
        ' 名前参照
        On Error Resume Next
        ws2.Cells(i, 1).Value = Application.WorksheetFunction.SumIf(Worksheets("名簿").Columns(1), ws2.Cells(i, 3), Worksheets("名簿").Columns(2))
        On Error GoTo 0
    Next
    
    If MR1 > 0 Then
        ws2.Cells(1, 7).Value = Month(ws1.Cells(2, 6).Value) & "月"
    End If
End Sub

Private Sub Step2_立替データ保存()
    Dim ws1 As Worksheet '集計したシート
    Dim ws3 As Worksheet '勤務データ
    
    Set ws1 = Worksheets("立替金集計")
    Set ws3 = p_wsMain
    
    Dim MR1 As Long, MR2 As Long
    Dim i As Long, j As Long
    
    MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    MR2 = ws3.Cells(Rows.Count, 2).End(xlUp).Row
    
    ' クリア
    ws3.Range("I2:K" & MR2).ClearContents
    
    ' 転記処理
    For i = 2 To MR2 ' 保存先
        For j = 2 To MR1 ' 保存元
            If ws3.Cells(i, 1).Value = ws1.Cells(j, 1).Value Then
                ws3.Cells(i, 9).Value = ws1.Cells(j, 5).Value '当番手当
                ws3.Cells(i, 9).NumberFormatLocal = "##,##"
                ws3.Cells(i, 10).Value = ws1.Cells(j, 6).Value '交通費
                ws3.Cells(i, 10).NumberFormatLocal = "##,##"
                Exit For
            End If
        Next j
    Next i
End Sub

Private Sub Step3_請求明細作成()
    Dim ws2 As Worksheet '請求明細
    Dim ws3 As Worksheet '契約情報
    Dim ws1 As Worksheet
    Set ws1 = p_wsMain
    Set ws2 = Worksheets("請求明細")
    Set ws3 = Worksheets("e-staffing TCnmhtの最新情報")
    
    Dim MR1 As Long, MR2 As Long, MR3 As Long
    Dim i As Long, j As Long
    
    Application.StandardFont = "游ゴシック"
    Application.StandardFontSize = 11
    
    MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    MR2 = ws2.Cells(Rows.Count, 7).End(xlUp).Row
    
    ' 勤務データの加工
    ws1.Range("L2").Value = "=+SUMIF('e-staffing TCnmhtの最新情報'!M:M,B2,'e-staffing TCnmhtの最新情報'!BC:BC)"
    ws1.Range("M2").Value = "=+SUMIFS(請求内訳!H:H,請求内訳!C:C,A2,請求内訳!J:J,""基本料金"")"
    ws1.Range("N2").Value = "=+(L2=1)*(ROUNDDOWN(E2*2,0)/2+ROUNDDOWN(F2*2,0)/2)+(L2=3)*MAX(0,D2-M2)"
    ws1.Range("O2").Value = "=+(L2=1)*(D2-ROUNDDOWN(E2*2,0)/2-ROUNDDOWN(f2*2,0)/2)+(L2=3)*MIN(D2,M2)"
    
    If MR1 >= 2 Then
        ws1.Range("L2:O2").Copy
        ws1.Range("L2:O" & MR1).PasteSpecial xlPasteAll
        ws1.Calculate
        ws1.Range("L2:O" & MR1).Copy
        ws1.Range("L2:O" & MR1).PasteSpecial xlPasteValues
    End If
    
    For i = 2 To MR1
        If ws1.Cells(i, 12).Value = 1 Then
            ws1.Cells(i, 14).Value = WorksheetFunction.RoundDown(ws1.Cells(i, 5).Value * 2, 0) / 2
        End If
    Next
    
    ' 請求明細のデータクリア
    If MR2 >= 7 Then ws2.Range("A7:AT" & MR2 + 100).ClearContents
    
    ' 氏名をコピー
    ws1.Range("C2:C" & MR1).Copy
    ws2.Range("G7").PasteSpecial xlPasteAll
    ws1.Range("B2:B" & MR1).Copy
    ws2.Range("F7").PasteSpecial xlPasteAll
    ws2.Range("H7").PasteSpecial xlPasteAll
    
    MR2 = ws2.Cells(Rows.Count, 7).End(xlUp).Row
    MR3 = ws3.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' 契約番号を書き込む
    For i = 7 To MR2
        Dim jc As Variant
        Dim ret As Variant
        jc = ws2.Cells(i, 6).Value
        On Error Resume Next
        ret = WorksheetFunction.Match(jc, ws3.Range("M1:M" & MR3), 0)
        If Err.Number = 0 Then
            ws2.Cells(i, 2).Value = ws3.Cells(ret, 1).Value
        End If
        On Error GoTo 0
    Next
    
    ' UALに絞り込む
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
    For j = lRow To 7 Step -1
        If ws2.Cells(j, 3).Value <> "ual-nmht" Then
            ws2.Rows(j).Delete
        End If
    Next j
    
    Application.Calculation = xlCalculationManual
    
    ' ■データの書込み
    MR2 = ws2.Cells(Rows.Count, 3).End(xlUp).Row
    
    For i = 7 To MR2
        ws2.Cells(i, 3).Value = BuildInvoiceDetailCode(CStr(ws2.Cells(i, 6).Value))
        ws2.Cells(i, 4).Value = 1
        ws2.Cells(i, 9).NumberFormatLocal = "@"
        ws2.Cells(i, 9).Value = Format(ws1.Range("K1").Value, "yyyy") & "/" & Format(ws1.Range("K1").Value, "mm")
        ws2.Cells(i, 10).Value = Format(ws1.Range("H1").Value, "yyyy/mm/dd")
        ws2.Cells(i, 11).Value = Format(ws1.Range("K1").Value, "yyyy/mm/dd")
    Next
    
    ' 数式セット
    ws2.Cells(7, 5).Value = "=+VLOOKUP(G7,名簿!A:B,2,0)"
    ws2.Cells(7, 12).Value = "=+SUMIF('e-staffing TCnmhtの最新情報'!A:A,B7,'e-staffing TCnmhtの最新情報'!BB:BB)"
    ws2.Cells(7, 13).Value = "=+SUMIF('e-staffing TCnmhtの最新情報'!A:A,B7,'e-staffing TCnmhtの最新情報'!BC:BC)"
    ws2.Cells(7, 14).Value = "=+ROUNDDOWN(SUMIF(勤務データ!B:B,H7,勤務データ!o:o),0)"
    ws2.Cells(7, 15).Value = "=+ROUNDDOWN((SUMIF(勤務データ!B:B,H7,勤務データ!O:O)-N7+.001)*2,0)*30"
    ws2.Cells(7, 16).Value = "=+IF(M7=1,L7*(N7+O7/60),IF(M7=3,L7,""error!""))"
    
    ' 時間外～立替金
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
    
    ' 数式全面展開
    If MR2 > 7 Then
        ws2.Range("E7").Copy ws2.Range("E8:E" & MR2)
        ws2.Range("L7:AN7").Copy ws2.Range("L8:AN" & MR2)
    End If
    
    ws2.Calculate
    
    ' 書式コピー
    ws2.Range("B7:AQ7").Copy
    ws2.Range("B8:AQ" & MR2).PasteSpecial xlPasteFormats
    
    ' 月額制修正
    Dim k As Long
    For k = 7 To MR2
        If ws2.Cells(k, 13).Value = 3 Then
            ws2.Cells(k, 20).Value = 0
            ws2.Cells(k, 21).Value = 0
            ws2.Cells(k, 26).Value = 0
            ws2.Cells(k, 27).Value = 0
        End If
    Next
    ws2.Calculate
    
    ' 値貼り付け
    ws2.Range("B7:AQ" & MR2).Copy
    ws2.Range("B7:AQ" & MR2).PasteSpecial xlValues
    ws2.Range("W7:X" & MR2).Copy
    ws2.Range("W7:X" & MR2).PasteSpecial xlValues
    
    ' 時間繰り上がりの修正
    For k = 7 To MR2
        If ws2.Cells(k, 24).Value > 30 Then
            ws2.Cells(k, 23).Value = ws2.Cells(k, 23).Value + 1
            ws2.Cells(k, 24).Value = 0
        End If
    Next
    
    ' 請求金額合計
    ws2.Cells(7, 39).NumberFormatLocal = "0_ "
    ws2.Cells(7, 39).Font.Size = 11
    ws2.Cells(7, 39).Value = "=+P7+Y7+AB7-AH7+AI7+AL7"
    If MR2 > 7 Then
        ws2.Range("AM7").Copy ws2.Range("AM8:AM" & MR2)
    End If
    ws2.Calculate
    ws2.Range("AM7:AM" & MR2).Copy
    ws2.Range("AM7:AM" & MR2).PasteSpecial xlPasteValues
    
    ' 未使用欄処理
    ws2.Range("E7:E" & MR2).Value = "nmht"
    ws2.Range("AF7:AF" & MR2).Value = 0
    ws2.Range("AF7:AG" & MR2).Value = 0
    ws2.Range("AI7:AL" & MR2).Value = 0
    ws2.Range("AP7:AQ" & MR2).Value = 0
    
End Sub

Private Sub Step4_請求ヘッダ作成()
    Dim ws2 As Worksheet '請求明細
    Dim ws3 As Worksheet '請求ヘッダ
    Set ws2 = Worksheets("請求明細")
    Set ws3 = Worksheets("請求ヘッダテンプレート (請求コメントあり)")
    
    Dim MR2 As Long, MR3 As Long
    Dim i As Long
    
    MR3 = ws3.Cells(Rows.Count, 7).End(xlUp).Row
    If MR3 >= 7 Then ws3.Range("B7:BG" & MR3).Delete
    
    ws3.Cells.NumberFormatLocal = "G/標準"
    
    ' 固定値・数式セット
    ws3.Range("B7").Value = "請H_2"
    ws3.Range("C7").Value = 1
    ws3.Range("D7").Value = "=+請求明細!C7"
    ws3.Range("E7").Value = "=+請求明細!I7"
    ws3.Range("F7").Value = "ual-nmht"
    ws3.Range("G7").Value = "=+請求明細!C7"
    ws3.Range("H7").Value = "=""【"" & 請求明細!G7 &""】""& VLOOKUP(請求明細!F7,'e-staffing TCnmhtの最新情報'!U:CC,61,0)"
    ws3.Range("I7").Value = "=+請求明細!J7"
    ws3.Range("J7").Value = "=+請求明細!K7"
    ws3.Range("K7").Value = "=+請求明細!K7"
    ws3.Range("L7").Value = "=+EOMONTH(K7,1)"
    ws3.Range("L7").NumberFormatLocal = "yyyy/mm/dd"
    ws3.Range("M7").Value = "=+請求明細!AM7"
    ws3.Range("P7").Value = "=+rounddown(M7*0.1,0)"
    ws3.Range("Q7").Value = "=+rounddown(M7*1.1,0)"
    ws3.Range("R7").Value = "=+請求明細!AN7"
    ws3.Range("S7").Value = "=+Q7+R7"
    ws3.Range("T7").Value = "りそな銀行"
    ws3.Range("U7").Value = "九段支店"
    ws3.Range("V7").Value = 1
    ws3.Range("W7").Value = "1416351"
    ws3.Range("AR7").Value = "管理部"
    ws3.Range("AS7").Value = "03-6222-9420"
    ws3.Range("AT7").Value = 1
    ws3.Range("BG7").Value = 0
    
    MR2 = ws2.Cells(Rows.Count, 7).End(xlUp).Row
    
    ' 展開
    ws3.Range("B7:BG7").Copy
    ws3.Range("B8:BG" & MR2).PasteSpecial xlPasteAll
    ws3.Range("B7:BG" & MR2).Copy
    ws3.Range("B7:BG" & MR2).PasteSpecial xlPasteValues
    
    ' 置換処理
    For i = 7 To MR2
        ws3.Cells(i, 8).Value = Replace(ws3.Cells(i, 8).Value, "ＮＵＬシステム購買室", "ユニアデックス株式会社")
        ws3.Cells(i, 8).Value = Replace(ws3.Cells(i, 8).Value, "。", "")
        ws3.Cells(i, 8).Value = Replace(ws3.Cells(i, 8).Value, "．", "")
    Next
End Sub

Private Sub Step5_請求データとの照合()
    Dim ws11 As Worksheet '勤務データ
    Dim ws12 As Worksheet '請求内訳
    Dim ws13 As Worksheet '名簿
    Dim wb2 As Workbook
    Dim ws21 As Worksheet
    
    Set ws11 = p_wsMain
    Set ws12 = Worksheets("請求内訳")
    Set ws13 = Worksheets("名簿")
    
    Dim OpenFileName As String
    OpenFileName = ws11.Range("M1").Value
    
    ' ファイルがなければスキップ
    If OpenFileName = "" Or Dir(OpenFileName) = "" Then
        ' ファイルがない場合は何もしない
        Exit Sub
    End If
    
    ' ファイルを開く
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
    
    MR11 = ws11.Cells(Rows.Count, 6).End(xlUp).Row 'コピー後再取得
    MR12 = ws12.Cells(Rows.Count, 1).End(xlUp).Row
    MR13 = ws13.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' 社員番号追加
    For i = 2 To MR11
        For j = 2 To MR13
            If ws11.Cells(i, 4).Value = ws13.Cells(j, 1).Value Then
                ws11.Cells(i, 1).Value = ws13.Cells(j, 2).Value
            End If
        Next
    Next
    
    ' 請求金額追加
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
    
    ' 合計行
    ws11.Cells(MR11 + 1, 11).Value = "=sum(K2:K" & MR11 & ")"
    ws11.Cells(MR11 + 1, 11).NumberFormatLocal = "##,#"
    ws11.Cells(MR11 + 1, 12).Value = "=J" & MR11 + 1 & "-K" & MR11 + 1
    ws11.Cells(MR11 + 1, 12).NumberFormatLocal = "##,#"
    
    ' 罫線
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
        p_wsMain.Activate 'メインに戻す
    End If
End Sub

Private Sub Step7_提出データファイル作成()
    Dim ws2 As Worksheet '請求明細
    Dim ws3 As Worksheet '請求ヘッダ
    Dim wb2 As Workbook '明細ファイル
    Dim wb3 As Workbook 'ヘッダファイル
    
    Set ws2 = Worksheets("請求明細")
    Set ws3 = Worksheets("請求ヘッダテンプレート (請求コメントあり)")
    
    Dim Hd_file As String, M_file As String
    
    ' 共通のタイムスタンプを使用
    Hd_file = p_FolderPath & "\" & "BH" & p_DateStr & p_TimeStr & ".txt"
    M_file = p_FolderPath & "\" & "BD" & p_DateStr & p_TimeStr & ".txt"
    
    ' 明細出力
    ws2.Copy
    Set wb2 = ActiveWorkbook
    wb2.Sheets(1).Cells.NumberFormatLocal = "@"
    wb2.SaveAs fileName:=M_file, FileFormat:=xlText
    
    ' 不要行削除
    wb2.Sheets(1).Columns(1).Delete
    wb2.Sheets(1).Rows("1:6").Delete
    wb2.SaveAs fileName:=M_file, FileFormat:=xlText
    wb2.Close SaveChanges:=False
    
    ' ヘッダ出力
    ws3.Copy
    Set wb3 = ActiveWorkbook
    wb3.SaveAs fileName:=Hd_file, FileFormat:=xlText
    
    ' 不要行削除
    wb3.Sheets(1).Columns(1).Delete
    wb3.Sheets(1).Rows("1:6").Delete
    wb3.SaveAs fileName:=Hd_file, FileFormat:=xlText
    wb3.Close SaveChanges:=False
End Sub

Private Function BuildInvoiceDetailCode(jobCode As String) As String
    Dim closingDateVal As Variant
    Dim trimmedJobCode As String

    closingDateVal = p_wsMain.Range("K1").Value
    If Not IsDate(closingDateVal) Then
        Err.Raise 2101, , "勤務データシートのK1セルが日付ではないため、請求書明細コードを生成できません。"
    End If

    trimmedJobCode = Trim$(jobCode)
    If trimmedJobCode = "" Then
        Err.Raise 2102, , "Jobコードが空白のため、請求書明細コードを生成できません。"
    End If

    BuildInvoiceDetailCode = "BNT" & Format(CDate(closingDateVal), "yyyymmdd") & trimmedJobCode
End Function

Private Function ResolveWebTCJobCode(ws As Worksheet, headerRow As Long) As String
    Dim primaryCode As String
    Dim fallbackCode As String

    primaryCode = Trim$(CStr(ws.Cells(headerRow, 5).Value))
    If primaryCode <> "" Then
        ResolveWebTCJobCode = primaryCode
        Exit Function
    End If

    fallbackCode = Trim$(CStr(ws.Cells(headerRow, 9).Value))
    ResolveWebTCJobCode = fallbackCode
End Function

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
            tcCounter = tcCounter + 1
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

    Dim wsKintai As Worksheet
    On Error Resume Next
    Set wsKintai = ThisWorkbook.Worksheets("勤怠 (就業場所区分あり)")
    On Error GoTo 0

    If Not wsKintai Is Nothing Then
        Dim lastRowKintai As Long
        lastRowKintai = wsKintai.Cells(wsKintai.Rows.Count, 2).End(xlUp).Row
        If lastRowKintai >= 7 Then
            wsKintai.Range("B7:QS" & lastRowKintai).ClearContents
        End If

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

    Dim holidayDict As Object
    Set holidayDict = BuildHolidayDict()

    For j = 1 To lastRow
        If wsWebTC.Cells(j, 1).Value = "H" And wsWebTC.Cells(j, 2).Value = contractNo Then
            jobCode = ResolveWebTCJobCode(wsWebTC, j)
            foundH = True
        ElseIf wsWebTC.Cells(j, 1).Value = "D" And foundH Then
            Dim rowCategory As String
            rowCategory = CStr(wsWebTC.Cells(j, 3).Value)

            If rowCategory = "" Then
                Dim rowDateVal As Variant
                rowDateVal = wsWebTC.Cells(j, 2).Value
                If IsDate(rowDateVal) Then
                    If holidayDict.Exists(Format(CDate(rowDateVal), "yyyy/mm/dd")) Then
                        rowCategory = "2"
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
    If Trim$(jobCode) = "" Then Err.Raise 2006, , contractNo & ": Jobコードが取得できません。"

    fields(8) = jobCode
    fields(3) = BuildInvoiceDetailCode(jobCode)
    fields(4) = contractNo
    fields(5) = CStr(tcNumber)
    fields(6) = CStr(wsEStaffing.Cells(contractRow, 5).Value)
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
                dict.Add key, CStr(ws.Cells(i, 1).Value)
            End If
        End If
    Next i

    Set BuildHolidayDict = dict
End Function
Private Sub Step9_Method()
    ' ZIP作成処理
    Dim objFSO As Object
    Dim objShell As Object
    Dim objStream As Object
    Dim objCompress As Object
    Dim strCompressFilePath As String
    Dim objSourceFilePaths As Collection
    Dim objValue As Variant
    Dim nCounter As Integer
    Dim bResult As Boolean
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("Shell.Application")
    Set objSourceFilePaths = New Collection
    
    ' ZIPファイルパス
    strCompressFilePath = p_FolderPath & "\nmht_" & p_DateStr & p_TimeStr & ".zip"
    
    ' 格納するファイル（3つすべて）
    objSourceFilePaths.Add p_FolderPath & "\" & "BH" & p_DateStr & p_TimeStr & ".txt"
    objSourceFilePaths.Add p_FolderPath & "\" & "BD" & p_DateStr & p_TimeStr & ".txt"
    objSourceFilePaths.Add p_FolderPath & "\" & "BT" & p_DateStr & p_TimeStr & ".txt"
    
    ' 既存ZIP削除
    If objFSO.FileExists(strCompressFilePath) Then
        objFSO.DeleteFile strCompressFilePath
    End If
    
    ' 空ZIP作成
    Set objStream = objFSO.CreateTextFile(strCompressFilePath, True)
    objStream.Write "PK" & Chr(5) & Chr(6) & String(18, 0)
    objStream.Close
    Set objStream = Nothing
    
    ' コピー（圧縮）
    Set objCompress = objShell.Namespace(objFSO.GetAbsolutePathName(strCompressFilePath))
    
    For Each objValue In objSourceFilePaths
        If objFSO.FileExists(objValue) Then
            objCompress.CopyHere objValue
            nCounter = nCounter + 1
            ' 完了待機
            Do While objCompress.Items().Count < nCounter
                DoEvents
            Loop
        End If
    Next
    
    ' オブジェクト破棄
    Set objCompress = Nothing
    Set objShell = Nothing
    Set objFSO = Nothing
    
    ' 元ファイルの削除（BH, BD, BT すべて削除）
    On Error Resume Next
    Kill p_FolderPath & "\" & "BH" & p_DateStr & p_TimeStr & ".txt"
    Kill p_FolderPath & "\" & "BD" & p_DateStr & p_TimeStr & ".txt"
    Kill p_FolderPath & "\" & "BT" & p_DateStr & p_TimeStr & ".txt"
    On Error GoTo 0
End Sub



