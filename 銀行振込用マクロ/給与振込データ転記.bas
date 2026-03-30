'=======================================================
' 給与振込データ転記マクロ
'
' Q-meisai CSVとjinjer CSVから給与振込一覧シートへ
' データを自動転記するマクロ
'
' 【VBA貼り付け手順】
'   1. Alt+F11 でVBエディタを開く
'   2. 左ツリーで「VBAProject」を右クリック
'      → 「挿入」→「標準モジュール」
'   3. このファイルの内容を全部コピーして貼り付け
'   4. VBエディタを閉じる（Alt+F4）
'   5. ファイルを「.xlsm」形式で保存し直す
'   6. Alt+F8 →「ボタン追加」を実行（初回のみ）
'=======================================================

Option Explicit

'-------------------------------------------------------
' メインマクロ：給与振込データ転記
'-------------------------------------------------------
Sub 給与振込データ転記()

    '--- シートの取得 ---
    Dim wsOutput  As Worksheet   ' 給与振込一覧
    
    On Error Resume Next
    Set wsOutput  = ThisWorkbook.Sheets("給与振込一覧")
    On Error GoTo 0
    
    If wsOutput Is Nothing Then
        MsgBox "「給与振込一覧」シートが見つかりません。" & Chr(10) & _
               "シート名を確認してください。", vbExclamation
        Exit Sub
    End If
    
    '--- 支給日の入力 ---
    Dim payDateStr As String
    payDateStr = InputBox("支給日を入力してください。" & Chr(10) & _
                          "（例：2026/3/25）", "支給日の指定")
    
    If payDateStr = "" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If
    
    ' 日付の妥当性チェック
    If Not IsDate(payDateStr) Then
        MsgBox "正しい日付形式で入力してください。" & Chr(10) & _
               "（例：2026/3/25）", vbExclamation
        Exit Sub
    End If
    
    Dim payDate As Date
    payDate = CDate(payDateStr)
    
    '===========================================================
    ' ① Q-meisai CSVファイルの選択・読み込み → Dictionary
    '   キー: 社員番号  値: 配列(氏名, 口座1振込額)
    '===========================================================
    Dim qmeisaiPath As String
    qmeisaiPath = Application.GetOpenFilename( _
        FileFilter:="CSVファイル (*.csv), *.csv", _
        Title:="Q-meisai CSVファイルを選択してください")
    
    If qmeisaiPath = "False" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If
    
    Dim qmeisaiDict As Object
    Set qmeisaiDict = CreateObject("Scripting.Dictionary")
    
    ' Q-meisaiの社員番号順を保持する配列
    Dim qEmpOrder() As String
    Dim qOrderCount As Long
    qOrderCount = 0
    
    Dim fileNo As Integer
    fileNo = FreeFile
    
    Open qmeisaiPath For Input As #fileNo
    
    ' ヘッダー行から列位置を特定
    Dim qHeaderLine As String
    Line Input #fileNo, qHeaderLine
    
    Dim qHeaders() As String
    qHeaders = SplitCSVLine(qHeaderLine)
    
    Dim colQEmpNo As Integer:   colQEmpNo = -1
    Dim colQName As Integer:    colQName = -1
    Dim colQAmount As Integer:  colQAmount = -1
    
    Dim qh As Integer
    For qh = 0 To UBound(qHeaders)
        Select Case Trim(qHeaders(qh))
            Case "社員番号"
                colQEmpNo = qh
            Case "氏名"
                colQName = qh
            Case "口座1振込額"
                colQAmount = qh
        End Select
    Next qh
    
    ' 必須列が見つからない場合
    If colQEmpNo = -1 Or colQName = -1 Or colQAmount = -1 Then
        Close #fileNo
        MsgBox "Q-meisai CSVのヘッダーが想定と異なります。" & Chr(10) & _
               "以下の列が必要です：" & Chr(10) & _
               "・社員番号" & Chr(10) & _
               "・氏名" & Chr(10) & _
               "・口座1振込額", vbExclamation
        Exit Sub
    End If
    
    ' データ行の読み込み
    Dim qDataLine As String
    Dim qFields() As String
    Do While Not EOF(fileNo)
        Line Input #fileNo, qDataLine
        If Trim(qDataLine) = "" Then GoTo NextQmeisaiLine
        
        qFields = SplitCSVLine(qDataLine)
        
        Dim qEmpNo As String
        qEmpNo = Trim(qFields(colQEmpNo))
        If qEmpNo = "" Then GoTo NextQmeisaiLine
        
        ' 氏名・振込額を配列で格納
        Dim qInfo(0 To 1) As String
        qInfo(0) = Trim(qFields(colQName))      ' 氏名
        qInfo(1) = Trim(qFields(colQAmount))     ' 口座1振込額
        
        qmeisaiDict(qEmpNo) = qInfo
        
        ' 順序を保持
        ReDim Preserve qEmpOrder(qOrderCount)
        qEmpOrder(qOrderCount) = qEmpNo
        qOrderCount = qOrderCount + 1
        
NextQmeisaiLine:
    Loop
    Close #fileNo
    
    If qmeisaiDict.Count = 0 Then
        MsgBox "Q-meisai CSVからデータを読み込めませんでした。" & Chr(10) & _
               "ファイルの内容を確認してください。", vbExclamation
        Exit Sub
    End If
    
    '===========================================================
    ' ② jinjer CSVファイルの選択・読み込み → Dictionary
    '   キー: 社員番号  値: 配列(銀行名, 支店名, 口座番号, 口座名義カナ)
    '===========================================================
    Dim jinjerPath As String
    jinjerPath = Application.GetOpenFilename( _
        FileFilter:="CSVファイル (*.csv), *.csv", _
        Title:="jinjer CSVファイルを選択してください")
    
    If jinjerPath = "False" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If
    
    Dim jinjerDict As Object
    Set jinjerDict = CreateObject("Scripting.Dictionary")
    
    fileNo = FreeFile
    
    Open jinjerPath For Input As #fileNo
    
    ' ヘッダー行から列位置を特定
    Dim headerLine As String
    Line Input #fileNo, headerLine
    
    Dim headers() As String
    headers = SplitCSVLine(headerLine)
    
    Dim colEmpNo As Integer:     colEmpNo = -1
    Dim colBankName As Integer:  colBankName = -1
    Dim colBranch As Integer:    colBranch = -1
    Dim colAcctNo As Integer:    colAcctNo = -1
    Dim colAcctName As Integer:  colAcctName = -1
    
    Dim h As Integer
    For h = 0 To UBound(headers)
        Select Case Trim(headers(h))
            Case "社員番号"
                colEmpNo = h
            Case "振込先銀行名(銀行振込口座1)"
                colBankName = h
            Case "振込先銀行支店名(銀行振込口座1)"
                colBranch = h
            Case "口座番号(銀行振込口座1)"
                colAcctNo = h
            Case "口座名義（ﾌﾘｶﾞﾅ）(銀行振込口座1)"
                colAcctName = h
        End Select
    Next h
    
    ' 必須列が見つからない場合
    If colEmpNo = -1 Or colBankName = -1 Or colBranch = -1 Or _
       colAcctNo = -1 Or colAcctName = -1 Then
        Close #fileNo
        MsgBox "jinjer CSVのヘッダーが想定と異なります。" & Chr(10) & _
               "以下の列が必要です：" & Chr(10) & _
               "・社員番号" & Chr(10) & _
               "・振込先銀行名(銀行振込口座1)" & Chr(10) & _
               "・振込先銀行支店名(銀行振込口座1)" & Chr(10) & _
               "・口座番号(銀行振込口座1)" & Chr(10) & _
               "・口座名義（ﾌﾘｶﾞﾅ）(銀行振込口座1)", vbExclamation
        Exit Sub
    End If
    
    ' データ行の読み込み
    Dim dataLine As String
    Dim fields() As String
    Do While Not EOF(fileNo)
        Line Input #fileNo, dataLine
        If Trim(dataLine) = "" Then GoTo NextJinjerLine
        
        fields = SplitCSVLine(dataLine)
        
        Dim empNo As String
        empNo = Trim(fields(colEmpNo))
        If empNo = "" Then GoTo NextJinjerLine
        
        ' 口座情報を配列で格納
        Dim acctInfo(0 To 3) As String
        acctInfo(0) = Trim(fields(colBankName))     ' 金融機関名
        acctInfo(1) = Trim(fields(colBranch))        ' 支店名
        acctInfo(2) = Trim(fields(colAcctNo))        ' 口座番号
        acctInfo(3) = Trim(fields(colAcctName))      ' 口座名義カナ
        
        jinjerDict(empNo) = acctInfo
        
NextJinjerLine:
    Loop
    Close #fileNo
    
    If jinjerDict.Count = 0 Then
        MsgBox "jinjer CSVからデータを読み込めませんでした。" & Chr(10) & _
               "ファイルの内容を確認してください。", vbExclamation
        Exit Sub
    End If

    '===========================================================
    ' ②-2 S-meisai CSVファイルの読み込み（任意）
    '   キー: 社員番号  値: 口座1振込額（文字列）
    '===========================================================
    Dim smeisaiDict As Object
    Set smeisaiDict = CreateObject("Scripting.Dictionary")

    Dim hasSmeisai As Boolean
    hasSmeisai = (MsgBox("今月はS-meisaiファイルがありますか？", _
                  vbYesNo + vbQuestion, "S-meisai確認") = vbYes)

    If hasSmeisai Then
        Dim smeisaiPath As String
        smeisaiPath = Application.GetOpenFilename( _
            FileFilter:="CSVファイル (*.csv), *.csv", _
            Title:="S-meisai CSVファイルを選択してください")

        If smeisaiPath = "False" Then
            MsgBox "キャンセルされました。S-meisaiなしで続行します。", vbInformation
            hasSmeisai = False
        Else
            fileNo = FreeFile
            Open smeisaiPath For Input As #fileNo

            Dim sHeaderLine As String
            Line Input #fileNo, sHeaderLine
            Dim sHeaders() As String
            sHeaders = SplitCSVLine(sHeaderLine)

            Dim colSEmpNo As Integer:   colSEmpNo = -1
            Dim colSAmount As Integer:  colSAmount = -1

            Dim sh As Integer
            For sh = 0 To UBound(sHeaders)
                Select Case Trim(sHeaders(sh))
                    Case "社員番号":    colSEmpNo = sh
                    Case "口座1振込額": colSAmount = sh
                End Select
            Next sh

            If colSEmpNo = -1 Or colSAmount = -1 Then
                Close #fileNo
                MsgBox "S-meisai CSVのヘッダーが想定と異なります。" & Chr(10) & _
                       "以下の列が必要です：社員番号、口座1振込額", vbExclamation
                Exit Sub
            End If

            Dim sDataLine As String
            Dim sFields() As String
            Do While Not EOF(fileNo)
                Line Input #fileNo, sDataLine
                If Trim(sDataLine) = "" Then GoTo NextSmeisaiLine

                sFields = SplitCSVLine(sDataLine)
                Dim sEmpNo As String
                sEmpNo = Trim(sFields(colSEmpNo))
                If sEmpNo = "" Then GoTo NextSmeisaiLine

                smeisaiDict(sEmpNo) = Trim(sFields(colSAmount))

NextSmeisaiLine:
            Loop
            Close #fileNo
        End If
    End If

    '===========================================================
    ' ②-3 退職者口座補完CSVの読み込み（任意）
    '   jinjerに口座情報がない退職者用のマスタ
    '   キー: 社員番号  値: 配列(銀行名, 支店名, 口座番号, 口座名義カナ)
    '===========================================================
    Dim suppDict As Object
    Set suppDict = CreateObject("Scripting.Dictionary")

    Dim hasSupp As Boolean
    hasSupp = (MsgBox("退職者の口座補完CSVがありますか？", _
               vbYesNo + vbQuestion, "退職者口座補完CSV") = vbYes)

    If hasSupp Then
        Dim suppPath As String
        suppPath = Application.GetOpenFilename( _
            FileFilter:="CSVファイル (*.csv), *.csv", _
            Title:="退職者口座補完CSVを選択してください")

        If suppPath = "False" Then
            MsgBox "キャンセルされました。補完CSVなしで続行します。", vbInformation
            hasSupp = False
        Else
            fileNo = FreeFile
            Open suppPath For Input As #fileNo

            Dim suppHeaderLine As String
            Line Input #fileNo, suppHeaderLine
            Dim suppHeaders() As String
            suppHeaders = SplitCSVLine(suppHeaderLine)

            Dim colSupEmpNo   As Integer: colSupEmpNo   = -1
            Dim colSupBank    As Integer: colSupBank    = -1
            Dim colSupBranch  As Integer: colSupBranch  = -1
            Dim colSupAcctNo  As Integer: colSupAcctNo  = -1
            Dim colSupAcctNm  As Integer: colSupAcctNm  = -1

            Dim sph As Integer
            For sph = 0 To UBound(suppHeaders)
                Select Case Trim(suppHeaders(sph))
                    Case "社員番号"
                        colSupEmpNo  = sph
                    Case "振込先銀行名(銀行振込口座1)", "銀行名"
                        colSupBank   = sph
                    Case "振込先銀行支店名(銀行振込口座1)", "支店名"
                        colSupBranch = sph
                    Case "口座番号(銀行振込口座1)", "口座番号"
                        colSupAcctNo = sph
                    Case "口座名義（ﾌﾘｶﾞﾅ）(銀行振込口座1)", "口座名義カナ"
                        colSupAcctNm = sph
                End Select
            Next sph

            If colSupEmpNo = -1 Or colSupBank = -1 Or colSupBranch = -1 Or _
               colSupAcctNo = -1 Or colSupAcctNm = -1 Then
                Close #fileNo
                MsgBox "退職者口座補完CSVのヘッダーが想定と異なります。" & Chr(10) & _
                       "以下の列が必要です（jinjer形式または簡略形式）：" & Chr(10) & _
                       "・社員番号" & Chr(10) & _
                       "・銀行名（または 振込先銀行名(銀行振込口座1)）" & Chr(10) & _
                       "・支店名（または 振込先銀行支店名(銀行振込口座1)）" & Chr(10) & _
                       "・口座番号（または 口座番号(銀行振込口座1)）" & Chr(10) & _
                       "・口座名義カナ（または 口座名義（ﾌﾘｶﾞﾅ）(銀行振込口座1)）", vbExclamation
                Exit Sub
            End If

            Dim suppDataLine As String
            Dim suppFields() As String
            Do While Not EOF(fileNo)
                Line Input #fileNo, suppDataLine
                If Trim(suppDataLine) = "" Then GoTo NextSuppLine

                suppFields = SplitCSVLine(suppDataLine)
                Dim suppEmpNo As String
                suppEmpNo = Trim(suppFields(colSupEmpNo))
                If suppEmpNo = "" Then GoTo NextSuppLine

                Dim suppInfo(0 To 3) As String
                suppInfo(0) = Trim(suppFields(colSupBank))
                suppInfo(1) = Trim(suppFields(colSupBranch))
                suppInfo(2) = Trim(suppFields(colSupAcctNo))
                suppInfo(3) = Trim(suppFields(colSupAcctNm))
                suppDict(suppEmpNo) = suppInfo

NextSuppLine:
            Loop
            Close #fileNo
        End If
    End If

    '===========================================================
    ' ③ 給与振込一覧シートへの転記
    '===========================================================
    
    '--- 従業員コード「7777777」の行を検索 ---
    Dim specialRow As Long
    specialRow = 0
    Dim searchRow As Long
    Dim searchLastRow As Long
    searchLastRow = wsOutput.Cells(Rows.Count, 3).End(xlUp).Row
    If searchLastRow >= 7 Then
        For searchRow = 7 To searchLastRow
            If CStr(wsOutput.Cells(searchRow, 3).Value) = "7777777" Then
                specialRow = searchRow
                Exit For
            End If
        Next searchRow
    End If

    Dim clearMsg As String
    clearMsg = ""
    Dim clearFrom As Long
    Dim clearTo As Long

    If specialRow > 0 Then
        Dim availableRows As Long
        availableRows = specialRow - 7  ' 7777777より上に使える行数

        If qOrderCount > availableRows Then
            ' 不足分だけ7777777の上に行を挿入
            Dim insertCount As Long
            insertCount = qOrderCount - availableRows
            wsOutput.Rows(specialRow & ":" & (specialRow + insertCount - 1)).Insert Shift:=xlDown
            specialRow = specialRow + insertCount  ' 挿入後の7777777位置を更新
        ElseIf qOrderCount < availableRows Then
            ' 余剰行の内容をクリア
            clearFrom = 7 + qOrderCount
            clearTo = specialRow - 1
            wsOutput.Range("A" & clearFrom & ":J" & clearTo).ClearContents
            clearMsg = clearTo - clearFrom + 1 & " 行のデータをクリアしました（" & clearFrom & "行〜" & clearTo & "行）"
        End If
    Else
        '--- 既存データをクリア（Row7以降、新データ行より下） ---
        Dim outLastRow As Long
        outLastRow = wsOutput.Cells(Rows.Count, 1).End(xlUp).Row
        If outLastRow >= 7 + qOrderCount Then
            clearFrom = 7 + qOrderCount
            clearTo = outLastRow
            wsOutput.Range("A" & clearFrom & ":J" & clearTo).ClearContents
            clearMsg = clearTo - clearFrom + 1 & " 行のデータをクリアしました（" & clearFrom & "行〜" & clearTo & "行）"
        End If
    End If

    ' 書き込み終了予定行を記録（支給額0スキップ後のクリア用）
    Dim expectedEndRow As Long
    If specialRow > 0 Then
        expectedEndRow = specialRow - 1
    Else
        expectedEndRow = 6 + qOrderCount
    End If

    '--- Q-meisaiデータを順番に処理 → 給与振込一覧シートに転記 ---
    Dim outputRow As Long
    outputRow = 7  ' データ開始行
    
    Dim totalCount As Long:  totalCount = 0
    Dim errCount As Long:    errCount = 0
    Dim errMsg As String:    errMsg = ""
    
    Dim idx As Long
    For idx = 0 To qOrderCount - 1
        Dim curEmpNo As String
        curEmpNo = qEmpOrder(idx)
        
        ' Q-meisaiから氏名・振込額を取得
        Dim qData As Variant
        qData = qmeisaiDict(curEmpNo)
        Dim qName As String:      qName = qData(0)
        Dim qAmountStr As String: qAmountStr = qData(1)

        ' S-meisaiがある場合、Q-meisaiと合算
        Dim finalAmount As Variant
        If hasSmeisai And smeisaiDict.Exists(curEmpNo) Then
            Dim sAmountStr As String: sAmountStr = smeisaiDict(curEmpNo)
            Dim qAmt As Double: qAmt = 0
            Dim sAmt As Double: sAmt = 0
            If IsNumeric(qAmountStr) Then qAmt = CDbl(qAmountStr)
            If IsNumeric(sAmountStr) Then sAmt = CDbl(sAmountStr)
            finalAmount = qAmt + sAmt
        Else
            finalAmount = qAmountStr
        End If
        
        ' 支給額が0の場合はスキップ
        Dim skipAmt As Double
        skipAmt = 0
        If IsNumeric(finalAmount) Then skipAmt = CDbl(finalAmount)
        If skipAmt = 0 Then GoTo NextEmployee

        ' jinjerから口座情報を取得
        Dim bankName As String:   bankName = ""
        Dim branchName As String: branchName = ""
        Dim acctNumber As String: acctNumber = ""
        Dim acctHolder As String: acctHolder = ""
        
        If jinjerDict.Exists(curEmpNo) Then
            Dim info As Variant
            info = jinjerDict(curEmpNo)
            bankName   = info(0)
            branchName = info(1)
            acctNumber = info(2)
            acctHolder = info(3)
        ElseIf hasSupp And suppDict.Exists(curEmpNo) Then
            ' jinjerになければ退職者口座補完CSVを参照
            Dim suppInfoData As Variant
            suppInfoData = suppDict(curEmpNo)
            bankName   = suppInfoData(0)
            branchName = suppInfoData(1)
            acctNumber = suppInfoData(2)
            acctHolder = suppInfoData(3)
        Else
            errCount = errCount + 1
            errMsg = errMsg & "・社員番号「" & curEmpNo & "」(" & qName & _
                     ") の口座情報がjinjer・補完CSV双方に見つかりません。" & Chr(10)
        End If
        
        ' 給与振込一覧シートに転記
        wsOutput.Cells(outputRow, 1).Value = Format(payDate, "yyyy/m/d")  ' 支給日
        wsOutput.Cells(outputRow, 2).Value = "給与"                       ' 振込種別（固定）
        wsOutput.Cells(outputRow, 3).Value = curEmpNo                     ' 従業員コード
        wsOutput.Cells(outputRow, 4).Value = qName                        ' 従業員名
        ' 金融機関名の表記統一
        If bankName = "中日信金" Then bankName = "中日信用金庫"

        ' 金融機関名に「銀行」を付加
        ' （末尾が 銀行／労金／農協／信金／信組／信用金庫／信用組合 でなければ付加）
        Dim displayBankName As String
        Dim bankSuffixes As Variant
        bankSuffixes = Array("銀行", "労金", "農協", "信金", "信組", "信用金庫", "信用組合", "農業協同組合")
        Dim bsIdx As Integer
        Dim needsBankSuffix As Boolean
        needsBankSuffix = True
        For bsIdx = 0 To UBound(bankSuffixes)
            If Right(bankName, Len(bankSuffixes(bsIdx))) = bankSuffixes(bsIdx) Then
                needsBankSuffix = False
                Exit For
            End If
        Next bsIdx
        If needsBankSuffix Then
            displayBankName = bankName & "銀行"
        Else
            displayBankName = bankName
        End If
        wsOutput.Cells(outputRow, 5).Value = displayBankName              ' 金融機関名
        ' 支店名に「支店」を付加
        ' （ゆうちょ銀行以外 かつ 「本店」を含まない かつ 「営業部」で終わらない場合）
        Dim displayBranch As String
        If displayBankName <> "ゆうちょ銀行" And InStr(branchName, "本店") = 0 _
           And Right(branchName, 3) <> "営業部" Then
            displayBranch = branchName & "支店"
        Else
            displayBranch = branchName
        End If
        wsOutput.Cells(outputRow, 6).Value = displayBranch                ' 支店名
        acctNumber = NormalizeAccountNumber(acctNumber)

        wsOutput.Cells(outputRow, 7).Value = "普通"                       ' 口座種別（固定）
        With wsOutput.Cells(outputRow, 8)
            .NumberFormatLocal = "@"
            .Value = acctNumber                                           ' 口座番号（文字列として保持）
        End With
        wsOutput.Cells(outputRow, 9).Value = acctHolder                   ' 口座名義
        wsOutput.Cells(outputRow, 10).Value = finalAmount                  ' 支給額
        
        outputRow = outputRow + 1
        totalCount = totalCount + 1
NextEmployee:
    Next idx

    ' 支給額0スキップにより余剰行が生じた場合はクリア
    If outputRow <= expectedEndRow Then
        wsOutput.Range("A" & outputRow & ":J" & expectedEndRow).ClearContents
        If clearMsg <> "" Then clearMsg = clearMsg & Chr(10)
        clearMsg = clearMsg & (expectedEndRow - outputRow + 1) & _
                   " 行を支給額0のためクリアしました（" & outputRow & "行〜" & expectedEndRow & "行）"
    End If
    
    '--- 振込件数・振込人数・支給額合計を更新（Row4） ---
    wsOutput.Cells(4, 1).Value = totalCount & " 件"
    wsOutput.Cells(4, 2).Value = totalCount & " 人"
    ' 支給額合計は合計計算
    If totalCount > 0 Then
        Dim totalAmount As Double
        totalAmount = 0
        Dim calcRow As Long
        For calcRow = 7 To 7 + totalCount - 1
            If IsNumeric(wsOutput.Cells(calcRow, 10).Value) Then
                totalAmount = totalAmount + CDbl(wsOutput.Cells(calcRow, 10).Value)
            End If
        Next calcRow
        wsOutput.Cells(4, 3).Value = Format(totalAmount, "#,##0") & " 円"
    Else
        wsOutput.Cells(4, 3).Value = "0 円"
    End If
    
    '===========================================================
    ' ③-2 データ行の空白行を削除（行7以降）
    '===========================================================
    Dim delRow As Long
    For delRow = wsOutput.Cells(Rows.Count, 1).End(xlUp).Row To 7 Step -1
        If Trim(CStr(wsOutput.Cells(delRow, 3).Value)) = "" Then
            wsOutput.Rows(delRow).Delete
        End If
    Next delRow

    '===========================================================
    ' ④ CSV保存（楽たすインポート用）
    '===========================================================
    Dim csvSavedPath As String
    csvSavedPath = ""

    If totalCount > 0 Then
        Dim csvFileName As String
        csvFileName = Month(payDate) & "月振込額.csv"

        Dim csvSavePath As String
        csvSavePath = Application.GetSaveAsFilename( _
            InitialFileName:=csvFileName, _
            FileFilter:="CSVファイル (*.csv), *.csv", _
            Title:="CSVの保存先を選択してください（楽たすインポート用）")

        If csvSavePath <> "False" Then
            ' 保存先フォルダが存在しない場合は自動作成
            Dim fso As Object
            Set fso = CreateObject("Scripting.FileSystemObject")
            Call EnsureFolderExists(fso.GetParentFolderName(csvSavePath), fso)

            Dim csvNo As Integer
            csvNo = FreeFile
            On Error Resume Next
            Open csvSavePath For Output As #csvNo
            If Err.Number <> 0 Then
                On Error GoTo 0
                MsgBox "CSVファイルへの書き込みに失敗しました。" & Chr(10) & Chr(10) & _
                       "原因として考えられること：" & Chr(10) & _
                       "・同名のファイルが Excel や メモ帳 で開かれている" & Chr(10) & _
                       "・保存先フォルダへの書き込み権限がない" & Chr(10) & Chr(10) & _
                       "保存先: " & csvSavePath, vbExclamation
                Exit Sub
            End If
            On Error GoTo 0

            ' シートの内容をそのまま出力（行1〜実際の最終データ行）
            Dim csvEndRow As Long
            csvEndRow = wsOutput.Cells(Rows.Count, 1).End(xlUp).Row
            Dim csvRow As Long
            For csvRow = 1 To csvEndRow
                Dim csvLine As String
                csvLine = ""
                Dim csvCol As Integer
                For csvCol = 1 To 10
                    Dim cellVal As String
                    Dim cellRaw As Variant
                    cellRaw = wsOutput.Cells(csvRow, csvCol).Value
                    If IsEmpty(cellRaw) Then
                        cellVal = ""
                    Else
                        cellVal = CStr(cellRaw)
                    End If
                    ' カンマ・改行・ダブルクォートを含む場合はクォートで囲む
                    If InStr(cellVal, ",") > 0 Or InStr(cellVal, """") > 0 _
                       Or InStr(cellVal, Chr(10)) > 0 Then
                        cellVal = """" & Replace(cellVal, """", """""") & """"
                    End If
                    If csvCol = 1 Then
                        csvLine = cellVal
                    Else
                        csvLine = csvLine & "," & cellVal
                    End If
                Next csvCol
                Print #csvNo, csvLine
            Next csvRow

            Close #csvNo
            csvSavedPath = csvSavePath
        End If
    End If

    '--- 完了メッセージ ---
    Dim msg As String
    msg = "給与振込データの転記が完了しました！" & Chr(10) & Chr(10) & _
          "転記件数: " & totalCount & " 件" & Chr(10) & _
          "支給日: " & Format(payDate, "yyyy/m/d") & Chr(10) & _
          "Q-meisai読込件数: " & qmeisaiDict.Count & " 件" & Chr(10) & _
          "jinjer読込件数: " & jinjerDict.Count & " 件"

    If csvSavedPath <> "" Then
        msg = msg & Chr(10) & Chr(10) & _
              "CSV保存先: " & Chr(10) & csvSavedPath
    End If

    If clearMsg <> "" Then
        msg = msg & Chr(10) & Chr(10) & "【クリア情報】" & Chr(10) & clearMsg
    End If

    If errCount > 0 Then
        msg = msg & Chr(10) & Chr(10) & _
              "⚠ 以下の問題が見つかりました（" & errCount & "件）:" & Chr(10) & errMsg
        MsgBox msg, vbExclamation, "転記完了（要確認）"
    Else
        MsgBox msg, vbInformation, "転記完了"
    End If
    
End Sub


'-------------------------------------------------------
' 口座番号を正規化（先頭ゼロ保持）
'-------------------------------------------------------
Private Function NormalizeAccountNumber(ByVal rawAccount As String) As String
    Dim s As String
    s = Trim(rawAccount)

    If Left$(s, 1) = "'" Then s = Mid$(s, 2)
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    s = Replace(s, "-", "")

    If s = "" Then
        NormalizeAccountNumber = ""
        Exit Function
    End If

    If IsNumeric(s) Then
        If Len(s) < 7 Then
            s = Right$("0000000" & s, 7)
        End If
    End If

    NormalizeAccountNumber = s
End Function


'-------------------------------------------------------
' CSV行をカンマ区切りで分割（ダブルクォート対応）
'-------------------------------------------------------
Private Function SplitCSVLine(ByVal csvLine As String) As String()
    Dim result() As String
    Dim fieldCount As Integer
    Dim i As Long
    Dim inQuote As Boolean
    Dim currentField As String
    
    fieldCount = 0
    inQuote = False
    currentField = ""
    
    For i = 1 To Len(csvLine)
        Dim ch As String
        ch = Mid(csvLine, i, 1)
        
        If ch = """" Then
            inQuote = Not inQuote
        ElseIf ch = "," And Not inQuote Then
            ReDim Preserve result(fieldCount)
            result(fieldCount) = currentField
            fieldCount = fieldCount + 1
            currentField = ""
        Else
            currentField = currentField & ch
        End If
    Next i
    
    ' 最後のフィールド
    ReDim Preserve result(fieldCount)
    result(fieldCount) = currentField
    
    SplitCSVLine = result
End Function


'-------------------------------------------------------
' フォルダを再帰的に作成（存在しない階層をまとめて作る）
'-------------------------------------------------------
Private Sub EnsureFolderExists(ByVal folderPath As String, ByVal fso As Object)
    If fso.FolderExists(folderPath) Then Exit Sub
    EnsureFolderExists fso.GetParentFolderName(folderPath), fso
    fso.CreateFolder folderPath
End Sub


'-------------------------------------------------------
' 補助マクロ：「データ転記」ボタンをシートに追加
' ★ 初回のみ実行してください
'-------------------------------------------------------
Sub ボタン追加()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("給与振込一覧")
    
    ' 既存ボタンがあれば削除
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "btnDataTransfer" Then shp.Delete
    Next shp
    
    ' ボタンを追加
    Dim btn As Shape
    Set btn = ws.Shapes.AddFormControl(xlButtonControl, 380, 3, 160, 28)
    btn.Name = "btnDataTransfer"
    btn.TextFrame.Characters.Text = "給与振込データ転記"
    btn.TextFrame.Characters.Font.Size = 11
    btn.TextFrame.Characters.Font.Bold = True
    btn.OnAction = "給与振込データ転記"
    
    MsgBox "「給与振込データ転記」ボタンをシートに追加しました！", vbInformation
End Sub
