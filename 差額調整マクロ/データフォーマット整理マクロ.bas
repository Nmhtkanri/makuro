Attribute VB_Name = "データフォーマット整理マクロ"
Option Explicit

'------------------------------------------------------------------------------
' 給与明細処理マクロ v6
' - 差額調整は「給与明細(前月実績)」を元に計算
' - 基本給変更リスト(A:社員番号, B:氏名, C:前月基礎時給, D:前月標準時間, E:前月みなし給, F:前月基本給)を参照
' - Step4で当月固定給（基本給/みなし給/各種手当）をデータベース値で上書き
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' メインメニュー表示
'------------------------------------------------------------------------------
Sub 給与明細処理メニュー()
    Dim choice As String
    
    choice = InputBox( _
        "【給与明細処理メニュー】" & vbCrLf & vbCrLf & _
        "1. データベースCSV取り込み" & vbCrLf & _
        "2. 給与明細CSV取り込み" & vbCrLf & _
        "3. 差額調整計算" & vbCrLf & _
        "4. 基本給・みなし給上書き処理" & vbCrLf & _
        "5. 退職者処理" & vbCrLf & _
        "6. データ加工（0埋め処理）" & vbCrLf & _
        "7. 型変換・L列処理" & vbCrLf & _
        "8. CSV出力" & vbCrLf & _
        "9. 全ステップ一括実行" & vbCrLf & vbCrLf & _
        "番号を入力してください:", _
        "給与明細処理")
    
    Select Case choice
        Case "1"
            Call Step1_データベースCSV取り込み
        Case "2"
            Call Step2_給与明細CSV取り込み
        Case "3"
            Call Step3_差額調整計算
        Case "4"
            Call Step4_基本給みなし給上書き
        Case "5"
            Call Step5_退職者処理
        Case "6"
            Call Step6_データ加工
        Case "7"
            Call Step6_5_型変換とL列処理
        Case "8"
            Call Step7_CSV出力
        Case "9"
            Call Step8_全ステップ一括実行
        Case ""
            ' キャンセル
        Case Else
            MsgBox "1～9の番号を入力してください。", vbExclamation
    End Select
End Sub

'------------------------------------------------------------------------------
' Step1: データベースCSV取り込み（2ファイル同時選択・追記対応）
'------------------------------------------------------------------------------
Sub Step1_データベースCSV取り込み()
    Dim filePaths As Variant
    Dim wsDB As Worksheet
    Dim wb As Workbook
    Dim wbCSV As Workbook
    Dim lastRowDB As Long
    Dim lastRowCSV As Long
    Dim lastCol As Long
    Dim i As Long
    Dim totalCount As Long
    Dim fileCount As Long
    
    filePaths = Application.GetOpenFilename( _
        FileFilter:="CSVファイル (*.csv),*.csv", _
        Title:="データベースCSVファイルを選択してください（2つまで選択可）", _
        MultiSelect:=True)
    
    If Not IsArray(filePaths) Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If
    
    fileCount = UBound(filePaths) - LBound(filePaths) + 1
    If fileCount > 2 Then
        MsgBox "選択できるファイルは2つまでです。" & vbCrLf & _
               "選択されたファイル数: " & fileCount & " 件", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set wsDB = wb.Sheets("データベース")
    On Error GoTo 0
    
    If wsDB Is Nothing Then
        Set wsDB = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsDB.Name = "データベース"
    Else
        wsDB.Cells.Clear
    End If
    
    totalCount = 0
    
    For i = LBound(filePaths) To UBound(filePaths)
        Set wbCSV = Workbooks.Open(filePaths(i), Local:=True)
        
        With wbCSV.Sheets(1)
            lastRowCSV = .Cells(.Rows.Count, 1).End(xlUp).Row
            lastCol = 65
            
            If i = LBound(filePaths) Then
                .Range(.Cells(1, 1), .Cells(lastRowCSV, lastCol)).Copy wsDB.Range("A1")
                totalCount = totalCount + (lastRowCSV - 1)
            Else
                lastRowDB = wsDB.Cells(wsDB.Rows.Count, 1).End(xlUp).Row + 1
                
                If lastRowCSV >= 2 Then
                    .Range(.Cells(2, 1), .Cells(lastRowCSV, lastCol)).Copy wsDB.Cells(lastRowDB, 1)
                    totalCount = totalCount + (lastRowCSV - 1)
                End If
            End If
        End With
        
        wbCSV.Close SaveChanges:=False
    Next i
    
    Application.ScreenUpdating = True
    
    Dim resultMsg As String
    resultMsg = "データベースCSVの取り込みが完了しました。" & vbCrLf & vbCrLf & _
                "取り込みファイル数: " & fileCount & " 件" & vbCrLf & _
                "取り込みデータ件数: " & totalCount & " 件"
    
    resultMsg = resultMsg & vbCrLf & vbCrLf & "【取り込んだファイル】"
    For i = LBound(filePaths) To UBound(filePaths)
        resultMsg = resultMsg & vbCrLf & (i - LBound(filePaths) + 1) & ". " & Dir(filePaths(i))
    Next i
    
    MsgBox resultMsg, vbInformation
End Sub

'------------------------------------------------------------------------------
' Step2: 給与明細CSV取り込み
'------------------------------------------------------------------------------
Sub Step2_給与明細CSV取り込み()
    Dim filePath As String
    Dim wsMeisai As Worksheet
    Dim wb As Workbook
    Dim wbCSV As Workbook
    Dim lastRow As Long
    Dim lastCol As Long
    
    filePath = Application.GetOpenFilename( _
        FileFilter:="CSVファイル (*.csv),*.csv", _
        Title:="給与明細CSVファイルを選択してください")
    
    If filePath = "False" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set wsMeisai = wb.Sheets("給与明細")
    On Error GoTo 0
    
    If wsMeisai Is Nothing Then
        Set wsMeisai = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsMeisai.Name = "給与明細"
    Else
        wsMeisai.Cells.Clear
    End If
    
    Set wbCSV = Workbooks.Open(filePath, Local:=True)
    
    With wbCSV.Sheets(1)
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Copy wsMeisai.Range("A1")
    End With
    
    wbCSV.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
    MsgBox "給与明細CSVの取り込みが完了しました。" & vbCrLf & _
           "取り込み件数: " & (lastRow - 1) & " 件", vbInformation
End Sub

'------------------------------------------------------------------------------
' Step3: 差額調整計算（前月実績ベース）
'------------------------------------------------------------------------------
Sub Step3_差額調整計算()
    Dim wsMeisai As Worksheet, wsDB As Worksheet, wsChange As Worksheet
    Dim lastRowMeisai As Long, lastRowDB As Long, lastRowChange As Long
    Dim dbMap As Object, chgMap As Object
    Dim i As Long, dbRow As Long
    Dim empKey As String, salaryTitle As String
    Dim m0 As Double, n0 As Double, s0 As Double, t0 As Double, w0 As Double
    Dim actualAmt As Double, provisionalAmt As Double
    Dim monthlyVarAmt As Double, prevMinashiAmt As Double, prevBasicAmt As Double
    Dim info As Variant
    Dim rawPrevAT As String, rawPrevAQ As String, rawPrevMinashi As String
    Dim rawPrevBasic As String
    
    Dim processedCount As Long, hourlyCount As Long, monthlyCount As Long
    Dim dbNotFoundCount As Long, hourlyFallbackCount As Long, monthlyFallbackCount As Long
    
    On Error Resume Next
    Set wsMeisai = ThisWorkbook.Sheets("給与明細")
    Set wsDB = ThisWorkbook.Sheets("データベース")
    Set wsChange = ThisWorkbook.Sheets("基本給変更リスト")
    On Error GoTo 0
    
    If wsMeisai Is Nothing Then
        MsgBox "給与明細シートがありません。先にStep2を実行してください。", vbExclamation
        Exit Sub
    End If
    If wsDB Is Nothing Then
        MsgBox "データベースシートがありません。先にStep1を実行してください。", vbExclamation
        Exit Sub
    End If
    
    lastRowMeisai = wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Row
    lastRowDB = wsDB.Cells(wsDB.Rows.Count, 63).End(xlUp).Row
    Set dbMap = BuildEmpRowMap(wsDB, 63, lastRowDB)
    
    Set chgMap = CreateObject("Scripting.Dictionary")
    If Not wsChange Is Nothing Then
        lastRowChange = wsChange.Cells(wsChange.Rows.Count, 1).End(xlUp).Row
        Set chgMap = BuildChangeMap(wsChange, lastRowChange)
    End If
    
    Application.ScreenUpdating = False
    
    For i = 2 To lastRowMeisai
        empKey = NormalizeEmpID(wsMeisai.Cells(i, 1).Value)
        If empKey = "" Then GoTo ContinueLoop
        
        If Not dbMap.Exists(empKey) Then
            dbNotFoundCount = dbNotFoundCount + 1
            GoTo ContinueLoop
        End If
        
        dbRow = CLng(dbMap(empKey))
        salaryTitle = Trim$(CStr(wsDB.Cells(dbRow, 35).Value))
        
        m0 = ToDbl(wsMeisai.Cells(i, 13).Value)
        n0 = ToDbl(wsMeisai.Cells(i, 14).Value)
        s0 = ToDbl(wsMeisai.Cells(i, 19).Value)
        t0 = ToDbl(wsMeisai.Cells(i, 20).Value)
        w0 = ToDbl(wsMeisai.Cells(i, 23).Value)
        
        rawPrevAT = ""
        rawPrevAQ = ""
        rawPrevMinashi = ""
        rawPrevBasic = ""
        If chgMap.Exists(empKey) Then
            info = chgMap(empKey)
            rawPrevAT = CStr(info(0))
            rawPrevAQ = CStr(info(1))
            rawPrevMinashi = CStr(info(2))
            rawPrevBasic = CStr(info(3))
        End If
        
        If salaryTitle = "時給制" Then
            hourlyCount = hourlyCount + 1
            ' 時給制はX列（調整手当）も差額調整の実績額に含める
            actualAmt = GetActualAmountFromMeisai(wsMeisai, i) + ToDbl(wsMeisai.Cells(i, 24).Value)
            
            If rawPrevAT <> "" And rawPrevAQ <> "" Then
                provisionalAmt = ToDbl(rawPrevAT) * ToDbl(rawPrevAQ)
            Else
                provisionalAmt = ToDbl(wsDB.Cells(dbRow, 46).Value) * ToDbl(wsDB.Cells(dbRow, 43).Value)
                hourlyFallbackCount = hourlyFallbackCount + 1
            End If
        Else
            monthlyCount = monthlyCount + 1
            
            If rawPrevMinashi <> "" Then
                prevMinashiAmt = ToDbl(rawPrevMinashi)
            Else
                prevMinashiAmt = n0
                monthlyFallbackCount = monthlyFallbackCount + 1
            End If
            
            ' 月給制は固定給差分を打ち消し、変動分（O/P/Q）のみを差額調整に反映
            monthlyVarAmt = ToDbl(wsMeisai.Cells(i, 15).Value) + _
                            ToDbl(wsMeisai.Cells(i, 16).Value) + _
                            ToDbl(wsMeisai.Cells(i, 17).Value)
            If rawPrevBasic <> "" Then
                prevBasicAmt = ToDbl(rawPrevBasic)
            Else
                prevBasicAmt = m0
            End If
            
            provisionalAmt = prevBasicAmt + prevMinashiAmt + s0 + t0 + w0
            actualAmt = m0 + n0 + monthlyVarAmt + s0 + t0 + w0
        End If
        
        wsMeisai.Cells(i, 27).Value = actualAmt - provisionalAmt
        processedCount = processedCount + 1
        
ContinueLoop:
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "差額調整計算が完了しました。" & vbCrLf & _
           "処理件数: " & processedCount & " 件" & vbCrLf & _
           "時給制: " & hourlyCount & " 件（変更リスト未設定フォールバック: " & hourlyFallbackCount & " 件）" & vbCrLf & _
           "月給制: " & monthlyCount & " 件（前月みなし給/基本給未設定フォールバック: " & monthlyFallbackCount & " 件）" & vbCrLf & _
           "DB未マッチ: " & dbNotFoundCount & " 件", vbInformation
End Sub

'------------------------------------------------------------------------------
' Step4: 基本給・みなし給・手当 上書き（当月値）
'------------------------------------------------------------------------------
Sub Step4_基本給みなし給上書き()
    Dim wsMeisai As Worksheet, wsDB As Worksheet, wsDate As Worksheet
    Dim lastRowMeisai As Long, lastRowDB As Long
    Dim dbMap As Object
    Dim i As Long, dbRow As Long
    Dim empNo As String, empKey As String
    Dim matchCount As Long, notFoundCount As Long, futureSkipCount As Long
    Dim notFoundList As String, futureSkipList As String
    Dim yearVal As Long, monthVal As Long, dayVal As Long
    Dim targetDate As Date, effectiveDate As Date
    
    On Error Resume Next
    Set wsMeisai = ThisWorkbook.Sheets("給与明細")
    Set wsDB = ThisWorkbook.Sheets("データベース")
    Set wsDate = ThisWorkbook.Sheets("年月日設定")
    On Error GoTo 0
    
    If wsMeisai Is Nothing Then
        MsgBox "給与明細シートがありません。先にStep2を実行してください。", vbExclamation
        Exit Sub
    End If
    If wsDB Is Nothing Then
        MsgBox "データベースシートがありません。先にStep1を実行してください。", vbExclamation
        Exit Sub
    End If
    If wsDate Is Nothing Then
        MsgBox "年月日設定シートがありません。上書きを中断します。", vbExclamation
        Exit Sub
    End If
    
    yearVal = Val(wsDate.Cells(2, 1).Value)
    monthVal = Val(wsDate.Cells(2, 2).Value)
    dayVal = Val(wsDate.Cells(2, 3).Value)
    
    If yearVal = 0 Or monthVal = 0 Or dayVal = 0 Then
        MsgBox "年月日設定シートのA2:年、B2:月、C2:日を入力してください。", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    targetDate = DateSerial(yearVal, monthVal, dayVal)
    On Error GoTo 0
    If targetDate = 0 Then
        MsgBox "年月日設定シートの日付が不正です。", vbExclamation
        Exit Sub
    End If
    
    lastRowMeisai = wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Row
    lastRowDB = wsDB.Cells(wsDB.Rows.Count, 63).End(xlUp).Row
    Set dbMap = BuildEmpRowMap(wsDB, 63, lastRowDB)
    
    Application.ScreenUpdating = False
    
    For i = 2 To lastRowMeisai
        empNo = Trim$(CStr(wsMeisai.Cells(i, 1).Value))
        empKey = NormalizeEmpID(wsMeisai.Cells(i, 1).Value)
        If empKey = "" Then GoTo ContinueStep4
        
        If dbMap.Exists(empKey) Then
            dbRow = CLng(dbMap(empKey))
            
            ' DB H列（8列目）が基準日より未来なら当月上書きをスキップ
            If IsDate(wsDB.Cells(dbRow, 8).Value) Then
                effectiveDate = CDate(wsDB.Cells(dbRow, 8).Value)
                If effectiveDate > targetDate Then
                    futureSkipCount = futureSkipCount + 1
                    If futureSkipList <> "" Then futureSkipList = futureSkipList & ", "
                    futureSkipList = futureSkipList & empNo
                    GoTo ContinueStep4
                End If
            End If
            
            wsMeisai.Cells(i, 13).Value = ToDbl(wsDB.Cells(dbRow, 42).Value)
            wsMeisai.Cells(i, 14).Value = ToDbl(wsDB.Cells(dbRow, 49).Value)
            wsMeisai.Cells(i, 19).Value = ToDbl(wsDB.Cells(dbRow, 38).Value)
            wsMeisai.Cells(i, 20).Value = ToDbl(wsDB.Cells(dbRow, 39).Value)
            wsMeisai.Cells(i, 23).Value = ToDbl(wsDB.Cells(dbRow, 40).Value)
            wsMeisai.Cells(i, 24).Value = ToDbl(wsDB.Cells(dbRow, 37).Value)
            
            matchCount = matchCount + 1
        Else
            notFoundCount = notFoundCount + 1
            If notFoundList <> "" Then notFoundList = notFoundList & ", "
            notFoundList = notFoundList & empNo
        End If
        
ContinueStep4:
    Next i
    
    Application.ScreenUpdating = True
    
    Dim msg As String
    msg = "当月固定給の上書きが完了しました。" & vbCrLf & _
          "基準日: " & Format$(targetDate, "yyyy/mm/dd") & vbCrLf & _
          "上書き件数: " & matchCount & " 件" & vbCrLf & _
          "未来日スキップ件数: " & futureSkipCount & " 件" & vbCrLf & _
          "未マッチ件数: " & notFoundCount & " 件"
    
    If futureSkipCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "【未来日でスキップした社員番号】" & vbCrLf & futureSkipList
    End If
    
    If notFoundCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "【未マッチ社員番号】" & vbCrLf & notFoundList
    End If
    
    MsgBox msg, vbInformation
End Sub

'------------------------------------------------------------------------------
' Step5: 退職者処理
'------------------------------------------------------------------------------
Sub Step5_退職者処理()
    Dim wsMeisai As Worksheet
    Dim wsRetired As Worksheet
    Dim lastRowMeisai As Long
    Dim lastRowRetired As Long
    Dim i As Long, j As Long
    Dim empNoMeisai As String, empKeyMeisai As String
    Dim empNoRetired As String, empKeyRetired As String
    Dim processedCount As Long
    Dim processedList As String
    
    On Error Resume Next
    Set wsMeisai = ThisWorkbook.Sheets("給与明細")
    Set wsRetired = ThisWorkbook.Sheets("退職者リスト")
    On Error GoTo 0
    
    If wsMeisai Is Nothing Then
        MsgBox "給与明細シートがありません。" & vbCrLf & _
               "先にStep2を実行してください。", vbExclamation
        Exit Sub
    End If
    
    If wsRetired Is Nothing Then
        MsgBox "退職者リストシートがありません。" & vbCrLf & _
               "「退職者リスト」シートを作成してください。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    lastRowMeisai = wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Row
    lastRowRetired = wsRetired.Cells(wsRetired.Rows.Count, 1).End(xlUp).Row
    
    processedCount = 0
    processedList = ""
    
    For i = 2 To lastRowMeisai
        empNoMeisai = Trim(CStr(wsMeisai.Cells(i, 1).Value))
        empKeyMeisai = NormalizeEmpID(wsMeisai.Cells(i, 1).Value)
        If empKeyMeisai = "" Then GoTo ContinueStep5
        
        For j = 2 To lastRowRetired
            empNoRetired = Trim(CStr(wsRetired.Cells(j, 1).Value))
            empKeyRetired = NormalizeEmpID(wsRetired.Cells(j, 1).Value)
            
            If empKeyRetired <> "" And empKeyRetired = empKeyMeisai Then
                wsMeisai.Cells(i, 13).Value = 0
                wsMeisai.Cells(i, 14).Value = 0
                
                processedCount = processedCount + 1
                If processedList <> "" Then processedList = processedList & ", "
                processedList = processedList & empNoMeisai
                Exit For
            End If
        Next j
ContinueStep5:
    Next i
    
    Application.ScreenUpdating = True
    
    Dim resultMsg As String
    resultMsg = "退職者処理が完了しました。" & vbCrLf & _
                "処理件数: " & processedCount & " 件"
    
    If processedCount > 0 Then
        resultMsg = resultMsg & vbCrLf & vbCrLf & _
                    "【処理した社員番号】" & vbCrLf & processedList
    End If
    
    MsgBox resultMsg, vbInformation
End Sub

'------------------------------------------------------------------------------
' Step6: データ加工（0埋め処理）
'------------------------------------------------------------------------------
Sub Step6_データ加工()
    Dim wsMeisai As Worksheet
    Dim wsDB As Worksheet
    Dim lastRowMeisai As Long
    Dim lastRowDB As Long
    Dim dbMap As Object
    Dim i As Long
    Dim empKey As String
    Dim dbRow As Long
    Dim koyouKeitai As String
    
    On Error Resume Next
    Set wsMeisai = ThisWorkbook.Sheets("給与明細")
    Set wsDB = ThisWorkbook.Sheets("データベース")
    On Error GoTo 0
    
    If wsMeisai Is Nothing Then
        MsgBox "給与明細シートがありません。" & vbCrLf & _
               "先にStep2を実行してください。", vbExclamation
        Exit Sub
    End If
    
    If wsDB Is Nothing Then
        MsgBox "データベースシートがありません。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    lastRowMeisai = wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Row
    lastRowDB = wsDB.Cells(wsDB.Rows.Count, 63).End(xlUp).Row
    Set dbMap = BuildEmpRowMap(wsDB, 63, lastRowDB)
    
    For i = 2 To lastRowMeisai
        wsMeisai.Cells(i, 15).Value = 0
        wsMeisai.Cells(i, 16).Value = 0
        wsMeisai.Cells(i, 17).Value = 0
        
        empKey = NormalizeEmpID(wsMeisai.Cells(i, 1).Value)
        If empKey = "" Then GoTo ContinueStep6
        
        If dbMap.Exists(empKey) Then
            dbRow = CLng(dbMap(empKey))
            koyouKeitai = Trim(CStr(wsDB.Cells(dbRow, 35).Value))
            
            If koyouKeitai = "時給制" Then
                wsMeisai.Cells(i, 14).Value = 0
            End If
        End If
ContinueStep6:
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "データ加工が完了しました。" & vbCrLf & _
           "O列～Q列を0に設定しました。" & vbCrLf & _
           "時給制の場合はN列も0に設定しました。" & vbCrLf & _
           "処理件数: " & (lastRowMeisai - 1) & " 件", vbInformation
End Sub

'------------------------------------------------------------------------------
' Step6_5: 型変換とL列マイナス記号処理
'------------------------------------------------------------------------------
Sub Step6_5_型変換とL列処理()
    Dim wsMeisai As Worksheet
    
    On Error Resume Next
    Set wsMeisai = ThisWorkbook.Sheets("給与明細")
    On Error GoTo 0
    
    If wsMeisai Is Nothing Then
        MsgBox "給与明細シートがありません。" & vbCrLf & _
               "先にStep2を実行してください。", vbExclamation
        Exit Sub
    End If
    
    Call 型を標準にする(wsMeisai)
    Call L列マイナス記号処理(wsMeisai)
    
    MsgBox "型変換とL列処理が完了しました。", vbInformation
End Sub

'------------------------------------------------------------------------------
' 処理A: 型を標準にする（AS列＝45列目を除く）
'------------------------------------------------------------------------------
Private Sub 型を標準にする(ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim rng As Range
    
    Application.ScreenUpdating = False
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If col <> 45 Then
            Set rng = ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col))
            rng.NumberFormat = "General"
            On Error Resume Next
            rng.Value = rng.Value
            On Error GoTo 0
        End If
    Next col
    
    Application.ScreenUpdating = True
End Sub

'------------------------------------------------------------------------------
' 処理B: L列の"-"を0にする（確認ポップアップ付き）
'------------------------------------------------------------------------------
Private Sub L列マイナス記号処理(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim empNo As String
    Dim empName As String
    Dim targetList As String
    Dim targetCount As Long
    Dim targetRows() As Long
    Dim result As VbMsgBoxResult
    
    Application.ScreenUpdating = False
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    targetCount = 0
    targetList = ""
    ReDim targetRows(1 To lastRow)
    
    For i = 2 To lastRow
        cellValue = Trim(CStr(ws.Cells(i, 12).Value))
        
        If InStr(cellValue, "-") > 0 Then
            targetCount = targetCount + 1
            targetRows(targetCount) = i
            
            empNo = Trim(CStr(ws.Cells(i, 1).Value))
            empName = Trim(CStr(ws.Cells(i, 2).Value))
            
            If targetList <> "" Then targetList = targetList & vbCrLf
            targetList = targetList & empNo & "  " & empName & "  （現在値: " & cellValue & "）"
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    If targetCount = 0 Then
        MsgBox "L列（有給休暇残）に ""-"" が含まれるデータはありませんでした。", vbInformation
        Exit Sub
    End If
    
    Dim confirmMsg As String
    confirmMsg = "【L列に ""-"" が含まれる社員が " & targetCount & " 件見つかりました】" & vbCrLf & vbCrLf & _
                 targetList & vbCrLf & vbCrLf & _
                 "上記のL列の値を 0 に変換します。" & vbCrLf & vbCrLf & _
                 "本当に実行しますか？"
    
    result = MsgBox(confirmMsg, vbQuestion + vbYesNo, "L列マイナス記号処理の確認")
    
    If result = vbYes Then
        Application.ScreenUpdating = False
        
        For i = 1 To targetCount
            ws.Cells(targetRows(i), 12).Value = 0
        Next i
        
        Application.ScreenUpdating = True
        
        MsgBox "L列の ""-"" を 0 に変換しました。" & vbCrLf & _
               "処理件数: " & targetCount & " 件", vbInformation
    Else
        MsgBox "L列マイナス記号処理をスキップしました。", vbInformation
    End If
End Sub

'------------------------------------------------------------------------------
' Step7: CSV出力
'------------------------------------------------------------------------------
Sub Step7_CSV出力()
    Dim wsMeisai As Worksheet
    Dim savePath As String
    Dim fileName As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim wbNew As Workbook
    
    On Error Resume Next
    Set wsMeisai = ThisWorkbook.Sheets("給与明細")
    On Error GoTo 0
    
    If wsMeisai Is Nothing Then
        MsgBox "給与明細シートがありません。" & vbCrLf & _
               "先にStep2を実行してください。", vbExclamation
        Exit Sub
    End If
    
    fileName = "給与明細_加工済_" & Format(Date, "yyyymmdd") & ".csv"
    
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=fileName, _
        FileFilter:="CSVファイル (*.csv),*.csv", _
        Title:="CSVファイルの保存先を選択してください")
    
    If savePath = "False" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    lastRow = wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Row
    lastCol = wsMeisai.Cells(1, wsMeisai.Columns.Count).End(xlToLeft).Column
    
    Set wbNew = Workbooks.Add
    wsMeisai.Range(wsMeisai.Cells(1, 1), wsMeisai.Cells(lastRow, lastCol)).Copy wbNew.Sheets(1).Range("A1")
    
    wbNew.SaveAs fileName:=savePath, FileFormat:=xlCSV, Local:=True
    wbNew.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "CSVファイルを出力しました。" & vbCrLf & _
           "保存先: " & savePath, vbInformation
End Sub

'------------------------------------------------------------------------------
' Step8: 全ステップ一括実行（文言更新）
'------------------------------------------------------------------------------
Sub Step8_全ステップ一括実行()
    Dim result As VbMsgBoxResult
    
    result = MsgBox("全ステップを一括実行します。" & vbCrLf & vbCrLf & _
                    "1. データベースCSV取り込み" & vbCrLf & _
                    "2. 給与明細CSV取り込み" & vbCrLf & _
                    "3. 差額調整計算（前月実績ベース）" & vbCrLf & _
                    "4. 当月固定給上書き（基本給・みなし給・手当）" & vbCrLf & _
                    "5. 退職者処理" & vbCrLf & _
                    "6. データ加工（0埋め処理）" & vbCrLf & _
                    "7. 型変換・L列処理" & vbCrLf & _
                    "8. CSV出力" & vbCrLf & vbCrLf & _
                    "実行しますか?", _
                    vbQuestion + vbYesNo, "一括実行確認")
    
    If result = vbNo Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If
    
    Call Step1_データベースCSV取り込み
    If Not SheetExists("データベース") Then
        MsgBox "Step1が完了しませんでした。処理を中断します。", vbExclamation
        Exit Sub
    End If
    
    Call Step2_給与明細CSV取り込み
    If Not SheetExists("給与明細") Then
        MsgBox "Step2が完了しませんでした。処理を中断します。", vbExclamation
        Exit Sub
    End If
    
    Call Step3_差額調整計算
    Call Step4_基本給みなし給上書き
    Call Step5_退職者処理
    Call Step6_データ加工
    Call Step6_5_型変換とL列処理
    Call Step7_CSV出力
    
    MsgBox "全ステップの処理が完了しました。", vbInformation
End Sub

'------------------------------------------------------------------------------
' シート存在チェック関数
'------------------------------------------------------------------------------
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

'------------------------------------------------------------------------------
' ヘルパー関数
'------------------------------------------------------------------------------
Private Function BuildEmpRowMap(ByVal ws As Worksheet, ByVal empCol As Long, ByVal lastRow As Long) As Object
    Dim dic As Object
    Dim r As Long
    Dim key As String
    
    Set dic = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        key = NormalizeEmpID(ws.Cells(r, empCol).Value)
        If key <> "" Then
            If Not dic.Exists(key) Then
                dic.Add key, r
            End If
        End If
    Next r
    Set BuildEmpRowMap = dic
End Function

Private Function BuildChangeMap(ByVal ws As Worksheet, ByVal lastRow As Long) As Object
    Dim dic As Object
    Dim r As Long
    Dim key As String
    Dim cVal As String, dVal As String, eVal As String, fVal As String
    
    Set dic = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        key = NormalizeEmpID(ws.Cells(r, 1).Value)
        If key <> "" Then
            cVal = Trim$(CStr(ws.Cells(r, 3).Value))
            dVal = Trim$(CStr(ws.Cells(r, 4).Value))
            eVal = Trim$(CStr(ws.Cells(r, 5).Value))
            fVal = Trim$(CStr(ws.Cells(r, 6).Value))
            dic(key) = Array(cVal, dVal, eVal, fVal)
        End If
    Next r
    Set BuildChangeMap = dic
End Function

Private Function NormalizeEmpID(ByVal v As Variant) As String
    Dim s As String
    
    If IsError(v) Then Exit Function
    
    s = Trim$(CStr(v))
    If s = "" Then Exit Function
    
    On Error Resume Next
    s = StrConv(s, vbNarrow)
    On Error GoTo 0
    
    s = Replace(s, " ", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, ChrW(12288), "")
    s = Replace(s, ",", "")
    
    If IsNumeric(s) Then
        If InStr(1, s, "E", vbTextCompare) > 0 Or InStr(1, s, ".") > 0 Then
            s = Format$(CDbl(s), "0")
        End If
    End If
    
    NormalizeEmpID = s
End Function

Private Function ToDbl(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If Trim$(CStr(v)) = "" Then Exit Function
    ToDbl = Val(Replace(CStr(v), ",", ""))
End Function

Private Function GetActualAmountFromMeisai(ByVal ws As Worksheet, ByVal rowNo As Long) As Double
    GetActualAmountFromMeisai = _
        ToDbl(ws.Cells(rowNo, 13).Value) + _
        ToDbl(ws.Cells(rowNo, 14).Value) + _
        ToDbl(ws.Cells(rowNo, 15).Value) + _
        ToDbl(ws.Cells(rowNo, 16).Value) + _
        ToDbl(ws.Cells(rowNo, 17).Value) + _
        ToDbl(ws.Cells(rowNo, 19).Value) + _
        ToDbl(ws.Cells(rowNo, 20).Value) + _
        ToDbl(ws.Cells(rowNo, 23).Value)
End Function


