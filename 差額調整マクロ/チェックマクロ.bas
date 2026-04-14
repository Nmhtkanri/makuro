Attribute VB_Name = "チェックマクロ"
Option Explicit

'------------------------------------------------------------------------------
' 基本給チェックモジュール
' - 給与明細(M列)を「前月給与明細(M列)」と突合
' - 給与明細(M列)を「データベース(AP列)」と突合
' - DB H列が基準日より未来なら「未来日スキップ対象」として判定を分離
'------------------------------------------------------------------------------

Public Sub 基本給チェック実行()
    Dim wsMeisai As Worksheet
    Dim wsDB As Worksheet
    Dim wsDate As Worksheet
    Dim wsPrev As Worksheet
    Dim wsOut As Worksheet
    
    Dim prevSheetName As String
    Dim outSheetName As String
    
    Dim targetDate As Date
    Dim yearVal As Long, monthVal As Long, dayVal As Long
    
    Dim dbMap As Object, prevMap As Object
    Dim lastMeisai As Long, lastDB As Long, lastPrev As Long
    
    Dim i As Long, outRow As Long
    Dim empNo As String, empKey As String, empName As String
    Dim dbRow As Long, prevRow As Long
    
    Dim mCurrent As Double, mPrev As Double, mDB As Double
    Dim hasDB As Boolean, hasPrev As Boolean
    Dim hasDBDate As Boolean, isFuture As Boolean
    Dim dbDate As Date
    
    Dim prevJudge As String, dbJudge As String, totalJudge As String, note As String
    
    Dim totalCount As Long, okCount As Long, ngCount As Long, warnCount As Long
    Dim prevMissingCount As Long, prevMismatchCount As Long
    Dim dbMissingCount As Long, dbMismatchCount As Long, futureCount As Long
    
    On Error Resume Next
    Set wsMeisai = ThisWorkbook.Sheets("給与明細（当月）")
    Set wsDB = ThisWorkbook.Sheets("データベース")
    Set wsDate = ThisWorkbook.Sheets("年月日設定")
    On Error GoTo 0
    
    If wsMeisai Is Nothing Then
        MsgBox "給与明細（当月）シートがありません。", vbExclamation
        Exit Sub
    End If
    If wsDB Is Nothing Then
        MsgBox "データベースシートがありません。", vbExclamation
        Exit Sub
    End If
    If wsDate Is Nothing Then
        MsgBox "年月日設定シートがありません。", vbExclamation
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
        MsgBox "年月日設定の日付が不正です。", vbExclamation
        Exit Sub
    End If
    
    prevSheetName = InputBox( _
        "前月給与明細シート名を入力してください。" & vbCrLf & _
        "（社員番号=A列、基本給=M列）", _
        "基本給チェック", "前月給与明細")
    If prevSheetName = "" Then Exit Sub
    
    On Error Resume Next
    Set wsPrev = ThisWorkbook.Sheets(prevSheetName)
    On Error GoTo 0
    If wsPrev Is Nothing Then
        MsgBox "前月給与明細シートが見つかりません: " & prevSheetName, vbExclamation
        Exit Sub
    End If
    
    outSheetName = "基本給チェック結果"
    
    lastMeisai = wsMeisai.Cells(wsMeisai.Rows.Count, 1).End(xlUp).Row
    lastDB = wsDB.Cells(wsDB.Rows.Count, 63).End(xlUp).Row
    lastPrev = wsPrev.Cells(wsPrev.Rows.Count, 1).End(xlUp).Row
    
    Set dbMap = Chk_BuildEmpMap(wsDB, 63, lastDB)   ' DB BK
    Set prevMap = Chk_BuildEmpMap(wsPrev, 1, lastPrev)
    
    Application.ScreenUpdating = False
    Set wsOut = Chk_PrepareOutputSheet(outSheetName)
    
    wsOut.Cells(1, 1).Value = "No"
    wsOut.Cells(1, 2).Value = "社員番号"
    wsOut.Cells(1, 3).Value = "氏名"
    wsOut.Cells(1, 4).Value = "当月M(給与明細)"
    wsOut.Cells(1, 5).Value = "前月M(前月給与明細)"
    wsOut.Cells(1, 6).Value = "DB_AP(当月基本給)"
    wsOut.Cells(1, 7).Value = "DB_H(適用日)"
    wsOut.Cells(1, 8).Value = "基準日"
    wsOut.Cells(1, 9).Value = "未来日スキップ対象"
    wsOut.Cells(1, 10).Value = "前月比較"
    wsOut.Cells(1, 11).Value = "DB比較"
    wsOut.Cells(1, 12).Value = "総合判定"
    wsOut.Cells(1, 13).Value = "メモ"
    wsOut.Cells(1, 14).Value = "照合キー"
    wsOut.Cells(1, 15).Value = "DB未検出詳細"
    
    outRow = 2
    
    For i = 2 To lastMeisai
        empNo = Trim$(CStr(wsMeisai.Cells(i, 1).Value))
        empKey = Chk_NormalizeEmpID(wsMeisai.Cells(i, 1).Value)
        If empKey = "" Then GoTo ContinueLoop
        
        empName = Trim$(CStr(wsMeisai.Cells(i, 2).Value))
        mCurrent = Chk_ToDbl(wsMeisai.Cells(i, 13).Value) ' M
        
        hasDB = dbMap.Exists(empKey)
        hasPrev = prevMap.Exists(empKey)
        hasDBDate = False
        isFuture = False
        note = ""
        
        If hasDB Then
            dbRow = CLng(dbMap(empKey))
            mDB = Chk_ToDbl(wsDB.Cells(dbRow, 42).Value) ' AP
            
            If IsDate(wsDB.Cells(dbRow, 8).Value) Then
                hasDBDate = True
                dbDate = CDate(wsDB.Cells(dbRow, 8).Value)
                isFuture = (dbDate > targetDate)
                If isFuture Then futureCount = futureCount + 1
            End If
        Else
            mDB = 0
            dbMissingCount = dbMissingCount + 1
        End If
        
        If hasPrev Then
            prevRow = CLng(prevMap(empKey))
            mPrev = Chk_ToDbl(wsPrev.Cells(prevRow, 13).Value) ' M
        Else
            mPrev = 0
            prevMissingCount = prevMissingCount + 1
        End If
        
        If hasPrev Then
            If Chk_IsSameAmount(mCurrent, mPrev) Then
                prevJudge = "一致"
            Else
                prevJudge = "不一致"
                prevMismatchCount = prevMismatchCount + 1
            End If
        Else
            prevJudge = "前月未検出"
        End If
        
        If hasDB Then
            If isFuture Then
                dbJudge = "比較対象外(未来日)"
            ElseIf Chk_IsSameAmount(mCurrent, mDB) Then
                dbJudge = "一致"
            Else
                dbJudge = "不一致"
                dbMismatchCount = dbMismatchCount + 1
            End If
        Else
            dbJudge = "DB未検出"
        End If
        
        ' 総合判定:
        ' NG: DB未検出 / 前月不一致 / (未来日でないDB不一致)
        ' 要確認: 未来日スキップ対象でDB比較対象外
        ' OK: それ以外で一致
        If (Not hasDB) Or (prevJudge = "不一致") Or ((dbJudge = "不一致") And (Not isFuture)) Then
            totalJudge = "NG"
            ngCount = ngCount + 1
        ElseIf isFuture Then
            totalJudge = "要確認"
            warnCount = warnCount + 1
        Else
            totalJudge = "OK"
            okCount = okCount + 1
        End If
        
        If isFuture Then
            note = "DB H列が基準日より未来のため上書きスキップ対象"
        End If
        
        wsOut.Cells(outRow, 1).Value = outRow - 1
        wsOut.Cells(outRow, 2).Value = empNo
        wsOut.Cells(outRow, 3).Value = empName
        wsOut.Cells(outRow, 4).Value = mCurrent
        If hasPrev Then wsOut.Cells(outRow, 5).Value = mPrev
        If hasDB Then wsOut.Cells(outRow, 6).Value = mDB
        If hasDBDate Then wsOut.Cells(outRow, 7).Value = dbDate
        wsOut.Cells(outRow, 8).Value = targetDate
        wsOut.Cells(outRow, 9).Value = IIf(isFuture, "YES", "NO")
        wsOut.Cells(outRow, 10).Value = prevJudge
        wsOut.Cells(outRow, 11).Value = dbJudge
        wsOut.Cells(outRow, 12).Value = totalJudge
        wsOut.Cells(outRow, 13).Value = note
        wsOut.Cells(outRow, 14).Value = empKey
        If Not hasDB Then
            wsOut.Cells(outRow, 15).Value = "明細A(raw)=" & Chk_ValueToText(wsMeisai.Cells(i, 1).Value) & " / key=" & empKey
        End If
        
        Select Case totalJudge
            Case "NG"
                wsOut.Rows(outRow).Interior.Color = RGB(255, 235, 235)
            Case "要確認"
                wsOut.Rows(outRow).Interior.Color = RGB(255, 248, 220)
            Case Else
                wsOut.Rows(outRow).Interior.Color = RGB(235, 255, 235)
        End Select
        
        outRow = outRow + 1
        totalCount = totalCount + 1
        
ContinueLoop:
    Next i
    
    If outRow > 2 Then
        With wsOut.Range(wsOut.Cells(2, 4), wsOut.Cells(outRow - 1, 6))
            .NumberFormat = "#,##0"
        End With
        With wsOut.Range(wsOut.Cells(2, 7), wsOut.Cells(outRow - 1, 8))
            .NumberFormat = "yyyy/mm/dd"
        End With
    End If
    
    wsOut.Rows(1).Font.Bold = True
    wsOut.Rows(1).Interior.Color = RGB(220, 230, 241)
    wsOut.Columns("A:O").AutoFit
    
    Application.ScreenUpdating = True
    
    MsgBox "基本給チェックが完了しました。" & vbCrLf & vbCrLf & _
           "対象件数: " & totalCount & " 件" & vbCrLf & _
           "OK: " & okCount & " 件" & vbCrLf & _
           "要確認(未来日): " & warnCount & " 件" & vbCrLf & _
           "NG: " & ngCount & " 件" & vbCrLf & vbCrLf & _
           "前月未検出: " & prevMissingCount & " 件" & vbCrLf & _
           "前月不一致: " & prevMismatchCount & " 件" & vbCrLf & _
           "DB未検出: " & dbMissingCount & " 件" & vbCrLf & _
           "DB不一致(未来日除く): " & dbMismatchCount & " 件" & vbCrLf & _
           "未来日スキップ対象: " & futureCount & " 件", vbInformation
End Sub

Private Function Chk_PrepareOutputSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    
    Set Chk_PrepareOutputSheet = ws
End Function

Private Function Chk_BuildEmpMap(ByVal ws As Worksheet, ByVal empCol As Long, ByVal lastRow As Long) As Object
    Dim dic As Object
    Dim r As Long
    Dim key As String
    
    Set dic = CreateObject("Scripting.Dictionary")
    
    For r = 2 To lastRow
        key = Chk_NormalizeEmpID(ws.Cells(r, empCol).Value)
        If key <> "" Then
            If Not dic.Exists(key) Then
                dic.Add key, r
            End If
        End If
    Next r
    
    Set Chk_BuildEmpMap = dic
End Function

Private Function Chk_ToDbl(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If Trim$(CStr(v)) = "" Then Exit Function
    Chk_ToDbl = Val(Replace(CStr(v), ",", ""))
End Function

Private Function Chk_IsSameAmount(ByVal a As Double, ByVal b As Double) As Boolean
    Chk_IsSameAmount = (Abs(a - b) < 0.01)
End Function

Private Function Chk_NormalizeEmpID(ByVal v As Variant) As String
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
    
    Chk_NormalizeEmpID = s
End Function

Private Function Chk_ValueToText(ByVal v As Variant) As String
    If IsError(v) Then
        Chk_ValueToText = "#ERROR"
    Else
        Chk_ValueToText = CStr(v)
    End If
End Function


