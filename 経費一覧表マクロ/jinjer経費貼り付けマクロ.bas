Attribute VB_Name = "jinjer経費貼り付けマクロ"
Option Explicit

Sub ファイルからデータ貼り付けA2()
    Dim fileDialog As fileDialog
    Dim selectedFile As String
    Dim sourceWb As Workbook
    Dim targetWs As Worksheet
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' エクスプローラーでファイルを選択
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
   With fileDialog
        .Title = "コピー元のファイルを選択してください"
        .filters.Clear
        .filters.Add "Excel/CSVファイル", "*.xlsx; *.xls; *.xlsm; *.csv"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            MsgBox "ファイルが選択されませんでした。", vbExclamation
            Exit Sub
        End If
    End With
    
    ' 貼り付け先のシートを設定
    On Error Resume Next
    Set targetWs = ThisWorkbook.Worksheets("経費統合一覧表")
    On Error GoTo 0
    
    If targetWs Is Nothing Then
        MsgBox "「e-staffing情報貼り付けシート」が見つかりません。", vbCritical
        Exit Sub
    End If
    
    ' 選択したファイルを開く
    Application.ScreenUpdating = False
    Set sourceWb = Workbooks.Open(selectedFile)
    Set sourceWs = sourceWb.Worksheets(1) ' 最初のシートを対象
    
    ' コピー元のデータ範囲を取得（A2から最終行・最終列まで）
    lastRow = sourceWs.Cells(sourceWs.rows.Count, 1).End(xlUp).Row
    lastCol = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column
    
    If lastRow >= 2 Then
        Dim pastedRows As Long
        Dim movedRows As Long
        pastedRows = lastRow - 1

        ' A2から最終行・最終列までをコピー
        sourceWs.Range(sourceWs.Cells(2, 1), sourceWs.Cells(lastRow, lastCol)).Copy
        
        ' 貼り付け先のA2に貼り付け
        targetWs.Range("A2").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False

        movedRows = MovePastedJinjerCustomerBillsToAH(targetWs, 2, pastedRows)
        
        MsgBox "データをA2に貼り付けました。" & vbCrLf & _
               "行数: " & pastedRows & " 列数: " & lastCol & vbCrLf & _
               "顧客請求AH振り分け: " & movedRows & " 行", vbInformation
    Else
        MsgBox "コピーするデータがありません（A2以降にデータがない）。", vbExclamation
    End If
    
    ' ファイルを閉じる
    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
End Sub

' ============================================================
'  貼り付け済みjinjer行の顧客請求金額を AH列のみに寄せる
'  - 既にAHへ寄っている行は加算しない
'  - 夜間当番/顧客対応当番系は対象外
'  - D/M/N/P と AH の二重計上を避ける
' ============================================================
Private Function MovePastedJinjerCustomerBillsToAH(ByVal ws As Worksheet, _
                                                   ByVal firstRow As Long, _
                                                   ByVal rowCount As Long) As Long
    Dim r As Long, movedCount As Long
    For r = firstRow To firstRow + rowCount - 1
        If MoveOnePastedJinjerCustomerBillToAH(ws, r) Then movedCount = movedCount + 1
    Next r
    MovePastedJinjerCustomerBillsToAH = movedCount
End Function

Private Function MoveOnePastedJinjerCustomerBillToAH(ByVal ws As Worksheet, _
                                                     ByVal rowNum As Long) As Boolean
    Dim cBillType As Long, cDetail As Long, cMemoReq As Long, cExpenseType As Long
    Dim cMemoLine As Long, cTotal As Long, cSubTotal As Long, cFare As Long
    Dim cAmount As Long, cCustomerBill As Long

    cBillType = FindHeaderColOrDefault(ws, Array("請求区分"), 10)
    cDetail = FindHeaderColOrDefault(ws, Array("内訳"), 8)
    cMemoReq = FindHeaderColOrDefault(ws, Array("備考(申請書)"), 5)
    cExpenseType = FindHeaderColOrDefault(ws, Array("費用種別"), 11)
    cMemoLine = FindHeaderColOrDefault(ws, Array("備考(明細)"), 20)
    cTotal = FindHeaderColOrDefault(ws, Array("合計"), 4)
    cSubTotal = FindHeaderColOrDefault(ws, Array("小計"), 13)
    cFare = FindHeaderColOrDefault(ws, Array("金額(交通費)"), 14)
    cAmount = FindHeaderColOrDefault(ws, Array("金額"), 16)
    cCustomerBill = FindHeaderColOrDefault(ws, Array("顧客請求費", "顧客請求分"), 34)

    If Trim$(CStr(ws.Cells(rowNum, cBillType).Value)) <> "顧客請求" Then Exit Function

    Dim judgeText As String
    judgeText = CStr(ws.Cells(rowNum, cDetail).Value) & " " & _
                CStr(ws.Cells(rowNum, cMemoReq).Value) & " " & _
                CStr(ws.Cells(rowNum, cExpenseType).Value) & " " & _
                CStr(ws.Cells(rowNum, cMemoLine).Value)
    If IsJinjerNightDutyText(judgeText) Then Exit Function

    Dim hasAH As Boolean
    hasAH = (Trim$(CStr(ws.Cells(rowNum, cCustomerBill).Value)) <> "")

    Dim billAmt As Variant
    billAmt = FirstNonEmptyCellValue(ws, rowNum, Array(cTotal, cAmount, cSubTotal, cFare))

    If Not hasAH Then
        If Trim$(CStr(billAmt)) = "" Then Exit Function
        ws.Cells(rowNum, cCustomerBill).Value = billAmt
    End If

    ws.Cells(rowNum, cTotal).ClearContents
    ws.Cells(rowNum, cSubTotal).ClearContents
    ws.Cells(rowNum, cFare).ClearContents
    ws.Cells(rowNum, cAmount).ClearContents
    MoveOnePastedJinjerCustomerBillToAH = True
End Function

Private Function FindHeaderColOrDefault(ByVal ws As Worksheet, _
                                        ByVal headerNames As Variant, _
                                        ByVal defaultCol As Long) As Long
    Dim lastCol As Long, c As Long, i As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For i = LBound(headerNames) To UBound(headerNames)
        For c = 1 To lastCol
            If Trim$(CStr(ws.Cells(1, c).Value)) = CStr(headerNames(i)) Then
                FindHeaderColOrDefault = c
                Exit Function
            End If
        Next c
    Next i

    FindHeaderColOrDefault = defaultCol
End Function

Private Function FirstNonEmptyCellValue(ByVal ws As Worksheet, _
                                        ByVal rowNum As Long, _
                                        ByVal cols As Variant) As Variant
    Dim i As Long, v As Variant
    For i = LBound(cols) To UBound(cols)
        v = ws.Cells(rowNum, CLng(cols(i))).Value
        If Trim$(CStr(v)) <> "" Then
            FirstNonEmptyCellValue = v
            Exit Function
        End If
    Next i
    FirstNonEmptyCellValue = ""
End Function

Private Function IsJinjerNightDutyText(ByVal textValue As String) As Boolean
    Dim keys As Variant
    keys = Array("夜間当番", "24時間準直当番", "準直当番", "深夜出動", _
                 "顧客当番", "顧客対応当番", "オンコール")

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        If InStr(1, textValue, CStr(keys(i)), vbTextCompare) > 0 Then
            IsJinjerNightDutyText = True
            Exit Function
        End If
    Next i
End Function

