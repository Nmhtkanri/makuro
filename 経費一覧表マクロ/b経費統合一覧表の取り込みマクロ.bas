Attribute VB_Name = "b経費統合一覧表取込マクロ"
' ==========================================
'  本社経費 → 経費統合一覧表
' ==========================================
Public Sub Append_本社経費_to_経費統合一覧表()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsDst As Worksheet
    
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set wsSrc = wb.Worksheets("本社経費")
    Set wsDst = wb.Worksheets("経費統合一覧表")
    On Error GoTo ErrHandler
    
    If wsSrc Is Nothing Then
        MsgBox "「本社経費」シートが見つかりません。", vbExclamation
        GoTo FinallyExit
    End If
    
    If wsDst Is Nothing Then
        MsgBox "「経費統合一覧表」シートが見つかりません。", vbExclamation
        GoTo FinallyExit
    End If

    ' =========================================================
    ' 1. 本社経費側の列位置（固定）
    ' =========================================================
    ' ヘッダー: 申請日(1), 申請者(2), 申請タイトル(3), 合計金額(4), 日付(5), 経費科目(6), 内容(7), 金額(8), 備考(9), 社員番号(10)
    Dim cDateApp As Long: cDateApp = 1      ' 申請日
    Dim cName As Long: cName = 2            ' 申請者
    Dim cTitle As Long: cTitle = 3          ' 申請タイトル
    Dim cTotal As Long: cTotal = 4          ' 合計金額（使用しない）
    Dim cDateUse As Long: cDateUse = 5      ' 日付
    Dim cSubj As Long: cSubj = 6            ' 経費科目
    Dim cCont As Long: cCont = 7            ' 内容
    Dim cAmt As Long: cAmt = 8              ' 金額 ← これを転記
    Dim cMemo As Long: cMemo = 9            ' 備考
    Dim cEmpID As Long: cEmpID = 10         ' 社員番号

    ' =========================================================
    ' 2. 最終行の取得
    ' =========================================================
    Dim srcLastRow As Long
    srcLastRow = wsSrc.Cells(wsSrc.rows.Count, cName).End(xlUp).Row
    If srcLastRow < 2 Then
        MsgBox "本社経費シートにデータがありません。", vbInformation
        GoTo FinallyExit
    End If

    Dim rowsCount As Long
    rowsCount = srcLastRow - 1

    ' =========================================================
    ' 3. 追記先の開始行
    ' =========================================================
    Dim dstStartRow As Long
    dstStartRow = Application.WorksheetFunction.Max( _
                    wsDst.Cells(wsDst.rows.Count, 1).End(xlUp).Row, _
                    wsDst.Cells(wsDst.rows.Count, 2).End(xlUp).Row) + 1
    If dstStartRow < 2 Then dstStartRow = 2

    ' =========================================================
    ' 4. 配列準備（行数 × 34列）
    ' =========================================================
    Dim arr() As Variant
    ReDim arr(1 To rowsCount, 1 To 34)

    ' =========================================================
    ' 5. データ転記ループ
    ' =========================================================
    Dim i As Long, r As Long
    Dim sG As String, sI As String
    
    For i = 1 To rowsCount
        r = i + 1 ' データは2行目から
        
        ' A列(1): 社員番号 ← 社員番号(10列目)
        arr(i, 1) = CStr(wsSrc.Cells(r, cEmpID).value)
        
        ' B列(2): 氏名 ← 申請者(2列目)
        arr(i, 2) = CStr(wsSrc.Cells(r, cName).value)
        
        ' C列(3): 申請日 ← 申請日(1列目)
        arr(i, 3) = FormatDateStr(wsSrc.Cells(r, cDateApp).value)
        
        ' D列(4): 合計 ← 金額(8列目) ★ここがポイント
        arr(i, 4) = wsSrc.Cells(r, cAmt).value
        
        arr(i, 5) = CStr(wsSrc.Cells(r, cTitle).value)
        
        ' F列(6): 利用日 ← 日付(5列目)
        arr(i, 6) = FormatDateStr(wsSrc.Cells(r, cDateUse).value)
        
        ' G列(7): 交通機関 ← （空欄）
        
        ' H列(8): 内訳 ← 申請タイトル(3列目)
        arr(i, 8) = CStr(wsSrc.Cells(r, cSubj).value)
        
        arr(i, 16) = wsSrc.Cells(r, cAmt).value
        
        sG = "": sI = ""
        sG = CStr(wsSrc.Cells(r, cCont).value)
        sI = CStr(wsSrc.Cells(r, cMemo).value)
        If sG <> "" And sI <> "" Then
            arr(i, 20) = sG & " / " & sI
        ElseIf sG <> "" Then
            arr(i, 20) = sG
        Else
            arr(i, 20) = sI
        End If
        
    Next i

    ' =========================================================
    ' 6. シートへの書き出し
    ' =========================================================
    With wsDst
        Dim targetRange As Range
        Set targetRange = .Range(.Cells(dstStartRow, 1), .Cells(dstStartRow + rowsCount - 1, 34))
        
        ' 書式を文字列に設定
        targetRange.NumberFormat = "@"
        
        ' 一気に貼り付け
        targetRange.value = arr
    End With

    MsgBox "本社経費データの追記が完了しました！" & vbCrLf & _
           "件数: " & rowsCount & " 件" & vbCrLf & _
           "開始行: " & dstStartRow & " 行目", vbInformation

FinallyExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "本社経費取り込みエラー：" & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume FinallyExit
End Sub

' --- 日付を文字列に変換する補助関数 ---
Private Function FormatDateStr(v As Variant) As String
    If IsDate(v) Then
        FormatDateStr = Format$(v, "yyyy/mm/dd")
    Else
        FormatDateStr = CStr(v)
    End If
End Function
