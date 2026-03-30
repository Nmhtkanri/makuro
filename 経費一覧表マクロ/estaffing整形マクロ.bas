Attribute VB_Name = "estaffing整形マクロ"
Option Explicit

' === ヘッダー無し（1行目からデータ）版：1シート完結版 ===
Public Sub Export_EStaffing_SelectedColumns()
    Dim ws As Worksheet
    Dim lastRow As Long, startRow As Long, n As Long
    Dim outRow As Long, i As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' ★ソース兼出力シート（1枚で完結）
    Set ws = ThisWorkbook.Worksheets("e-staffing_出力")
    
    ' ヘッダー行数（ソースはヘッダー無しなので 0）
    Const SRC_HDR_ROWS As Long = 0
    startRow = SRC_HDR_ROWS + 1          ' = 1 行目から読む
    
    ' 最終行（E=名前列を基準）
    lastRow = ws.Cells(ws.rows.Count, "E").End(xlUp).Row
    n = lastRow - startRow + 1
    If n <= 0 Then GoTo FinallyExit

    ' ★ここで元データを配列に読み込む（この段階ではまだシートは消さない）
    Dim arrE, arrF, arrI, arrJ, arrK, arrL, arrO
    arrE = ws.Range("E" & startRow & ":E" & lastRow).value
    arrF = ws.Range("F" & startRow & ":F" & lastRow).value
    arrI = ws.Range("I" & startRow & ":I" & lastRow).value
    arrJ = ws.Range("J" & startRow & ":J" & lastRow).value
    arrK = ws.Range("K" & startRow & ":K" & lastRow).value
    arrL = ws.Range("L" & startRow & ":L" & lastRow).value
    arrO = ws.Range("O" & startRow & ":O" & lastRow).value

    ' 出力配列
    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To 7)

    outRow = 0
    For i = 1 To n
        ' 空行スキップなど必要ならここで調整
        If NzText(arrE(i, 1)) <> "" Or NzText(arrO(i, 1)) <> "" Then
            outRow = outRow + 1
            outArr(outRow, 1) = NzText(arrE(i, 1))           ' 名前（E）
            outArr(outRow, 2) = NormalizeDate(arrF(i, 1))    ' 日付（F）
            outArr(outRow, 3) = NzText(arrI(i, 1))           ' 出発（I）
            outArr(outRow, 4) = NzText(arrJ(i, 1))           ' 到着（J）
            outArr(outRow, 5) = NzText(arrK(i, 1))           ' 手段（K）
            outArr(outRow, 6) = NzText(arrL(i, 1))           ' 内訳（L）
            outArr(outRow, 7) = NormalizeAmount(arrO(i, 1))  ' 金額（O）
        End If
    Next

    ' ★ここでシートを一旦クリアして、整形済みデータを書き戻す
    ws.Cells.Clear

    If outRow > 0 Then
        ' 見出し
        ws.Range("A1").Resize(1, 7).value = _
            Array("名前", "日付", "出発", "到着", "手段", "内訳", "金額")
        
        ' データ本体
        ws.Range("A2").Resize(outRow, 7).value = outArr
        
        With ws
            .Columns("A:G").AutoFit
            .Range("A1:G1").Font.Bold = True
            .Range("B2:B" & (outRow + 1)).NumberFormatLocal = "yyyy/m/d"
            .Range("G2:G" & (outRow + 1)).NumberFormatLocal = "#,##0;[赤]-#,##0"
        End With
    Else
        ' 条件に合う行が1件もなかった場合、見出しだけ出すならこちら
        ws.Range("A1").Resize(1, 7).value = _
            Array("名前", "日付", "出発", "到着", "手段", "内訳", "金額")
        ws.Range("A1:G1").Font.Bold = True
    End If

FinallyExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "Export_EStaffing_SelectedColumns でエラー：" & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume FinallyExit
End Sub

Private Function NzText(v) As String
    If IsError(v) Then
        NzText = ""
    ElseIf IsNull(v) Or isEmpty(v) Then
        NzText = ""
    Else
        NzText = CStr(v)
    End If
End Function

Private Function NormalizeDate(v) As Variant
    ' 日付が文字でも数値でもだいたい受ける
    If IsDate(v) Then
        NormalizeDate = CDate(v)
    ElseIf IsNumeric(v) Then
        ' たとえば 20251001 や 45500 系のシリアルにも対応
        If v > 30000 And v < 60000 Then
            NormalizeDate = DateSerial(1899, 12, 30) + CLng(v)
        Else
            ' 20251001 形式なら文字にして日付化
            Dim s As String: s = CStr(v)
            If Len(s) = 8 Then
                NormalizeDate = DateSerial(Left$(s, 4), Mid$(s, 5, 2), Right$(s, 2))
            Else
                NormalizeDate = v ' そのまま
            End If
        End If
    Else
        ' "2025/10/01" みたいな文字は DateValue で拾う
        On Error Resume Next
        NormalizeDate = DateValue(CStr(v))
        If Err.Number <> 0 Then
            NormalizeDate = v
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Function

Private Function NormalizeAmount(v) As Variant
    ' 金額を数値に。カンマや￥が混ざっててもOK
    Dim s As String
    s = NzText(v)
    If s = "" Then
        NormalizeAmount = Empty
        Exit Function
    End If
    s = Replace(s, ",", "")
    s = Replace(s, "\", "")
    s = Replace(s, "￥", "")
    s = Replace(s, "円", "")
    If IsNumeric(s) Then
        NormalizeAmount = CDbl(s)
    Else
        NormalizeAmount = v
    End If
End Function

