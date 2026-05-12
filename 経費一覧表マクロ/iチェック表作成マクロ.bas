Attribute VB_Name = "iチェック表作成マクロ"
Option Explicit

' ============================================================
'  経費差分チェック表を集計シートの M～O 列に出力
' ============================================================
'  M 列 = 経費統合一覧表 の D 列 + D 空欄行の AH 列を社員番号でグループ集計
'  N 列 = C 列（集計値）- M 列（明細合計）の差分
'  O 列 = 判定（OK / 差分あり）
'  → 差分が 0 でない行は赤太字で強調
'
'  AH 列は顧客請求費用。D 列が空欄の行だけ足し、二重計上を避ける。
' ============================================================

Private Const SH_SUM As String = "集計"
Private Const SH_SRC As String = "経費統合一覧表"
Private Const COL_M As Long = 13   ' 明細合計
Private Const COL_N As Long = 14   ' 差分 C - M
Private Const COL_O As Long = 15   ' 判定

Public Sub Setup_経費チェック_集計シート()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsSum As Worksheet, wsSrc As Worksheet

    On Error Resume Next
    Set wsSum = wb.Worksheets(SH_SUM)
    Set wsSrc = wb.Worksheets(SH_SRC)
    On Error GoTo ErrHandler

    If wsSum Is Nothing Then
        MsgBox "「" & SH_SUM & "」シートが見つかりません。", vbExclamation
        GoTo FinallyExit
    End If
    If wsSrc Is Nothing Then
        MsgBox "「" & SH_SRC & "」シートが見つかりません。", vbExclamation
        GoTo FinallyExit
    End If

    ' 集計シートの最終行（A 列ベース）
    Dim lastRow As Long
    lastRow = wsSum.Cells(wsSum.rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "集計シートにデータがありません。先に集計を実行してください。", vbExclamation
        GoTo FinallyExit
    End If

    ' M～O 列をいったんクリア（古い値を消す）
    wsSum.Range(wsSum.Cells(1, COL_M), wsSum.Cells(lastRow, COL_O)).Clear

    ' ヘッダー
    wsSum.Cells(1, COL_M).value = "経費明細合計（D+D空欄時AH）"
    wsSum.Cells(1, COL_N).value = "差分（C - M）"
    wsSum.Cells(1, COL_O).value = "判定"
    With wsSum.Range(wsSum.Cells(1, COL_M), wsSum.Cells(1, COL_O))
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)  ' 薄い黄色
        .HorizontalAlignment = xlCenter
    End With

    ' 経費統合一覧表をVBAで再集計（文字列金額も数値化する）
    Dim detailTotals As Object
    Set detailTotals = BuildDetailTotals(wsSrc)

    Dim r As Long
    For r = 2 To lastRow
        Dim empNo As String
        Dim summaryTotal As Double, detailTotal As Double, diff As Double

        empNo = NormalizeIdForCheck(wsSum.Cells(r, 1).Value)
        If empNo <> "" And detailTotals.Exists(empNo) Then
            detailTotal = CDbl(detailTotals(empNo))
        Else
            detailTotal = 0
        End If

        summaryTotal = ParseAmountForCheck(wsSum.Cells(r, 3).Value)
        diff = summaryTotal - detailTotal

        wsSum.Cells(r, COL_M).Value = detailTotal
        wsSum.Cells(r, COL_N).Value = diff
        If detailTotal = 0 Then
            wsSum.Cells(r, COL_O).Value = "明細なし"
        ElseIf diff = 0 Then
            wsSum.Cells(r, COL_O).Value = "OK"
        Else
            wsSum.Cells(r, COL_O).Value = "差分あり"
        End If
    Next r

    ' 数値書式（M, N列）
    wsSum.Range(wsSum.Cells(2, COL_M), wsSum.Cells(lastRow, COL_N)).NumberFormat = "#,##0"

    ' 判定列（O 列）の中央寄せ
    wsSum.Range(wsSum.Cells(2, COL_O), wsSum.Cells(lastRow, COL_O)).HorizontalAlignment = xlCenter

    ' 条件付き書式: 差分が 0 でない行（N 列）→ 赤太字
    Dim rngN As Range
    Set rngN = wsSum.Range(wsSum.Cells(2, COL_N), wsSum.Cells(lastRow, COL_N))
    rngN.FormatConditions.Delete
    With rngN.FormatConditions.Add(Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=0")
        .Font.Color = RGB(192, 0, 0)
        .Font.Bold = True
        .StopIfTrue = False
    End With

    ' 条件付き書式: 判定列 → 「差分あり」赤、「OK」緑、「明細なし」灰色
    Dim rngO As Range
    Set rngO = wsSum.Range(wsSum.Cells(2, COL_O), wsSum.Cells(lastRow, COL_O))
    rngO.FormatConditions.Delete

    With rngO.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""差分あり""")
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(192, 0, 0)
        .Font.Bold = True
        .StopIfTrue = False
    End With
    With rngO.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""OK""")
        .Font.Color = RGB(0, 97, 0)
        .Interior.Color = RGB(198, 239, 206)
        .StopIfTrue = False
    End With
    With rngO.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""明細なし""")
        .Font.Color = RGB(128, 128, 128)
        .Interior.Color = RGB(217, 217, 217)
        .StopIfTrue = False
    End With

    ' 列幅
    wsSum.Columns(COL_M).ColumnWidth = 18
    wsSum.Columns(COL_N).ColumnWidth = 14
    wsSum.Columns(COL_O).ColumnWidth = 12

    ' 集計件数の確認
    Dim diffCount As Long, okCount As Long, emptyCount As Long
    For r = 2 To lastRow
        Select Case wsSum.Cells(r, COL_O).value
            Case "差分あり": diffCount = diffCount + 1
            Case "OK":       okCount = okCount + 1
            Case "明細なし": emptyCount = emptyCount + 1
        End Select
    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "経費差分チェック表を集計シートのM～O列に出力しました。" & vbCrLf & vbCrLf & _
           "対象社員数: " & (lastRow - 1) & " 名" & vbCrLf & _
           "  [OK]      : " & okCount & " 名" & vbCrLf & _
           "  [差分あり]: " & diffCount & " 名" & vbCrLf & _
           "  [明細なし]: " & emptyCount & " 名" & vbCrLf & vbCrLf & _
           "差分のある行を優先的に確認してください。", vbInformation

FinallyExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "経費チェック表作成エラー: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume FinallyExit
End Sub

Private Function BuildDetailTotals(ByVal wsSrc As Worksheet) As Object
    Dim totals As Object
    Set totals = CreateObject("Scripting.Dictionary")

    Dim lastSrcRow As Long
    lastSrcRow = wsSrc.Cells(wsSrc.rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastSrcRow
        Dim empNo As String
        empNo = NormalizeIdForCheck(wsSrc.Cells(r, 1).Value)
        If empNo <> "" Then
            Dim amount As Double
            If IsBlankForCheck(wsSrc.Cells(r, 4).Value) Then
                amount = ParseAmountForCheck(wsSrc.Cells(r, 34).Value)
            Else
                amount = ParseAmountForCheck(wsSrc.Cells(r, 4).Value)
            End If

            If amount <> 0 Then
                If Not totals.Exists(empNo) Then totals.Add empNo, 0#
                totals(empNo) = CDbl(totals(empNo)) + amount
            End If
        End If
    Next r

    Set BuildDetailTotals = totals
End Function

Private Function NormalizeIdForCheck(ByVal v As Variant) As String
    Dim s As String, i As Long, ch As String, out As String
    s = CStr(v)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    NormalizeIdForCheck = out
End Function

Private Function ParseAmountForCheck(ByVal v As Variant) As Double
    Dim s As String
    s = CStr(v)
    s = Replace(s, ",", "")
    s = Replace(s, "￥", "")
    s = Replace(s, "\", "")
    s = Replace(s, "円", "")
    s = Replace(s, "(", "-")
    s = Replace(s, ")", "")
    s = Replace(s, "（", "-")
    s = Replace(s, "）", "")
    s = Trim$(s)

    If s = "" Then
        ParseAmountForCheck = 0
    ElseIf IsNumeric(s) Then
        ParseAmountForCheck = CDbl(s)
    Else
        ParseAmountForCheck = 0
    End If
End Function

Private Function IsBlankForCheck(ByVal v As Variant) As Boolean
    IsBlankForCheck = (Trim$(CStr(v)) = "")
End Function
