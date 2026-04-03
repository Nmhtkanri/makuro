Attribute VB_Name = "Module4"
Option Explicit

Public Sub Update_TeleworkAllowance_HoursGE4()
    Const SRC_SHEET As String = "勤務時間帯一覧"
    Const DST_SHEET As String = "テレワーク手当"
    Const SRC_COL_EMP As Long = 1   ' A: 従業員番号
    Const SRC_COL_TW  As Long = 8   ' H: テレワーク（文字列に時間が入る想定）
    Const DST_COL_EMP As Long = 1   ' A: 従業員番号
    Const DST_COL_AMT As Long = 3   ' C: テレワーク（支給額）
    Const RATE As Double = 400      ' 1件あたり400円

    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim lastSrc As Long, lastDst As Long
    Dim r As Long, emp As String, v As Variant
    Dim sums As Object, dstIndex As Object
    Dim updated As Long, unmatched As Long
    Dim hours As Double

    On Error GoTo ErrH
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsSrc = ThisWorkbook.Worksheets(SRC_SHEET)
    Set wsDst = ThisWorkbook.Worksheets(DST_SHEET)

    lastSrc = GetLastRow(wsSrc, SRC_COL_EMP)
    lastDst = GetLastRow(wsDst, DST_COL_EMP)

    Set sums = CreateObject("Scripting.Dictionary")

    ' --- H列の文字列から数字を抜き出して 4以上なら1件カウント ---
    For r = 2 To lastSrc
        emp = Trim$(CStr(wsSrc.Cells(r, SRC_COL_EMP).value))
        If Len(emp) > 0 Then
            v = wsSrc.Cells(r, SRC_COL_TW).value
            If Len(Trim$(CStr(v))) > 0 Then
                hours = ExtractFirstNumber(CStr(v)) ' 文字列から先頭の数値を抽出（全角→半角対応）
                If hours >= 4 Then
                    If Not sums.Exists(emp) Then sums(emp) = 0#
                    sums(emp) = sums(emp) + 1#
                End If
            End If
        End If
    Next

    ' 宛先C列クリア（ヘッダー除く）
    If lastDst >= 2 Then
        wsDst.Range(wsDst.Cells(2, DST_COL_AMT), wsDst.Cells(lastDst, DST_COL_AMT)).ClearContents
    End If

    ' 宛先の従業員番号→行index
    Set dstIndex = CreateObject("Scripting.Dictionary")
    For r = 2 To lastDst
        emp = Trim$(CStr(wsDst.Cells(r, DST_COL_EMP).value))
        If Len(emp) > 0 Then
            If Not dstIndex.Exists(emp) Then dstIndex(emp) = r
        End If
    Next

    ' 書き込み
    Dim key As Variant, amt As Double, rowDst As Long
    For Each key In sums.Keys
        amt = sums(key) * RATE
        If dstIndex.Exists(key) Then
            rowDst = CLng(dstIndex(key))
            wsDst.Cells(rowDst, DST_COL_AMT).value = amt
            updated = updated + 1
        Else
            Debug.Print "未掲載の従業員（テレワーク手当に行なし）:", key, "金額", amt
            unmatched = unmatched + 1
        End If
    Next

    MsgBox "テレワーク手当の更新完了。" & vbCrLf & _
           "更新件数: " & updated & "　スキップ: " & unmatched, vbInformation

FINALLY:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
ErrH:
    MsgBox "エラー: " & Err.Number & " / " & Err.Description, vbExclamation
    Resume FINALLY
End Sub

' 文字列の中から最初に出てくる数値（整数/小数）を抜き出して返す
' 全角→半角にしてから正規表現で抽出（例： "４時間"→4, "4.5h"→4.5）
Private Function ExtractFirstNumber(ByVal s As String) As Double
    Dim re As Object, m As Object
    Dim t As String
    On Error GoTo EH

    ' 全角→半角（数字とピリオド）
    t = StrConv(s, vbNarrow)

    Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = "(\d+(\.\d+)?)"
        .Global = False
        .IgnoreCase = True
    End With

    If re.Test(t) Then
        Set m = re.Execute(t)(0)
        ExtractFirstNumber = CDbl(m.SubMatches(0)) ' m.Value でもOK
    Else
        ExtractFirstNumber = 0#
    End If
    Exit Function
EH:
    ExtractFirstNumber = 0#
End Function

Private Function GetLastRow(ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).row
    If r < 2 Then
        GetLastRow = 1
    Else
        GetLastRow = r
    End If
End Function



