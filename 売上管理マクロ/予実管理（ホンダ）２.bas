Attribute VB_Name = "Module3"
Sub freeeデータ加工()
Attribute freeeデータ加工.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 貼り付けたfreeeのデータをデータベースに整理する。
''
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.StatusBar = "開始"

Dim ws1 As Worksheet 'ボタンがあるシート
Dim ws2 As Worksheet '集計表を書き出すシート
Dim ws3 As Worksheet '名簿シート

Dim wb1 As Workbook 'ボタンのあるブック

Dim c1 As String
c1 = ""

Set wb1 = ActiveWorkbook
Set ws1 = ActiveSheet
Set ws2 = Worksheets("集計表")
Set ws3 = Worksheets("名簿")

ws1.Range("A1").Select
Selection.Delete Shift:=xlUp
ws1.Columns("A:A").Select
Selection.Insert Shift:=xlToRight

MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row


For i = 1 To MR1


If ws1.Cells(i, 2).Value = "従業員別" Then
c1 = ws1.Cells(i - 8, 2).Value
End If

If ws1.Cells(i, 2).Value = "売上高 計" Then
c1 = ""
Else
ws1.Cells(i, 1).Value = c1
End If
ThisWorkbook.Activate
Application.StatusBar = c1


Next



Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = "終了"



End Sub

Sub 予実表の作成()
'
' 予実表を作成する。
''
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.StatusBar = "開始"

Dim ws1 As Worksheet 'ボタンがあるシート
Dim ws2 As Worksheet '集計表
Dim ws3 As Worksheet '操作画面

Dim Mn As Integer '対象人数
Dim coln As Integer '予実表の行数


Set ws1 = ActiveSheet
Set ws2 = Worksheets("集計表")
Set ws3 = Worksheets("操作画面")



MR1 = ws1.Cells(Rows.Count, 1).End(xlUp).Row
ws1.Range("A3:Z" & Format(MR1)).Clear
'ws1.Range("A3:Z" & Format(MR1)).ClearFormats

MR2 = ws2.Cells(Rows.Count, 2).End(xlUp).Row '集計表の行数

Mn = ws3.Range("D1").Value
coln = 3

For i = 2 To MR2
'For i = 300 To 301
'For i = 3 To 10


Application.StatusBar = "データ収集中" & i & "/" & MR2 & " " & ws2.Cells(i, 1).Value & " " & ws2.Cells(i, 4).Value & " " & ws2.Cells(i, 3).Value & " を処理中"

ws1.Cells(coln, 1).Value = ws2.Cells(i, 1).Value '部門
ws1.Cells(coln, 2).Value = ws2.Cells(i, 4).Value '小分類
ws1.Cells(coln, 3).Value = ws2.Cells(i, 2).Value '社員番号
ws1.Cells(coln, 4).Value = ws2.Cells(i, 3).Value '氏名
ws1.Cells(coln, 5).Value = "計画"
ws1.Cells(coln, 6).Value = ws2.Cells(i, 5).Value  '科目
'ws1.Cells(coln, 7).Value = "=+SUMIFS(集計表!f:f,集計表!$B:$B,$C" & Format(coln) & ",集計表!$E:$E,$F" & Format(coln) & ")" '４月～３月
ws1.Cells(coln, 7).Value = "=+SUMIFS(集計表!F:F,集計表!$B:$B,$C" & Format(coln) & ",集計表!$E:$E,$F" & Format(coln) & ",集計表!$A:$A,予実表!$A" & Format(coln) & ")" '４月～３月

coln = coln + 1

ws1.Cells(coln, 1).Value = ws2.Cells(i, 1).Value '部門
ws1.Cells(coln, 2).Value = ws2.Cells(i, 4).Value '小分類
ws1.Cells(coln, 3).Value = ws2.Cells(i, 2).Value '社員番号
ws1.Cells(coln, 4).Value = ws2.Cells(i, 3).Value '氏名
ws1.Cells(coln, 5).Value = "実績"
ws1.Cells(coln, 6).Value = ws2.Cells(i, 5).Value  '科目
If ws2.Cells(i, 5).Value = "総受注金額" Then
ws1.Cells(coln, 7).Value = "=+SUMIFS(freeeデータ!C:C,freeeデータ!$A:$A,""売上高*"",freeeデータ!$B:$B,$D" & Format(coln) & ")*(G" & Format(coln - 1) & "<>0)" '４月～３月
Else
'ws1.Cells(coln, 7).Value = "=+SUMIF(freeeデータ!$B:$B,$D" & Format(coln) & ",freeeデータ!C:C)-SUMIFS(freeeデータ!C:C,freeeデータ!$A:$A,""売上高*"",freeeデータ!$B:$B,$D" & Format(coln) & ")" '４月～３月"
ws1.Cells(coln, 7).Value = "=+(SUMIF(freeeデータ!$B:$B,$D" & Format(coln) & ",freeeデータ!C:C)-SUMIFS(freeeデータ!C:C,freeeデータ!$A:$A,""売上高*"",freeeデータ!$B:$B,$D" & Format(coln) & "))*(G" & Format(coln - 1) & "<>0)" '４月～３月"
End If
coln = coln + 1

ws1.Cells(coln, 1).Value = ws2.Cells(i, 1).Value '部門
ws1.Cells(coln, 2).Value = ws2.Cells(i, 4).Value '小分類
ws1.Cells(coln, 3).Value = ws2.Cells(i, 2).Value '社員番号
ws1.Cells(coln, 4).Value = ws2.Cells(i, 3).Value '氏名
ws1.Cells(coln, 5).Value = "予実差"
ws1.Cells(coln, 6).Value = ws2.Cells(i, 5).Value  '科目
ws1.Cells(coln, 7).Value = "=+G" & Format(coln - 1) & "-G" & Format(coln - 2) '４月～３月
coln = coln + 1

If ws2.Cells(i, 5).Value = "総経費" Then
ws1.Cells(coln, 1).Value = ws2.Cells(i, 1).Value '部門
ws1.Cells(coln, 2).Value = ws2.Cells(i, 4).Value '小分類
ws1.Cells(coln, 3).Value = ws2.Cells(i, 2).Value '社員番号
ws1.Cells(coln, 4).Value = ws2.Cells(i, 3).Value '氏名
ws1.Cells(coln, 5).Value = "計画"
ws1.Cells(coln, 6).Value = "粗利"  '科目
'ws1.Cells(coln, 7).Value = "=+SUMIFS(G:G,$C:$C,$C" & Format(coln) & ",$F:$F,""総受注金額"",$E:$E,""計画"")-SUMIFS(G:G,$C:$C,$C" & Format(coln) & ",$F:$F,""総経費"",$E:$E,""計画"")" '４月～３月
ws1.Cells(coln, 7).Value = "=+(SUMIFS(G:G,$C:$C,$C" & Format(coln) & ",$F:$F,""総受注金額"",$E:$E,""計画"")-SUMIFS(G:G,$C:$C,$C" & Format(coln) & ",$F:$F,""総経費"",$E:$E,""計画""))*(G" & Format(coln - 3) & "<>0)" '４月～３月
coln = coln + 1

ws1.Cells(coln, 1).Value = ws2.Cells(i, 1).Value '部門
ws1.Cells(coln, 2).Value = ws2.Cells(i, 4).Value '小分類
ws1.Cells(coln, 3).Value = ws2.Cells(i, 2).Value '社員番号
ws1.Cells(coln, 4).Value = ws2.Cells(i, 3).Value '氏名
ws1.Cells(coln, 5).Value = "実績"
ws1.Cells(coln, 6).Value = "粗利"  '科目
'ws1.Cells(coln, 7).Value = "=+SUMIFS(G:G,$C:$C,$C" & Format(coln) & ",$F:$F,""総受注金額"",$E:$E,""実績"")-SUMIFS(G:G,$C:$C,$C" & Format(coln) & ",$F:$F,""総経費"",$E:$E,""実績"")" '４月～３月
ws1.Cells(coln, 7).Value = "=+(SUMIFS(G:G,$C:$C,$C" & Format(coln) & ",$F:$F,""総受注金額"",$E:$E,""実績"")-SUMIFS(G:G,$C:$C,$C" & Format(coln) & ",$F:$F,""総経費"",$E:$E,""実績""))*(G" & Format(coln - 4) & "<>0)" '４月～３月
coln = coln + 1

ws1.Cells(coln, 1).Value = ws2.Cells(i, 1).Value '部門
ws1.Cells(coln, 2).Value = ws2.Cells(i, 4).Value '小分類
ws1.Cells(coln, 3).Value = ws2.Cells(i, 2).Value '社員番号
ws1.Cells(coln, 4).Value = ws2.Cells(i, 3).Value '氏名
ws1.Cells(coln, 5).Value = "予実差"
ws1.Cells(coln, 6).Value = "粗利"  '科目
ws1.Cells(coln, 7).Value = "=+G" & Format(coln - 1) & "-G" & Format(coln - 2) '４月～３月
coln = coln + 1

End If
Next


MR12 = ws1.Cells(Rows.Count, 1).End(xlUp).Row

ws1.Range("G3:G" & Format(MR12)).Select
Selection.AutoFill Destination:=ws1.Range("G3:R" & Format(MR12)), Type:=xlFillDefault
    
ws1.Range("S3").Value = "=+SUM(G3:R3)"
ws1.Range("S3").Select
    Selection.Copy
ws1.Range("S3:S" & Format(MR12)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    
    
    
'Columns("G:S").Select
'Selection.Style = "Comma [0]"

'着地点の作成
Application.StatusBar = "着地点作成中"

ws1.Cells(3, 20).Value = "　計画分"
ws1.Cells(4, 20).Value = "　実績分"
ws1.Cells(5, 20).Value = "　着地点"
'1Q
ws1.Cells(3, 21).Value = "=+SUM(OFFSET($U3,0,-14+MIN($T$1,3)):OFFSET($U3,0,-15+3))*($T$1<3)"
ws1.Cells(4, 21).Value = "=+SUM(OFFSET($U4,0,-15):OFFSET($U4,0,-15+MIN($T$1,3)))"
ws1.Cells(5, 21).Value = "=SUM(U3:U4)"
'2Q
ws1.Cells(3, 22).Value = "=+SUM(OFFSET($U3,0,-14+MIN($T$1,6)):OFFSET($U3,0,-15+6))*($T$1<6)"
ws1.Cells(4, 22).Value = "=+SUM(OFFSET($U4,0,-15):OFFSET($U4,0,-15+MIN($T$1,6)))"
ws1.Cells(5, 22).Value = "=SUM(V3:V4)"
'3Q
ws1.Cells(3, 23).Value = "=+SUM(OFFSET($U3,0,-14+MIN($T$1,9)):OFFSET($U3,0,-15+9))*($T$1<9)"
ws1.Cells(4, 23).Value = "=+SUM(OFFSET($U4,0,-15):OFFSET($U4,0,-15+MIN($T$1,9)))"
ws1.Cells(5, 23).Value = "=SUM(W3:W4)"
'4Q
ws1.Cells(3, 24).Value = "=+SUM(OFFSET($U3,0,-14+MIN($T$1,12)):OFFSET($U3,0,-15+12))*($T$1<12)"
ws1.Cells(4, 24).Value = "=+SUM(OFFSET($U4,0,-15):OFFSET($U4,0,-15+MIN($T$1,12)))"
ws1.Cells(5, 24).Value = "=SUM(X3:X4)"

ws1.Cells(5, 25).Value = "=+X5-S3"


Range("T3:Y5").Select
Selection.AutoFill Destination:=Range("T3:Y" & Format(MR12)), Type:=xlFillDefault

Columns("D:Y").Select
Selection.Columns.AutoFit
Selection.Style = "Comma [0]"

Application.StatusBar = "罫線作成中"

    Range("A3:X5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlHairline
    End With
    
    Selection.Copy
    
    Range("A6:X" & Format(MR12)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = "終了"

End Sub
