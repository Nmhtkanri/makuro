Attribute VB_Name = "Module1"
Sub ワークシート一覧()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.StatusBar = "開始"

Dim ws0 As Worksheet 'マクロが開くシート
Dim ws1 As Worksheet 'ボタンのあるブック
Dim ws2 As Worksheet '操作画面シート
Dim wb0 As Workbook 'マクロが開くブック
Dim wb1 As Workbook 'ボタンのあるブック
Dim s_name(100) As String

Set wb1 = ActiveWorkbook
Set ws2 = Worksheets("操作画面")

ActiveSheet.Range("A:B").AutoFilter Field:=2
Columns("A:B").Select
Selection.Clear


MR1 = ws2.Cells(Rows.Count, 1).End(xlUp).Row

For i = 3 To MR1
Application.StatusBar = ws2.Range("B" & Format(i)).Value
If ws2.Range("B" & Format(i)).Value = ws2.Range("B" & Format(i - 1)).Value Then GoTo label1
Workbooks.Open Filename:=ws2.Range("B" & Format(i)).Value, UpdateLinks:=False
Set wb0 = ActiveWorkbook
Set ws0 = wb0.Worksheets(ws2.Range("C" & Format(i)).Value)
sc = sc + Sheets.Count
k = i - 2
For J = 1 To Sheets.Count
s_name(k + J) = Sheets(J).Name
Next J
k = k + J
wb0.Close
label1:
Next i

wb1.Activate
For i = 2 To k
Cells(i, 1).Value = s_name(i)
Cells(i, 2).Value = "=+COUNTIF(操作画面!C:C,A" & Format(i) & ")"
Next

ActiveSheet.Range("$A$1:$B$47").AutoFilter Field:=2, Criteria1:="0"

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = "終了"

End Sub
Sub まとめ()


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.StatusBar = "開始"

Dim ws0 As Worksheet 'マクロが開くシート
Dim ws1 As Worksheet 'まとめシート
Dim ws2 As Worksheet '操作画面シート
Dim wb0 As Workbook 'マクロが開くブック
Dim wb1 As Workbook 'ボタンのあるブック
Dim ws_name As String '売上のあるワークシート名
Dim L_num As Integer 'DELL売上のある行番号


Set wb1 = ActiveWorkbook
Set ws1 = Worksheets("まとめ")
Set ws2 = Worksheets("操作画面")

MR1 = ws1.Cells(Rows.Count, 1).End(xlUp).Row

Application.StatusBar = "予実表作成"

ws1.Cells(3, 4).Value = "=+SUMIFS(予実表!G:G,予実表!$A:$A,$A3,予実表!$F:$F,$B3,予実表!$E:$E,まとめ!$C3)"
ws1.Cells(4, 4).Value = "=+SUMIFS(予実表!G:G,予実表!$A:$A,$A4,予実表!$F:$F,$B4,予実表!$E:$E,まとめ!$C4)"
ws1.Cells(5, 4).Value = "=+D4-D3"
ws1.Cells(6, 4).Value = "=+SUMIFS(予実表!G:G,予実表!$A:$A,$A6,予実表!$F:$F,$B6,予実表!$E:$E,まとめ!$C6)"
ws1.Cells(7, 4).Value = "=+SUMIFS(予実表!G:G,予実表!$A:$A,$A7,予実表!$F:$F,$B7,予実表!$E:$E,まとめ!$C7)"
ws1.Cells(8, 4).Value = "=+D4-D3"
ws1.Cells(9, 4).Value = "=+D3-D6"
ws1.Cells(10, 4).Value = "=+D4-D7"
ws1.Cells(11, 4).Value = "=+D5-D8"

Range("D3:D11").Select
Selection.AutoFill Destination:=Range("D3:O38"), Type:=xlFillDefault

ws1.Cells(3, 16).Value = "=SUM(D3:O3)"

Range("P3").Select
Selection.AutoFill Destination:=Range("P3:P38"), Type:=xlFillDefault

'DELL分修正

Application.StatusBar = "DELL分修正"


ws_name = ws2.Range("G3").Value

For i = 3 To MR1
If ws1.Cells(i, 1).Value = "DELL" Then
If ws1.Cells(i, 2).Value = "総受注金額" Then
If ws1.Cells(i, 3).Value = "計画" Then
L_num = i
End If
End If
End If
Next

Workbooks.Open Filename:=ws2.Range("F3").Value, UpdateLinks:=False
Set wb0 = ActiveWorkbook
Set ws0 = wb0.Worksheets(ws_name)

ws0.Select

For i = 1 To 12
ws1.Cells(L_num, i + 3).Value = ws0.Cells(2, i + 8).Value
Next

wb0.Close

ws1.Cells(L_num + 1, 4).Value = "=+SUMIFS(freeeデータ!C:C,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"
ws1.Cells(L_num + 1, 5).Value = "=+SUMIFS(freeeデータ!D:D,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"
ws1.Cells(L_num + 1, 6).Value = "=+SUMIFS(freeeデータ!E:E,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"
ws1.Cells(L_num + 1, 7).Value = "=+SUMIFS(freeeデータ!F:F,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"
ws1.Cells(L_num + 1, 8).Value = "=+SUMIFS(freeeデータ!G:G,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"
ws1.Cells(L_num + 1, 9).Value = "=+SUMIFS(freeeデータ!H:H,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"
ws1.Cells(L_num + 1, 10).Value = "=+SUMIFS(freeeデータ!I:I,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"
ws1.Cells(L_num + 1, 11).Value = "=+SUMIFS(freeeデータ!J:J,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"
ws1.Cells(L_num + 1, 12).Value = "=+SUMIFS(freeeデータ!K:K,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"
ws1.Cells(L_num + 1, 13).Value = "=+SUMIFS(freeeデータ!L:L,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"
ws1.Cells(L_num + 1, 14).Value = "=+SUMIFS(freeeデータ!M:M,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"
ws1.Cells(L_num + 1, 15).Value = "=+SUMIFS(freeeデータ!N:N,freeeデータ!A:A,""売上高*"",freeeデータ!B:B,""[UESP担当者]"")"


'ホンダ分修正

Application.StatusBar = "ホンダ分修正"


ws_name = ws2.Range("G4").Value

For i = 2 To MR1
If ws1.Cells(i, 1).Value = "ホンダ" Then
If ws1.Cells(i, 2).Value = "総受注金額" Then
If ws1.Cells(i, 3).Value = "計画" Then
L_num = i
End If
End If
End If
Next

Workbooks.Open Filename:=ws2.Range("F4").Value, UpdateLinks:=False
Set wb0 = ActiveWorkbook
Set ws0 = wb0.Worksheets(ws_name)

ws0.Select

For i = 1 To 12
ws1.Cells(L_num, i + 3).Value = ws0.Cells(2, i + 8).Value
Next

wb0.Close

ws1.Cells(L_num + 1, 4).Value = "=+SUMIFS('freeeデータ (ホンダ)'!C:C,'freeeデータ (ホンダ)'!B:B,""売上高"")"
ws1.Cells(L_num + 1, 5).Value = "=+SUMIFS('freeeデータ (ホンダ)'!D:D,'freeeデータ (ホンダ)'!B:B,""売上高"")"
ws1.Cells(L_num + 1, 6).Value = "=+SUMIFS('freeeデータ (ホンダ)'!E:E,'freeeデータ (ホンダ)'!B:B,""売上高"")"
ws1.Cells(L_num + 1, 7).Value = "=+SUMIFS('freeeデータ (ホンダ)'!F:F,'freeeデータ (ホンダ)'!B:B,""売上高"")"
ws1.Cells(L_num + 1, 8).Value = "=+SUMIFS('freeeデータ (ホンダ)'!G:G,'freeeデータ (ホンダ)'!B:B,""売上高"")"
ws1.Cells(L_num + 1, 9).Value = "=+SUMIFS('freeeデータ (ホンダ)'!H:H,'freeeデータ (ホンダ)'!B:B,""売上高"")"
ws1.Cells(L_num + 1, 10).Value = "=+SUMIFS('freeeデータ (ホンダ)'!I:I,'freeeデータ (ホンダ)'!B:B,""売上高"")"
ws1.Cells(L_num + 1, 11).Value = "=+SUMIFS('freeeデータ (ホンダ)'!J:J,'freeeデータ (ホンダ)'!B:B,""売上高"")"
ws1.Cells(L_num + 1, 12).Value = "=+SUMIFS('freeeデータ (ホンダ)'!K:K,'freeeデータ (ホンダ)'!B:B,""売上高"")"
ws1.Cells(L_num + 1, 13).Value = "=+SUMIFS('freeeデータ (ホンダ)'!L:L,'freeeデータ (ホンダ)'!B:B,""売上高"")"
ws1.Cells(L_num + 1, 14).Value = "=+SUMIFS('freeeデータ (ホンダ)'!M:M,'freeeデータ (ホンダ)'!B:B,""売上高"")"
ws1.Cells(L_num + 1, 15).Value = "=+SUMIFS('freeeデータ (ホンダ)'!N:N,'freeeデータ (ホンダ)'!B:B,""売上高"")"



'着地点の作成
Application.StatusBar = "着地点作成中"

ws1.Cells(3, 18).Value = "　計画分"
ws1.Cells(4, 18).Value = "　実績分"
ws1.Cells(5, 18).Value = "　着地点"
'1Q
ws1.Cells(3, 19).Value = "=+SUM(OFFSET($S3,0,-15+MIN($Q$1,3)):OFFSET($S3,0,-16+3))*($Q$1<3)"
ws1.Cells(4, 19).Value = "=+SUM(OFFSET($S4,0,-16):OFFSET($S4,0,-16+MIN($Q$1,3)))"
ws1.Cells(5, 19).Value = "=SUM(S3:S4)"
'2Q
ws1.Cells(3, 20).Value = "=+SUM(OFFSET($S3,0,-15+MIN($Q$1,6)):OFFSET($S3,0,-16+6))*($Q$1<6)"
ws1.Cells(4, 20).Value = "=+SUM(OFFSET($S4,0,-16):OFFSET($S4,0,-16+MIN($Q$1,6)))"
ws1.Cells(5, 20).Value = "=SUM(T3:T4)"
'03Q
ws1.Cells(3, 21).Value = "=+SUM(OFFSET($S3,0,-15+MIN($Q$1,9)):OFFSET($S3,0,-16+9))*($Q$1<9)"
ws1.Cells(4, 21).Value = "=+SUM(OFFSET($S4,0,-16):OFFSET($S4,0,-16+MIN($Q$1,9)))"
ws1.Cells(5, 21).Value = "=SUM(U3:U4)"
'4Q
ws1.Cells(3, 22).Value = "=+SUM(OFFSET($S3,0,-15+MIN($Q$1,12)):OFFSET($S3,0,-16+12))*($Q$1<12)"
ws1.Cells(4, 22).Value = "=+SUM(OFFSET($S4,0,-16):OFFSET($S4,0,-16+MIN($Q$1,12)))"
ws1.Cells(5, 22).Value = "=SUM(V3:V4)"

ws1.Cells(5, 23).Value = "=+P3-V5"


Range("R3:W5").Select
Selection.AutoFill Destination:=Range("R3:W38"), Type:=xlFillDefault

Columns("D:V").Select
Selection.Columns.AutoFit
Selection.Style = "Comma [0]"

Range("A1").Select

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = "終了"

End Sub

Sub ホンダ売上()


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.StatusBar = "開始"

Dim ws0 As Worksheet 'マクロが開くシート
Dim ws1 As Worksheet 'まとめシート
Dim ws2 As Worksheet '操作画面シート
Dim wb0 As Workbook 'マクロが開くブック
Dim wb1 As Workbook 'ボタンのあるブック
Dim ws_name As String '売上のあるワークシート名
Dim L_num As Integer 'DELL売上のある行番号


Set wb1 = ActiveWorkbook
Set ws1 = Worksheets("まとめ")
Set ws2 = Worksheets("操作画面")

ws_name = ws2.Range("G4").Value

MR1 = ws1.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To MR1
If ws1.Cells(i, 1).Value = "ホンダ" Then
If ws1.Cells(i, 2).Value = "総受注金額" Then
If ws1.Cells(i, 3).Value = "計画" Then
L_num = i
End If
End If
End If
Next

Workbooks.Open Filename:=ws2.Range("F4").Value, UpdateLinks:=False
Set wb0 = ActiveWorkbook
Set ws0 = wb0.Worksheets(ws_name)

ws0.Select

For i = 1 To 12
ws1.Cells(L_num, i + 3).Value = ws0.Cells(2, i + 8).Value
Next

wb0.Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = "終了"

End Sub

Sub 社員番号のチェック()

'予実表作成に当たり社員番号がもれなく記載されているかをチェック

Application.ScreenUpdating = False
Application.DisplayAlerts = False
ThisWorkbook.Activate
Application.StatusBar = "開始"

Dim ws0 As Worksheet 'マクロが開くシート
Dim ws1 As Worksheet 'ボタンがあるシート
Set wb1 = ActiveWorkbook
Set ws1 = ActiveSheet

MR1 = ws1.Cells(Rows.Count, 1).End(xlUp).Row

For i = 3 To MR1
Application.StatusBar = ws1.Range("B" & Format(i)).Value & "/" & ws1.Range("C" & Format(i)).Value
Workbooks.Open Filename:=ws1.Range("B" & Format(i)).Value, UpdateLinks:=False
Set wb0 = ActiveWorkbook
Set ws0 = wb0.Worksheets(ws1.Range("C" & Format(i)).Value)
MR0 = ws0.Cells(Rows.Count, 1).End(xlUp).Row
ws0.Activate

For J = 1 To MR0
If ws0.Cells(J, 1).Value = "" Then
If ws0.Cells(J, 2).Value <> "" Then
MsgBox "社員番号記載漏れ発見！"
Exit Sub
End If
End If
Next
wb0.Close
Next
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = "終了"

End Sub
Sub 名簿のコピー()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim ws0 As Worksheet
Dim wb1 As Workbook

Set ws0 = ActiveSheet
genbo = Range("H1").Value


Workbooks.Open Filename:=genbo, UpdateLinks:=False
Set wb1 = ActiveWorkbook


Range("A:D").Copy
ws0.Range("A1").PasteSpecial Paste:=xlPasteValues

wb1.Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub


Sub データ収集()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.StatusBar = "開始"
Application.Calculation = xlCalculationManual

Dim ws0 As Worksheet '書き込み先シート
Dim ws1 As Worksheet 'ボタンがあるシート
Dim ws2 As Worksheet '売上データシート
Dim ws3 As Worksheet '経費データシート

Dim wb0 As Workbook '書き込み先ブック
Dim wb1 As Workbook 'ボタンのあるブック
Dim wb2 As Workbook '売上データブック
Dim wb3 As Workbook '経費データブック

Dim bias As Integer '各表の列方向の項目の位置のズレを補正する変数
Dim ws2_col As Integer '集計表に書き出すときの列番号

Dim FCell As Range  '検索先の場所
Dim FFrst As Range  '項目検索
Dim wRow As Long    '該当列

Set wb1 = ActiveWorkbook
Set ws1 = ActiveSheet


Application.StatusBar = "売上データ読み込み中"
Workbooks.Open Filename:=ws1.Range("G" & Format(3)).Value, UpdateLinks:=False
Set wb2 = ActiveWorkbook
Set ws2 = wb2.Worksheets(ws1.Range("H" & Format(3)).Value)
Application.StatusBar = "経費データ読み込み中"
Workbooks.Open Filename:=ws1.Range("G" & Format(4)).Value, UpdateLinks:=False
Set wb3 = ActiveWorkbook
Set ws3 = wb3.Worksheets(ws1.Range("H" & Format(4)).Value)


wb1.Activate

this_month = ws1.Range("G1").Value & "月"

MR1 = ws1.Cells(Rows.Count, 1).End(xlUp).Row
ws2_col = 2


For i = 3 To MR1
bias = ws1.Cells(i, 4).Value
Workbooks.Open Filename:=ws1.Range("B" & Format(i)).Value, UpdateLinks:=False
Set wb0 = ActiveWorkbook
Set ws0 = wb0.Worksheets(ws1.Range("C" & Format(i)).Value)
MR0 = ws0.Cells(Rows.Count, 1).End(xlUp).Row

T_mon = WorksheetFunction.Match(this_month, ws0.Range("A1:Z1"), 0)

On Error Resume Next
For J = 7 To MR0

If J Mod 8 = 7 Then
If ws0.Cells(J, 5).Value = "転出" Then
J = J + 11
GoTo label1
End If
End If

pct = J / MR0
pct = Format(pct, "0.0%")
ThisWorkbook.Activate
Application.StatusBar = wb0.Name & " ワークブックの" & ws0.Name & " 処理中 " & pct & "終了"

bk1 = "[" & wb3.Name & "]"

'If J Mod 8 = 0 Then '★派遣金額
'If ws1.Cells(i, 3).Value = "UAL業務委託" Then GoTo label0
'ws0.Cells(J, T_mon).Value = "=+SUMIFS('[" & wb2.Name & "]" & ws2.Name & "'!$H:$H,'[" & wb2.Name & "]" & ws2.Name & "'!$B:$B,A" & J & ",'[" & wb2.Name & "]" & ws2.Name & "'!$F:$F,""売上高"")/1.1"
'ws0.Cells(J, T_mon).Value = Round(ws0.Cells(J, T_mon).Value, 0)
'If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
'End If
label0:

'社員番号よりアドレスを求める(1ユーザー1回のみ）
If J Mod 8 = 7 Then
    Set FCell = ws3.Range("A:A").Find(ws0.Cells(J, 1).Value)
    If FCell Is Nothing Or Trim(ws0.Cells(J, 1).Value) = 0 Then
        wRow = 0
    Else
        wRow = FCell.Row
    End If
End If

If J Mod 8 = 0 Then '★賃金
'ws0.Cells(J, T_mon).Value = "=+SUMIF(" & bk1 & "総一覧!$A:$A,A" & J & "," & bk1 & "総一覧!$L:$L)-SUMIF(" & bk1 & "総一覧!$A:$A,A" & J & "," & bk1 & "総一覧!$I:$I)"
'ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value
If wRow <> 0 Then
ws0.Cells(J, T_mon).Value = ""
ws0.Cells(J, T_mon).Value = ws3.Cells(wRow, 166) '源泉対象額
End If

If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
k_hoken = ws0.Cells(J, T_mon).Value * 6 / 1000
End If

If J Mod 8 = 1 Then '★通勤定期代+テレワーク手当→テレワーク手当は除外
'ws0.Cells(J, T_mon).Value = "=+SUMIF(" & bk1 & "総一覧!$A:$A,A" & J & "," & bk1 & "総一覧!$K:$K)+SUMIF(" & bk1 & "総一覧!$A:$A,A" & J & "," & bk1 & "総一覧!$I:$I)"
'ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value
ws0.Cells(J, T_mon).Value = ""
ws0.Cells(J, T_mon).Value = ws3.Cells(wRow, 93) '非課税通勤費
ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value '+ ws3.Cells(wRow, 76) 'テレワーク手当

If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
End If

If J Mod 8 = 3 Then '★健康保険＋介護保険
'ws0.Cells(J, T_mon).Value = "=+SUMIF(" & bk1 & "総一覧!$A:$A,A" & J & "," & bk1 & "総一覧!$M:$M)+SUMIF(" & bk1 & "総一覧!$A:$A,A" & J & "," & bk1 & "総一覧!$P:$P)"
'ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value * -1
ws0.Cells(J, T_mon).Value = ""
ws0.Cells(J, T_mon).Value = ws3.Cells(wRow, 138) '健康保険料
ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value + ws3.Cells(wRow, 139) '介護保険料

If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
End If

If J Mod 8 = 4 Then '★厚生年金
'ws0.Cells(J, T_mon).Value = "=+SUMIF(" & bk1 & "総一覧!$A:$A,A" & J & "," & bk1 & "総一覧!$N:$N)"
'ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value * -1
ws0.Cells(J, T_mon).Value = ""
ws0.Cells(J, T_mon).Value = ws3.Cells(wRow, 140) '厚生年金保険料

'If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
End If

If J Mod 8 = 5 Then '★労災保険
'ws0.Cells(J, T_mon).Value = "=+SUMIF(" & bk1 & "総一覧!$A:$A,A" & J & "," & bk1 & "総一覧!$O:$O)"
'ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value * -1
ws0.Cells(J, T_mon).Value = ""
ws0.Cells(J, T_mon).Value = ws3.Cells(wRow, 180) '事業主労災保険料

If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
End If

If J Mod 8 = 6 Then '★雇用保険
'ws0.Cells(J, T_mon).Value = k_hoken
'ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value
ws0.Cells(J, T_mon).Value = ""
ws0.Cells(J, T_mon).Value = ws3.Cells(wRow, 137) '雇用保険料
If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
End If

'★交通費その他、紹介料は変更していない。変更するならj Mod 2 = 10


label1:
Next
'wb2.Close
'wb3.Close

wb1.Activate
ws2.Activate


    ws0.Activate
    
    
 MR01 = ws0.Cells(Rows.Count, 11).End(xlUp).Row
   
   
   
 For kk = 1 To MR01
    Cells(kk, T_mon).Select
    With Selection.Font
        .ColorIndex = 1
        .TintAndShade = 0
    End With
Next

Next


ws1.Range("C1").Value = Now & "現在"

' 開いたファイルを閉じる
wb2.Close SaveChanges:=False
wb3.Close SaveChanges:=False


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = "終了"
Application.Calculation = xlCalculationAutomatic


End Sub
Sub データ収集_業務委託()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.StatusBar = "開始"

Dim ws0 As Worksheet '書き込み先シート
Dim ws1 As Worksheet 'ボタンがあるシート
Dim ws2 As Worksheet '売上データシート
Dim ws3 As Worksheet '経費データシート

Dim wb0 As Workbook '書き込み先ブック
Dim wb1 As Workbook 'ボタンのあるブック
Dim wb2 As Workbook '売上データブック
Dim wb3 As Workbook '経費データブック

Dim bias As Integer '各表の列方向の項目の位置のズレを補正する変数
Dim ws2_col As Integer '集計表に書き出すときの列番号


Set wb1 = ActiveWorkbook
Set ws1 = ActiveSheet


'Application.StatusBar = "売上データ読み込み中"
'Workbooks.Open Filename:=ws1.Range("G" & Format(3)).Value, UpdateLinks:=False
'Set wb2 = ActiveWorkbook
'Set ws2 = wb2.Worksheets(ws1.Range("H" & Format(3)).Value)
Application.StatusBar = "経費データ読み込み中"
Workbooks.Open Filename:=ws1.Range("G" & Format(4)).Value, UpdateLinks:=False
Set wb3 = ActiveWorkbook
Set ws3 = wb3.Worksheets(ws1.Range("H" & Format(4)).Value)


wb1.Activate

this_month = ws1.Range("G1").Value & "月"

MR1 = ws1.Cells(Rows.Count, 7).End(xlUp).Row
ws2_col = 2


For i = 21 To MR1
bias = ws1.Cells(i, 9).Value
Workbooks.Open Filename:=ws1.Range("G" & Format(i)).Value, UpdateLinks:=False
Set wb0 = ActiveWorkbook
Set ws0 = wb0.Worksheets(ws1.Range("H" & Format(i)).Value)
MR0 = ws0.Cells(Rows.Count, 1).End(xlUp).Row

T_mon = WorksheetFunction.Match(this_month, ws0.Range("A1:Z1"), 0)

On Error Resume Next
For J = 7 To MR0

If J Mod 8 = 7 Then
If ws0.Cells(J, 5).Value = "転出" Then
J = J + 11
GoTo label1
End If
End If

pct = J / MR0
pct = Format(pct, "0.0%")
ThisWorkbook.Activate
Application.StatusBar = wb0.Name & " ワークブックの" & ws0.Name & " 処理中 " & pct & "終了"



'If j Mod 12 = 6 Then
'ws0.Cells(j, T_mon).Value = "=+SUMIFS('[" & wb2.Name & "]" & ws2.Name & "'!$H:$H,'[" & wb2.Name & "]" & ws2.Name & "'!$B:$B,A" & j & ",'[" & wb2.Name & "]" & ws2.Name & "'!$F:$F,""売上高"")/1.1"
'ws0.Cells(j, T_mon).Value = Round(ws0.Cells(j, T_mon).Value, 0)
'If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = ""
'End If

If J Mod 8 = 0 Then
ws0.Cells(J, T_mon).Value = "=+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & J & ",[給与Ver3.xlsm]総一覧!$L:$L)"
ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value
If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
k_hoken = ws0.Cells(J, T_mon).Value * 6 / 1000
End If

If J Mod 8 = 1 Then
ws0.Cells(J, T_mon).Value = "=+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & J & ",[給与Ver3.xlsm]総一覧!$K:$K)"
ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value
If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
End If

If J Mod 8 = 2 Then
ws0.Cells(J, T_mon).Value = "=+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & J & ",[給与Ver3.xlsm]総一覧!$M:$M)+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & J & ",[給与Ver3.xlsm]総一覧!$P:$P)"
ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value * -1
If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
End If

If J Mod 8 = 3 Then
ws0.Cells(J, T_mon).Value = "=+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & J & ",[給与Ver3.xlsm]総一覧!$N:$N)"
ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value * -1
If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
End If

If J Mod 8 = 4 Then
ws0.Cells(J, T_mon).Value = "=+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & J & ",[給与Ver3.xlsm]総一覧!$O:$O)"
ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value * -1
If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
End If

If J Mod 8 = 5 Then
ws0.Cells(J, T_mon).Value = k_hoken
ws0.Cells(J, T_mon).Value = ws0.Cells(J, T_mon).Value
If ws0.Cells(J, T_mon).Value = 0 Then ws0.Cells(J, T_mon).Value = ""
End If




'ws2_col = ws2_col + 1

label1:
Next
'wb2.Close
'wb3.Close

wb1.Activate
ws2.Activate


    ws0.Activate
    
    
 MR01 = ws0.Cells(Rows.Count, 11).End(xlUp).Row
   
   
   
 For kk = 1 To MR01
    Cells(kk, T_mon).Select
    With Selection.Font
        .ColorIndex = 1
        .TintAndShade = 0
    End With
Next

Next


ws1.Range("C1").Value = Now & "現在"

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = "終了"



End Sub
