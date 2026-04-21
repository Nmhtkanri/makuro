Attribute VB_Name = "Module1"

Private Function NormalizeSheetNameKey(ByVal sheetName As String) As String
    Dim key As String
    key = CStr(sheetName)
    key = Replace$(key, vbTab, "")
    key = Replace$(key, Chr$(160), "")
    key = Replace$(key, ChrW$(&H3000), "")
    key = Replace$(key, " ", "")
    key = Replace$(key, "（", "(")
    key = Replace$(key, "）", ")")
    NormalizeSheetNameKey = LCase$(Trim$(key))
End Function

Private Function TryGetSheetByNameLooseLocal(ByVal wb As Workbook, ByVal targetName As String, ByRef wsOut As Worksheet) As Boolean
    Dim targetKey As String
    targetKey = NormalizeSheetNameKey(targetName)
    If targetKey = "" Then Exit Function

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If NormalizeSheetNameKey(ws.Name) = targetKey Then
            Set wsOut = ws
            TryGetSheetByNameLooseLocal = True
            Exit Function
        End If
    Next ws
End Function

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
For j = 1 To Sheets.Count
s_name(k + j) = Sheets(j).Name
Next j
k = k + j
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
Dim targetBookPath As String


Set wb1 = ActiveWorkbook
On Error Resume Next
Set ws1 = wb1.Worksheets("まとめ")
Set ws2 = wb1.Worksheets("操作画面")
On Error GoTo 0
If ws1 Is Nothing Or ws2 Is Nothing Then
    MsgBox "実行元ブックに「まとめ」または「操作画面」シートが見つかりません。" & vbCrLf & _
           "現在のアクティブブック: " & wb1.Name, vbExclamation
    GoTo SafeExit_Matome
End If

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

ws1.Range("D3:D11").AutoFill Destination:=ws1.Range("D3:O38"), Type:=xlFillDefault

ws1.Cells(3, 16).Value = "=SUM(D3:O3)"

ws1.Range("P3").AutoFill Destination:=ws1.Range("P3:P38"), Type:=xlFillDefault

'DELL分修正

Application.StatusBar = "DELL分修正"


ws_name = Trim$(CStr(ws2.Range("G3").Value))
targetBookPath = Trim$(CStr(ws2.Range("F3").Value))
L_num = 0

For i = 3 To MR1
If ws1.Cells(i, 1).Value = "DELL" Then
If ws1.Cells(i, 2).Value = "総受注金額" Then
If ws1.Cells(i, 3).Value = "計画" Then
L_num = i
End If
End If
End If
Next
If L_num = 0 Then
    MsgBox "まとめシートで DELL / 総受注金額 / 計画 の行が見つかりません。", vbExclamation
    GoTo SafeExit_Matome
End If
If ws_name = "" Then
    MsgBox "操作画面G3（DELLのシート名）が空です。", vbExclamation
    GoTo SafeExit_Matome
End If
If targetBookPath = "" Then
    MsgBox "操作画面F3（DELLのブックパス）が空です。", vbExclamation
    GoTo SafeExit_Matome
End If

Workbooks.Open Filename:=targetBookPath, UpdateLinks:=False
Set wb0 = ActiveWorkbook
On Error Resume Next
Set ws0 = wb0.Worksheets(ws_name)
On Error GoTo 0
If ws0 Is Nothing Then
    MsgBox "売上シートが見つかりません。" & vbCrLf & _
           "操作画面シート名: " & ws_name & vbCrLf & _
           "対象ブック: " & targetBookPath, vbExclamation
    wb0.Close SaveChanges:=False
    Set wb0 = Nothing
    GoTo SafeExit_Matome
End If

ws0.Select

For i = 1 To 12
ws1.Cells(L_num, i + 3).Value = ws0.Cells(2, i + 8).Value
Next

wb0.Close SaveChanges:=False
Set wb0 = Nothing

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


ws_name = Trim$(CStr(ws2.Range("G4").Value))
targetBookPath = Trim$(CStr(ws2.Range("F4").Value))
L_num = 0

For i = 2 To MR1
If ws1.Cells(i, 1).Value = "ホンダ" Then
If ws1.Cells(i, 2).Value = "総受注金額" Then
If ws1.Cells(i, 3).Value = "計画" Then
L_num = i
End If
End If
End If
Next
If L_num = 0 Then
    MsgBox "まとめシートで ホンダ / 総受注金額 / 計画 の行が見つかりません。", vbExclamation
    GoTo SafeExit_Matome
End If
If ws_name = "" Then
    MsgBox "操作画面G4（ホンダのシート名）が空です。", vbExclamation
    GoTo SafeExit_Matome
End If
If targetBookPath = "" Then
    MsgBox "操作画面F4（ホンダのブックパス）が空です。", vbExclamation
    GoTo SafeExit_Matome
End If

Workbooks.Open Filename:=targetBookPath, UpdateLinks:=False
Set wb0 = ActiveWorkbook
On Error Resume Next
Set ws0 = wb0.Worksheets(ws_name)
On Error GoTo 0
If ws0 Is Nothing Then
    MsgBox "売上シートが見つかりません。" & vbCrLf & _
           "操作画面シート名: " & ws_name & vbCrLf & _
           "対象ブック: " & targetBookPath, vbExclamation
    wb0.Close SaveChanges:=False
    Set wb0 = Nothing
    GoTo SafeExit_Matome
End If

ws0.Select

For i = 1 To 12
ws1.Cells(L_num, i + 3).Value = ws0.Cells(2, i + 8).Value
Next

wb0.Close SaveChanges:=False
Set wb0 = Nothing

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


ws1.Range("R3:W5").AutoFill Destination:=ws1.Range("R3:W38"), Type:=xlFillDefault

ws1.Columns("D:V").AutoFit
ws1.Columns("D:V").Style = "Comma [0]"

wb1.Activate
ws1.Activate
ws1.Range("A1").Select

SafeExit_Matome:
On Error Resume Next
If Not wb0 Is Nothing Then wb0.Close SaveChanges:=False
On Error GoTo 0
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.StatusBar = "終了"

End Sub

Sub 一括実行_データ収集とまとめ()

Dim phase As String
Dim controlWb As Workbook
Set controlWb = ActiveWorkbook
On Error GoTo ErrHandler

phase = "データ収集"
Application.StatusBar = "一括実行: データ収集"
Call データ収集

If Not controlWb Is Nothing Then controlWb.Activate
phase = "まとめ"
Application.StatusBar = "一括実行: まとめ"
Call まとめ

Application.StatusBar = False
MsgBox "データ収集 → まとめ の実行が完了しました。", vbInformation
Exit Sub

ErrHandler:
Application.StatusBar = False
MsgBox "一括実行中にエラーが発生しました。工程: " & phase & vbCrLf & _
       "Err " & Err.Number & ": " & Err.Description, vbExclamation

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
Set ws1 = wb1.Worksheets("まとめ")
Set ws2 = wb1.Worksheets("操作画面")

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
If Not controlWb Is Nothing Then controlWb.Activate
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
Set ws0 = Nothing
If Not TryGetSheetByNameLooseLocal(wb0, CStr(ws1.Range("C" & Format(i)).Value), ws0) Then
MsgBox "操作画面C" & i & "のシートが見つかりません: " & ws1.Range("C" & Format(i)).Value, vbExclamation
wb0.Close SaveChanges:=False
GoTo label2
End If
MR0 = ws0.Cells(Rows.Count, 1).End(xlUp).Row
ws0.Activate

For j = 1 To MR0
If ws0.Cells(j, 1).Value = "" Then
If ws0.Cells(j, 2).Value <> "" Then
MsgBox "社員番号記載漏れ発見！"
Exit Sub
End If
End If
Next
wb0.Close
label2:
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
Dim ws3Kihon As Worksheet '経費データシート（基本給参照用）

Dim wb0 As Workbook '書き込み先ブック
Dim wb1 As Workbook 'ボタンのあるブック
Dim wb2 As Workbook '売上データブック
Dim wb3 As Workbook '経費データブック

Dim bias As Integer '各表の列方向の項目の位置のズレを補正する変数
Dim ws2_col As Integer '集計表に書き出すときの列番号

Dim FCell As Range  '検索先の場所
Dim FCellBase As Range  '基本給参照シート検索先
Dim FFrst As Range  '項目検索
Dim wRow As Long    '該当列
Dim wRowBase As Long '基本給参照シート該当列

Dim wsDell As Worksheet   ' DELL委託人件費（FS）シート
Dim wsDellCandidate As Worksheet
Dim T_mon2 As Long        ' DELLシート側の今月列
Dim salesSheetName As String
Dim expenseSheetName As String
Dim targetSheetName As String
Const BASIC_PAY_COL As Long = 13 'M列（基本給）

Set wb1 = ActiveWorkbook
Set ws1 = ActiveSheet

Application.StatusBar = "売上データ読み込み中"
Workbooks.Open Filename:=ws1.Range("G" & Format(3)).Value, UpdateLinks:=False
Set wb2 = ActiveWorkbook
salesSheetName = Trim$(CStr(ws1.Range("H" & Format(3)).Value))
Set ws2 = Nothing
If Not TryGetSheetByNameLooseLocal(wb2, salesSheetName, ws2) Then
    MsgBox "操作画面H3の売上データシートが見つかりません: " & salesSheetName, vbExclamation
    GoTo SafeExit
End If
If ws2 Is Nothing Then
    MsgBox "操作画面H3の売上データシートが見つかりません: " & salesSheetName, vbExclamation
    GoTo SafeExit
End If
Application.StatusBar = "経費データ読み込み中"
Workbooks.Open Filename:=ws1.Range("G" & Format(4)).Value, UpdateLinks:=False
Set wb3 = ActiveWorkbook
expenseSheetName = Trim$(CStr(ws1.Range("H" & Format(4)).Value))
Set ws3 = Nothing
If Not TryGetSheetByNameLooseLocal(wb3, expenseSheetName, ws3) Then
    MsgBox "操作画面H4の経費データシートが見つかりません: " & expenseSheetName, vbExclamation
    GoTo SafeExit
End If
If ws3 Is Nothing Then
    MsgBox "操作画面H4の経費データシートが見つかりません: " & expenseSheetName, vbExclamation
    GoTo SafeExit
End If
Set ws3Kihon = ws3

Dim kihonSheetName As String
kihonSheetName = Trim$(CStr(ws1.Range("I4").Value))
If kihonSheetName <> "" Then
    Set ws3Kihon = Nothing
    If Not TryGetSheetByNameLooseLocal(wb3, kihonSheetName, ws3Kihon) Then
        MsgBox "I4の基本給参照シートが見つかりません。H4シートを使用します。", vbExclamation
        Set ws3Kihon = ws3
    End If
    If ws3Kihon Is Nothing Then
        MsgBox "I4の基本給参照シートが見つかりません。H4シートを使用します。", vbExclamation
        Set ws3Kihon = ws3
    End If
End If

wb1.Activate

this_month = ws1.Range("G1").Value & "月"

' === 給与明細データシートの列確認（初回のみ） ===
If i = 3 Then
    Debug.Print "=== 経費データシートの列確認 ==="
    Debug.Print "93列目のヘッダー: " & ws3.Cells(1, 93).Value
    Debug.Print "137列目のヘッダー: " & ws3.Cells(1, 137).Value
    Debug.Print "138列目のヘッダー: " & ws3.Cells(1, 138).Value
    Debug.Print "139列目のヘッダー: " & ws3.Cells(1, 139).Value
    Debug.Print "140列目のヘッダー: " & ws3.Cells(1, 140).Value
    Debug.Print "基本給列(M=13)ヘッダー: " & ws3.Cells(1, BASIC_PAY_COL).Value
    Debug.Print "180列目のヘッダー: " & ws3.Cells(1, 180).Value
End If

MR1 = ws1.Cells(Rows.Count, 1).End(xlUp).Row
ws2_col = 2

For i = 3 To MR1
    bias = ws1.Cells(i, 4).Value
    Workbooks.Open Filename:=ws1.Range("B" & Format(i)).Value, UpdateLinks:=False
    Set wb0 = ActiveWorkbook
    targetSheetName = Trim$(CStr(ws1.Range("C" & Format(i)).Value))
    Set ws0 = Nothing
    If Not TryGetSheetByNameLooseLocal(wb0, targetSheetName, ws0) Then
        MsgBox "操作画面C" & i & "のシートが見つかりません: " & targetSheetName, vbExclamation
        wb0.Close SaveChanges:=False
        GoTo NextWorkbook
    End If
    If ws0 Is Nothing Then
        MsgBox "操作画面C" & i & "のシートが見つかりません: " & targetSheetName, vbExclamation
        wb0.Close SaveChanges:=False
        GoTo NextWorkbook
    End If
    MR0 = ws0.Cells(Rows.Count, 1).End(xlUp).Row

    ' ===== DELL委託人件費（FS）シート設定 =====
    Set wsDell = Nothing
    On Error Resume Next
    Set wsDell = wb0.Worksheets("DELL委託人件費（FS）")
    If wsDell Is Nothing Then Set wsDell = wb0.Worksheets("DELL委託人件費(FS）")
    On Error GoTo 0
    If wsDell Is Nothing Then
        For Each wsDellCandidate In wb0.Worksheets
            If InStr(1, wsDellCandidate.Name, "DELL委託人件費", vbTextCompare) > 0 And _
               InStr(1, wsDellCandidate.Name, "FS", vbTextCompare) > 0 Then
                Set wsDell = wsDellCandidate
                Exit For
            End If
        Next wsDellCandidate
    End If

    T_mon2 = 0
    If Not wsDell Is Nothing Then
        Dim mCell As Range
        Set mCell = wsDell.Rows(1).Find(What:=this_month, LookIn:=xlValues, _
                                        LookAt:=xlWhole, SearchOrder:=xlByColumns, _
                                        SearchDirection:=xlNext)
        If Not mCell Is Nothing Then
            T_mon2 = mCell.Column
            Debug.Print "FSシート月列発見: " & T_mon2 & "列目 (" & this_month & ")"
        Else
            Debug.Print "FSシートで月列が見つかりません: " & this_month
            Debug.Print "FSシート1行目の値: " & wsDell.Range("A1:L1").Value2(1, 1) & "~" & wsDell.Range("A1:L1").Value2(1, 12)
        End If
    End If
    ' ===== ここまで DELLシート設定 =====

    ' 月列を検索（エラーハンドリング付き）
    On Error Resume Next
    T_mon = WorksheetFunction.Match(this_month, ws0.Range("A1:Z1"), 0) '何月かをとってきている
    If Err.Number <> 0 Then
        ' 月が見つからない場合は処理をスキップ
        Err.Clear
        GoTo NextWorkbook
    End If
    On Error GoTo 0

    ' 顧客対応当番手当の列を検索
    On Error Resume Next
    Dim dutyCol As Long
    Dim v As Variant
    dutyCol = 0
    v = Application.Match("顧客対応当番手当", ws3.Rows(1), 0)
    If IsError(v) Then v = Application.Match("顧客対応当番", ws3.Rows(1), 0)
    If IsError(v) Then v = Application.Match("顧客対応", ws3.Rows(1), 0)
    If Not IsError(v) Then dutyCol = CLng(v)
    Dim isUALSheet As Boolean
    isUALSheet = (InStr(1, ws0.Name, "UAL常駐", vbTextCompare) > 0)

    For j = 3 To MR0
        ' 進捗表示
        Dim pct As String
        pct = Format(j / MR0, "0.0%")
        Application.StatusBar = wb0.Name & " ワークブックの" & ws0.Name & " 処理中 " & pct & "終了"

        ' ブロック頭で社員番号を探す＆転出スキップ
        If isUALSheet Then
            Dim ualItemName As String
            ualItemName = Trim$(CStr(ws0.Cells(j, 11).Value))

            If ualItemName = "総受注金額" Then
                wRow = 0
                wRowBase = 0
                If Trim$(CStr(ws0.Cells(j, 1).Value)) <> "" Then
                    Set FCell = ws3.Columns(1).Find(What:=CStr(ws0.Cells(j, 1).Value), _
                                                    LookIn:=xlValues, LookAt:=xlWhole)
                    Set FCellBase = ws3Kihon.Columns(1).Find(What:=CStr(ws0.Cells(j, 1).Value), _
                                                             LookIn:=xlValues, LookAt:=xlWhole)
                    If Not FCell Is Nothing Then
                        wRow = FCell.Row
                        If Not FCellBase Is Nothing Then wRowBase = FCellBase.Row
                        If wRowBase = 0 Then wRowBase = wRow
                    End If
                End If
            End If

            Select Case ualItemName
                Case "賃金"
                    ws0.Cells(j, T_mon).Value = ""
                    If wRowBase > 0 Then
                        ws0.Cells(j, T_mon).Value = ws3Kihon.Cells(wRowBase, BASIC_PAY_COL).Value
                    ElseIf wRow > 0 Then
                        ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, BASIC_PAY_COL).Value
                    End If
                    If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
                Case "顧客対応当番", "顧客対応当番手当", "交通費"
                    ws0.Cells(j, T_mon).Value = ""
                    If wRow > 0 And dutyCol > 0 Then
                        ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, dutyCol).Value
                    End If
                    If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
                Case "通勤定期代"
                    ws0.Cells(j, T_mon).Value = ""
                    If wRow > 0 Then ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, 93).Value
                    If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
                Case "健康保険"
                    ws0.Cells(j, T_mon).Value = ""
                    If wRow > 0 Then
                        ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, 138).Value + ws3.Cells(wRow, 139).Value
                    End If
                    If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
                Case "厚生年金"
                    ws0.Cells(j, T_mon).Value = ""
                    If wRow > 0 Then ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, 140).Value
                    If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
                Case "労災保険"
                    ws0.Cells(j, T_mon).Value = ""
                    If wRow > 0 Then ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, 180).Value
                    If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
                Case "雇用保険"
                    ws0.Cells(j, T_mon).Value = ""
                    If wRow > 0 Then ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, 137).Value
                    If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
            End Select

            GoTo label1
        End If
        If (j - 3) Mod 8 = 0 Then
            wRow = 0
            wRowBase = 0
            
            ' 転出の場合はスキップ
            If ws0.Cells(j, 5).Value = "転出" Then
                j = j + 7
                GoTo label1
            End If
            
            ' 現在処理中の社員番号
            Dim currentEmpNo As String
            currentEmpNo = CStr(ws0.Cells(j, 1).Value)
            
            ' 給与明細データから社員番号を検索
            If Trim(currentEmpNo) <> "" Then
                Set FCell = ws3.Columns(1).Find(What:=currentEmpNo, _
                                                LookIn:=xlValues, LookAt:=xlWhole)
                Set FCellBase = ws3Kihon.Columns(1).Find(What:=currentEmpNo, _
                                                         LookIn:=xlValues, LookAt:=xlWhole)
            End If

            If Not FCell Is Nothing And Trim(currentEmpNo) <> "" Then
                wRow = FCell.Row
                If Not FCellBase Is Nothing Then wRowBase = FCellBase.Row
                If wRowBase = 0 Then wRowBase = wRow
                Debug.Print "社員番号 " & currentEmpNo & " 経費データ行: " & wRow
                ' === 給与明細データの値確認 ===
                Debug.Print "  基本給(M=13): " & ws3.Cells(wRow, BASIC_PAY_COL).Value
                If wRowBase > 0 Then Debug.Print "  基本給参照(M=13): " & ws3Kihon.Cells(wRowBase, BASIC_PAY_COL).Value
                Debug.Print "  非課税通勤費(93列): " & ws3.Cells(wRow, 93).Value
                Debug.Print "  健康保険料(138列): " & ws3.Cells(wRow, 138).Value
                Debug.Print "  介護保険料(139列): " & ws3.Cells(wRow, 139).Value
                Debug.Print "  厚生年金保険料(140列): " & ws3.Cells(wRow, 140).Value
                Debug.Print "  事業主労災保険料(180列): " & ws3.Cells(wRow, 180).Value
                Debug.Print "  雇用保険料(137列): " & ws3.Cells(wRow, 137).Value
                If dutyCol > 0 Then Debug.Print "  当番手当(" & dutyCol & "列): " & ws3.Cells(wRow, dutyCol).Value
            Else
                Debug.Print "社員番号 " & currentEmpNo & " 経費データで見つからず"
            End If
            
            ' === DELL委託人件費（FS）シートへの書き込み ===
            If Not wsDell Is Nothing And T_mon2 > 0 And wRow > 0 Then
                ' FSシートで該当する社員番号の行を検索
                Dim fEmpFS As Range
                Set fEmpFS = wsDell.Columns(1).Find(What:=currentEmpNo, _
                                                   LookIn:=xlValues, LookAt:=xlWhole)
                
                If Not fEmpFS Is Nothing Then
                    Dim baseRowFS As Long
                    baseRowFS = fEmpFS.Row  ' 社員番号の行を基準とする
                    Debug.Print "FSシート社員番号 " & currentEmpNo & " 発見: " & baseRowFS & "行目"
                    
                    ' 社員番号行から各項目行に書き込み（SIシートの計算ロジックを適用）
                    ' 賃金（社員番号行+1）- 源泉対象額
                    If baseRowFS + 1 <= wsDell.Rows.Count Then
                        wsDell.Cells(baseRowFS + 1, T_mon2).Value = ""
                        If wRowBase > 0 Then
                            wsDell.Cells(baseRowFS + 1, T_mon2).Value = ws3Kihon.Cells(wRowBase, BASIC_PAY_COL).Value '基本給（基本給参照シート優先）
                        Else
                            wsDell.Cells(baseRowFS + 1, T_mon2).Value = ws3.Cells(wRow, BASIC_PAY_COL).Value '基本給
                        End If
                        If wsDell.Cells(baseRowFS + 1, T_mon2).Value = 0 Then wsDell.Cells(baseRowFS + 1, T_mon2).Value = "￥0"
                    End If
                    
                    ' 当番手当（社員番号行+2）
                    If baseRowFS + 2 <= wsDell.Rows.Count Then
                        wsDell.Cells(baseRowFS + 2, T_mon2).Value = ""
                        If dutyCol > 0 Then
                            wsDell.Cells(baseRowFS + 2, T_mon2).Value = ws3.Cells(wRow, dutyCol).Value
                        End If
                        If wsDell.Cells(baseRowFS + 2, T_mon2).Value = 0 Then wsDell.Cells(baseRowFS + 2, T_mon2).Value = "￥0"
                    End If
                    
                    ' 交通費（社員番号行+3）- 非課税通勤費
                    If baseRowFS + 3 <= wsDell.Rows.Count Then
                        wsDell.Cells(baseRowFS + 3, T_mon2).Value = ""
                        wsDell.Cells(baseRowFS + 3, T_mon2).Value = ws3.Cells(wRow, 93).Value '非課税通勤費
                        If wsDell.Cells(baseRowFS + 3, T_mon2).Value = 0 Then wsDell.Cells(baseRowFS + 3, T_mon2).Value = "￥0"
                    End If
                    
                    ' 健保＋介護（社員番号行+4）- 健康保険料＋介護保険料
                    If baseRowFS + 4 <= wsDell.Rows.Count Then
                        wsDell.Cells(baseRowFS + 4, T_mon2).Value = ""
                        wsDell.Cells(baseRowFS + 4, T_mon2).Value = ws3.Cells(wRow, 138).Value '健康保険料
                        wsDell.Cells(baseRowFS + 4, T_mon2).Value = wsDell.Cells(baseRowFS + 4, T_mon2).Value + ws3.Cells(wRow, 139).Value '介護保険料
                        If wsDell.Cells(baseRowFS + 4, T_mon2).Value = 0 Then wsDell.Cells(baseRowFS + 4, T_mon2).Value = "￥0"
                    End If
                    
                    ' 厚生年金（社員番号行+5）- 厚生年金保険料
                    If baseRowFS + 5 <= wsDell.Rows.Count Then
                        wsDell.Cells(baseRowFS + 5, T_mon2).Value = ""
                        wsDell.Cells(baseRowFS + 5, T_mon2).Value = ws3.Cells(wRow, 140).Value '厚生年金保険料
                        If wsDell.Cells(baseRowFS + 5, T_mon2).Value = 0 Then wsDell.Cells(baseRowFS + 5, T_mon2).Value = "￥0"
                    End If
                    
                    ' 労災（社員番号行+6）- 事業主労災保険料
                    If baseRowFS + 6 <= wsDell.Rows.Count Then
                        wsDell.Cells(baseRowFS + 6, T_mon2).Value = ""
                        wsDell.Cells(baseRowFS + 6, T_mon2).Value = ws3.Cells(wRow, 180).Value '事業主労災保険料
                        If wsDell.Cells(baseRowFS + 6, T_mon2).Value = 0 Then wsDell.Cells(baseRowFS + 6, T_mon2).Value = "￥0"
                    End If
                    
                    ' 雇用（社員番号行+7）- 雇用保険料
                    If baseRowFS + 7 <= wsDell.Rows.Count Then
                        wsDell.Cells(baseRowFS + 7, T_mon2).Value = ""
                        wsDell.Cells(baseRowFS + 7, T_mon2).Value = ws3.Cells(wRow, 137).Value '雇用保険料
                        If wsDell.Cells(baseRowFS + 7, T_mon2).Value = 0 Then wsDell.Cells(baseRowFS + 7, T_mon2).Value = "￥0"
                    End If
                Else
                    Debug.Print "FSシートで社員番号 " & currentEmpNo & " が見つかりません"
                End If
            End If
            ' === FSシートへの書き込み終了 ===
        End If

        ' === DPTシート（ws0）への書き込み（SIシートの計算ロジックを適用） ===
        ' 賃金（余り1）- 源泉対象額
        If (j - 3) Mod 8 = 1 Then
            ws0.Cells(j, T_mon).Value = ""
            If wRowBase > 0 Then
                ws0.Cells(j, T_mon).Value = ws3Kihon.Cells(wRowBase, BASIC_PAY_COL).Value '基本給（基本給参照シート優先）
            ElseIf wRow > 0 Then
                ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, BASIC_PAY_COL).Value '基本給
            End If
            If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
        End If

        ' 顧客対応当番手当（余り2）
        If (j - 3) Mod 8 = 2 Then
            ws0.Cells(j, T_mon).Value = ""
            If wRow > 0 And dutyCol > 0 Then
                ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, dutyCol).Value
            End If
            If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
        End If

        ' 交通費（余り3）- 非課税通勤費
        If (j - 3) Mod 8 = 3 Then
            ws0.Cells(j, T_mon).Value = ""
            If wRow > 0 Then
                ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, 93).Value '非課税通勤費
            End If
            If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
        End If

        ' 健康保険＋介護（余り4）- 健康保険料＋介護保険料
        If (j - 3) Mod 8 = 4 Then
            ws0.Cells(j, T_mon).Value = ""
            If wRow > 0 Then
                ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, 138).Value '健康保険料
                ws0.Cells(j, T_mon).Value = ws0.Cells(j, T_mon).Value + ws3.Cells(wRow, 139).Value '介護保険料
            End If
            If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
        End If

        ' 厚生年金（余り5）- 厚生年金保険料
        If (j - 3) Mod 8 = 5 Then
            ws0.Cells(j, T_mon).Value = ""
            If wRow > 0 Then
                ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, 140).Value '厚生年金保険料
            End If
            If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
        End If

        ' 労災保険（余り6）- 事業主労災保険料
        If (j - 3) Mod 8 = 6 Then
            ws0.Cells(j, T_mon).Value = ""
            If wRow > 0 Then
                ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, 180).Value '事業主労災保険料
            End If
            If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
        End If

        ' 雇用保険（余り7）- 雇用保険料
        If (j - 3) Mod 8 = 7 Then
            ws0.Cells(j, T_mon).Value = ""
            If wRow > 0 Then
                ws0.Cells(j, T_mon).Value = ws3.Cells(wRow, 137).Value '雇用保険料
            End If
            If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = "￥0"
        End If

label1:
    Next j

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


NextWorkbook:
Next

ws1.Range("C1").Value = Now & "現在"

SafeExit:
' 開いたファイルを閉じる
If Not wb2 Is Nothing Then wb2.Close SaveChanges:=False
If Not wb3 Is Nothing Then wb3.Close SaveChanges:=False

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
For j = 7 To MR0

If j Mod 8 = 7 Then
If ws0.Cells(j, 5).Value = "転出" Then
j = j + 11
GoTo label1
End If
End If

pct = j / MR0
pct = Format(pct, "0.0%")
If Not controlWb Is Nothing Then controlWb.Activate
Application.StatusBar = wb0.Name & " ワークブックの" & ws0.Name & " 処理中 " & pct & "終了"



'If j Mod 12 = 6 Then
'ws0.Cells(j, T_mon).Value = "=+SUMIFS('[" & wb2.Name & "]" & ws2.Name & "'!$H:$H,'[" & wb2.Name & "]" & ws2.Name & "'!$B:$B,A" & j & ",'[" & wb2.Name & "]" & ws2.Name & "'!$F:$F,""売上高"")/1.1"
'ws0.Cells(j, T_mon).Value = Round(ws0.Cells(j, T_mon).Value, 0)
'If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = ""
'End If

If j Mod 8 = 0 Then
ws0.Cells(j, T_mon).Value = "=+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & j & ",[給与Ver3.xlsm]総一覧!$M:$M)"
ws0.Cells(j, T_mon).Value = ws0.Cells(j, T_mon).Value
If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = ""
k_hoken = ws0.Cells(j, T_mon).Value * 6 / 1000
End If

If j Mod 8 = 1 Then
ws0.Cells(j, T_mon).Value = "=+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & j & ",[給与Ver3.xlsm]総一覧!$K:$K)"
ws0.Cells(j, T_mon).Value = ws0.Cells(j, T_mon).Value
If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = ""
End If

If j Mod 8 = 2 Then
ws0.Cells(j, T_mon).Value = "=+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & j & ",[給与Ver3.xlsm]総一覧!$M:$M)+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & j & ",[給与Ver3.xlsm]総一覧!$P:$P)"
ws0.Cells(j, T_mon).Value = ws0.Cells(j, T_mon).Value * -1
If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = ""
End If

If j Mod 8 = 3 Then
ws0.Cells(j, T_mon).Value = "=+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & j & ",[給与Ver3.xlsm]総一覧!$N:$N)"
ws0.Cells(j, T_mon).Value = ws0.Cells(j, T_mon).Value * -1
If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = ""
If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = ""
End If

If j Mod 8 = 4 Then
ws0.Cells(j, T_mon).Value = "=+SUMIF([給与Ver3.xlsm]総一覧!$A:$A,A" & j & ",[給与Ver3.xlsm]総一覧!$O:$O)"
ws0.Cells(j, T_mon).Value = ws0.Cells(j, T_mon).Value * -1
If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = ""
End If

If j Mod 8 = 5 Then
ws0.Cells(j, T_mon).Value = k_hoken
ws0.Cells(j, T_mon).Value = ws0.Cells(j, T_mon).Value
If ws0.Cells(j, T_mon).Value = 0 Then ws0.Cells(j, T_mon).Value = ""
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

