Attribute VB_Name = "Module2"
Option Explicit

' === quiet実行用（MsgBoxを抑止して自動実行するためのフラグ）===
Private gQuiet As Boolean

' ============================================================
' quiet実行（MsgBoxなし。COM経由の自動実行・検証用）
' ============================================================
Public Sub 経費インポートCSV作成_Quiet()
    gQuiet = True
    On Error GoTo QuietDone
    経費インポートCSV作成
QuietDone:
    gQuiet = False
End Sub

Sub 経費インポートCSV作成()
    '=============================================================
    ' 経費一覧表 → jinjerインポート用CSV 変換マクロ
    '
    ' 【処理概要】
    ' 集計シートのデータを読み取り、jinjerにインポートできる
    ' CSV形式に変換して このブックのあるフォルダ に保存します。
    '
    ' 【マッピングルール】(2026-07-16 ヘッダー見直し・11列)
    '  jinjer CSV列        ← 集計シートの列
    '  A: 社員番号          ← A列
    '  B: 氏名              ← B列
    '  C: 夜間当番手当      ← R/S・T/Uの内訳が「夜間当番手当」の金額
    '  D: 定常外業務対応手当 ← R/S・T/Uの内訳が「定常外業務対応手当」の金額
    '  E: 支給過不足調整    ← 0（旧「過不足調整」から名称変更）
    '  F: 非課税通勤費      ← V列（旧「課税通勤費」列は削除）
    '  G: 立替金（顧客請求分）← W列
    '  H: 立替金            ← X列
    '  I: その他            ← Y列（その他経費）
    '  J: その他手当        ← R/S・T/Uの内訳が「その他手当」の金額（その他とは別物）
    '  K: 現物支給          ← 0
    '=============================================================

    Dim wsSource As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim csvPath As String
    Dim fileNum As Integer
    Dim csvLine As String
    
    '--- 経費一覧表シートを取得 ---
    Set wsSource = ThisWorkbook.Sheets("集計")
    
    '--- 最終行を取得（A列で判定）---
    lastRow = wsSource.Cells(wsSource.rows.Count, "A").End(xlUp).Row
    
    '--- データがあるか確認 ---
    If lastRow < 2 Then
        If Not gQuiet Then MsgBox "集計シートにデータがありません。", vbExclamation
        Exit Sub
    End If

    '--- CSV保存先パス ---
    ' このブックのあるフォルダ（月フォルダに置く運用のため自動的に正しい月へ出る）
    csvPath = ThisWorkbook.Path & "\jinjer_経費インポート_" & Format(Date, "yyyymmdd") & ".csv"

    '--- CSVファイルを作成 ---
    fileNum = FreeFile
    Open csvPath For Output As #fileNum

    '--- ヘッダー行を書き込み（2026-07-16 見直し・11列。jinjerテンプレート「経費インポート用」と一致させること）---
    Print #fileNum, "社員番号,氏名,夜間当番手当,定常外業務対応手当," & _
                    "支給過不足調整,非課税通勤費," & _
                    "立替金（顧客請求分）,立替金,その他,その他手当,現物支給"

    '--- データ行を書き込み ---
    Dim empNo As String      ' 社員番号
    Dim empName As String     ' 氏名
    Dim nightDuty As Double   ' 夜間当番手当（R/S・T/U）
    Dim teijoDuty As Double   ' 定常外業務対応手当（R/S・T/U）
    Dim sonotaTeate As Double ' その他手当（R/S・T/U）
    Dim transport As Double   ' 非課税通勤費（V列）
    Dim custBill As Double    ' 立替金（顧客請求分）（W列）
    Dim tatekaeTax As Double  ' 立替金（X列）
    Dim otherExp As Double    ' その他（Y列）

    For i = 2 To lastRow
        '--- 集計シートからデータ取得 ---
        empNo = Trim(CStr(wsSource.Cells(i, "A").value & ""))     ' A列：社員番号
        empName = Trim(CStr(wsSource.Cells(i, "B").value & ""))   ' B列：氏名

        ' 社員番号が空欄の行はスキップ
        If empNo = "" Then GoTo NextRow

        nightDuty = 0
        teijoDuty = 0
        sonotaTeate = 0
        AddAllowanceFromPair wsSource.Cells(i, "R").value, wsSource.Cells(i, "S").value, nightDuty, teijoDuty, sonotaTeate
        AddAllowanceFromPair wsSource.Cells(i, "T").value, wsSource.Cells(i, "U").value, nightDuty, teijoDuty, sonotaTeate

        transport = ValJP(wsSource.Cells(i, "V").value)   ' V列：非課税通勤費
        custBill = ValJP(wsSource.Cells(i, "W").value)    ' W列：立替金（顧客請求分）
        tatekaeTax = ValJP(wsSource.Cells(i, "X").value)  ' X列：立替金
        otherExp = ValJP(wsSource.Cells(i, "Y").value)    ' Y列：その他

        '--- CSV行を組み立て ---
        ' カンマ区切りで値を連結
        ' 文字列項目はダブルクォートで囲む
        csvLine = EscapeCSV(empNo) & "," & _
                  EscapeCSV(empName) & "," & _
                  NumText(nightDuty) & "," & _
                  NumText(teijoDuty) & "," & _
                  "0" & "," & _
                  NumText(transport) & "," & _
                  NumText(custBill) & "," & _
                  NumText(tatekaeTax) & "," & _
                  NumText(otherExp) & "," & _
                  NumText(sonotaTeate) & "," & _
                  "0"

        '--- 書き込み ---
        Print #fileNum, csvLine

NextRow:
    Next i

    '--- ファイルを閉じる ---
    Close #fileNum

    '--- 完了メッセージ ---
    If Not gQuiet Then MsgBox "jinjerインポート用CSVを作成しました！" & vbCrLf & vbCrLf & _
           "保存先: " & csvPath & vbCrLf & _
           "対象: " & (lastRow - 1) & " 件", vbInformation

End Sub

'=============================================================
' R/S・T/U の内訳名と金額をCSV出力用の手当に振り分ける
'=============================================================
Private Sub AddAllowanceFromPair(ByVal allowanceName As Variant, _
                                 ByVal amountValue As Variant, _
                                 ByRef nightDuty As Double, _
                                 ByRef teijoDuty As Double, _
                                 ByRef sonotaTeate As Double)
    Dim nameText As String
    Dim amount As Double

    nameText = Trim$(CStr(allowanceName))
    amount = ValJP(amountValue)

    Select Case nameText
        Case "夜間当番手当"
            nightDuty = nightDuty + amount
        Case "定常外業務対応手当"
            teijoDuty = teijoDuty + amount
        Case "その他手当"
            sonotaTeate = sonotaTeate + amount
    End Select
End Sub

'=============================================================
' 日本語Excel表示の金額を数値化する
'=============================================================
Private Function ValJP(ByVal v As Variant) As Double
    If IsError(v) Or isEmpty(v) Then Exit Function
    
    Dim s As String
    s = Trim$(CStr(v))
    s = StrConv(s, vbNarrow)
    s = Replace(s, "\", "")
    s = Replace(s, "円", "")
    s = Replace(s, ",", "")
    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    
    If Len(s) >= 2 And Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        s = "-" & Mid$(s, 2, Len(s) - 2)
    End If
    
    If s <> "" And IsNumeric(s) Then ValJP = CDbl(s)
End Function

'=============================================================
' CSVに出す数値文字列を整える
'=============================================================
Private Function NumText(ByVal v As Double) As String
    If v = 0 Then
        NumText = "0"
    Else
        Dim s As String
        s = Format$(v, "0.########")
        Do While InStr(1, s, ".", vbBinaryCompare) > 0 And Right$(s, 1) = "0"
            s = Left$(s, Len(s) - 1)
        Loop
        If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
        NumText = s
    End If
End Function

'=============================================================
' CSV用エスケープ関数
' カンマやダブルクォートを含む文字列を安全にCSV出力するための関数
'
' 【やっていること】
' ① 値の中にダブルクォート（"）がある場合、""に置き換える
'    （CSVのルールで、"は""と書く決まり）
' ② 値全体をダブルクォートで囲んで返す
'    例："山田 太郎" → """山田 太郎"""ではなく"山田 太郎"
'    例："株式会社""ABC""" のようにクォート含む場合も安全
'=============================================================
Private Function EscapeCSV(ByVal val As String) As String
    ' ダブルクォートを2つに置き換え
    val = Replace(val, """", """""")
    ' 全体をダブルクォートで囲む
    EscapeCSV = """" & val & """"
End Function


