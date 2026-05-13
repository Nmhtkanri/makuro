Attribute VB_Name = "Module2"
Sub 経費インポートCSV作成()
    '=============================================================
    ' 経費一覧表 → jinjerインポート用CSV 変換マクロ
    '
    ' 【処理概要】
    ' 経費一覧表シートのデータを読み取り、jinjerにインポートできる
    ' CSV形式に変換して Z:\jinjer移行\共有 に保存します。
    '
    ' 【マッピングルール】
    '  jinjer CSV列    ← 経費一覧表の列
    '  A: 社員番号      ← A列
    '  B: 氏名          ← B列
    '  C: 夜間当番手当  ← F列（手当2＝夜間＋RINK）
    '  D: 営業手当      ← 0
    '  E: 現場管理費    ← 0
    '  F: テレワーク手当 ← J列
    '  G: 定常外業務対応手当 ← 空欄（手入力用）
    '  H: 家賃手当      ← 0
    '  I: その他手当    ← 0
    '  J: 過不足調整    ← 0
    '  K: 課税通勤費    ← 0
    '  L: 非課税通勤費  ← H列（交通費）※立替金ありの人は0
    '  M: 立替金(顧客請求分) ← G列
    '  N: 立替金        ← X列（V+X+Y合算済み）※立替金ありの人のみ
    '  O: その他        ← I列 ※立替金ありの人は0
    '
    ' 【立替金の判定】
    '  X列（非課税精算・立替金）に値がある場合：
    '    → N列にX列の値を入れる
    '    → L列（非課税通勤費）を0にする
    '    → O列（その他）を0にする
    '  X列が0または空欄の場合：
    '    → N列は0
    '    → L列にH列（交通費）を入れる
    '    → O列にI列（その他）を入れる
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
        MsgBox "経費一覧表にデータがありません。", vbExclamation
        Exit Sub
    End If
    
    '--- CSV保存先パス ---
    ' ファイル名に年月を付けて保存
    csvPath = "Z:\NMHT総務関係\freee\作業データ\jinjer_経費インポート_" & Format(Date, "yyyymmdd") & ".csv"
    
    '--- 保存先フォルダの存在確認 ---
    If Dir("Z:\jinjer移行\共有", vbDirectory) = "" Then
        MsgBox "保存先フォルダが見つかりません。" & vbCrLf & _
               "Z:\jinjer移行\共有 を確認してください。", vbExclamation
        Exit Sub
    End If
    
    '--- CSVファイルを作成 ---
    fileNum = FreeFile
    Open csvPath For Output As #fileNum
    
    '--- ヘッダー行を書き込み ---
    Print #fileNum, "社員番号,氏名,夜間当番手当,営業手当,現場管理費," & _
                    "テレワーク手当,定常外業務対応手当,家賃手当,その他手当," & _
                    "過不足調整,課税通勤費,非課税通勤費," & _
                    "立替金（顧客請求分）,立替金,その他"
    
    '--- データ行を書き込み ---
    Dim empNo As String      ' 社員番号
    Dim empName As String     ' 氏名
    Dim nightDuty As Variant  ' 夜間当番手当（F列：手当2）
    Dim telework As Variant   ' テレワーク手当（J列）
    Dim transport As Variant  ' 交通費（H列）
    Dim custBill As Variant   ' 顧客請求分（G列）
    Dim otherExp As Variant   ' その他（I列）
    Dim tatekaeTax As Variant ' 非課税精算・立替金（X列）= V+X+Y合算済み
    Dim hasTA As Boolean      ' 立替金有無フラグ
    
    ' CSV出力用の値
    Dim csvL As Variant       ' 非課税通勤費
    Dim csvN As Variant       ' 立替金
    Dim csvO As Variant       ' その他
    
    For i = 2 To lastRow
        '--- 経費一覧表からデータ取得 ---
        empNo = Trim(CStr(wsSource.Cells(i, "A").value & ""))     ' A列：社員番号
        empName = Trim(CStr(wsSource.Cells(i, "B").value & ""))   ' B列：氏名
        
        ' 社員番号が空欄の行はスキップ
        If empNo = "" Then GoTo NextRow
        
        nightDuty = val(wsSource.Cells(i, "F").value & "")  ' F列：手当2（夜間＋RINK）
        custBill = val(wsSource.Cells(i, "G").value & "")   ' G列：顧客請求分
        transport = val(wsSource.Cells(i, "H").value & "")  ' H列：交通費
        otherExp = val(wsSource.Cells(i, "I").value & "")   ' I列：その他
        telework = val(wsSource.Cells(i, "J").value & "")   ' J列：テレワーク手当
        tatekaeTax = val(wsSource.Cells(i, "X").value & "") ' X列：非課税精算（立替金）※合算済み
        
        '--- 立替金の有無で分岐 ---
        ' X列に値がある → 立替金あり
        hasTA = (tatekaeTax <> 0)
        
        If hasTA Then
            ' 立替金ありの場合
            csvL = 0               ' 非課税通勤費 → 0
            csvN = tatekaeTax      ' 立替金 → X列の値（V+X+Y合算済み）
            csvO = 0               ' その他 → 0
        Else
            ' 立替金なしの場合
            csvL = transport       ' 非課税通勤費 → H列（交通費）
            csvN = 0               ' 立替金 → 0
            csvO = otherExp        ' その他 → I列
        End If
        
        '--- CSV行を組み立て ---
        ' カンマ区切りで値を連結
        ' 文字列項目はダブルクォートで囲む
        csvLine = EscapeCSV(empNo) & "," & _
                  EscapeCSV(empName) & "," & _
                  nightDuty & "," & _
                  "0" & "," & _
                  "0" & "," & _
                  telework & "," & _
                  "" & "," & _
                  "0" & "," & _
                  "0" & "," & _
                  "0" & "," & _
                  "0" & "," & _
                  csvL & "," & _
                  custBill & "," & _
                  csvN & "," & _
                  csvO
        
        '--- 書き込み ---
        Print #fileNum, csvLine
        
NextRow:
    Next i
    
    '--- ファイルを閉じる ---
    Close #fileNum
    
    '--- 完了メッセージ ---
    MsgBox "jinjerインポート用CSVを作成しました！" & vbCrLf & vbCrLf & _
           "保存先: " & csvPath & vbCrLf & _
           "対象: " & (lastRow - 1) & " 件", vbInformation

End Sub

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


