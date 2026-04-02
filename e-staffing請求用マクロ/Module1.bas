Attribute VB_Name = "Module1"
Option Explicit

' ==========================================================
'  準備・取込用コード群（ActiveSheetで動作）
'  ※それぞれ個別のボタンに登録して使用してください
' ==========================================================

Sub 取込み()
    ' WebTimeCardの勤怠処理ダウンロードデータを取り込む(TXT)
    Dim ws1 As Worksheet
    Dim myFile As Variant
    Dim myFile1 As String

    Set ws1 = ActiveSheet

    myFile = Application.GetOpenFilename("TEXTファイル(*.txt),*.txt")

    If VarType(myFile) = vbBoolean Then
        MsgBox "キャンセルされました"
        Exit Sub
    Else
        MsgBox myFile & " が選択されました"
    End If

    Workbooks.Open myFile
    ws1.Cells.ClearContents
    ActiveWorkbook.ActiveSheet.UsedRange.Copy Destination:=ws1.Range("A1")

    myFile1 = Dir(myFile)
    Workbooks(myFile1).Close SaveChanges:=False
End Sub

Sub 取込み2()
    ' WebTimeCardの勤怠処理ダウンロードデータを取り込む(CSV - 立替金)
    Dim ws1 As Worksheet
    Dim myFile As Variant
    Dim myFile1 As String

    Set ws1 = ActiveSheet

    myFile = Application.GetOpenFilename("CSVファイル(*.csv),*.csv")

    If VarType(myFile) = vbBoolean Then
        MsgBox "キャンセルされました"
        Exit Sub
    End If

    MsgBox Dir(myFile) & " が選択されました"

    Workbooks.Open myFile
    ws1.Cells.ClearContents
    ActiveWorkbook.ActiveSheet.UsedRange.Copy Destination:=ws1.Range("A1")

    myFile1 = Dir(myFile)
    Workbooks(myFile1).Close SaveChanges:=False
End Sub

Sub 取込み3()
    ' e-Staffing契約データを取り込む(CSV - TCnmht)
    Dim ws1 As Worksheet
    Dim myFile As Variant
    Dim myFile1 As String

    Set ws1 = ActiveSheet

    myFile = Application.GetOpenFilename("CSVファイル(*.csv),*.csv")

    If VarType(myFile) = vbBoolean Then
        MsgBox "キャンセルされました"
        Exit Sub
    End If

    MsgBox Dir(myFile) & " が選択されました"

    Workbooks.Open myFile
    ws1.Cells.ClearContents
    ActiveWorkbook.ActiveSheet.UsedRange.Copy Destination:=ws1.Range("A1")

    myFile1 = Dir(myFile)
    Workbooks(myFile1).Close SaveChanges:=False
End Sub
Sub 名簿2作成()
    ' 名簿シートからデータをActiveSheetにコピー
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim MR1 As Long, i As Long
    
    Set ws1 = ActiveSheet
    ' ※「名簿 (2)」というシートが存在する必要があります
    Set ws2 = Worksheets("名簿 (2)")
    
    MR1 = ws1.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To MR1
        Application.StatusBar = "処理中：" & i & "/" & MR1
        ws2.Cells(i, 1).Value = ws1.Cells(i, 2).Value
        ws2.Cells(i, 2).Value = ws1.Cells(i, 1).Value
        ws2.Cells(i, 3).Value = ws1.Cells(i, 3).Value
    Next

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub webTCデータの抽出()
    ' webTC_dataシートからActiveSheetへデータを集計・展開
    ' ※これを実行するとActiveSheetのA列～H列がいったんクリアされ、再構築されます
    
    Dim STRTTIME As Date
    STRTTIME = Now
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim HC(200) As Integer
    Dim n1 As Integer, n2 As Integer, n3 As Integer
    Dim MR1 As Long, MR2 As Long, MR11 As Long
    Dim hc_cnt As Integer, hc_cnt_max As Integer
    Dim i As Long
    Dim dtt As Double, endt As String
    Dim JOB_C As Variant, S_name As Variant
    Dim hantei As Double
    
    Set ws1 = ActiveSheet
    Set ws2 = Worksheets("webTC_data")
    Set ws3 = Worksheets("名簿")
    
    MR1 = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    MR2 = ws2.Cells(Rows.Count, 2).End(xlUp).Row
    
    ' 画面クリア
    ws1.Range("A2:H" & Format(MR1)).Clear
    
    hc_cnt = 1
    For i = 1 To MR2
        If ws2.Cells(i, 1).Value = "H" Then
            HC(hc_cnt) = i
            hc_cnt = hc_cnt + 1
        End If
    Next
    
    hc_cnt_max = hc_cnt
    HC(hc_cnt_max) = MR2
    
    hc_cnt = 1
    
    On Error GoTo errlab1
    
    For i = 1 To hc_cnt_max
        ' 残時間計算
        dtt = Now - STRTTIME
        If i > 0 Then
            endt = Format(dtt * (hc_cnt_max / i - 1), "hh:nn:ss")
        End If
        
        JOB_C = ws2.Cells(HC(i), 5).Value
        S_name = ws2.Cells(HC(i), 6).Value
        
        Application.StatusBar = "データ抽出中：" & S_name & " " & i & "/" & hc_cnt_max & "件　残時間：" & endt
        
        ' 名簿参照用の数式セット
        ws1.Cells(hc_cnt + 1, 1).Value = "=+SUMIF(名簿!A:A,C" & Format(hc_cnt + 1) & ",名簿!B:B)"
        
        ws1.Cells(hc_cnt + 1, 2).Value = JOB_C
        ws1.Cells(hc_cnt + 1, 3).Value = S_name
        
        n1 = HC(hc_cnt) + 1
        n2 = HC(hc_cnt + 1) - 1
        n3 = Int((n1 + n2) / 2 + 1.5)
        
        ws1.Cells(hc_cnt + 1, 4).Value = Application.WorksheetFunction.Sum(ws2.Range("J" & n1 & ":J" & n2)) * 24
        ws1.Cells(hc_cnt + 1, 4).NumberFormatLocal = "0.00"
        
        ws1.Cells(hc_cnt + 1, 5).Value = Application.WorksheetFunction.Sum(ws2.Range("M" & n1 & ":M" & n2)) * 24
        ws1.Cells(hc_cnt + 1, 5).NumberFormatLocal = "0.00"
        
        hantei = Application.WorksheetFunction.Sum(ws2.Range("J" & n3 & ":J" & n2)) * 24
        If hantei = 0 Then
            ws1.Cells(hc_cnt + 1, 8).Value = "未承認"
        End If
        
        ws1.Cells(hc_cnt + 1, 6).Value = Application.WorksheetFunction.Sum(ws2.Range("O" & n1 & ":O" & n2)) * 24
        ws1.Cells(hc_cnt + 1, 6).NumberFormatLocal = "0.00"
        
        ws1.Cells(hc_cnt + 1, 7).Value = Application.WorksheetFunction.Sum(ws2.Range("N" & n1 & ":N" & n2)) * 24
        ws1.Cells(hc_cnt + 1, 7).NumberFormatLocal = "0.00"
        
        hc_cnt = hc_cnt + 1
    Next
    
errlab1:
    ' エラー時または終了時の処理
    If hc_cnt + 1 <= ws1.Rows.Count Then
        ws1.Cells(hc_cnt + 1, 1).Value = ""
    End If
    ws1.Cells(1, 8).Value = ws2.Cells(2, 2).Value
    
    MR11 = ws1.Cells(Rows.Count, 2).End(xlUp).Row
    If MR11 >= 2 Then
        ' 値貼り付けで数式を固定化（元コードの仕様）
        ws1.Range("A2:A" & MR11).Value = ws1.Range("A2:A" & MR11).Value
    End If
    
    Application.StatusBar = "終了"
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub スタッフコード()
    ' e-staffingの最新情報シートと照合してスタッフコード(C列)などを更新
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim MR1 As Long, MR2 As Long
    Dim i As Long, j As Long
    
    Set ws1 = ActiveSheet
    Set ws2 = Worksheets("e-staffing TCnmhtの最新情報")
    
    MR1 = ws1.Cells(Rows.Count, 1).End(xlUp).Row
    MR2 = ws2.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To MR1
        Application.StatusBar = "処理中：" & i & "/" & MR1
        ' 元コードは2重ループで探索（※データ量が多いと遅くなる可能性がありますが、元のロジックを維持）
        For j = 2 To MR2
            If ws1.Cells(i, 1).Value = ws2.Cells(j, 22).Value Then
                ws1.Cells(i, 3).Value = ws2.Cells(j, 21).Value
            Else
                ' 一致しない場合、元コードでは空欄にしています
                ' ※既に名前が入っているセルを消してしまうリスクがあるため注意してください
                ' ws1.Cells(i, 3).Value = ""
            End If
        Next
    Next
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

