Attribute VB_Name = "estaffing貼り付けマクロ"
Sub ファイルからデータ貼り付け()
    Dim fileDialog As fileDialog
    Dim selectedFile As String
    Dim sourceWb As Workbook
    Dim targetWs As Worksheet
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' エクスプローラーでファイルを選択
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "コピー元のファイルを選択してください"
        .filters.Clear
        .filters.Add "Excel/CSVファイル", "*.xlsx; *.xls; *.xlsm; *.csv"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        Else
            MsgBox "ファイルが選択されませんでした。", vbExclamation
            Exit Sub
        End If
    End With
    
    ' 貼り付け先のシートを設定
    On Error Resume Next
    Set targetWs = ThisWorkbook.Worksheets("e-staffing_出力")
    On Error GoTo 0
    
    If targetWs Is Nothing Then
        MsgBox "「e-staffing_出力シート」が見つかりません。", vbCritical
        Exit Sub
    End If
    
    ' 選択したファイルを開く
    Application.ScreenUpdating = False
    Set sourceWb = Workbooks.Open(selectedFile)
    Set sourceWs = sourceWb.Worksheets(1) ' 最初のシートを対象
    
    ' コピー元のデータ範囲を取得（A2から最終行・最終列まで）
    lastRow = sourceWs.Cells(sourceWs.rows.Count, 1).End(xlUp).Row
    lastCol = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column
    
    If lastRow >= 2 Then
        ' A2から最終行・最終列までをコピー
        sourceWs.Range(sourceWs.Cells(2, 1), sourceWs.Cells(lastRow, lastCol)).Copy
        
        ' 貼り付け先のA1に貼り付け
        targetWs.Range("A1").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        
        MsgBox "データを貼り付けました。" & vbCrLf & _
               "行数: " & (lastRow - 1) & " 列数: " & lastCol, vbInformation
    Else
        MsgBox "コピーするデータがありません（A2以降にデータがない）。", vbExclamation
    End If
    
    ' ファイルを閉じる
    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
End Sub

