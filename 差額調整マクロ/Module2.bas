Attribute VB_Name = "Module2"
'------------------------------------------------------------------------------
' Step2: jinjer_給与支給控除項目一覧表
'------------------------------------------------------------------------------
Sub jinjer_給与支給控除項目一覧表()
    Dim filePath As String
    Dim wsTarget As Worksheet
    Dim wb As Workbook
    Dim wbCSV As Workbook
    Dim lastRow As Long
    Dim lastCol As Long
    
    filePath = Application.GetOpenFilename( _
        FileFilter:="CSVファイル (*.csv),*.csv", _
        Title:="給与明細CSVファイルを選択してください")
    
    If filePath = "False" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set wsTarget = wb.Sheets("jinjer後に列削除")
    On Error GoTo 0
    
    ' シートが存在しない場合はエラーメッセージを表示
    If wsTarget Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "「jinjer後に列削除」シートが見つかりません。" & vbCrLf & _
               "シートを作成してから再度実行してください。", vbExclamation
        Exit Sub
    End If
    
    ' シートの既存データをクリア
    wsTarget.Cells.Clear
    
    Set wbCSV = Workbooks.Open(filePath, Local:=True)
    
    With wbCSV.Sheets(1)
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Copy wsTarget.Range("A1")
    End With
    
    wbCSV.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    
    MsgBox "給与明細CSVの取り込みが完了しました。" & vbCrLf & _
           "取り込み先: jinjer後に列削除" & vbCrLf & _
           "取り込み件数: " & (lastRow - 1) & " 件", vbInformation
End Sub

