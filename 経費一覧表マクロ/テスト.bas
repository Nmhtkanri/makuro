Attribute VB_Name = "テスト"
Sub CheckProblemRows()
    Dim ws As Worksheet
    Set ws = Worksheets("経費統合一覧表")
    
    ' マッチしなかった行番号
    Dim rows As Variant
    rows = Array(1996, 2013, 2047, 2120, 2131, 2141, 2174)
    
    Debug.Print "=== 問題行のA列・B列確認 ==="
    
    Dim i As Long, r As Long
    For i = LBound(rows) To UBound(rows)
        r = rows(i)
        Debug.Print "行" & r & ":"
        Debug.Print "  A列（社員番号）= [" & ws.Cells(r, 1).value & "] 長さ=" & Len(CStr(ws.Cells(r, 1).value))
        Debug.Print "  B列（名前）= [" & ws.Cells(r, 2).value & "]"
    Next i
End Sub
