Attribute VB_Name = "Module3"
Public Sub FixDisplay_勤務時間帯一覧()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("勤務時間帯一覧")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    With ws
        .Columns(3).NumberFormat = "yyyy/m/d" ' C: 日付（数値）
        .Columns(4).NumberFormat = "[h]:mm"   ' D: 開始（数値）
        .Columns(5).NumberFormat = "[h]:mm"   ' E: 終了（数値）
        .Columns(6).NumberFormat = "[h]:mm"   ' F: 休憩開始（数値）
        .Columns(7).NumberFormat = "[h]:mm"   ' G: 休憩終了（数値）
        .Columns(8).NumberFormat = "[h]:mm"   ' H: テレワーク時間（数値）
    End With
End Sub

