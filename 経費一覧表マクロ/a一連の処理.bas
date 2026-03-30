Attribute VB_Name = "a一連の処理"
' === ワンボタンで全部（立替 → eスタッフ → 本社経費 → 社員付番 → 重複削除） ===
Public Sub Append_全部_一括処理()
    
    Dim startTime As Double
    startTime = Timer
    
    ' 1. 立替精算一覧 取り込み
    Append_立替精算一覧_to_経費統合一覧表
    
    ' 2. e-staffing 取り込み
    Append_e_staffing_出力_to_経費統合一覧表
    
    ' 3. 本社経費 取り込み（Freeeデータ）
    Append_本社経費_to_経費統合一覧表
    
    ' 4. 社員番号の付与（マスタとの照合）
    AssignEmployeeNo_ByName_集計toJinjer False
    
    ' 5. 重複削除
    RemoveDuplicates_A_D_F_AndLog
    
    Dim elapsed As Double
    elapsed = Timer - startTime
    
    MsgBox "全ての処理が完了しました！" & vbCrLf & vbCrLf & _
           "処理時間: " & Format(elapsed, "0.0") & " 秒", vbInformation

End Sub

