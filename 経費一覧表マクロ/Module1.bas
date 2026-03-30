Attribute VB_Name = "Module1"
Public wb10 As Workbook
Public ws11 As Worksheet

Public w交通費 As Single
Public w経費1 As Single
Public w経費2 As Single
Public w経費3 As Single
Public w経費4 As Single
Public w経費5 As Single
Sub s申請者(w申請者 As Variant)
    
    
    Dim Fcell, FRCel As Range
    Dim wr As Long
    
    w交通費 = 0
    w経費1 = 0
    w経費2 = 0
    w経費3 = 0
    w経費4 = 0
    w経費5 = 0
    
    Set Fcell = ws04.Range("A:A").Find(w申請者)
    If Not Fcell Is Nothing Then
        '1回目
        Set FRCel = Fcell
        
        wr = Fcell.Row
        
        Call sデータセット(wr)
        '2回目以降
        Do
            Set Fcell = ws04.Range("A:A").FindNext(Fcell)
            '一回りしたら終了
            If Fcell.Address = FRCel.Address Then Exit Do
            wr = Fcell.Row
            Call sデータセット(wr)
         Loop
    End If
    
    
End Sub
Sub sデータセット(wr As Long)
    
    Select Case True
        Case ws04.Cells(wr, 7) = "電車・バス"
            w交通費 = w交通費 + ws04.Cells(wr, 9)
        Case ws04.Cells(wr, 7) = "タクシー"
            w交通費 = w交通費 + ws04.Cells(wr, 9)
        Case InStr(ws04.Cells(wr, 8), "RINK日当") > 0
            w経費1 = w経費1 + ws04.Cells(wr, 12)
        Case InStr(ws04.Cells(wr, 8), "顧客対応当番手当") > 0
            w経費1 = w経費1 + ws04.Cells(wr, 12)
        Case InStr(ws04.Cells(wr, 8), "テレワーク手当") > 0
            w経費2 = w経費2 + ws04.Cells(wr, 12)
        Case ws04.Cells(wr, 8) = "その他経費"
            w経費5 = w経費5 + ws04.Cells(wr, 12)

       Case Else
    End Select
End Sub

