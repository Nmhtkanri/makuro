Attribute VB_Name = "dRY列振り分けマクロ"
Option Explicit

' === シート名 ===
Private Const SH_SUM As String = "集計"

' =================================================================
' Distribute_To_RY_Columns
'   集計シートの C～K 列の値を R～Y 列に振り分ける
'
'   F列(手当2)           → R列(夜間当番手当) + S列(金額)
'   H列(交通費)          → V列(金額)
'   G列(顧客請求)        → W列(金額)
'   I列(非課税立替)      → X列(金額)
'   J列(その他)          → Y列(金額)
'   K列(テレワーク手当)  → 廃止（転記しない）
' =================================================================
Public Sub Distribute_To_RY_Columns()
    Dim ws As Worksheet
    Set ws = Worksheets(SH_SUM)
    
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastR
        ' A列が空行ならスキップ
        If Trim$(CStr(ws.Cells(r, 1).Value)) <> "" Then
            Dim vF As Double, vH As Double, vG As Double, vI As Double, vJ As Double
            vF = ValJP(ws.Cells(r, 6).Value2)  ' F: 手当2（顧客+RINK）
            vH = ValJP(ws.Cells(r, 8).Value2)  ' H: 交通費
            vG = ValJP(ws.Cells(r, 7).Value2)  ' G: 顧客請求
            vI = ValJP(ws.Cells(r, 9).Value2)  ' I: 非課税立替
            vJ = ValJP(ws.Cells(r, 10).Value2) ' J: その他(会議費・消耗品など)
            
            ' R/S: 手当2(F列)に値があれば「夜間当番手当」＋金額
            If vF <> 0 Then
                ws.Cells(r, 18).Value = "夜間当番手当"  ' R列
                ws.Cells(r, 19).Value = vF                ' S列
            End If
            
            ' V: 交通費(H列)に値があればその値を記入
            If vH <> 0 Then
                ws.Cells(r, 22).Value = vH                ' V列
            End If
            
            ' W: 顧客請求(G列)に値があればその値を記入
            If vG <> 0 Then
                ws.Cells(r, 23).Value = vG                ' W列
            End If
            
            ' X: 非課税立替(I列)に値があればその値を記入
            If vI <> 0 Then
                ws.Cells(r, 24).Value = vI                ' X列
            End If
            
            ' Y: その他(J列)に値があればその値を記入
            If vJ <> 0 Then
                ws.Cells(r, 25).Value = vJ                ' Y列
            End If
        End If
    Next r
    
    MsgBox "R～Y列への振り分けが完了しました。", vbInformation
End Sub

' ========= ヘルパー =========
' 「\」「円」「,」「全角数字」「()のマイナス」などを考慮して数値化
Private Function ValJP(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Dim s As String: s = CStr(v)
    s = Trim$(s)
    s = StrConv(s, vbNarrow)          ' 全角→半角
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
