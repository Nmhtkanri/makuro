Attribute VB_Name = "I追加金額マクロ"
Option Explicit

Private Const SH_SUM As String = "集計"
Private Const SH_SETTING As String = "仕訳データ"

Sub Merge_Q_To_XY()
    Dim wsSUM As Worksheet, wsSET As Worksheet
    Dim dicF As Object, dicH As Object
    Dim lastSetR As Long, lastA As Long, lastB As Long, lastE As Long
    Dim sr As Long, lastR As Long, r As Long
    Dim bVal As String, eVal As String, empNo As String
    Dim vQ As Double, vX As Double, vY As Double
    Dim vS As Double, vU As Double, vV As Double, vW As Double

    Set wsSUM = Worksheets(SH_SUM)
    Set wsSET = Worksheets(SH_SETTING)
    Set dicF = CreateObject("Scripting.Dictionary")
    Set dicH = CreateObject("Scripting.Dictionary")

    ' ① 設定シートの最終行を各列で取得
    lastA = wsSET.Cells(wsSET.rows.Count, 1).End(xlUp).Row
    lastB = wsSET.Cells(wsSET.rows.Count, 2).End(xlUp).Row
    lastE = wsSET.Cells(wsSET.rows.Count, 5).End(xlUp).Row
    lastSetR = lastA
    If lastB > lastSetR Then lastSetR = lastB
    If lastE > lastSetR Then lastSetR = lastE

    ' ② 設定シートからB列・E列の社員番号を読み込む
    For sr = 2 To lastSetR
        bVal = Trim$(CStr(wsSET.Cells(sr, 2).value))
        If bVal <> "" Then
            If Not dicF.Exists(bVal) Then dicF.Add bVal, True
        End If

        eVal = Trim$(CStr(wsSET.Cells(sr, 5).value))
        If eVal <> "" Then
            If Not dicH.Exists(eVal) Then dicH.Add eVal, True
        End If
    Next sr

    ' ③ 集計シートを処理
    lastR = wsSUM.Cells(wsSUM.rows.Count, 1).End(xlUp).Row

    For r = 2 To lastR
        If Trim$(CStr(wsSUM.Cells(r, 1).value)) <> "" Then

            empNo = Trim$(CStr(wsSUM.Cells(r, 1).value))
            vQ = ValJP(wsSUM.Cells(r, 17).Value2)

            If dicH.Exists(empNo) Then
    vV = ValJP(wsSUM.Cells(r, 22).Value2)  ' V列
    vW = ValJP(wsSUM.Cells(r, 23).Value2)  ' W列
    vX = ValJP(wsSUM.Cells(r, 24).Value2)  ' X列
    vY = ValJP(wsSUM.Cells(r, 25).Value2)  ' Y列
    wsSUM.Cells(r, 24).value = vV + vW + vX + vY
    wsSUM.Cells(r, 22).ClearContents  ' V列
    wsSUM.Cells(r, 23).ClearContents  ' W列
    wsSUM.Cells(r, 25).ClearContents  ' Y列

            ElseIf dicF.Exists(empNo) Then
                vX = ValJP(wsSUM.Cells(r, 24).Value2)
                vV = ValJP(wsSUM.Cells(r, 22).Value2)
                If vV <> 0 Then
                    wsSUM.Cells(r, 24).value = vX + vV
                    wsSUM.Cells(r, 22).ClearContents
                End If

            ElseIf vQ <> 0 Then
                vX = ValJP(wsSUM.Cells(r, 24).Value2)
                vY = ValJP(wsSUM.Cells(r, 25).Value2)
                wsSUM.Cells(r, 24).value = vX + vY
                wsSUM.Cells(r, 25).value = vQ
            End If

        End If
    Next r

    MsgBox "処理が完了しました。", vbInformation
End Sub

Private Function ValJP(ByVal v As Variant) As Double
    If IsError(v) Or isEmpty(v) Then Exit Function
    Dim s As String: s = CStr(v)
    s = Trim$(s)
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
