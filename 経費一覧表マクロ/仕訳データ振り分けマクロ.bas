Attribute VB_Name = "仕訳データ振り分けマクロ"
Option Explicit

' ============================================================
' 仕訳データ手当振り分け_RS_TU
' ------------------------------------------------------------
' 仕訳データシートの通信手当(H,I列)・定常外業務対応手当(K,L列)を
' 集計シートのR/S(内訳1/金額1), T/U(内訳2/金額2)に振り分ける。
' ルール:
'   - 社員番号で突合(Trim+CStr, 先頭ゼロ対応)
'   - 同一社員の複数行は加算
'   - R空→R=手当名,S=金額 / R同名→S加算 / R別名→T側判定
'   - T空→T=手当名,U=金額 / T同名→U加算 / 両方別名→スキップ+ログ
'   - V/W列は絶対に触らない
' ============================================================
Public Sub 仕訳データ手当振り分け_RS_TU()

    Dim calcMode As Long
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual

    ' --- シート取得 ---
    Dim wsShiwake As Worksheet, wsSyukei As Worksheet
    On Error Resume Next
    Set wsShiwake = ThisWorkbook.Worksheets("仕訳データ")
    Set wsSyukei = ThisWorkbook.Worksheets("集計")
    On Error GoTo ErrHandler

    If wsShiwake Is Nothing Or wsSyukei Is Nothing Then
        MsgBox "必須シートが見つかりません" & vbCrLf & _
               "(仕訳データ / 集計)", vbCritical
        GoTo Cleanup
    End If

    ' --- 集計シート 社員番号→行番号 辞書 ---
    Dim empDict As Object
    Set empDict = CreateObject("Scripting.Dictionary")
    Dim lastRowS As Long, r As Long
    lastRowS = wsSyukei.Cells(wsSyukei.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRowS
        Dim empKey As String
        empKey = NormId(wsSyukei.Cells(r, 1).Value)
        If empKey <> "" Then
            If Not empDict.Exists(empKey) Then empDict.Add empKey, r
        End If
    Next r

    ' --- 仕訳データ 社員単位で手当金額を集計 ---
    Dim dictTsushin As Object:  Set dictTsushin = CreateObject("Scripting.Dictionary")
    Dim dictTeijo As Object:    Set dictTeijo = CreateObject("Scripting.Dictionary")
    Dim lastRowJ As Long

    ' 通信手当: I列(9)基準
    lastRowJ = wsShiwake.Cells(wsShiwake.Rows.Count, 9).End(xlUp).Row
    For r = 2 To lastRowJ
        Dim idT As String, amtT As Double
        idT = NormId(wsShiwake.Cells(r, 9).Value)
        amtT = Val(CStr(wsShiwake.Cells(r, 8).Value))
        If idT <> "" And amtT <> 0 Then
            If dictTsushin.Exists(idT) Then
                dictTsushin(idT) = dictTsushin(idT) + amtT
            Else
                dictTsushin.Add idT, amtT
            End If
        End If
    Next r

    ' 定常外業務対応手当: L列(12)基準
    lastRowJ = wsShiwake.Cells(wsShiwake.Rows.Count, 12).End(xlUp).Row
    For r = 2 To lastRowJ
        Dim idD As String, amtD As Double
        idD = NormId(wsShiwake.Cells(r, 12).Value)
        amtD = Val(CStr(wsShiwake.Cells(r, 11).Value))
        If idD <> "" And amtD <> 0 Then
            If dictTeijo.Exists(idD) Then
                dictTeijo(idD) = dictTeijo(idD) + amtD
            Else
                dictTeijo.Add idD, amtD
            End If
        End If
    Next r

    ' --- ログシート準備 ---
    Dim wsLog As Worksheet
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets("仕訳データ振り分けログ")
    On Error GoTo ErrHandler
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsLog.Name = "仕訳データ振り分けログ"
    Else
        wsLog.Cells.Clear
    End If
    wsLog.Cells(1, 1).Value = "社員番号"
    wsLog.Cells(1, 2).Value = "手当名"
    wsLog.Cells(1, 3).Value = "金額"
    wsLog.Cells(1, 4).Value = "対象行"
    wsLog.Cells(1, 5).Value = "結果"
    Dim logRow As Long: logRow = 2

    ' --- 振り分け処理 ---
    Dim hitCount As Long: hitCount = 0

    ' 1) 通信手当を先に処理
    Dim kTsushin As Variant
    For Each kTsushin In dictTsushin.Keys
        Dim empNoT As String: empNoT = CStr(kTsushin)
        Dim valT As Double: valT = dictTsushin(kTsushin)
        If empDict.Exists(empNoT) Then
            Dim tgtRowT As Long: tgtRowT = empDict(empNoT)
            Dim resT As String
            resT = WriteToRS_TU(wsSyukei, tgtRowT, "通信手当", valT)
            hitCount = hitCount + 1
            wsLog.Cells(logRow, 1).Value = empNoT
            wsLog.Cells(logRow, 2).Value = "通信手当"
            wsLog.Cells(logRow, 3).Value = valT
            wsLog.Cells(logRow, 4).Value = tgtRowT
            wsLog.Cells(logRow, 5).Value = resT
            logRow = logRow + 1
        Else
            wsLog.Cells(logRow, 1).Value = empNoT
            wsLog.Cells(logRow, 2).Value = "通信手当"
            wsLog.Cells(logRow, 3).Value = valT
            wsLog.Cells(logRow, 4).Value = ""
            wsLog.Cells(logRow, 5).Value = "突合不可"
            logRow = logRow + 1
        End If
    Next kTsushin

    ' 2) 定常外業務対応手当を後に処理
    Dim kTeijo As Variant
    For Each kTeijo In dictTeijo.Keys
        Dim empNoD As String: empNoD = CStr(kTeijo)
        Dim valD As Double: valD = dictTeijo(kTeijo)
        If empDict.Exists(empNoD) Then
            Dim tgtRowD As Long: tgtRowD = empDict(empNoD)
            Dim resD As String
            resD = WriteToRS_TU(wsSyukei, tgtRowD, "定常外業務対応手当", valD)
            hitCount = hitCount + 1
            wsLog.Cells(logRow, 1).Value = empNoD
            wsLog.Cells(logRow, 2).Value = "定常外業務対応手当"
            wsLog.Cells(logRow, 3).Value = valD
            wsLog.Cells(logRow, 4).Value = tgtRowD
            wsLog.Cells(logRow, 5).Value = resD
            logRow = logRow + 1
        Else
            wsLog.Cells(logRow, 1).Value = empNoD
            wsLog.Cells(logRow, 2).Value = "定常外業務対応手当"
            wsLog.Cells(logRow, 3).Value = valD
            wsLog.Cells(logRow, 4).Value = ""
            wsLog.Cells(logRow, 5).Value = "突合不可"
            logRow = logRow + 1
        End If
    Next kTeijo

    MsgBox "処理完了: " & hitCount & "件の手当を振り分けました" & vbCrLf & _
           "(ログ: 仕訳データ振り分けログ シート)", vbInformation

Cleanup:
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ============================================================
' WriteToRS_TU - R/S or T/U に手当名+金額を書き込む
' 戻り値: "R書き込み" / "R加算" / "T書き込み" / "T加算" / "スキップ(両スロット使用済み)"
' ============================================================
Private Function WriteToRS_TU(ws As Worksheet, ByVal rowNum As Long, _
                               ByVal teate As String, ByVal amt As Double) As String
    Const COL_R As Long = 18  ' R列 = 内訳1
    Const COL_S As Long = 19  ' S列 = 内訳金額1
    Const COL_T As Long = 20  ' T列 = 内訳2
    Const COL_U As Long = 21  ' U列 = 内訳金額2

    Dim valR As String: valR = Trim$(CStr(ws.Cells(rowNum, COL_R).Value))
    Dim valTSlot As String: valTSlot = Trim$(CStr(ws.Cells(rowNum, COL_T).Value))

    ' --- R/S 判定 ---
    If valR = "" Or valR = "0" Then
        ws.Cells(rowNum, COL_R).Value = teate
        ws.Cells(rowNum, COL_S).Value = amt
        WriteToRS_TU = "R書き込み"
        Exit Function
    End If
    If valR = teate Then
        ws.Cells(rowNum, COL_S).Value = Val(CStr(ws.Cells(rowNum, COL_S).Value)) + amt
        WriteToRS_TU = "R加算"
        Exit Function
    End If

    ' --- T/U 判定 ---
    If valTSlot = "" Or valTSlot = "0" Then
        ws.Cells(rowNum, COL_T).Value = teate
        ws.Cells(rowNum, COL_U).Value = amt
        WriteToRS_TU = "T書き込み"
        Exit Function
    End If
    If valTSlot = teate Then
        ws.Cells(rowNum, COL_U).Value = Val(CStr(ws.Cells(rowNum, COL_U).Value)) + amt
        WriteToRS_TU = "T加算"
        Exit Function
    End If

    ' --- 両方とも別の手当名で埋まっている → スキップ ---
    WriteToRS_TU = "スキップ(両スロット使用済み)"
End Function

' ============================================================
' NormId - 社員番号を正規化 (Trim + 数値の場合は整数文字列化)
' ============================================================
Private Function NormId(ByVal v As Variant) As String
    Dim s As String: s = Trim$(CStr(v))
    If s = "" Or s = "0" Then NormId = "": Exit Function
    If IsNumeric(s) Then
        NormId = CStr(CLng(Val(s)))
    Else
        NormId = s
    End If
End Function
