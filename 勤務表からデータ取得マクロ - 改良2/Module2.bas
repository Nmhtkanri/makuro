Attribute VB_Name = "Module2"
' startS,endS: 開始/終了の時刻シリアル（D/E列）
' breakMin: 休憩の長さ（分）。F列に一時的に入ってる時刻から算出する想定
' outStartS,outEndS: 休憩開始/終了の時刻シリアル（F/G列に書く値）
' 戻り値: True=割り当て成功, False=入れる場所なし（そのまま空欄に）
' ====== 30分単位スナップ ======
Private Function SnapTo30(ByVal mins As Long) As Long
    ' 15分未満は切り捨て、15分以上は切り上げ → 30分刻みに
    SnapTo30 = 30 * ((mins + 15) \ 30)
End Function

' startS,endS: 開始/終了の時刻シリアル（D/E）
' breakMin   : 休憩の長さ（分）? Fに一時的に入っている値から ToMinutes() で算出済み
' outStartS,outEndS: 休憩開始/終了の時刻シリアル（F/G に書く値）
Private Function AssignBreakAvoidNight( _
    ByVal startS As Double, ByVal endS As Double, ByVal breakMin As Long, _
    ByRef outStartS As Double, ByRef outEndS As Double) As Boolean

    Const DAY As Long = 24 * 60
    Const FORBID_A As Long = 22 * 60   ' 22:00
    Const FORBID_B As Long = 29 * 60   ' 29:00 (=翌5:00)
    Const LUNCH_A  As Long = 12 * 60   ' 12:00
    Const LUNCH_B  As Long = 13 * 60   ' 13:00

'    If breakMin <= 0 Or startS = 0 Or endS = 0 Then Exit Function
    If breakMin <= 0 Or IsEmpty(startS) = True Or IsEmpty(endS) = True Then Exit Function

    ' 勤務区間 [s,e] （日跨ぎ対応）
    Dim s As Long, e As Long
    s = ToMinutes(startS)
    e = ToMinutes(endS)
    If e <= s Then e = e + DAY

    ' ---- 固定ルール：1:00 は昼、1:30 は終了33:00なら 31:00?32:30 ----
    If breakMin = 60 Then
        ' 昼(12:00?13:00)が [s,e] に収まるなら最優先
        If (LUNCH_A >= s) And ((LUNCH_A + 60) <= e) Then
            outStartS = FromMinutes(LUNCH_A)   ' ★Modなし
            outEndS = FromMinutes(LUNCH_A + 60)
            AssignBreakAvoidNight = True
            Exit Function
        End If
    ElseIf breakMin = 90 Then
        ' 終了が 33:00 なら 31:00?32:30 を優先
        If (e Mod DAY) = (33 * 60) Then
            If (31 * 60) >= s And (32 * 60 + 30) <= e Then
                outStartS = FromMinutes(31 * 60)
                outEndS = FromMinutes(31 * 60 + 90)
                AssignBreakAvoidNight = True
                Exit Function
            End If
        End If
    End If

    ' ---- 昼に部分的にでも入れられるか（60未満は12:00からの長さで）----
    If breakMin < 60 Then
        Dim need As Long: need = breakMin
        Dim c1 As Long, c2 As Long
        c1 = Application.Max(LUNCH_A, s)
        c2 = Application.Min(LUNCH_A + need, e)
        If c2 - c1 >= need Then
            outStartS = FromMinutes(LUNCH_A)
            outEndS = FromMinutes(LUNCH_A + need)
            AssignBreakAvoidNight = True
            Exit Function
        End If
    End If

    ' ---- 夜間禁止帯を除いた許可区間（最大2区間）を作る ----
    Dim segA1 As Long, segA2 As Long, segB1 As Long, segB2 As Long
    segA1 = s: segA2 = e: segB1 = 0: segB2 = 0

    If Not (e <= FORBID_A Or s >= FORBID_B) Then
        ' 左側（禁止帯より前）
        If s < FORBID_A Then
            segA1 = s
            segA2 = Application.Min(e, FORBID_A)
        Else
            segA1 = 0: segA2 = 0
        End If
        ' 右側（禁止帯より後）
        If e > FORBID_B Then
            segB1 = Application.Max(s, FORBID_B)
            segB2 = e
        End If
        If (segA2 <= segA1) And (segB2 <= segB1) Then Exit Function
    End If

    ' ---- 一番長い許可区間の中央に配置し、:00/:30 にスナップ ----
    Dim best1 As Long, best2 As Long
    best1 = 0: best2 = 0
    If (segA2 - segA1) > (best2 - best1) Then best1 = segA1: best2 = segA2
    If (segB2 - segB1) > (best2 - best1) Then best1 = segB1: best2 = segB2

    If (best2 - best1) >= breakMin Then
        Dim tryStart As Long
        tryStart = best1 + ((best2 - best1 - breakMin) \ 2)
        tryStart = SnapTo30(tryStart)   ' 30分刻みに揃える
        ' はみ出し補正（右端を超えないように）
        If tryStart + breakMin > best2 Then tryStart = best2 - breakMin
        ' 最終書き出し（★Mod なし）
        outStartS = FromMinutes(tryStart)
        outEndS = FromMinutes(tryStart + breakMin)
        AssignBreakAvoidNight = True
        Exit Function
    End If

    AssignBreakAvoidNight = False
End Function


Public Sub AllocateBreaks_AvoidNight()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("勤務時間帯一覧")

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastR < 2 Then Exit Sub

    Dim r As Long
    For r = 2 To lastR
        Dim sS As Double, eS As Double, lenS As Double
        sS = ws.Cells(r, 4).value   ' 開始(D)
        eS = ws.Cells(r, 5).value   ' 終了(E)
        lenS = ws.Cells(r, 6).value ' 休憩の長さ（今は一時的にFに入ってる）

        If IsEmpty(sS) = True Or IsEmpty(eS) = True Or lenS = 0 Then
            ' いずれか欠けてたら消して次へ
            ws.Cells(r, 6).value = "": ws.Cells(r, 7).value = ""
        Else
            Dim brMin As Long
            brMin = ToMinutes(lenS)

            Dim outS As Double, outE As Double
            If AssignBreakAvoidNight(sS, eS, brMin, outS, outE) Then
                ws.Cells(r, 6).value = outS ' 休憩開始
                ws.Cells(r, 7).value = outE ' 休憩終了
            Else
                ' 入れられない場合はいったん空欄のまま（後で手当て）
                ws.Cells(r, 6).value = ""
                ws.Cells(r, 7).value = ""
            End If
        End If
    Next r
End Sub

' ====== 時刻→分 ヘルパ ======
Private Function FromMinutes(ByVal mins As Long) As Double
    ' 分 → Excel時刻シリアル(1日=1)
    FromMinutes = mins / (24# * 60#)
End Function

Private Function ToMinutes(ByVal serialTime As Double) As Long
    ' Excel時刻シリアル → 分
    ' 空/非数でも0を返す安全版にしとく
    If IsError(serialTime) Or IsEmpty(serialTime) Then
        ToMinutes = 0
    Else
        ToMinutes = CLng(Round(CDbl(serialTime) * 24# * 60#))
    End If
End Function

Private Function Span(ByVal a As Long, ByVal b As Long) As Long
    ' 区間長（b>aなら差、そうでなければ0）
    If b > a Then Span = b - a Else Span = 0
End Function

