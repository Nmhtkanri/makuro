VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const SRC_TW_SHEET As String = "勤務時間帯一覧"
Private Const SRC_GRP_SHEET As String = "打刻グループ"
Private Const DST_SHEET As String = "汎用データひな型"

Private Const COL_A As Long = 1
Private Const COL_B As Long = 2
Private Const COL_C As Long = 3
Private Const COL_D As Long = 4
Private Const COL_U As Long = 21
Private Const COL_V As Long = 22
Private Const COL_W As Long = 23
Private Const COL_AP As Long = 42
Private Const COL_AQ As Long = 43
Private Const COL_CT As Long = 98
Private Const COL_DS As Long = 123  ' DS列（打刻区分ID）


Public Sub Run_TeleworkToTemplate()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Finally_

    Dim wsSrc As Worksheet, wsGrp As Worksheet, wsDst As Worksheet
    Set wsSrc = GetSheetByNameLike(SRC_TW_SHEET)
    Set wsGrp = GetSheetByNameLike(SRC_GRP_SHEET)
    Set wsDst = GetSheetByNameLike(DST_SHEET)
    If wsSrc Is Nothing Or wsGrp Is Nothing Or wsDst Is Nothing Then
        MsgBox "必要シートが見つからないよ。" & vbCrLf & "・" & SRC_TW_SHEET & " / " & SRC_GRP_SHEET & " / " & DST_SHEET, vbExclamation
        GoTo Finally_
    End If

    Dim hdrRowGrp As Long: hdrRowGrp = FindHeaderRow(wsGrp, 1, 10)
    If hdrRowGrp = 0 Then hdrRowGrp = 1

    Dim cgEmpId As Long: cgEmpId = FindCol(wsGrp, hdrRowGrp, Array("社員番号", "従業員番号", "従業員ID", "社員ID"))
    If cgEmpId = 0 Then cgEmpId = 1
    Dim cGid As Long: cGid = 4

    Dim grpDict As Object: Set grpDict = CreateObject("Scripting.Dictionary")
    grpDict.CompareMode = 1
    BuildGroupDict wsGrp, hdrRowGrp, cgEmpId, cGid, grpDict

    ClearTargetColumns wsDst, Array(COL_A, COL_B, COL_C, COL_D, COL_U, COL_V, COL_W, COL_AP, COL_AQ, COL_CT, COL_DS)
    SetGeneralFormat wsDst, Array(COL_A, COL_B, COL_C, COL_D, COL_U, COL_V, COL_W, COL_AP, COL_AQ, COL_CT, COL_DS)

    Dim cName As Long: cName = 2
    Dim cEmpId As Long: cEmpId = 1
    Dim colDate As Long: colDate = 3
    Dim cStart As Long: cStart = 4
    Dim cEnd As Long: cEnd = 5
    Dim cBrkS As Long: cBrkS = 6
    Dim cBrkR As Long: cBrkR = 7
    Dim cHoliday As Long: cHoliday = 10
    Dim cTelework As Long: cTelework = 8  ' H列（テレワーク時間）

    Dim lastSrcRow As Long: lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, cName).End(xlUp).row
    If lastSrcRow < 2 Then GoTo Finally_

    ' ループ内で使う変数はすべてここで宣言
    Dim nm As String
    Dim d As Variant
    Dim dStr As String
    Dim gid As String
    Dim empID As String
    Dim st As String, ed As String, bs As String, br As String
    Dim holTxt As String
    Dim codeU As String
    Dim tw As String
    Dim twHours As Double

    Dim outRow As Long: outRow = 2
    Dim r As Long
    For r = 2 To lastSrcRow
        nm = NzText(wsSrc.Cells(r, cName).value)
        If nm <> "" Then
            d = wsSrc.Cells(r, colDate).value
            dStr = NormalizeDateStr(d)
            If dStr <> "" Then

                empID = NzText(wsSrc.Cells(r, cEmpId).value)
                If grpDict.Exists(NormalizeEmployeeId(empID)) Then
                    gid = NzText(grpDict(NormalizeEmployeeId(empID)))
                Else
                    gid = ""
                End If
                st = NormalizeTimeStr(wsSrc.Cells(r, cStart).value)
                ed = NormalizeTimeStr(wsSrc.Cells(r, cEnd).value)
                bs = NormalizeTimeStr(wsSrc.Cells(r, cBrkS).value)
                br = NormalizeTimeStr(wsSrc.Cells(r, cBrkR).value)

                holTxt = NzText(wsSrc.Cells(r, cHoliday).value)

                wsDst.Cells(outRow, COL_A).value = nm
                wsDst.Cells(outRow, COL_B).value = empID
                wsDst.Cells(outRow, COL_C).value = dStr
                wsDst.Cells(outRow, COL_D).value = gid

                codeU = MapHolidayCode(holTxt)
                If codeU <> "" Then
                    wsDst.Cells(outRow, COL_U).value = Val(codeU)
                Else
                    wsDst.Cells(outRow, COL_U).ClearContents
                End If

                With wsDst.Cells(outRow, COL_V)
                    .NumberFormat = "@"
                    .value = st
                End With
                With wsDst.Cells(outRow, COL_W)
                    .NumberFormat = "@"
                    .value = ed
                End With
                With wsDst.Cells(outRow, COL_AP)
                    .NumberFormat = "@"
                    .value = bs
                End With
                With wsDst.Cells(outRow, COL_AQ)
                    .NumberFormat = "@"
                    .value = br
                End With

                ' CT列：欠勤のときだけ 1
                If holTxt = "欠勤" Then
                    wsDst.Cells(outRow, COL_CT).value = 1
                Else
                    wsDst.Cells(outRow, COL_CT).ClearContents
                End If

                ' DS列：テレワーク時間が4時間以上ならTRUE、それ以外（空欄・4時間未満）はFALSE
                twHours = GetTimeAsHours(wsSrc.Cells(r, cTelework).value)
                If twHours >= 4 Then
                    wsDst.Cells(outRow, COL_DS).value = "TRUE"
                Else
                    wsDst.Cells(outRow, COL_DS).value = "FALSE"
                End If

                outRow = outRow + 1
            End If
        End If
    Next r

Finally_:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub ClearTargetColumns(ByVal ws As Worksheet, ByVal cols As Variant)
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        ws.Range(ws.Cells(2, cols(i)), ws.Cells(ws.Rows.Count, cols(i))).ClearContents
    Next i
End Sub

Private Sub SetGeneralFormat(ByVal ws As Worksheet, ByVal cols As Variant)
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        ws.Columns(cols(i)).NumberFormat = "@"
    Next i
End Sub

Private Function GetSheetByNameLike(ByVal mustContain As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, mustContain, vbTextCompare) = 0 Then
            Set GetSheetByNameLike = ws
            Exit Function
        End If
    Next
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, mustContain, vbTextCompare) > 0 Then
            Set GetSheetByNameLike = ws
            Exit Function
        End If
    Next
End Function

Private Function FindHeaderRow(ByVal ws As Worksheet, ByVal rowStart As Long, ByVal rowEnd As Long) As Long
    Dim r As Long
    For r = rowStart To rowEnd
        If Application.WorksheetFunction.CountA(ws.Rows(r)) > 0 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next
    FindHeaderRow = 0
End Function

Private Function FindCol(ByVal ws As Worksheet, ByVal hdrRow As Long, ByVal candidates As Variant) As Long
    Dim lastCol As Long: lastCol = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, i As Long
    For c = 1 To lastCol
        Dim v As String: v = NzText(ws.Cells(hdrRow, c).value)
        If v <> "" Then
            For i = LBound(candidates) To UBound(candidates)
                If InStr(1, v, CStr(candidates(i)), vbTextCompare) > 0 Then
                    FindCol = c
                    Exit Function
                End If
            Next i
        End If
    Next c
    FindCol = 0
End Function

Private Sub BuildGroupDict(ByVal ws As Worksheet, ByVal hdrRow As Long, ByVal cEmpId As Long, ByVal cGid As Long, ByRef dict As Object)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, cEmpId).End(xlUp).row
    Dim r As Long
    For r = hdrRow + 1 To lastRow
        Dim empID As String: empID = NormalizeEmployeeId(ws.Cells(r, cEmpId).value)
        If empID <> "" Then
            Dim key As String: key = empID
            If Not dict.Exists(key) Then dict.Add key, NzText(ws.Cells(r, cGid).value)
        End If
    Next r
End Sub

Private Function NormalizeEmployeeId(ByVal v As Variant) As String
    Dim t As String: t = NzText(v)
    t = Replace(t, "　", "")
    t = Replace(t, " ", "")
    NormalizeEmployeeId = t
End Function

Private Function NormalizeName(ByVal s As String) As String
    Dim t As String: t = UCase$(s)
    t = Replace(t, "　", "")
    t = Replace(t, " ", "")
    NormalizeName = t
End Function

Private Function NzText(ByVal v As Variant) As String
    If IsError(v) Then
        NzText = ""
    ElseIf IsNull(v) Or IsEmpty(v) Then
        NzText = ""
    Else
        NzText = Trim(CStr(v))
    End If
End Function

Private Function NormalizeDateStr(ByVal v As Variant) As String
    On Error GoTo Fail
    If IsDate(v) Then
        NormalizeDateStr = Format$(CDate(v), "yyyy/m/d")
        Exit Function
    End If
    Dim s As String: s = NzText(v)
    If s <> "" And IsDate(s) Then
        NormalizeDateStr = Format$(CDate(s), "yyyy/m/d")
        Exit Function
    End If
Fail:
    NormalizeDateStr = ""
End Function

Private Function GetTimeAsHours(ByVal v As Variant) As Double
    ' 時間値をDouble（時間単位）に変換。空・無効の場合は -1 を返す
    On Error GoTo Fail
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then GoTo Fail

    Dim s As String: s = NzText(v)
    If s = "" Then GoTo Fail

    ' 数値の場合
    If IsNumeric(v) Then
        Dim dbl As Double: dbl = CDbl(v)
        If dbl < 0 Then GoTo Fail
        If dbl < 1 Then
            ' 時刻シリアル値（例: 4時間 = 4/24 ? 0.1667）
            GetTimeAsHours = dbl * 24
        ElseIf dbl <= 24 Then
            ' 時間数として直接入力（例: 4 → 4時間）
            GetTimeAsHours = dbl
        Else
            ' 日付+時刻のシリアル値（日付部分を除いた時刻部分のみ）
            GetTimeAsHours = (dbl - Int(dbl)) * 24
        End If
        Exit Function
    End If

    ' "H:MM" または "HH:MM" 形式のテキストの場合
    If s Like "*:*" Then
        Dim parts() As String: parts = Split(s, ":")
        If UBound(parts) >= 1 Then
            If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
                GetTimeAsHours = CLng(parts(0)) + CLng(parts(1)) / 60
                Exit Function
            End If
        End If
    End If

    ' 日付として解釈できる場合
    If IsDate(v) Then
        Dim dDate As Double: dDate = CDbl(CDate(v))
        GetTimeAsHours = (dDate - Int(dDate)) * 24
        Exit Function
    End If

Fail:
    GetTimeAsHours = -1  ' 空・無効を示す
End Function

Private Function NormalizeTimeStr(ByVal v As Variant) As String
    On Error GoTo Fail

    Dim s As String: s = NzText(v)

    If s <> "" Then
        If s Like "*:*" Then
            Dim parts() As String
            parts = Split(s, ":")
            If UBound(parts) >= 1 Then
                If IsNumeric(parts(0)) And IsNumeric(parts(1)) Then
                    Dim hrs As Long, mins As Long
                    hrs = CLng(parts(0))
                    mins = CLng(parts(1))
                    If hrs >= 24 Then
                        NormalizeTimeStr = Format$(hrs, "00") & ":" & Format$(mins, "00")
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    If IsNumeric(v) And Not IsEmpty(v) Then
        Dim dblVal As Double
        dblVal = CDbl(v)
        If dblVal >= 1 Then
            Dim totalHours As Long, totalMins As Long
            totalHours = Int(dblVal * 24)
            totalMins = Int((dblVal * 24 - totalHours) * 60)
            NormalizeTimeStr = Format$(totalHours, "00") & ":" & Format$(totalMins, "00")
            Exit Function
        Else
            NormalizeTimeStr = Format$(dblVal, "hh:mm")
            Exit Function
        End If
    End If

    If IsDate(v) Then
        NormalizeTimeStr = Format$(CDbl(v), "hh:mm")
        Exit Function
    End If

    If s = "" Then
        NormalizeTimeStr = ""
        Exit Function
    End If

    If IsDate(s) Then
        NormalizeTimeStr = Format$(CDate(s), "hh:mm")
    Else
        NormalizeTimeStr = s
    End If
    Exit Function

Fail:
    NormalizeTimeStr = ""
End Function

Private Function MapHolidayCode(ByVal s As String) As String
    Dim t As String: t = NormalizeJP(NzText(s))
    If t = "" Then Exit Function

    Dim mp As Object
    Set mp = CreateObject("Scripting.Dictionary")
    mp.CompareMode = 1

    mp("法休") = "0"
    mp("法定休日") = "0"
    mp("法定") = "0"
    mp("所休") = "1"
    mp("所定休日") = "1"
    mp("所定") = "1"
    mp("シフト休") = "1"
    mp("シフト") = "1"
    mp("シフト休暇") = "1"

    If mp.Exists(t) Then
        MapHolidayCode = mp(t)
        Exit Function
    End If

    Dim k As Variant
    For Each k In mp.Keys
        If t Like "*" & k & "*" Then
            MapHolidayCode = mp(k)
            Exit Function
        End If
    Next
End Function

Private Function NormalizeJP(ByVal s As String) As String
    Dim r As String: r = NzText(s)
    r = Trim$(r)
    r = Replace(r, "　", "")
    r = Replace(r, " ", "")
    r = Replace(r, "（", "(")
    r = Replace(r, "）", ")")
    On Error Resume Next
    r = StrConv(r, vbWide)
    r = StrConv(r, vbKatakana)
    On Error GoTo 0
    NormalizeJP = r
End Function

Public Sub ExportTemplateToCSV()
    Dim wsDst As Worksheet
    Set wsDst = GetSheetByNameLike(DST_SHEET)

    If wsDst Is Nothing Then
        MsgBox "「" & DST_SHEET & "」シートが見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim savePath As String
    Dim fileName As String
    fileName = "汎用データひな型_" & Format$(Now, "yyyymmdd_hhmmss") & ".csv"

    If ThisWorkbook.Path <> "" Then
        savePath = ThisWorkbook.Path & "\" & fileName
    Else
        savePath = Environ("USERPROFILE") & "\Desktop\" & fileName
    End If

    Dim finalPath As Variant
    finalPath = Application.GetSaveAsFilename( _
        InitialFileName:=savePath, _
        FileFilter:="CSVファイル (*.csv),*.csv", _
        Title:="CSVファイルの保存先を選択")

    If finalPath = False Then
        MsgBox "保存がキャンセルされました。", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler

    wsDst.Copy

    With ActiveWorkbook
        .SaveAs fileName:=finalPath, FileFormat:=xlCSV, Local:=True
        .Close SaveChanges:=False
    End With

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "CSV保存が完了しました！" & vbCrLf & vbCrLf & finalPath, vbInformation
    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。" & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub Run_TeleworkToTemplate_AndExport()
    Call Run_TeleworkToTemplate
    Call ExportTemplateToCSV
End Sub





