Attribute VB_Name = "hSAP経費取り込みマクロ"
Option Explicit

' ============================================================
'  SAP_経費 シート → 経費統合一覧表 への取込（34列マッピング）
' ============================================================
'  - 入力: SAP_経費 シート（13列、Paste_SAP経費_From_File で投入）
'  - 出力: 経費統合一覧表 シート（34列）
'  - 社員番号紐付け: 集計シート(A=社員番号, B=氏名)、姓+名 / 名+姓 両方で照合
'  - 重複削除: 行わない（ユーザー要望）
' ============================================================

Private Const SH_SRC As String = "SAP_経費"
Private Const SH_DST As String = "経費統合一覧表"
Private Const SH_MAP As String = "集計"

' SAP_経費 列インデックス（固定）
Private Const C_SEI As Long = 1            ' A: 姓
Private Const C_MEI As Long = 2            ' B: 名
Private Const C_AMT As Long = 3            ' C: 費用合計
Private Const C_VENDOR As Long = 4         ' D: 業者名
Private Const C_DATE As Long = 5           ' E: 費用エントリ日
Private Const C_DESC As Long = 6           ' F: 説明
Private Const C_APPR As Long = 7           ' G: 費用シート承認日
Private Const C_BU As Long = 8             ' H: 事業単位
Private Const C_CC As Long = 9             ' I: コストセンター
Private Const C_CUR As Long = 10           ' J: 通貨
Private Const C_ID As Long = 11            ' K: 費用シート ID
Private Const C_LOC As Long = 12           ' L: 勤務地
Private Const C_STATUS As Long = 13        ' M: 費用シートのステータス

Public Sub Append_SAP経費_to_経費統合一覧表()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsSrc As Worksheet, wsDst As Worksheet, wsMap As Worksheet

    On Error Resume Next
    Set wsSrc = wb.Worksheets(SH_SRC)
    Set wsDst = wb.Worksheets(SH_DST)
    Set wsMap = wb.Worksheets(SH_MAP)
    On Error GoTo ErrHandler

    If wsSrc Is Nothing Then
        MsgBox "「" & SH_SRC & "」シートが見つかりません。", vbExclamation
        GoTo FinallyExit
    End If
    If wsDst Is Nothing Then
        MsgBox "「" & SH_DST & "」シートが見つかりません。", vbExclamation
        GoTo FinallyExit
    End If

    ' 1. SAP_経費 の最終行（A列=姓 で判定）
    Dim srcLastRow As Long
    srcLastRow = wsSrc.Cells(wsSrc.rows.Count, C_SEI).End(xlUp).Row
    If srcLastRow < 2 Then
        MsgBox "「" & SH_SRC & "」シートにデータがありません。", vbInformation
        GoTo FinallyExit
    End If

    Dim rowsCount As Long
    rowsCount = srcLastRow - 1

    ' 2. 集計シートから 名前→社員番号 マップ作成（姓+名 / 名+姓 両対応）
    Dim dictEmp As Object: Set dictEmp = CreateObject("Scripting.Dictionary")
    dictEmp.CompareMode = 1
    BuildEmployeeNameMap wsMap, dictEmp

    ' 3. 経費統合一覧表 の追記開始行（A列とB列の最大）
    Dim dstStartRow As Long
    dstStartRow = Application.WorksheetFunction.Max( _
                    wsDst.Cells(wsDst.rows.Count, 1).End(xlUp).Row, _
                    wsDst.Cells(wsDst.rows.Count, 2).End(xlUp).Row) + 1
    If dstStartRow < 2 Then dstStartRow = 2

    ' 4. SAP データを一括読込
    Dim v As Variant
    v = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(srcLastRow, 13)).value

    ' 5. 34列の出力配列を構築
    Dim out() As Variant
    ReDim out(1 To rowsCount, 1 To 34)

    Dim i As Long
    Dim sei As String, mei As String, fullName As String
    Dim empNo As String, dt As String, apprDt As String
    Dim amt As Variant, vendorTxt As String, descTxt As String, ccTxt As String, sapId As String
    Dim memoTxt As String

    For i = 1 To rowsCount
        sei = SafeStr(v(i, C_SEI))
        mei = SafeStr(v(i, C_MEI))
        fullName = Trim$(sei & " " & mei)

        ' 社員番号: 姓+名 → 名+姓 の順で照合
        empNo = LookupEmpNo(dictEmp, sei, mei)

        ' 日付正規化
        dt = NormalizeDateStr(v(i, C_DATE))
        apprDt = NormalizeDateStr(v(i, C_APPR))

        ' 金額
        amt = NormalizeAmount(v(i, C_AMT))

        ' テキスト系
        vendorTxt = SafeStr(v(i, C_VENDOR))
        descTxt = SafeStr(v(i, C_DESC))
        ccTxt = SafeStr(v(i, C_CC))
        sapId = SafeStr(v(i, C_ID))

        ' 備考(T列): 説明 / コストセンター / 費用シートID
        memoTxt = descTxt
        If ccTxt <> "" Then memoTxt = memoTxt & " / CC: " & ccTxt
        If sapId <> "" Then memoTxt = memoTxt & " / ID: " & sapId

        ' === 34列マッピング ===
        out(i, 1) = empNo                ' A: 社員番号
        out(i, 2) = fullName             ' B: 氏名
        out(i, 3) = apprDt               ' C: 申請日 ← 承認日
        ' E(5): 内容（空）
        out(i, 6) = dt                   ' F: 利用日 ← 費用エントリ日
        ' G(7): 交通機関（空）
        out(i, 8) = TruncateText(vendorTxt, 80) ' H: 内訳 ← 業者名（80文字制限）
        ' I-M (9-13): 空
        ' O(15): 空
        ' Q-S (17-19): 空
        out(i, 20) = memoTxt             ' T: 備考
        ' U-AG (21-33): 空
        If IsNightDutyVendor(vendorTxt) Then
            out(i, 4) = amt              ' D: 合計（夜間当番手当として集計）
        Else
            out(i, 34) = amt             ' AH: 顧客請求額（SAP通常行）
        End If
    Next i

    ' 6. 経費統合一覧表 へ書き込み（文字列扱いで貼り付け）
    With wsDst
        Dim dstRange As Range
        Set dstRange = .Range(.Cells(dstStartRow, 1), .Cells(dstStartRow + rowsCount - 1, 34))
        dstRange.NumberFormat = "@"
        dstRange.value = out

        ' 日付列(C, F)だけ書式を強制（Excel の自動シリアル変換を抑止）
        .Range(.Cells(dstStartRow, 3), .Cells(dstStartRow + rowsCount - 1, 3)).NumberFormat = "yyyy/mm/dd"
        .Range(.Cells(dstStartRow, 6), .Cells(dstStartRow + rowsCount - 1, 6)).NumberFormat = "yyyy/mm/dd"
    End With

    MsgBox "SAP 経費の追記が完了しました。" & vbCrLf & _
           "件数: " & rowsCount & " 行" & vbCrLf & _
           "開始行: " & dstStartRow & " 行目", vbInformation

FinallyExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "SAP 経費取込エラー: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume FinallyExit
End Sub

' ============================================================
'  ヘルパー: 業者名が夜間当番系か判定
'  - 設定シートの夜間当番手当キーワードと同じ考え方で、SAP取込時点の二重計上を防ぐ
' ============================================================
Private Function IsNightDutyVendor(ByVal vendorTxt As String) As Boolean
    Dim keys As Variant
    keys = Array("夜間当番", "24時間準直当番", "準直当番", "深夜出動", _
                 "顧客当番", "顧客対応当番", "オンコール")

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        If InStr(1, vendorTxt, CStr(keys(i)), vbTextCompare) > 0 Then
            IsNightDutyVendor = True
            Exit Function
        End If
    Next i
End Function

' ============================================================
'  ヘルパー: 集計シートから 氏名→社員番号 マップを作る
'  - 姓+名（スペースなし） と 名+姓（スペースなし） の両方をキー登録
' ============================================================
Private Sub BuildEmployeeNameMap(ByVal wsMap As Worksheet, ByVal dict As Object)
    If wsMap Is Nothing Then Exit Sub

    Dim mapLastA As Long, mapLastB As Long, mapLast As Long
    mapLastA = wsMap.Cells(wsMap.rows.Count, 1).End(xlUp).Row
    mapLastB = wsMap.Cells(wsMap.rows.Count, 2).End(xlUp).Row
    mapLast = Application.WorksheetFunction.Max(mapLastA, mapLastB)
    If mapLast < 2 Then Exit Sub

    Dim m As Variant
    m = wsMap.Range(wsMap.Cells(2, 1), wsMap.Cells(mapLast, 2)).value

    Dim i As Long, empNo As String, nm As String, key As String
    For i = 1 To UBound(m, 1)
        empNo = SafeStr(m(i, 1))
        nm = SafeStr(m(i, 2))
        If empNo = "" Or nm = "" Then GoTo NextOne

        ' 元のフルネーム（スペース除去）
        key = NormKey(nm)
        If key <> "" Then
            If Not dict.Exists(key) Then dict.Add key, empNo
        End If

        ' 集計シート側の氏名は通常「姓 名」想定。スペースで分割して名+姓 の逆順キーも登録しておく
        Dim parts() As String
        Dim swapKey As String
        Dim normalized As String
        normalized = Replace(nm, ChrW(&H3000), " ")
        normalized = Trim$(normalized)
        If InStr(normalized, " ") > 0 Then
            parts = Split(normalized, " ")
            If UBound(parts) >= 1 Then
                swapKey = NormKey(parts(UBound(parts)) & parts(0))
                If swapKey <> "" And swapKey <> key Then
                    If Not dict.Exists(swapKey) Then dict.Add swapKey, empNo
                End If
            End If
        End If
NextOne:
    Next i
End Sub

' ============================================================
'  ヘルパー: 姓と名から社員番号を引く
'  - 姓+名 → 名+姓 の順で試す（イレギュラー: SAP 側で姓名逆登録に対応）
' ============================================================
Private Function LookupEmpNo(ByVal dict As Object, _
                              ByVal sei As String, _
                              ByVal mei As String) As String
    Dim k1 As String, k2 As String
    k1 = NormKey(sei & mei)
    If k1 <> "" Then
        If dict.Exists(k1) Then
            LookupEmpNo = CStr(dict(k1))
            Exit Function
        End If
    End If
    k2 = NormKey(mei & sei)
    If k2 <> "" Then
        If dict.Exists(k2) Then
            LookupEmpNo = CStr(dict(k2))
            Exit Function
        End If
    End If
    LookupEmpNo = ""
End Function

' ============================================================
'  共通ヘルパー
' ============================================================
Private Function NormKey(ByVal s As String) As String
    Dim t As String
    t = CStr(s)
    t = Replace(t, ChrW(&H3000), " ")  ' 全角スペース → 半角
    t = Trim$(t)
    t = Replace(t, " ", "")             ' すべての空白除去
    NormKey = t
End Function

Private Function SafeStr(v As Variant) As String
    If IsError(v) Then
        SafeStr = ""
    ElseIf IsNull(v) Or isEmpty(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
End Function

Private Function NormalizeDateStr(v As Variant) As String
    ' SAP の日付は "2026/4/22" や "2026/4/30 12:17" 形式で来る
    ' yyyy/mm/dd に正規化して返す（時刻は捨てる）
    If IsDate(v) Then
        NormalizeDateStr = Format$(CDate(v), "yyyy/mm/dd")
        Exit Function
    End If

    Dim s As String: s = SafeStr(v)
    If s = "" Then
        NormalizeDateStr = ""
        Exit Function
    End If

    ' 時刻部分を切り捨てる
    Dim spIdx As Long: spIdx = InStr(s, " ")
    If spIdx > 0 Then s = Left$(s, spIdx - 1)

    ' "2026/4/22" → "2026/04/22"
    Dim parts() As String
    parts = Split(s, "/")
    If UBound(parts) = 2 Then
        Dim y As String, mo As String, d As String
        y = parts(0)
        mo = parts(1)
        d = parts(2)
        If Len(mo) = 1 Then mo = "0" & mo
        If Len(d) = 1 Then d = "0" & d
        NormalizeDateStr = y & "/" & mo & "/" & d
    Else
        NormalizeDateStr = s
    End If
End Function

Private Function NormalizeAmount(v As Variant) As Variant
    Dim s As String
    s = SafeStr(v)
    If s = "" Then
        NormalizeAmount = Empty
        Exit Function
    End If
    s = Replace(s, ",", "")
    s = Replace(s, "\", "")
    s = Replace(s, "￥", "")
    s = Replace(s, "円", "")
    If IsNumeric(s) Then
        NormalizeAmount = CDbl(s)
    Else
        NormalizeAmount = v
    End If
End Function

Private Function TruncateText(ByVal s As String, ByVal maxLen As Long) As String
    If Len(s) <= maxLen Then
        TruncateText = s
    Else
        TruncateText = Left$(s, maxLen)
    End If
End Function
