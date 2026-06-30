Attribute VB_Name = "hSAP経費貼り付けマクロ"
Option Explicit

' ============================================================
'  SAP Fieldglass からエクスポートした経費 CSV を SAP_経費 シートに貼り付け
' ============================================================
'  - 文字コード: UTF-8 BOM 付き（SAP Reports の標準出力）
'  - 区切り: カンマ、テキスト修飾子: ダブルクォート
'  - 想定ヘッダー（13列）:
'    姓, 名, 費用合計, 業者名, 費用エントリ日, 説明, 費用シート承認日,
'    事業単位, コストセンター, 通貨, 費用シート ID, 勤務地, 費用シートのステータス
' ============================================================

Public Sub Paste_SAP経費_From_File()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 1. ファイル選択
    Dim csvPath As Variant
    csvPath = Application.GetOpenFilename( _
        "SAP 経費 CSV (*.csv),*.csv", , _
        "SAP からエクスポートした経費 CSV を選択してください")
    If csvPath = False Then GoTo FinallyExit

    ' 2. SAP_経費 シートを取得
    Dim wsDst As Worksheet
    On Error Resume Next
    Set wsDst = ThisWorkbook.Worksheets("SAP_経費")
    On Error GoTo ErrHandler
    If wsDst Is Nothing Then
        MsgBox "「SAP_経費」シートが見つかりません。", vbExclamation
        GoTo FinallyExit
    End If

    ' 3. UTF-8 として CSV を開く（OpenText / Origin:=65001 = UTF-8 codepage）
    Dim wbCSV As Workbook
    Workbooks.OpenText fileName:=CStr(csvPath), _
        Origin:=65001, _
        startRow:=1, _
        dataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=True, _
        Space:=False, _
        Other:=False, _
        Local:=False
    Set wbCSV = ActiveWorkbook

    Dim wsSrc As Worksheet
    Set wsSrc = wbCSV.Worksheets(1)

    ' 4. ソース範囲取得
    Dim srcLastRow As Long, srcLastCol As Long
    srcLastRow = wsSrc.Cells(wsSrc.rows.Count, 1).End(xlUp).Row
    srcLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    If srcLastRow < 2 Then
        MsgBox "CSV にデータ行がありません。", vbInformation
        wbCSV.Close SaveChanges:=False
        GoTo FinallyExit
    End If

    ' 5. SAP_経費 シートの既存データをクリア（ヘッダー除く）
    Dim dstLastRow As Long
    dstLastRow = wsDst.Cells(wsDst.rows.Count, 1).End(xlUp).Row
    If dstLastRow >= 2 Then
        wsDst.Range(wsDst.Cells(2, 1), wsDst.Cells(dstLastRow, 20)).Clear
    End If

    ' 6. ヘッダーを SAP_経費 の固定ヘッダーに揃え直す（CSV のヘッダー揺れ対策）
    Dim fixedHeaders As Variant
    fixedHeaders = Array("姓", "名", "費用合計", "業者名", "費用エントリ日", "説明", _
                         "費用シート承認日", "事業単位", "コストセンター", _
                         "通貨", "費用シート ID", "勤務地", "費用シートのステータス")

    Dim h As Long
    For h = 0 To UBound(fixedHeaders)
        wsDst.Cells(1, h + 1).value = fixedHeaders(h)
    Next h
    wsDst.Range("A1:M1").Font.Bold = True

    ' 7. ヘッダー名→CSV列インデックス マップを作成
    Dim csvColMap As Object
    Set csvColMap = CreateObject("Scripting.Dictionary")
    csvColMap.CompareMode = 1

    Dim c As Long, hdrName As String
    For c = 1 To srcLastCol
        hdrName = Trim$(CStr(wsSrc.Cells(1, c).value))
        If hdrName <> "" Then csvColMap(hdrName) = c
    Next c

    ' 8. CSV データを読み込んで配列に詰める
    Dim rowsToCopy As Long: rowsToCopy = srcLastRow - 1
    Dim outArr() As Variant
    ReDim outArr(1 To rowsToCopy, 1 To 13)

    Dim r As Long, i As Long
    Dim csvCol As Variant
    For r = 2 To srcLastRow
        i = r - 1
        For h = 0 To UBound(fixedHeaders)
            If csvColMap.Exists(CStr(fixedHeaders(h))) Then
                csvCol = csvColMap(CStr(fixedHeaders(h)))
                outArr(i, h + 1) = wsSrc.Cells(r, csvCol).value
            End If
        Next h
    Next r

    ' 8a. 夜間当番キーワードが説明(F列)のみにある場合は業者名(D列)へ転記
    '     下流の hSAP経費取り込みマクロ は業者名(D列)だけで夜間当番手当を判定するため、
    '     説明(F列)にしか無いと顧客請求分(AH)へ流れてしまうのを防ぐ。
    For i = 1 To rowsToCopy
        If IsNightDutyExpenseRow(outArr(i, 6), "") And Not IsNightDutyExpenseRow(outArr(i, 4), "") Then
            outArr(i, 4) = outArr(i, 6)
        End If
    Next i

    ' 8b. 夜間当番手当の税抜補正
    '     SAP は夜間当番手当（顧客対応当番／顧客当番）を税込で出力するため、
    '     C列(費用合計) を ÷1.1（四捨五入）して税抜に補正する。
    '     判定: D列(業者名) または F列(説明) に「顧客対応当番」または「顧客当番」を含む行
    Dim ndAmt As Double
    Dim ndCount As Long: ndCount = 0
    For i = 1 To rowsToCopy
        If IsNightDutyExpenseRow(outArr(i, 4), outArr(i, 6)) Then
            If TryParseAmount(outArr(i, 3), ndAmt) Then
                outArr(i, 3) = Application.WorksheetFunction.Round(ndAmt / 1.1, 0)
                ndCount = ndCount + 1
            End If
        End If
    Next i

    ' 9. SAP_経費 に書き出し
    With wsDst
        .Range(.Cells(2, 1), .Cells(rowsToCopy + 1, 13)).value = outArr
    End With

    ' 10. CSV ブックを閉じる
    wbCSV.Close SaveChanges:=False

    MsgBox "SAP 経費 CSV の取込が完了しました。" & vbCrLf & _
           "件数: " & rowsToCopy & " 行" & vbCrLf & _
           "夜間当番手当 税抜補正: " & ndCount & " 行", vbInformation

FinallyExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not wbCSV Is Nothing Then wbCSV.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "SAP 経費 CSV 取込エラー: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume FinallyExit
End Sub

' ============================================================
'  ヘルパー: 夜間当番手当（顧客対応当番／顧客当番）の行か判定
'  - D列(業者名) または F列(説明) に該当語を含むか
' ============================================================
Private Function IsNightDutyExpenseRow(ByVal vendor As Variant, ByVal descr As Variant) As Boolean
    Dim s As String
    s = SafeStr2(vendor) & " " & SafeStr2(descr)
    If InStr(1, s, "顧客対応", vbTextCompare) > 0 Then
        IsNightDutyExpenseRow = True
        Exit Function
    End If
    If InStr(1, s, "顧客当番", vbTextCompare) > 0 Then
        IsNightDutyExpenseRow = True
    End If
End Function

Private Function SafeStr2(ByVal v As Variant) As String
    If IsError(v) Then
        SafeStr2 = ""
    ElseIf IsNull(v) Or isEmpty(v) Then
        SafeStr2 = ""
    Else
        SafeStr2 = CStr(v)
    End If
End Function

Private Function TryParseAmount(ByVal v As Variant, ByRef outVal As Double) As Boolean
    Dim s As String
    s = SafeStr2(v)
    s = Replace(s, ",", "")
    s = Replace(s, "\", "")
    s = Replace(s, "￥", "")
    s = Replace(s, "円", "")
    s = Trim$(s)
    If s = "" Then
        TryParseAmount = False
        Exit Function
    End If
    If IsNumeric(s) Then
        outVal = CDbl(s)
        TryParseAmount = True
    Else
        TryParseAmount = False
    End If
End Function
