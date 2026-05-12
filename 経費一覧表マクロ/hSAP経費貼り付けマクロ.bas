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
    Workbooks.OpenText FileName:=CStr(csvPath), _
        Origin:=65001, _
        StartRow:=1, _
        DataType:=xlDelimited, _
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

    ' 9. SAP_経費 に書き出し
    With wsDst
        .Range(.Cells(2, 1), .Cells(rowsToCopy + 1, 13)).value = outArr
    End With

    ' 10. CSV ブックを閉じる
    wbCSV.Close SaveChanges:=False

    MsgBox "SAP 経費 CSV の取込が完了しました。" & vbCrLf & _
           "件数: " & rowsToCopy & " 行", vbInformation

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
