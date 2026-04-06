Attribute VB_Name = "GFREE整形幕r"
Sub Freeeデータ整形_列削除と名称変換_最終版()
    Dim wsData As Worksheet
    Dim wsMaster As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim empName As String
    Dim cleanName As String
    Dim cellVal As String
    Dim nameCol As Long
    
    ' 高速化のための辞書生成
    Dim dictMaster As Object
    Set dictMaster = CreateObject("Scripting.Dictionary")
    
    ' エラー処理
    On Error Resume Next
    Set wsData = ActiveSheet
    Set wsMaster = Worksheets("集計")
    On Error GoTo 0
    
    If wsMaster Is Nothing Then
        MsgBox "「集計」シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' =========================================================
    ' 0. 整形済みチェック（二重実行防止）
    ' =========================================================
    ' 1行目に「社員番号」列があれば、すでに整形済みと判断
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        If Trim(CStr(wsData.Cells(1, i).value)) = "社員番号" Then
            MsgBox "このシートは既に整形済みです。" & vbCrLf & _
                   "元データを再度貼り付けてから実行してください。", vbExclamation
            GoTo CleanUp
        End If
    Next i
    
    ' =========================================================
    ' 1. 残したい列を列名で特定（元データのヘッダーから探す）
    ' =========================================================
    Dim keepHeaders As Variant
    keepHeaders = Array("申請日", "申請者", "申請タイトル", "合計金額", _
                        "日付", "経費科目", "内容", "金額", "備考")
    
    ' 各ヘッダーの列番号を格納する配列
    Dim keepCols() As Long
    ReDim keepCols(LBound(keepHeaders) To UBound(keepHeaders))
    
    Dim headerName As String
    Dim foundCol As Long
    Dim j As Long
    
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    
    ' 残したい列の位置を探す
    For j = LBound(keepHeaders) To UBound(keepHeaders)
        keepCols(j) = 0 ' 初期化（見つからない場合は0）
        For i = 1 To lastCol
            headerName = Trim(CStr(wsData.Cells(1, i).value))
            If headerName = keepHeaders(j) Then
                keepCols(j) = i
                Exit For
            End If
        Next i
    Next j
    
    ' 必須列が見つからない場合は警告
    Dim missingCols As String
    missingCols = ""
    For j = LBound(keepHeaders) To UBound(keepHeaders)
        If keepCols(j) = 0 Then
            missingCols = missingCols & keepHeaders(j) & ", "
        End If
    Next j
    
    If Len(missingCols) > 0 Then
        missingCols = Left(missingCols, Len(missingCols) - 2)
        MsgBox "以下の列が元データに見つかりません：" & vbCrLf & _
               missingCols & vbCrLf & vbCrLf & _
               "元データのヘッダーを確認してください。", vbExclamation
        GoTo CleanUp
    End If
    
    ' =========================================================
    ' 2. 新しいシートに必要な列だけコピー
    ' =========================================================
    ' ★★★ 修正箇所1：最終行を正確に取得（全列の最大値を使用）★★★
    Dim tempLastRow As Long
    lastRow = 1
    
    ' まず、残したい列の最終行をチェック
    For j = LBound(keepCols) To UBound(keepCols)
        If keepCols(j) > 0 Then
            tempLastRow = wsData.Cells(wsData.rows.Count, keepCols(j)).End(xlUp).Row
            If tempLastRow > lastRow Then lastRow = tempLastRow
        End If
    Next j
    
    ' さらに、元データ全体の全列をチェック（A～Q列が空でもR列以降にデータがある場合に対応）
    For i = 1 To lastCol
        tempLastRow = wsData.Cells(wsData.rows.Count, i).End(xlUp).Row
        If tempLastRow > lastRow Then lastRow = tempLastRow
    Next i
    ' ★★★ 修正箇所1ここまで ★★★
    
    If lastRow < 2 Then
        MsgBox "データがありません。", vbInformation
        GoTo CleanUp
    End If
    
    ' 作業用配列を作成（行数 × 残す列数）
    Dim resultArr() As Variant
    ReDim resultArr(1 To lastRow, 1 To UBound(keepHeaders) + 1)
    
    Dim r As Long, c As Long
    
    ' ヘッダー行
    For c = LBound(keepHeaders) To UBound(keepHeaders)
        resultArr(1, c + 1) = keepHeaders(c)
    Next c
    
    ' データ行
    For r = 2 To lastRow
        For c = LBound(keepHeaders) To UBound(keepHeaders)
            If keepCols(c) > 0 Then
                resultArr(r, c + 1) = wsData.Cells(r, keepCols(c)).value
            End If
        Next c
    Next r
    
    ' =========================================================
    ' 3. 元のシートをクリアして結果を貼り付け
    ' =========================================================
    wsData.Cells.Clear
    wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, UBound(keepHeaders) + 1)).value = resultArr
    
    ' =========================================================
    ' 4. 名前列を特定（「申請者」列）
    ' =========================================================
    nameCol = 2 ' 整形後は2列目が「申請者」
    
    ' =========================================================
    ' 5. マスタ読み込み
    ' =========================================================
    Dim lastRowMaster As Long
    lastRowMaster = wsMaster.Cells(wsMaster.rows.Count, 2).End(xlUp).Row
    For i = 2 To lastRowMaster
        empName = wsMaster.Cells(i, 2).value
        cleanName = Replace(Replace(empName, " ", ""), "　", "")
        If cleanName <> "" And Not dictMaster.Exists(cleanName) Then
            dictMaster.Add cleanName, wsMaster.Cells(i, 1).value
        End If
    Next i
    
    ' =========================================================
    ' 6. 社員番号列を追加
    ' =========================================================
    Dim empNoCol As Long
    empNoCol = UBound(keepHeaders) + 2 ' 最後の列の次
    wsData.Cells(1, empNoCol).value = "社員番号"
    
    ' =========================================================
    ' 7. 名称変換（メールアドレス対応）
    ' =========================================================
    ' ★★★ 修正箇所2：lastRowを再取得しない（最初に取得した値を使い続ける）★★★
    ' 削除：lastRow = wsData.Cells(wsData.Rows.Count, nameCol).End(xlUp).Row
    ' lastRowは既に正確な値が入っているのでそのまま使用
    ' ★★★ 修正箇所2ここまで ★★★
    
    For i = 2 To lastRow
        cellVal = Trim(CStr(wsData.Cells(i, nameCol).value))
        
        If cellVal <> "" Then
            ' メールアドレスの@以降を削除
            If InStr(cellVal, "@") > 0 Then
                cellVal = Left(cellVal, InStr(cellVal, "@") - 1)
            End If
            
            ' 名称変換（部分一致で判定）
            If InStr(1, cellVal, "Tomono", vbTextCompare) > 0 Then
                wsData.Cells(i, nameCol).value = "友納 英彦"
            ElseIf InStr(1, cellVal, "maki.murayama", vbTextCompare) > 0 Then
                wsData.Cells(i, nameCol).value = "村山 真紀"
            ElseIf InStr(1, cellVal, "kazushi.mitani", vbTextCompare) > 0 Then
                wsData.Cells(i, nameCol).value = "三谷 一志"
            ElseIf InStr(1, cellVal, "kousei.shiokawa", vbTextCompare) > 0 Then
                wsData.Cells(i, nameCol).value = "塩川 浩生"
            ElseIf InStr(1, cellVal, "rina.hirano", vbTextCompare) > 0 Then
                wsData.Cells(i, nameCol).value = "平野 梨奈"
            End If
        End If
    Next i
    
    ' =========================================================
    ' 8. 空白埋め処理
    ' =========================================================
    ' ★★★ 修正箇所3：lastRowはそのまま使用（再取得不要）★★★
    ' これにより、申請者が空欄の最後の行にも正しく名前がコピーされる
    For i = 2 To lastRow
        ' 申請日（1列目）
        If wsData.Cells(i, 1).value = "" Then wsData.Cells(i, 1).value = wsData.Cells(i - 1, 1).value
        ' 申請者（2列目）
        If wsData.Cells(i, 2).value = "" Then wsData.Cells(i, 2).value = wsData.Cells(i - 1, 2).value
        ' 申請タイトル（3列目）
        If wsData.Cells(i, 3).value = "" Then wsData.Cells(i, 3).value = wsData.Cells(i - 1, 3).value
        ' 合計金額（4列目）
        If wsData.Cells(i, 4).value = "" Then wsData.Cells(i, 4).value = wsData.Cells(i - 1, 4).value
    Next i
    ' ★★★ 修正箇所3ここまで ★★★
    
    ' =========================================================
    ' 9. 申請日を日付形式に変換
    ' =========================================================
    With wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRow, 1))
        .NumberFormat = "yyyy/mm/dd"
    End With
    
    ' =========================================================
    ' 10. 社員番号付与
    ' =========================================================
    For i = 2 To lastRow
        empName = CStr(wsData.Cells(i, nameCol).value)
        cleanName = Replace(Replace(empName, " ", ""), "　", "")
        
        If dictMaster.Exists(cleanName) Then
            wsData.Cells(i, empNoCol).value = dictMaster(cleanName)
        Else
            wsData.Cells(i, empNoCol).value = "該当なし"
        End If
    Next i
    
   ' 列幅調整
    wsData.Columns("A:" & Split(wsData.Cells(1, empNoCol).Address, "$")(1)).AutoFit
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "整形完了！" & vbCrLf & _
           "・不要列を削除し、必要な9列を抽出しました" & vbCrLf & _
           "・名称変換、空白埋め、社員番号付与を行いました" & vbCrLf & vbCrLf & _
           "転記は「経費統合一覧表」シートから実行してください。", vbInformation

    Exit Sub

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

