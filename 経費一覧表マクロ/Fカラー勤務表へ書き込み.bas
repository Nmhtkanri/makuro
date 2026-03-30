Attribute VB_Name = "Fカラー勤務表へ書き込み"
Sub ProcessPathList()
    Dim wsPathList As Worksheet, wsKesssan As Worksheet, wsSkip As Worksheet
    Dim wbTarget As Workbook
    Dim i As Long, rowNum As Long
    Dim employeeID As String, filePath As String
    Dim skipData As Collection
    
    On Error GoTo ErrorHandler
    
    Set wsPathList = ThisWorkbook.Sheets("PathLis")
    Set wsKesssan = ThisWorkbook.Sheets("集計")
    Set wsSkip = ThisWorkbook.Sheets("スキップ")
    Set skipData = New Collection
    
    ' スキップシートをクリア
    wsSkip.Cells.Clear
    
    ' PathListの行数を確認
    rowNum = wsPathList.Cells(wsPathList.rows.Count, 1).End(xlUp).Row
    
    ' 各行を処理
    For i = 2 To rowNum ' ヘッダーをスキップ
        employeeID = wsPathList.Cells(i, 1).value
        filePath = wsPathList.Cells(i, 3).value
        
        ' 社員番号がない場合はスキップ
        If employeeID = "" Or filePath = "" Then
            GoTo NextIteration
        End If
        
        ' 集計シートから該当社員の行を検索
        Dim kessanRow As Long
        kessanRow = FindEmployeeRow(wsKesssan, employeeID)
        
        If kessanRow = 0 Then
            GoTo NextIteration
        End If
        
        ' 書き込む値があるか確認
        Dim hasValue As Boolean
        hasValue = HasValueToWrite(wsKesssan, kessanRow)
        
        If Not hasValue Then
            GoTo NextIteration
        End If
        
        ' ファイルを開く
        If Not FileExists(filePath) Then
            GoTo NextIteration
        End If
        
        Set wbTarget = Workbooks.Open(filePath)
        
        ' データを書き込み
        Call WriteDataToTarget(wsKesssan, wbTarget, kessanRow, employeeID, skipData)
        
        ' ファイルを保存して閉じる
        wbTarget.Save
        wbTarget.Close SaveChanges:=True
        
NextIteration:
    Next i
    
    ' スキップデータをスキップシートに記録
    If skipData.Count > 0 Then
        Call WriteSkipData(wsSkip, skipData)
    End If
    
    MsgBox "処理が完了しました。", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました：" & Err.Description, vbCritical
    If Not wbTarget Is Nothing Then
        wbTarget.Close SaveChanges:=False
    End If
End Sub

Function FindEmployeeRow(ws As Worksheet, employeeID As String) As Long
    Dim i As Long
    For i = 2 To ws.Cells(ws.rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 1).value = employeeID Then
            FindEmployeeRow = i
            Exit Function
        End If
    Next i
    FindEmployeeRow = 0
End Function

Function HasValueToWrite(wsKesssan As Worksheet, kessanRow As Long) As Boolean
    ' R, S, T, U, V, W, X, Y列のいずれかに値があるか確認
    Dim col As Long
    For col = 18 To 25 ' R~Y列
        If wsKesssan.Cells(kessanRow, col).value <> "" Then
            HasValueToWrite = True
            Exit Function
        End If
    Next col
    HasValueToWrite = False
End Function

Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function

Sub WriteDataToTarget(wsKesssan As Worksheet, wbTarget As Workbook, _
                      kessanRow As Long, employeeID As String, skipData As Collection)
    Dim wsTarget As Worksheet
    Dim srcCol As Long, tarRow As Long, tarCol As Long
    Dim srcCols As Variant, tarInfo As Variant
    Dim i As Long, value As Variant
    
    Dim vR As Variant, vS As Variant
    Dim vT As Variant, vU As Variant
    Dim destRow As Long
    Dim r As Long, c As Long
    Dim v As Variant
    Dim hasExactRS As Boolean, hasSameNameDiffAmtRS As Boolean
    Dim rsDiffRow As Long, rsExistingAmt As Variant
    Dim hasExactTU As Boolean, hasSameNameDiffAmtTU As Boolean
    Dim tuDiffRow As Long, tuExistingAmt As Variant
    Dim classMatch As Boolean
    
    ' 書き込みルール（イメージ）
    ' ・R/S, T/U は 62～64行の空き行に「セット」で書き込み
    '   - 内訳名：D～I列
    '   - 金額  ：J～K列
    ' ・同じ「内訳名＋金額」のセットがすでにあればスキップ（ログなし）
    ' ・同じ内訳名で金額だけ違う場合はスキップシートに記録して書き込まない
    ' ・V～Y は従来通りの行に書き込み
    
    Set wsTarget = wbTarget.ActiveSheet
    
    ' 集計シート側の値を先に取得
    vR = wsKesssan.Cells(kessanRow, 18).value  ' R列（内訳1）
    vS = wsKesssan.Cells(kessanRow, 19).value  ' S列（内訳金額1）
    vT = wsKesssan.Cells(kessanRow, 20).value  ' T列（内訳2）
    vU = wsKesssan.Cells(kessanRow, 21).value  ' U列（内訳金額2）
    
    '========================
    ' 内訳1 (R / S)
    '========================
    If (vR <> "" Or vS <> "") Then
        
        hasExactRS = False
        hasSameNameDiffAmtRS = False
        rsDiffRow = 0
        rsExistingAmt = Empty
        
        ' 62～64行を走査して、
        ' ・同じ内訳名＋同じ金額
        ' ・同じ内訳名＋違う金額
        ' を探す
        For r = 62 To 64
            classMatch = False
            
            ' 内訳名（D～I列）をチェック
            For c = 4 To 9
                If wsTarget.Cells(r, c).value = vR Then
                    classMatch = True
                    Exit For
                End If
            Next c
            
            If classMatch Then
                ' 金額（J～K列）をチェック
                For c = 10 To 11
                    v = wsTarget.Cells(r, c).value
                    If v <> "" Then
                        If v = vS Then
                            ' 内訳名＋金額 完全一致
                            hasExactRS = True
                            Exit For
                        Else
                            ' 内訳名は同じだが金額が違う
                            hasSameNameDiffAmtRS = True
                            rsDiffRow = r
                            rsExistingAmt = v
                            Exit For
                        End If
                    End If
                Next c
            End If
            
            If hasExactRS Or hasSameNameDiffAmtRS Then Exit For
        Next r
        
        If hasExactRS Then
            ' 完全一致 → 何もせず終了（重複スキップ）
            
        ElseIf hasSameNameDiffAmtRS Then
            ' 内訳名は同じで金額だけ違う → スキップシート用に記録
            skipData.Add Array( _
                employeeID, _
                "内訳1(金額不一致)", _
                "内訳:" & vR & " / 既存金額:" & rsExistingAmt & " / 新金額:" & vS, _
                rsDiffRow _
            )
        Else
            ' 同じ内訳＋金額も、同じ内訳＋別金額も無し
            ' → 62～64行の中から完全な空き行を探してセットで書き込み
            destRow = FindFirstEmptyRowForPair(wsTarget, 62, 64)
            
            If destRow > 0 Then
                ' 内訳名（R）→ D～I列（プルダウンから選択）
                If vR <> "" Then
                    Call SelectDropdownValue(wsTarget, destRow, 4, vR, _
                                             employeeID, skipData, "内訳1", wbTarget)
                End If
                
                ' 内訳金額（S）→ J～K列
                If vS <> "" Then
                    Call WriteRangeIfEmpty(wsTarget, destRow, 10, 10, vS, _
                                           employeeID, skipData, "内訳金額1")
                End If
            End If
        End If
    End If
    
    '========================
    ' 内訳2 (T / U)
    '========================
    If (vT <> "" Or vU <> "") Then
        
        hasExactTU = False
        hasSameNameDiffAmtTU = False
        tuDiffRow = 0
        tuExistingAmt = Empty
        
        ' 62～64行を走査（内訳2用）
        For r = 62 To 64
            classMatch = False
            
            ' 内訳名（D～I列）をチェック
            For c = 4 To 9
                If wsTarget.Cells(r, c).value = vT Then
                    classMatch = True
                    Exit For
                End If
            Next c
            
            If classMatch Then
                ' 金額（J～K列）をチェック
                For c = 10 To 11
                    v = wsTarget.Cells(r, c).value
                    If v <> "" Then
                        If v = vU Then
                            ' 内訳名＋金額 完全一致
                            hasExactTU = True
                            Exit For
                        Else
                            ' 内訳名は同じだが金額が違う
                            hasSameNameDiffAmtTU = True
                            tuDiffRow = r
                            tuExistingAmt = v
                            Exit For
                        End If
                    End If
                Next c
            End If
            
            If hasExactTU Or hasSameNameDiffAmtTU Then Exit For
        Next r
        
        If hasExactTU Then
            ' 完全一致 → 何もせず終了
            
        ElseIf hasSameNameDiffAmtTU Then
            ' 内訳名は同じで金額だけ違う → スキップシート用に記録
            skipData.Add Array( _
                employeeID, _
                "内訳2(金額不一致)", _
                "内訳:" & vT & " / 既存金額:" & tuExistingAmt & " / 新金額:" & vU, _
                tuDiffRow _
            )
        Else
            ' 同じ内訳＋金額も、同じ内訳＋別金額も無し
            ' → 空き行にセットで書き込み
            destRow = FindFirstEmptyRowForPair(wsTarget, 62, 64)
            
            If destRow > 0 Then
                ' 内訳名（T）→ D～I列（プルダウンから選択）
                If vT <> "" Then
                    Call SelectDropdownValue(wsTarget, destRow, 4, vT, _
                                             employeeID, skipData, "内訳2", wbTarget)
                End If
                
                ' 内訳金額（U）→ J～K列
                If vU <> "" Then
                    Call WriteRangeIfEmpty(wsTarget, destRow, 10, 11, vU, _
                                           employeeID, skipData, "内訳金額2")
                End If
            End If
        End If
    End If
    
    '========================
    ' 以下、V～Y は従来通りの固定行に書き込み
    '========================
    
    ' 通勤交通費 (V列→66行 J~K列)
    If wsKesssan.Cells(kessanRow, 22).value <> "" Then ' V列
        Call WriteRangeIfEmpty(wsTarget, 66, 10, 11, wsKesssan.Cells(kessanRow, 22).value, _
                               employeeID, skipData, "通勤交通費")
    End If
    
    ' 顧客請求分 (W列→67行 J~K列)
    If wsKesssan.Cells(kessanRow, 23).value <> "" Then ' W列
        Call WriteRangeIfEmpty(wsTarget, 67, 10, 11, wsKesssan.Cells(kessanRow, 23).value, _
                               employeeID, skipData, "顧客請求分")
    End If
    
    ' 非課税精算(立替金) (X列→68行 J~K列)
    If wsKesssan.Cells(kessanRow, 24).value <> "" Then ' X列
        Call WriteRangeIfEmpty(wsTarget, 68, 10, 11, wsKesssan.Cells(kessanRow, 24).value, _
                               employeeID, skipData, "非課税精算(立替金)")
    End If
    
    ' 非課税精算(その他) (Y列→69行 J~K列)
    If wsKesssan.Cells(kessanRow, 25).value <> "" Then ' Y列
        Call WriteRangeIfEmpty(wsTarget, 69, 10, 11, wsKesssan.Cells(kessanRow, 25).value, _
                               employeeID, skipData, "非課税精算(その他)")
    End If
End Sub



'====================================================
' 書き込み先シートの 62～64 行 D～I列・J～K列に
' 「内訳名（文字列）＋金額」の組み合わせが
' すでに存在するかをチェックする関数
'====================================================
Private Function ExistsPairInTarget( _
        ByVal ws As Worksheet, _
        ByVal firstRow As Long, ByVal lastRow As Long, _
        ByVal classValue As Variant, _
        ByVal amountValue As Variant _
    ) As Boolean

    Dim r As Long, c As Long
    Dim hasClass As Boolean, hasAmount As Boolean
    Dim v As Variant

    For r = firstRow To lastRow
        hasClass = False
        hasAmount = False

        '--- 内訳名：D～I列（4～9列） ---
        For c = 4 To 9
            v = ws.Cells(r, c).value
            If v = classValue Then
                hasClass = True
                Exit For
            End If
        Next c

        '--- 金額：J～K列（10～11列） ---
        For c = 10 To 11
            v = ws.Cells(r, c).value
            If v = amountValue Then
                hasAmount = True
                Exit For
            End If
        Next c

        ' 同じ行に「内訳名」と「金額」がそろっていれば同じセットとみなす
        If hasClass And hasAmount Then
            ExistsPairInTarget = True
            Exit Function
        End If
    Next r
End Function


'====================================================
' (2) 62～64 行の中で、D～I列 & J～K列 が
'     すべて空白の「空き行」を探す
'     見つかればその行番号、なければ 0 を返す
'====================================================
Private Function FindFirstEmptyRowForPair( _
        ByVal ws As Worksheet, _
        ByVal firstRow As Long, ByVal lastRow As Long _
    ) As Long

    Dim r As Long, c As Long
    Dim isEmpty As Boolean

    For r = firstRow To lastRow
        isEmpty = True
        For c = 4 To 11   ' D(4)～K(11)
            If Len(ws.Cells(r, c).value) > 0 Then
                isEmpty = False
                Exit For
            End If
        Next c

        If isEmpty Then
            FindFirstEmptyRowForPair = r
            Exit Function
        End If
    Next r
End Function


Sub SelectDropdownValue(wsTarget As Worksheet, tarRow As Long, tarCol As Long, _
                        value As Variant, employeeID As String, skipData As Collection, dataType As String, wbTarget As Workbook)
    Dim tarCell As Range
    Dim wsTable As Worksheet
    Dim i As Long
    Dim cellValue As Variant
    Dim found As Boolean
    
    Set tarCell = wsTarget.Cells(tarRow, tarCol)
    
    ' 結合セルが空か、または"0"か確認
    If tarCell.value <> "" And tarCell.value <> 0 Then
        ' スキップに記録
        skipData.Add Array(employeeID, dataType, value, tarRow)
        Exit Sub
    End If
    
    found = False
    
    ' ターゲットファイル内のテーブルシートを取得
    On Error Resume Next
    Set wsTable = wbTarget.Sheets("テーブル")
    On Error GoTo 0
    
    If wsTable Is Nothing Then
        ' テーブルシートがない場合は直接値を入力
        tarCell.value = value
        Exit Sub
    End If
    
    ' テーブルシートのQ列（17列目）から該当値を探す
    ' Q2からデータが始まる
    Dim lastRow As Long
    lastRow = wsTable.Cells(wsTable.rows.Count, 17).End(xlUp).Row
    
    For i = 2 To lastRow
        cellValue = wsTable.Cells(i, 17).value
        If cellValue = value Then
            tarCell.value = value
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        ' プルダウンに該当値がない場合はスキップに記録
        skipData.Add Array(employeeID, dataType, value, tarRow)
    End If
End Sub

Sub WriteRangeIfEmpty(wsTarget As Worksheet, tarRow As Long, startCol As Long, _
                      endCol As Long, value As Variant, employeeID As String, _
                      skipData As Collection, dataType As String)
    Dim tarCell As Range
    
    ' 左上のセルを取得
    Set tarCell = wsTarget.Cells(tarRow, startCol)
    
    ' 結合セルが空か、または"0"か確認（"0"の場合は上書き）
    If tarCell.value <> "" And tarCell.value <> 0 Then
        ' スキップに記録
        skipData.Add Array(employeeID, dataType, value, tarRow)
    Else
        ' 値を書き込み（空または0の場合）
        tarCell.value = value
    End If
End Sub

Sub WriteSkipData(wsSkip As Worksheet, skipData As Collection)
    Dim i As Long
    wsSkip.Cells(1, 1).value = "社員番号"
    wsSkip.Cells(1, 2).value = "データ種別"
    wsSkip.Cells(1, 3).value = "値"
    wsSkip.Cells(1, 4).value = "行番号"
    
    For i = 1 To skipData.Count
        wsSkip.Cells(i + 1, 1).value = skipData(i)(0)
        wsSkip.Cells(i + 1, 2).value = skipData(i)(1)
        wsSkip.Cells(i + 1, 3).value = skipData(i)(2)
        wsSkip.Cells(i + 1, 4).value = skipData(i)(3)
    Next i
End Sub
