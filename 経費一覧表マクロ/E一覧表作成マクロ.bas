Attribute VB_Name = "E一覧表作成マクロ"
    Option Explicit
    
    ' === 設定（必要に応じて変えてOK） ===
    Private Const SRC_SHEET As String = "立替精算一覧"          ' 取り込み元
    Private Const DST_SHEET As String = "経費統合一覧表"    ' 追記先
    Private Const SRC_START_ROW As Long = 2                     ' 見出し1行想定
    Private Const COPY_FIRST_COL As Long = 1                    ' A列=1
    Private Const COPY_LAST_COL As Long = 18                    ' P列=16R=18
    
    ' === 追加：e-staffing_出力 の定数 ===
    Private Const ESTF_SHEET As String = "e-staffing_出力"
    Private Const ESTF_START_ROW As Long = 2   ' ヘッダー1行想定
    Private Const KW_GUSET_DUTY As String = "顧客当番" ' 判定キーワード
    
    Public Sub Append_立替精算一覧_to_経費統合一覧表()
        On Error GoTo ErrHandler
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        
        Dim wb As Workbook
        Dim wsSrc As Worksheet, wsDst As Worksheet
        Set wb = ThisWorkbook
        Set wsSrc = wb.Worksheets(SRC_SHEET)
        Set wsDst = wb.Worksheets(DST_SHEET)
        
        ' --- 取り込み元の最終行 ---
        Dim srcLastRow As Long
        srcLastRow = LastUsedRowAcross(wsSrc, COPY_FIRST_COL, COPY_LAST_COL)
        If srcLastRow < SRC_START_ROW Then
            MsgBox "取り込み元（" & SRC_SHEET & "）にデータがありません。", vbInformation
            GoTo FinallyExit
        End If
        
        ' --- 追記先の開始行（A/Bの深い方+1） ---
        Dim dstLastA As Long, dstLastB As Long, dstStartRow As Long
        dstLastA = LastUsedRowInCol(wsDst, 1) ' A
        dstLastB = LastUsedRowInCol(wsDst, 2) ' B
        dstStartRow = Application.WorksheetFunction.Max(dstLastA, dstLastB) + 1
        If dstStartRow < SRC_START_ROW Then dstStartRow = SRC_START_ROW
        
        ' --- 転送範囲 ---
        Dim rowsToCopy As Long
        rowsToCopy = srcLastRow - SRC_START_ROW + 1
        
        ' --- Read all source data into array ---
        Dim v As Variant
        v = wsSrc.Range(wsSrc.Cells(SRC_START_ROW, COPY_FIRST_COL), _
                        wsSrc.Cells(srcLastRow, COPY_LAST_COL)).value

        ' --- Build 34-column output array ---
        Dim out() As Variant
        ReDim out(1 To rowsToCopy, 1 To 34)

        Dim r As Long
        Dim sP As String, sR As String, qVal As String

        For r = 1 To rowsToCopy
            ' A-H: src cols 1-8 -> dst cols 1-8
            out(r, 1) = v(r, 1)    ' A: -> A',
            out(r, 2) = v(r, 2)    ' B: -> B
            out(r, 3) = v(r, 3)    ' C: -> C
            out(r, 4) = v(r, 4)    ' D: -> D
            out(r, 5) = v(r, 5)    ' E: -> E
            out(r, 6) = v(r, 6)    ' F: -> F
            out(r, 7) = v(r, 7)    ' G: -> G
            out(r, 8) = v(r, 8)    ' H: -> H
            ' I-L (cols 9-12): skip (empty)
            ' M-S: src cols 9-15 -> dst cols 13-19
            out(r, 13) = v(r, 9)   ' I:->M',
            out(r, 14) = v(r, 10)  ' J:->N
            out(r, 15) = v(r, 11)  ' K:->O
            out(r, 16) = v(r, 12)  ' L:->P
            out(r, 17) = v(r, 13)  ' M:->Q
            out(r, 18) = v(r, 14)  ' N:->R
            out(r, 19) = v(r, 15)  ' O:->S
            ' T(20): P(16) + R(18) combined
            sP = ""
            sR = ""
            If Not IsEmpty(v(r, 16)) Then sP = CStr(v(r, 16))
            If Not IsEmpty(v(r, 18)) Then sR = CStr(v(r, 18))
            If sP <> "" And sR <> "" Then
                out(r, 20) = sP & " / " & sR
            ElseIf sP <> "" Then
                out(r, 20) = sP
            Else
                out(r, 20) = sR
            End If
            ' AH(34): Q(17) -> if contains customer billing, set D amount
            qVal = ""
            If Not IsEmpty(v(r, 17)) Then qVal = CStr(v(r, 17))
            If InStr(1, qVal, "顧客請求分", vbTextCompare) > 0 Then
                out(r, 34) = v(r, 4)
            End If
        Next r

        ' --- Write output ---
        With wsDst
            Dim dstRange As Range
            Set dstRange = .Range(.Cells(dstStartRow, 1), .Cells(dstStartRow + rowsToCopy - 1, 34))
            dstRange.NumberFormat = "@"
            dstRange.value = out
        End With
    
    
        MsgBox "追記完了！" & vbCrLf & _
               "元：" & SRC_SHEET & "（" & SRC_START_ROW & "～" & srcLastRow & "行、A～P列）" & vbCrLf & _
               "先：" & DST_SHEET & "（" & dstStartRow & "行から " & rowsToCopy & "行 追記）", vbInformation
    
FinallyExit:
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    
ErrHandler:
        MsgBox "エラー：" & Err.Number & vbCrLf & Err.Description, vbExclamation
        Resume FinallyExit
    End Sub
    
    ' スペース揺れを吸収（半角/全角・前後空白）
    Private Function SameKey(ByVal a As String, ByVal b As String) As Boolean
        Dim na As String, nb As String
        na = Replace(Replace(Trim$(a), " ", ""), "　", "")
        nb = Replace(Replace(Trim$(b), " ", ""), "　", "")
        SameKey = (na = nb)
    End Function
    
    
    ' === 補助：列ごとの最終行 ===
    Private Function LastUsedRowInCol(ws As Worksheet, ByVal colIndex As Long) As Long
        With ws
            If Application.WorksheetFunction.CountA(.Columns(colIndex)) = 0 Then
                LastUsedRowInCol = 1
            Else
                LastUsedRowInCol = .Cells(.rows.Count, colIndex).End(xlUp).Row
            End If
        End With
    End Function
    
    ' === 補助：複数列（A～P）のうち最も下の使用行を返す ===
    Private Function LastUsedRowAcross(ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long) As Long
        Dim c As Long, m As Long, t As Long
        For c = firstCol To lastCol
            t = LastUsedRowInCol(ws, c)
            If t > m Then m = t
        Next c
        If m = 0 Then m = 1
        LastUsedRowAcross = m
    End Function
    
    
    
    Public Sub Append_e_staffing_出力_to_経費統合一覧表()
        On Error GoTo ErrHandler
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    
        Dim wb As Workbook
        Dim wsSrc As Worksheet, wsDst As Worksheet
        Set wb = ThisWorkbook
        Set wsSrc = wb.Worksheets(ESTF_SHEET)
        Set wsDst = wb.Worksheets("経費統合一覧表")
    
        ' 元データの最終行（A～G）
        Dim srcLastRow As Long
        srcLastRow = LastUsedRowAcross(wsSrc, 1, 7)
        If srcLastRow < ESTF_START_ROW Then GoTo FinallyExit
    
        Dim srcRange As Range
        Set srcRange = wsSrc.Range(wsSrc.Cells(ESTF_START_ROW, 1), wsSrc.Cells(srcLastRow, 7))
    
        Dim v As Variant: v = srcRange.value
        Dim rowsTotal As Long: rowsTotal = UBound(v, 1)
    
        ' 有効行数を数える
        Dim r As Long, n As Long
        For r = 1 To rowsTotal
            If (CStr(v(r, 1)) <> "") Or (CStr(v(r, 2)) <> "") Or (CStr(v(r, 3)) <> "") Or _
               (CStr(v(r, 4)) <> "") Or (CStr(v(r, 5)) <> "") Or (CStr(v(r, 6)) <> "") Or _
               (CStr(v(r, 7)) <> "") Then
                n = n + 1
            End If
        Next r
        If n = 0 Then GoTo FinallyExit
    
        ' 追記開始行（A/Bの深い方 +1）
        Dim dstLastA As Long, dstLastB As Long, dstStartRow As Long
        dstLastA = LastUsedRowInCol(wsDst, 1) ' A
        dstLastB = LastUsedRowInCol(wsDst, 2) ' B
        dstStartRow = Application.WorksheetFunction.Max(dstLastA, dstLastB) + 1
        If dstStartRow < ESTF_START_ROW Then dstStartRow = ESTF_START_ROW
    
        ' --- Build employee name->ID lookup from shuukei sheet ---
        Dim wsMap As Worksheet
        On Error Resume Next
        Set wsMap = wb.Worksheets("集計")
        On Error GoTo ErrHandler
        Dim dictEmp As Object: Set dictEmp = CreateObject("Scripting.Dictionary")
        dictEmp.CompareMode = 1
        If Not wsMap Is Nothing Then
            Dim mapLast As Long
            mapLast = Application.Max(LastUsedRowInCol(wsMap, 1), LastUsedRowInCol(wsMap, 2))
            If mapLast >= 2 Then
                Dim mArr As Variant
                mArr = wsMap.Range(wsMap.Cells(2, 1), wsMap.Cells(mapLast, 2)).value
                Dim mi As Long
                For mi = 1 To UBound(mArr, 1)
                    Dim mNo As String, mNm As String, mKey As String
                    If Not IsEmpty(mArr(mi, 1)) Then mNo = CStr(mArr(mi, 1)) Else mNo = ""
                    If Not IsEmpty(mArr(mi, 2)) Then mNm = CStr(mArr(mi, 2)) Else mNm = ""
                    mKey = KeyName(mNm)
                    If mKey <> "" And mNo <> "" Then
                        If Not dictEmp.Exists(mKey) Then dictEmp.Add mKey, mNo
                    End If
                Next mi
            End If
        End If

        ' --- Build 34-column output array ---
        Dim out() As Variant
        ReDim out(1 To n, 1 To 34)

        Dim i As Long: i = 0
        Dim nm As String, dt As String, dep As String, arrv As String
        Dim method As String, detail As String, amt As String, empKey As String

        For r = 1 To rowsTotal
            If (CStr(v(r, 1)) <> "") Or (CStr(v(r, 2)) <> "") Or (CStr(v(r, 3)) <> "") Or _
               (CStr(v(r, 4)) <> "") Or (CStr(v(r, 5)) <> "") Or (CStr(v(r, 6)) <> "") Or _
               (CStr(v(r, 7)) <> "") Then

                i = i + 1

                nm = CStr(v(r, 1))     ' A: name
                If IsDate(v(r, 2)) Then
                    dt = Format$(CDate(v(r, 2)), "yyyy/mm/dd")
                Else
                    dt = CStr(v(r, 2))
                End If

                dep = CStr(v(r, 3))    ' C: departure
                arrv = CStr(v(r, 4))   ' D: arrival
                method = CStr(v(r, 5)) ' E: method
                detail = CStr(v(r, 6)) ' F: detail
                amt = CStr(v(r, 7))    ' G: amount

                ' A(1): employee number from name lookup
                empKey = KeyName(nm)
                If empKey <> "" And dictEmp.Exists(empKey) Then
                    out(i, 1) = CStr(dictEmp(empKey))
                End If
                ' B(2): name
                out(i, 2) = nm
                ' F(6): date
                out(i, 6) = dt
                ' G(7): method/transport
                out(i, 7) = method
                ' H(8): detail
                out(i, 8) = detail
                ' N(14): fare amount
                out(i, 14) = amt
                ' P(16): amount (same as N)
                out(i, 16) = amt
                ' AE(31): departure station
                out(i, 31) = dep
                ' AF(32): arrival station
                out(i, 32) = arrv
                ' AH(34): customer billing (all e-staffing is customer billing)
                out(i, 34) = amt
            End If
        Next r

        ' --- Write output ---
        With wsDst
            Dim dstRange As Range
            Set dstRange = .Range(.Cells(dstStartRow, 1), .Cells(dstStartRow + n - 1, 34))
            dstRange.NumberFormat = "@"
            dstRange.value = out
        End With
    
FinallyExit:
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    
ErrHandler:
        MsgBox "e-staffing_出力 取り込みエラー：" & Err.Number & vbCrLf & Err.Description, vbExclamation
        Resume FinallyExit
    End Sub
    
    
   ' ===== 社員番号付与：集計(A=社員番号, B=名前)→経費統合一覧表(B=名前→A=社員番号) =====

' 文字列比較用：前後・全角スペースを除去してキー化
Private Function KeyName(ByVal s As String) As String
    Dim t As String
    t = CStr(s)
    t = Trim$(Replace(t, ChrW(&H3000), " ")) ' 全角スペース→半角
    t = Replace(t, " ", "")                  ' スペース除去
    KeyName = t
End Function

Public Sub AssignEmployeeNo_ByName_集計toJinjer(Optional ByVal overwrite As Boolean = False)
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsMap As Worksheet, wsDst As Worksheet
    Set wsMap = wb.Worksheets("集計")
    Set wsDst = wb.Worksheets("経費統合一覧表")

    ' 集計側の最終行（A/Bのどちらかが入っている最下行）
    Dim mapLast As Long
    mapLast = Application.Max(LastUsedRowInCol(wsMap, 1), LastUsedRowInCol(wsMap, 2))
    If mapLast < 2 Then
        MsgBox "集計シートにデータがありません。", vbExclamation
        GoTo FinallyExit
    End If

    ' 集計 A2:B(mapLast) を辞書化（キー＝名前、値＝社員番号）
    Dim m As Variant, i As Long
    m = wsMap.Range(wsMap.Cells(2, 1), wsMap.Cells(mapLast, 2)).value

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' vbTextCompare

    Dim empNo As String, nm As String, key As String
    For i = 1 To UBound(m, 1)
        ' Null/Empty対策
        If Not isEmpty(m(i, 1)) And Not IsNull(m(i, 1)) Then empNo = CStr(m(i, 1)) Else empNo = ""
        If Not isEmpty(m(i, 2)) And Not IsNull(m(i, 2)) Then nm = CStr(m(i, 2)) Else nm = ""
        
        key = KeyName(nm)
        If key <> "" And empNo <> "" Then
            If Not dict.Exists(key) Then
                dict.Add key, empNo
            End If
        End If
    Next i
    
    If dict.Count = 0 Then
        MsgBox "集計シートに有効な社員番号と名前のペアがありません。", vbExclamation
        GoTo FinallyExit
    End If

    ' 付与対象の最終行（A列とB列の最大値を取得）
    Dim dstLast As Long
    dstLast = Application.Max(LastUsedRowInCol(wsDst, 1), LastUsedRowInCol(wsDst, 2))
    If dstLast < 2 Then
        MsgBox "経費統合一覧表シートにデータがありません。", vbExclamation
        GoTo FinallyExit
    End If

    ' A/B を配列で取得
    Dim arrA As Variant, arrB As Variant
    
    ' ★単一行対策：強制的に2次元配列化
    If dstLast = 2 Then
        ReDim arrA(1 To 1, 1 To 1)
        ReDim arrB(1 To 1, 1 To 1)
        arrA(1, 1) = wsDst.Cells(2, 1).value
        arrB(1, 1) = wsDst.Cells(2, 2).value
    Else
        arrA = wsDst.Range(wsDst.Cells(2, 1), wsDst.Cells(dstLast, 1)).value
        arrB = wsDst.Range(wsDst.Cells(2, 2), wsDst.Cells(dstLast, 2)).value
    End If

    Dim filled As Long, skipped As Long
    Dim rowCount As Long: rowCount = UBound(arrB, 1)
    
    For i = 1 To rowCount
        ' Null/Empty対策
        Dim currentName As String, currentNo As String
        
        ' ★エラー値のチェック追加
        If IsError(arrB(i, 1)) Then
            Debug.Print "スキップ（エラー値）: 行" & (i + 1)
            skipped = skipped + 1
            GoTo NextRow
        End If
        
        If Not isEmpty(arrB(i, 1)) And Not IsNull(arrB(i, 1)) Then
            currentName = CStr(arrB(i, 1))
        Else
            currentName = ""
        End If
        
        If Not isEmpty(arrA(i, 1)) And Not IsNull(arrA(i, 1)) Then
            currentNo = CStr(arrA(i, 1))
        Else
            currentNo = ""
        End If
        
        key = KeyName(currentName)
        
        ' ★数値のみの名前はスキップ（社員番号が誤って名前欄に入っている場合）
        If IsNumeric(currentName) Then
            Debug.Print "スキップ（数値のみ）: 行" & (i + 1) & " 名前=[" & currentName & "]"
            skipped = skipped + 1
            GoTo NextRow
        End If
        
        If key <> "" And dict.Exists(key) Then
            ' ★★★ 修正箇所：「該当なし」も上書き対象に追加 ★★★
            If overwrite Or Len(currentNo) = 0 Or currentNo = "該当なし" Then
                arrA(i, 1) = CStr(dict(key))
                filled = filled + 1
            Else
                skipped = skipped + 1
            End If
        ElseIf key <> "" Then
            ' マッチしない名前があった場合（デバッグ用）
            Debug.Print "マッチなし: 行" & (i + 1) & " 名前=[" & currentName & "] キー=[" & key & "]"
        End If
NextRow:
    Next i

    ' 文字列書式にしてから書き戻し
    With wsDst.Range(wsDst.Cells(2, 1), wsDst.Cells(dstLast, 1))
        .NumberFormat = "@"
        .value = arrA
    End With
    
    MsgBox "処理完了" & vbCrLf & _
           "付与: " & filled & " 件" & vbCrLf & _
           "スキップ: " & skipped & " 件" & vbCrLf & _
           "処理行数: " & rowCount & " 行", vbInformation

FinallyExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "社員番号付与エラー：" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "説明: " & Err.Description & vbCrLf & _
           "発生行: " & Erl & vbCrLf & _
           "処理中の行: " & (i + 1), vbCritical
    Resume FinallyExit
End Sub
    
    
    ' 文字列キー正規化（前後/全角スペースを吸収）
    Private Function NormKey(ByVal s As String) As String
        Dim t As String
        t = CStr(s)
        t = Trim$(Replace(t, ChrW(&H3000), " ")) ' 全角→半角スペース
        t = Replace(t, " ", "")                  ' すべての空白を除去
        NormKey = t
    End Function
    
    ' A+D+F の複合キー作成
    Private Function MakeKey3(ByVal a As String, ByVal d As String, ByVal f As String) As String
        MakeKey3 = NormKey(a) & "|" & NormKey(d) & "|" & NormKey(f)
    End Function
    
    ' === 重複削除＆ログ ===
    ' ・jinjer経費利用履歴の A,D,F が同一の重複行を削除
    ' ・削除前の行を「削除ログ」へ追記（2行目～）
    
    ' 既存の NormKey / MakeKey3 はそのまま再利用します
    ' （無ければ先に貼ってある版を使ってね）
    
    
    
    ' 既存の NormKey / MakeKey3 はそのまま再利用します
    ' 無い場合は前のメッセージの定義を貼ってください。
    
    Public Sub RemoveDuplicates_A_D_F_AndLog()
        On Error GoTo ErrHandler
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    
        Dim wb As Workbook: Set wb = ThisWorkbook
        Dim ws As Worksheet, wsLog As Worksheet, wsTate As Worksheet
        Set ws = wb.Worksheets("経費統合一覧表")
    
        ' --- 立替精算一覧 から対象者セット作成（A=社員番号, B=名前） ---
        On Error Resume Next
        Set wsTate = wb.Worksheets("立替精算一覧")
        On Error GoTo 0
    
        Dim dictEmp As Object, dictName As Object
        Set dictEmp = CreateObject("Scripting.Dictionary") ' 社員番号
        Set dictName = CreateObject("Scripting.Dictionary") ' 名前（正規化）
        dictEmp.CompareMode = 1
        dictName.CompareMode = 1
    
        If Not wsTate Is Nothing Then
            Dim tLast As Long
            tLast = Application.Max( _
                LastUsedRowInCol(wsTate, 1), _
                LastUsedRowInCol(wsTate, 2))
            If tLast >= 2 Then
                Dim i As Long, emp As String, nm As String
                For i = 2 To tLast
                    emp = Trim$(CStr(wsTate.Cells(i, 1).value))        ' A: 社員番号
                    nm = CStr(wsTate.Cells(i, 2).value)                 ' B: 名前
                    If LenB(emp) > 0 Then If Not dictEmp.Exists(emp) Then dictEmp.Add emp, True
                    nm = NormKey(nm)                                    ' 空白吸収
                    If LenB(nm) > 0 Then If Not dictName.Exists(nm) Then dictName.Add nm, True
                Next
            End If
        End If
    
        ' 対象者が一人も取れなければ、何も削除しないで終了（安全側）
        If dictEmp.Count = 0 And dictName.Count = 0 Then GoTo FinallyExit
        
        

    
        ' --- ログシート確保 ---
        On Error Resume Next
        Set wsLog = wb.Worksheets("削除ログ")
        On Error GoTo 0
        If wsLog Is Nothing Then
            Set wsLog = wb.Worksheets.Add(After:=ws)
            wsLog.Name = "削除ログ"
            Dim lastColHdr As Long
            lastColHdr = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            wsLog.Range(wsLog.Cells(1, 1), wsLog.Cells(1, lastColHdr)).value = _
                ws.Range(ws.Cells(1, 1), ws.Cells(1, lastColHdr)).value
        End If
    
        ' --- 範囲端点 ---
        Dim lastR As Long
        lastR = Application.Max(LastUsedRowInCol(ws, 1), LastUsedRowInCol(ws, 4), LastUsedRowInCol(ws, 6))
        If lastR < 2 Then GoTo FinallyExit
    
        Dim lastC As Long
        lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        ' --- e-staffing優先キー集合を構築 ---
       Dim pref As Object
       Set pref = BuildPreferredKeySet_Estaff(wb)
    
       ' --- 重複検出（対象者のみ／K="片道"は除外） ---
Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
seen.CompareMode = 1

' キーごとに保持行と優先フラグを持つ
Dim keepRow As Object: Set keepRow = CreateObject("Scripting.Dictionary")
Dim keepPref As Object: Set keepPref = CreateObject("Scripting.Dictionary")

Dim delRows As New Collection
Dim logs As New Collection

Dim r As Long
Dim a As String, bNm As String, d As String, f As String, kVal As String, key As String
Dim isTargetPerson As Boolean, curPref As Boolean

For r = 2 To lastR
    a = CStr(ws.Cells(r, 1).value)   ' A: 社員番号
    bNm = CStr(ws.Cells(r, 2).value) ' B: 名前
    d = CStr(ws.Cells(r, 4).value)   ' D
    f = CStr(ws.Cells(r, 6).value)   ' F
    kVal = CStr(ws.Cells(r, 15).value) ' K

    If LenB(a) = 0 And LenB(d) = 0 And LenB(f) = 0 Then GoTo ContinueNext
    If InStr(1, kVal, "片道", vbTextCompare) > 0 Then GoTo ContinueNext

    isTargetPerson = False
    If LenB(Trim$(a)) > 0 And dictEmp.Exists(Trim$(a)) Then
        isTargetPerson = True
    ElseIf dictName.Exists(NormKey(bNm)) Then
        isTargetPerson = True
    End If
    If Not isTargetPerson Then GoTo ContinueNext

    key = MakeKey3(a, d, f)

    ' e-staffing由来か（e側キー集合に含まれるか）
    curPref = pref.Exists(key)

    If Not seen.Exists(key) Then
        seen.Add key, True
        keepRow.Add key, r
        keepPref.Add key, curPref
    Else
        Dim keptR As Long, keptP As Boolean
        keptR = keepRow(key)
        keptP = keepPref(key)

     Select Case True
    Case keptP And Not curPref
        ' 既に e を保持 → 今回（その他）は削除
        logs.Add ws.Range(ws.Cells(r, 1), ws.Cells(r, lastC)).value
        delRows.Add r

    Case (Not keptP) And curPref
        ' 今回が e、保持はその他 → 保持側（その他）を削除し入替
        logs.Add ws.Range(ws.Cells(keptR, 1), ws.Cells(keptR, lastC)).value
        delRows.Add keptR
        keepRow(key) = r
        keepPref(key) = True

    Case keptP And curPref
        ' 両方 e → どちらも削除しない（先勝ち維持・今回も残す）
        ' 何もしない（ログ/削除に入れない）

    Case Else
        ' 両方その他 → 先勝ちで今回を削除
        logs.Add ws.Range(ws.Cells(r, 1), ws.Cells(r, lastC)).value
        delRows.Add r
End Select


    End If

ContinueNext:
Next r

    
        ' --- ログ追記 & 削除 ---
        If delRows.Count > 0 Then
            Dim logStart As Long: logStart = wsLog.Cells(wsLog.rows.Count, 1).End(xlUp).Row + 1
            If logStart < 2 Then logStart = 2
    
            Dim k As Long, c As Long
            Dim buf As Variant
            ReDim buf(1 To delRows.Count, 1 To lastC)
    
            For k = 1 To delRows.Count
                Dim oneRow As Variant
                oneRow = logs(k)
                For c = 1 To lastC
                    buf(k, c) = oneRow(1, c)
                Next c
            Next k
    
            wsLog.Range(wsLog.Cells(logStart, 1), wsLog.Cells(logStart + delRows.Count - 1, lastC)).value = buf
    
            ' 下から削除
            For r = delRows.Count To 1 Step -1
                ws.rows(delRows(r)).Delete
            Next r
        End If
    
FinallyExit:
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    
ErrHandler:
        MsgBox "重複削除エラー：" & Err.Number & vbCrLf & Err.Description, vbExclamation
        Resume FinallyExit
    End Sub
    
    
   ' === ワンボタンで全部（Freee → 立替 → eスタッフ → 社員付番 → 重複削除） ===
Public Sub Append_全部_Freee含む_一括処理()
    
    ' 1. Freeeデータの取り込み
    ' ※現在のアクティブシートが整形済みのFreeeデータである必要があります
    Dim ans As VbMsgBoxResult
    ans = MsgBox("現在開いているシートは、整形済みのFreeeデータですか？" & vbCrLf & _
                 "「はい」を押すと取り込みを開始します。", vbYesNo + vbQuestion)
    If ans = vbYes Then
        Append_Freee_to_経費統合一覧表
    End If
    
    ' 2. 立替精算一覧 取り込み
    Append_立替精算一覧_to_経費統合一覧表
    
    ' 3. e-staffing 取り込み
    Append_e_staffing_出力_to_経費統合一覧表
    
    ' 3.5 jinjer CSV 取り込み
    Append_jinjer_CSV_to_経費統合一覧表
    
    ' 4. 社員番号の付与（マスタとの照合）
    ' Freeeデータは既に番号が入っていますが、立替精算などのために再度回します
    AssignEmployeeNo_ByName_集計toJinjer False
    
    ' 5. 重複削除
    RemoveDuplicates_A_D_F_AndLog
    
    MsgBox "全ての処理が完了しました！", vbInformation
End Sub
' e-staffing_出力 から「社員番号 + 金額 + 日付」のキー集合を作る
' ・名前→社員番号は「集計」シート A:社員番号, B:氏名 を使用
' ・日付は "yyyy/mm/dd"、金額は文字列で揃える（君のAppend側に合わせる）
Private Function BuildPreferredKeySet_Estaff(ByVal wb As Workbook) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    
    Dim wsSrc As Worksheet, wsMap As Worksheet
    On Error Resume Next
    Set wsSrc = wb.Worksheets(ESTF_SHEET)
    Set wsMap = wb.Worksheets("集計")
    On Error GoTo 0
    If wsSrc Is Nothing Or wsMap Is Nothing Then
        Set BuildPreferredKeySet_Estaff = dict
        Exit Function
    End If
    
    ' 名前→社員番号マップ（KeyNameで空白吸収）
    Dim mapLast As Long: mapLast = Application.Max(LastUsedRowInCol(wsMap, 1), LastUsedRowInCol(wsMap, 2))
    If mapLast < 2 Then
        Set BuildPreferredKeySet_Estaff = dict
        Exit Function
    End If
    Dim nm2id As Object: Set nm2id = CreateObject("Scripting.Dictionary")
    nm2id.CompareMode = 1
    
    Dim r As Long, key As String, nm As String, emp As String
    For r = 2 To mapLast
        nm = KeyName(CStr(wsMap.Cells(r, 2).value)) ' B:氏名（正規化）
        emp = Trim$(CStr(wsMap.Cells(r, 1).value))   ' A:社員番号
        If nm <> "" And emp <> "" Then
            If Not nm2id.Exists(nm) Then nm2id.Add nm, emp
        End If
    Next
    
    ' e-staffing_出力の最終行（A:名前, B:日付, G:金額 を使う）
    Dim srcLast As Long: srcLast = LastUsedRowAcross(wsSrc, 1, 7)
    If srcLast < ESTF_START_ROW Then
        Set BuildPreferredKeySet_Estaff = dict
        Exit Function
    End If
    
    Dim nmRaw As String, dt As String, amt As String, empNo As String
    For r = ESTF_START_ROW To srcLast
        nmRaw = CStr(wsSrc.Cells(r, 1).value) ' A:名前
        If nmRaw <> "" Then
            empNo = ""
            Dim k As String: k = KeyName(nmRaw)
            If nm2id.Exists(k) Then empNo = CStr(nm2id(k))
            
            ' 日付（B）を君のAppendと同じ文字列化
            If IsDate(wsSrc.Cells(r, 2).value) Then
                dt = Format$(CDate(wsSrc.Cells(r, 2).value), "yyyy/mm/dd")
            Else
                dt = CStr(wsSrc.Cells(r, 2).value)
            End If
            
            ' 金額（G）文字列化（Appendでは D=金額 を文字列で入れてる）
            amt = CStr(wsSrc.Cells(r, 7).value)
            
            ' 社員番号が取れていないケースはスキップ（A空のキーは誤爆しやすい）
            If empNo <> "" And amt <> "" And dt <> "" Then
                key = MakeKey3(empNo, amt, dt)  ' 既存のMakeKey3を再利用
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        End If
    Next
    
    Set BuildPreferredKeySet_Estaff = dict
End Function

' ==========================================
'  修正版2：Freee整形済みデータ → jinjer経費利用履歴
'  (2次元配列を使用してエラーを回避・高速化)
' ==========================================
' ==========================================
'  Freee整形済みデータ → 経費統合一覧表
' ==========================================
Public Sub Append_Freee_to_経費統合一覧表()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wb As Workbook
    Dim wsSrc As Worksheet, wsDst As Worksheet
    
    ' Freeeデータ（アクティブシート）
    Set wsSrc = ActiveSheet
    
    Set wb = ThisWorkbook
    On Error Resume Next
    Set wsDst = wb.Worksheets("経費統合一覧表")
    On Error GoTo ErrHandler
    
    If wsDst Is Nothing Then
        MsgBox "「経費統合一覧表」シートが見つかりません。", vbExclamation
        GoTo FinallyExit
    End If

    ' =========================================================
    ' 1. Freee側の列位置を探す（整形後のヘッダー）
    ' =========================================================
    Dim cEmpID As Long      ' 社員番号
    Dim cName As Long       ' 申請者
    Dim cDateApp As Long    ' 申請日
    Dim cTitle As Long      ' 申請タイトル
    Dim cTotal As Long      ' 合計金額
    Dim cDateUse As Long    ' 日付
    Dim cSubj As Long       ' 経費科目
    Dim cCont As Long       ' 内容
    Dim cAmt As Long        ' 金額
    Dim cMemo As Long       ' 備考
    
    cEmpID = FindColByHeader(wsSrc, "社員番号")
    cName = FindColByHeader(wsSrc, "申請者")
    cDateApp = FindColByHeader(wsSrc, "申請日")
    cTitle = FindColByHeader(wsSrc, "申請タイトル")
    cTotal = FindColByHeader(wsSrc, "合計金額")
    cDateUse = FindColByHeader(wsSrc, "日付")
    cSubj = FindColByHeader(wsSrc, "経費科目")
    cCont = FindColByHeader(wsSrc, "内容")
    cAmt = FindColByHeader(wsSrc, "金額")
    cMemo = FindColByHeader(wsSrc, "備考")

    ' 必須列チェック
    If cEmpID = 0 Or cName = 0 Then
        MsgBox "Freeeデータに「社員番号」または「申請者」列が見つかりません。" & vbCrLf & _
               "先に整形処理を実行してください。", vbExclamation
        GoTo FinallyExit
    End If

    ' =========================================================
    ' 2. 最終行の取得
    ' =========================================================
    Dim srcLastRow As Long
    srcLastRow = wsSrc.Cells(wsSrc.rows.Count, cEmpID).End(xlUp).Row
    If srcLastRow < 2 Then
        MsgBox "取り込みデータがありません。", vbInformation
        GoTo FinallyExit
    End If

    Dim rowsCount As Long
    rowsCount = srcLastRow - 1

    ' =========================================================
    ' 3. 追記先の開始行
    ' =========================================================
    Dim dstStartRow As Long
    dstStartRow = Application.WorksheetFunction.Max( _
                    wsDst.Cells(wsDst.rows.Count, 1).End(xlUp).Row, _
                    wsDst.Cells(wsDst.rows.Count, 2).End(xlUp).Row) + 1
    If dstStartRow < 2 Then dstStartRow = 2

    ' =========================================================
    ' 4. 配列準備（行数 × 19列）
    ' =========================================================
    ' 経費統合一覧表の列構成：
    ' A(1):社員番号  B(2):氏名  C(3):申請日  D(4):合計
    ' E(5):備考(申請書)  F(6):利用日  G(7):交通機関  H(8):内訳
    ' I(9):請求区分ID  J(10):請求区分  K(11):費用種別  L(12):費用種別ID
    ' M(13):小計  N(14):金額(交通費)  O(15):往復  P(16):金額
    ' Q(17):単価  R(18):数量  S(19):人数  T(20):備考(明細)
    ' U(21):計上日(yyyy/mm/dd)  V(22)-AG(33):jinjer固有  AH(34):顧客請求費
    
    Dim arr() As Variant
    ReDim arr(1 To rowsCount, 1 To 34)

    ' =========================================================
    ' 5. データ転記ループ
    ' =========================================================
    Dim i As Long, r As Long
    
    For i = 1 To rowsCount
        r = i + 1 ' データは2行目から
        
        ' A列(1): 社員番号 ← 社員番号
        If cEmpID > 0 Then arr(i, 1) = CStr(wsSrc.Cells(r, cEmpID).value)
        
        ' B列(2): 氏名 ← 申請者
        If cName > 0 Then arr(i, 2) = CStr(wsSrc.Cells(r, cName).value)
        
        ' C列(3): 申請日 ← 申請日
        If cDateApp > 0 Then arr(i, 3) = FormatDateStr(wsSrc.Cells(r, cDateApp).value)
        
        ' D列(4): 合計 ← 金額（明細の金額）
        If cAmt > 0 Then arr(i, 4) = wsSrc.Cells(r, cAmt).value
        
        ' E列(5): 備考(申請書) ← （空欄）
        
        ' F列(6): 利用日 ← 日付
        If cDateUse > 0 Then arr(i, 6) = FormatDateStr(wsSrc.Cells(r, cDateUse).value)
        
        ' G列(7): 交通機関 ← （空欄）
        
        ' H列(8): 内訳 ← 申請タイトル
        If cTitle > 0 Then arr(i, 8) = CStr(wsSrc.Cells(r, cTitle).value)
        
        ' I列(9): 小計 ← （空欄）
        ' J列(10): 金額(交通費) ← （空欄）
        ' K列(11): 往復 ← （空欄）
        ' L列(12): 金額 ← （空欄）
        ' M列(13): 単価 ← （空欄）
        ' N列(14): 数量 ← （空欄）
        ' O列(15): 人数 ← （空欄）
        
        ' P列(16): 備考(明細) ← 備考
        If cMemo > 0 Then arr(i, 20) = CStr(wsSrc.Cells(r, cMemo).value)
        
        ' Q列(17): 計上日 ← （空欄）
        
        ' R列(18): 案件・目的 ← 内容
        If cCont > 0 Then arr(i, 5) = CStr(wsSrc.Cells(r, cCont).value)
        
        ' S列(19): 顧客請求費 ← （空欄）
        
    Next i

    ' =========================================================
    ' 6. シートへの書き出し
    ' =========================================================
    With wsDst
        Dim targetRange As Range
        Set targetRange = .Range(.Cells(dstStartRow, 1), .Cells(dstStartRow + rowsCount - 1, 34))
        
        ' 書式を文字列に設定
        targetRange.NumberFormat = "@"
        
        ' 一気に貼り付け
        targetRange.value = arr
    End With

    MsgBox "Freeeデータの追記が完了しました！" & vbCrLf & _
           "件数: " & rowsCount & " 件" & vbCrLf & _
           "開始行: " & dstStartRow & " 行目", vbInformation

FinallyExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "Freee取り込みエラー：" & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume FinallyExit
End Sub

' --- 補助関数はそのまま ---
Private Function FormatDateStr(v As Variant) As String
    If IsDate(v) Then
        FormatDateStr = Format$(v, "yyyy/mm/dd")
    Else
        FormatDateStr = CStr(v)
    End If
End Function

Private Function FindColByHeader(ws As Worksheet, headerName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If InStr(ws.Cells(1, c).value, headerName) > 0 Then
            FindColByHeader = c
            Exit Function
        End If
    Next c
    FindColByHeader = 0
End Function

' === jinjer CSV から経費統合一覧表へ追記 ===
Public Sub Append_jinjer_CSV_to_経費統合一覧表()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 1. CSVファイル選択
    Dim csvPath As Variant
    csvPath = Application.GetOpenFilename("CSVファイル (*.csv),*.csv", , "jinjer CSVを選択してください")
    If csvPath = False Then GoTo FinallyExit

    ' 2. CSVを開く
    Dim wbCSV As Workbook
    Set wbCSV = Workbooks.Open(fileName:=CStr(csvPath), Local:=True)
    Dim wsCSV As Worksheet
    Set wsCSV = wbCSV.Worksheets(1)

    ' 3. 経費統合一覧表シート
    Dim wsDst As Worksheet
    Set wsDst = ThisWorkbook.Worksheets("経費統合一覧表")

    ' 4. リネームマップ
    Dim renameMap As Object
    Set renameMap = CreateObject("Scripting.Dictionary")
    renameMap.CompareMode = 1
    renameMap.Add "申請者社員番号", "社員番号"
    renameMap.Add "申請者名", "氏名"

    ' 5. CSVヘッダー読み込み
    Dim csvLastCol As Long
    csvLastCol = wsCSV.Cells(1, wsCSV.Columns.Count).End(xlToLeft).Column
    Dim csvHeaders As Object
    Set csvHeaders = CreateObject("Scripting.Dictionary")
    csvHeaders.CompareMode = 1
    Dim c As Long, hdr As String
    For c = 1 To csvLastCol
        hdr = Trim$(CStr(wsCSV.Cells(1, c).value))
        If hdr <> "" Then csvHeaders.Add hdr, c
    Next c

    ' 6. 経費統合一覧表ヘッダー読み込み
    Dim dstLastCol As Long
    dstLastCol = wsDst.Cells(1, wsDst.Columns.Count).End(xlToLeft).Column
    Dim dstHeaders As Object
    Set dstHeaders = CreateObject("Scripting.Dictionary")
    dstHeaders.CompareMode = 1
    For c = 1 To dstLastCol
        hdr = Trim$(CStr(wsDst.Cells(1, c).value))
        If hdr <> "" Then dstHeaders.Add hdr, c
    Next c

    ' 7. 列対応辞書構築 (colMap: dstCol -> csvCol)
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Dim dstKey As Variant, csvName As String
    For Each dstKey In dstHeaders.Keys
        Dim csvKey As Variant
        Dim matched As Boolean: matched = False
        For Each csvKey In csvHeaders.Keys
            csvName = CStr(csvKey)
            Dim mappedName As String
            If renameMap.Exists(csvName) Then
                mappedName = renameMap(csvName)
            Else
                mappedName = csvName
            End If
            If mappedName = CStr(dstKey) Then
                colMap.Add dstHeaders(dstKey), csvHeaders(csvKey)
                matched = True
                Exit For
            End If
        Next csvKey
    Next dstKey

    ' 8. 書き込み開始行
    Dim dstStartRow As Long
    Dim dstLastA As Long, dstLastB As Long
    dstLastA = LastUsedRowInCol(wsDst, 1)
    dstLastB = LastUsedRowInCol(wsDst, 2)
    dstStartRow = Application.WorksheetFunction.Max(dstLastA, dstLastB) + 1
    If dstStartRow < 2 Then dstStartRow = 2

    ' 9. CSVデータループ
    Dim csvLastRow As Long
    csvLastRow = wsCSV.Cells(wsCSV.rows.Count, 1).End(xlUp).Row
    Dim wr As Long: wr = dstStartRow
    Dim r As Long
    For r = 2 To csvLastRow
        If Trim$(CStr(wsCSV.Cells(r, 1).value)) = "" And _
           Trim$(CStr(wsCSV.Cells(r, 2).value)) = "" Then GoTo NextCSVRow

        Dim dstCol As Variant
        For Each dstCol In colMap.Keys
            wsDst.Cells(wr, CLng(dstCol)).value = wsCSV.Cells(r, colMap(dstCol)).value
        Next dstCol
        wr = wr + 1
NextCSVRow:
    Next r

    ' 10. CSVを閉じる
    wbCSV.Close SaveChanges:=False

    MsgBox "jinjer CSV の追記が完了しました！" & vbCrLf & _
           "追記行数: " & (wr - dstStartRow) & " 行", vbInformation

FinallyExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not wbCSV Is Nothing Then wbCSV.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "jinjer CSV 取り込みエラー：" & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume FinallyExit
End Sub
