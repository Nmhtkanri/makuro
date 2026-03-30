Attribute VB_Name = "d集計シート"
        Option Explicit
        
        ' === シート名 ===
        Private Const SH_SUM As String = "集計"               ' 出力先（並べ替え後）
        Private Const SH_SRC As String = "経費統合一覧表" ' 取り込み元
        
        ' === 出力列（新レイアウト）===
        Private Const COL_EMP_NO As Long = 1   ' A: 社員番号
        Private Const COL_NAME  As Long = 2    ' B: 氏名
        Private Const COL_TOTAL As Long = 3    ' C: 合計（D+E+H+I+J）
        Private Const COL_GK    As Long = 4    ' D: 夜間当番手当
        Private Const COL_RINK  As Long = 5    ' E: RINK手当
        Private Const COL_ALLOW2 As Long = 6   ' F: 手当2（顧客＋RINK）=D+E
        Private Const COL_BILL  As Long = 7    ' G: 顧客請求分
        Private Const COL_TRANS As Long = 8    ' H: 交通費
        Private Const COL_NONTAX_TATEKAE As Long = 9  ' I: 非課税精算(立替金) ＜ NEW
        Private Const COL_ETC   As Long = 10   ' J: その他(精算済・ぴんなど)
        Private Const COL_TW    As Long = 11   ' K: テレワーク手当
        Private Const COL_DATE  As Long = 12   ' L: 請求日
        
        ' ================= エントリ（安全版） =================
        Public Sub Run_Safe_集計_新レイアウト()
            On Error GoTo ErrHandler
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
        
            Dim agg As Object, maxDate As Object, empList As Object, hitCount As Long
            Set agg = CreateObject("Scripting.Dictionary")   ' key -> Array( D,E,H,I,J,G )
            Set maxDate = CreateObject("Scripting.Dictionary")
            Set empList = CreateObject("Scripting.Dictionary")
        
            ' 1) 取り込み試行（ここで“何件取れたか”判定）
            hitCount = Collect_From_Jinjer(agg, maxDate, empList)
        
            If hitCount = 0 Then
                ' 0件なら何も壊さない（警告だけ出す）
                MsgBox "取り込み件数が0でした。出力シートは変更していません。" & vbCrLf & _
                       "・取り込み元シート名: " & SH_SRC & vbCrLf & _
                       "・見出し名の揺れ/金額0/キーワード未一致が原因の可能性があります。", vbExclamation
                GoTo FinallyExit
            End If
        
            ' 2) ここで初めてバックアップを作ってから、出力を作り直し
            BackupSheet SH_SUM
            Rewrite_Output_KeepAB agg, maxDate
        
FinallyExit:
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Exit Sub
ErrHandler:
            MsgBox "エラー: " & Err.Number & vbCrLf & Err.Description, vbExclamation
            Resume FinallyExit
        End Sub
        
Private Function Collect_From_Jinjer(ByRef agg As Object, ByRef maxDate As Object, ByRef empList As Object) As Long
    Dim ws As Worksheet
    If Not SheetExists(SH_SRC) Then
        MsgBox "取り込み元シートが見つかりません: " & SH_SRC, vbExclamation
        Collect_From_Jinjer = 0
        Exit Function
    End If
    Set ws = Worksheets(SH_SRC)

    Dim cEmpNo&, cName&, cUch&, cTrans&, cAmt&, cFareAmt&, cBooked&, cEStaff&
    cEmpNo = FindCol(ws, 1, Array("社員番号", "従業員番号", "社員ID"))
    cName = FindCol(ws, 1, Array("氏名", "名前"))
    cUch = FindCol(ws, 1, Array("内訳", "摘要", "内容"))
    cTrans = FindCol(ws, 1, Array("交通機関", "経路", "移動手段"), True)
    cAmt = FindCol(ws, 1, Array("合計", "小計", "金額"))
    cFareAmt = FindCol(ws, 1, Array("金額(交通費)", "交通費金額", "交通費（円）", "交通費(円)"), True)
    cBooked = FindCol(ws, 1, Array("計上日", "計上", "計上日付"), True)
    cEStaff = FindCol(ws, 1, Array("顧客請求費", "顧客請求費", "顧客対応", "顧客当番", "夜間当番", "24時間準直当番", "深夜出動", "24時間準直当番手当"), True)
    If cEStaff = 0 Then cEStaff = 19 ' S列フォールバック

    If cEmpNo = 0 Or cName = 0 Or cUch = 0 Or cAmt = 0 Then
        MsgBox "必須列が見つからないため取り込み中止。" & vbCrLf & _
               "社員番号/氏名/内訳/合計（金額）の見出しをご確認ください。", vbExclamation
        Collect_From_Jinjer = 0
        Exit Function
    End If

    Dim lastR&, r&, hits&
    lastR = ws.Cells(ws.rows.Count, cEmpNo).End(xlUp).Row
    If lastR < 2 Then
        Collect_From_Jinjer = 0
        Exit Function
    End If

    Dim kwTW, kwRink, kwGK, kwTrans, kwExcludeFromG, kwHotel, kwOnCall, kwDailyAllo
    Dim kwExcludeTrans ' NGワード用
    
    kwTW = Array("テレワーク", "在宅")
    kwRink = Array("RINK")
    kwGK = Array("夜間当番", "24時間準直当番", "準直当番", "深夜出動", "顧客当番", "顧客対応当番")
    
    ' 交通費キーワード（キーワードを網羅）
    kwTrans = Array("交通費", "電車", "バス", "タクシー", "地下鉄", "鉄道", "新幹線", "モノレール", _
                    "JR", "私鉄", "有料列車", "飛行機", "航空券", "定期", _
                    "ガソリン", "燃料", "駐車場", "パーキング", "高速", "ＥＴＣ", "ETC", _
                    "車両", "レンタカー", "移動", "旅費", "出張", "北総鉄道北総線", "自動車")
                    
    ' ★NGワード（これが入っていたら交通費にしない）
    kwExcludeTrans = Array("会議", "交際", "接待", "飲食", "手土産", "福利厚生", "親睦", "定期健康診断", "健康診断")
    
    kwExcludeFromG = Array("夜間当番", "RINK", "顧客当番", "顧客対応当番", "オンコール")
    
    ' ★キーワード定義を少し強化しておきます
    kwHotel = Array("宿", "ホテル", "宿泊", "イン")
    kwDailyAllo = Array("日当", "出張") ' 出張手当なども拾えるように
    kwOnCall = Array("オンコール")
    
    For r = 2 To lastR
        Dim empNo As String, empNm As String, key As String
        Dim desc As String, trans As String
        Dim amt As Double, fa As Double, estAmt As Double
        Dim estFilled As Boolean
        
        ' フラグ変数を宣言
        Dim isHotel As Boolean
        Dim isOnCall As Boolean
        Dim isDailyAllo As Boolean ' ★ここが重要
        Dim isTraffic As Boolean
        Dim isExcluded As Boolean

        empNo = NormalizeId(ws.Cells(r, cEmpNo).value)
        empNm = NormalizeName(ws.Cells(r, cName).value)
        If empNo = "" And empNm = "" Then GoTo NextR

        key = BuildEmpKey(empNo, empNm)
        If Not empList.Exists(key) Then empList.Add key, Array(empNo, empNm)

        ' 値の取得
        desc = NormalizeStr(ws.Cells(r, cUch).value)
        trans = IIf(cTrans > 0, NormalizeStr(ws.Cells(r, cTrans).value), "")
        estFilled = (cEStaff > 0 And Trim$(CStr(ws.Cells(r, cEStaff).value)) <> "")
        
        ' -------------------------------------------------------------
        ' ★判定ロジック修正箇所
        ' -------------------------------------------------------------
        isTraffic = (HitAny(desc, kwTrans) Or HitAny(trans, kwTrans)) ' 交通費系か
        isHotel = HitAny(desc, kwHotel)         ' 宿泊系か
        isOnCall = HitAny(desc, kwOnCall)       ' オンコールか
        isDailyAllo = HitAny(desc, kwDailyAllo) ' ★追加：日当系か（これを計算していませんでした）
        isExcluded = HitAny(desc, kwExcludeTrans) ' NGワードがあるか
        ' -------------------------------------------------------------

        ' --- G：S列（顧客請求費）だけを社員番号で合算 ---
        If cEStaff > 0 And empNo <> "" Then
            estAmt = ParseAmount(ws.Cells(r, cEStaff).value)
            If estAmt <> 0 Then
                If Not HitAny(desc, kwExcludeFromG) Then
                    If Not agg.Exists(key) Then agg.Add key, Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
                    Dim b0: b0 = agg(key)
                    b0(5) = b0(5) + estAmt
                    agg(key) = b0
                    hits = hits + 1
                End If
            End If
        End If

        ' --- 集計 ---
        amt = ParseAmount(ws.Cells(r, cAmt).value)
        If amt = 0 And cFareAmt > 0 Then
            fa = ParseAmount(ws.Cells(r, cFareAmt).value)
            If fa > 0 And isTraffic Then amt = fa
        End If

        If amt <> 0 Then
            If Not agg.Exists(key) Then agg.Add key, Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
            Dim bucket: bucket = agg(key)

            ' 0. 非課税精算(立替金) - 交通機関列が特定の交通費種別の場合
            Dim transKind As String
            If cTrans > 0 Then transKind = NormalizeStr(ws.Cells(r, cTrans).value) Else transKind = ""
            If transKind Like "*交通費*" Then
                bucket(6) = bucket(6) + amt   ' I: 非課税精算(立替金)
                GoTo Decided
            End If
            
            ' 優先順位順に判定
            
            ' 1. オンコール
            If isOnCall Then
                bucket(0) = bucket(0) + amt       ' D: 夜間当番手当
            
            ' 2. テレワーク
            ElseIf HitAny(desc, kwTW) Then
                If Not estFilled Then bucket(4) = bucket(4) + amt       ' J: テレワーク
            
            ' 3. RINK
            ElseIf HitAny(desc, kwRink) Then
                bucket(1) = bucket(1) + amt                             ' E: RINK
            
            ' 4. 夜間当番
            ElseIf HitAny(desc, kwGK) Then
                bucket(0) = bucket(0) + amt                             ' D: 夜間当番
            
            ' 5. ★交通費・宿泊・日当の合算（NGワードがない場合）
            ' ここで isHotel と isDailyAllo を判定に入れます
            ElseIf (isTraffic Or isHotel Or isDailyAllo) And Not isExcluded Then
                If Not estFilled Then bucket(2) = bucket(2) + amt       ' H: 交通費
            
            ' 6. その他
            Else
            ' S列（顧客請求費）に値がない場合のみ計上（二重計上防止）
             If Not estFilled Then bucket(3) = bucket(3) + amt       ' I: その他
            End If
            
Decided:
            agg(key) = bucket
            hits = hits + 1
        End If

        ' --- 計上日最大 ---
        If cBooked > 0 Then
            Dim d As Double: d = TryParseDate(ws.Cells(r, cBooked).value)
            If d > 0 Then DateMax maxDate, key, d
        End If
NextR:
    Next

    Collect_From_Jinjer = hits
End Function
   Private Sub Rewrite_Output_KeepAB(ByVal agg As Object, ByVal maxDate As Object)
    Dim ws As Worksheet
    If Not SheetExists(SH_SUM) Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = SH_SUM
    End If
    Set ws = Worksheets(SH_SUM)

    ' ヘッダー（C～Kだけ作る。A/Bは絶対に消さない）
    ws.Range(ws.Cells(1, 3), ws.Cells(1, 12)).ClearContents
    ws.Cells(1, 3).value = "合計"                    ' C列
    ws.Cells(1, 4).value = "夜間当番手当"         ' D列
    ws.Cells(1, 5).value = "RINK手当"                ' E列
    ws.Cells(1, 6).value = "手当2（夜間＋RINK）"     ' F列
    ws.Cells(1, 7).value = "顧客請求分"              ' G列
    ws.Cells(1, 8).value = "交通費"                  ' H列
    ws.Cells(1, 9).value = "非課税精算(立替金)"      ' I列 ＜ NEW
    ws.Cells(1, 10).value = "その他(会議費・消耗品など)" ' J列
    ws.Cells(1, 11).value = "テレワーク手当"         ' K列
    ws.Cells(1, 12).value = "請求日"                 ' L列
    ws.Columns(12).NumberFormatLocal = "yyyy/m/d"

    ' agg/maxDate を 社員番号ID単位に畳み直し
    Dim byId As Object: Set byId = CreateObject("Scripting.Dictionary")
    Dim dateById As Object: Set dateById = CreateObject("Scripting.Dictionary")
    Dim k, id$, arr, dmax#
    Dim m As Variant, ida As String
    Dim dmaxi As Double

    For Each k In agg.Keys
        id = ParseEmpNo(CStr(k)) ' "ID:123|NM:xxx" → "123"
        If id <> "" Then
            arr = agg(k) ' Array(D,E,H,I,J,G)
            If Not byId.Exists(id) Then byId.Add id, Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
            Dim cur: cur = byId(id)
            cur(0) = cur(0) + arr(0): cur(1) = cur(1) + arr(1): cur(2) = cur(2) + arr(2)
            cur(3) = cur(3) + arr(3): cur(4) = cur(4) + arr(4): cur(5) = cur(5) + arr(5)
            cur(6) = cur(6) + arr(6)
            byId(id) = cur
        End If
    Next

    For Each m In maxDate.Keys
        id = ParseEmpNo(CStr(m))
        If id <> "" Then
            dmaxi = maxDate(m)
            If Not dateById.Exists(id) Then
                dateById.Add id, dmaxi
            ElseIf dmaxi > dateById(id) Then
                dateById(id) = dmaxi
            End If
        End If
    Next

    ' A列の社員番号でC～Kを書き込み（A/Bはノータッチ）
    Dim lastR As Long: lastR = ws.Cells(ws.rows.Count, COL_EMP_NO).End(xlUp).Row
    ' --- 5) R～Y を埋める（A～K に出力済みの値から作る） ---
    
    ' 出力先の列（見出しで探しつつ、見つからなければ既定列にフォールバック）
    Dim cU1&, cU1Amt&, cU2&, cU2Amt&, cCommute&, cCustBillTatekae&, cNonTaxOther&
    cU1 = ColByHeader(ws, "内訳1", 18)                      ' R
    cU1Amt = ColByHeader(ws, "内訳金額1", 19)               ' S
    cU2 = ColByHeader(ws, "内訳2", 20)                      ' T
    cU2Amt = ColByHeader(ws, "内訳金額2", 21)               ' U
    cCommute = ColByHeader(ws, "通勤交通費", 22)            ' V
    cCustBillTatekae = ColByHeader(ws, "顧客請求分(立替金)", 23) ' W
    cNonTaxOther = ColByHeader(ws, "非課税精算(その他)", 25) ' Y
    
    Dim r As Long
    For r = 2 To lastR
        ' A列が空行ならスキップ
        If Trim$(CStr(ws.Cells(r, 1).value)) <> "" Then
            Dim vD As Double, vE As Double, vF As Double, vG As Double, vH As Double, vI As Double, vJ As Double
            vD = ValJP(ws.Cells(r, 4).Value2) ' D: 顧客当番対応手当
            vE = ValJP(ws.Cells(r, 5).Value2) ' E: RINK手当
            vF = ValJP(ws.Cells(r, COL_ALLOW2).Value2) ' F: 手当2（顧客+RINK）
            vG = ValJP(ws.Cells(r, 7).Value2) ' G: 顧客請求分
            vH = ValJP(ws.Cells(r, COL_TRANS).Value2) ' H: 交通費
            vI = ValJP(ws.Cells(r, COL_ETC).Value2) ' J: その他
            vJ = ValJP(ws.Cells(r, COL_TW).Value2) ' K: テレワーク手当
        End If
    Next r

    Dim Z As Long
    For Z = 2 To lastR
    Dim empNo As String: empNo = NormalizeId(ws.Cells(Z, 1).value)
    
    ' byIdに存在する場合のみ、D～K列を書き込む
    If empNo <> "" And byId.Exists(empNo) Then
        arr = byId(empNo)
        ws.Cells(Z, 4).value = arr(0)  ' D: 夜間当番手当
        ws.Cells(Z, 5).value = arr(1)  ' E: RINK手当
        ws.Cells(Z, COL_TRANS).value = arr(2)           ' H: 交通費
        ws.Cells(Z, COL_NONTAX_TATEKAE).value = arr(6)  ' I: 非課税精算(立替金) ＜ NEW
        ws.Cells(Z, COL_ETC).value = arr(3)             ' J: その他
        ws.Cells(Z, COL_TW).value = arr(4)              ' K: テレワーク手当
        ws.Cells(Z, COL_BILL).value = arr(5)            ' G: 客先請求
        If dateById.Exists(empNo) Then ws.Cells(Z, COL_DATE).value = dateById(empNo)  ' L: 請求日
    End If
    
    ' ★F列とC列の計算は条件の外に出す（A列に社員番号があれば計算する）
    If empNo <> "" Then
        ws.Cells(Z, COL_ALLOW2).value = Nz(ws.Cells(Z, COL_GK).value) + Nz(ws.Cells(Z, COL_RINK).value)  ' F: 手当2（D+E）
        ws.Cells(Z, COL_TOTAL).value = Nz(ws.Cells(Z, 6).value) + Nz(ws.Cells(Z, 7).value) + _
                                Nz(ws.Cells(Z, 8).value) + Nz(ws.Cells(Z, 9).value) + _
                                Nz(ws.Cells(Z, 10).value) + Nz(ws.Cells(Z, 11).value)  ' C: 合計（F+G+H+I+J+K）
    End If
Next Z
End Sub

' ===== キーから社員番号(ID)だけ抜く =====
' 期待形式: "ID:12345|NM:やつ" または "NM:やつ"（この場合は空を返す）
Public Function ParseEmpNo(ByVal key As String) As String
    Dim p As Long, q As Long, id As String
    p = InStr(1, key, "ID:")
    If p > 0 Then
        q = InStr(p + 3, key, "|NM:")
        If q > 0 Then
            id = Mid$(key, p + 3, q - (p + 3))
        Else
            id = Mid$(key, p + 3)
        End If
        ParseEmpNo = NormalizeId(id)
    Else
        ParseEmpNo = ""
    End If
End Function
            
            ' ===== ヘルパー =====
            Private Function FindColLike(ws As Worksheet, ByVal key As String) As Long
                Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                Dim c As Long, h As String
                For c = 1 To lastCol
                    h = CStr(ws.Cells(1, c).value)
                    If Len(h) > 0 Then
                        If InStr(h, key) > 0 Then FindColLike = c: Exit Function
                    End If
                Next
            End Function
            
            ' 「\」「円」「,」「全角数字」「()のマイナス」などを吸収して数値化
            Private Function ValJP(ByVal v As Variant) As Double
                If IsError(v) Or isEmpty(v) Then Exit Function
                Dim s As String: s = CStr(v)
                s = Trim$(s)
                s = StrConv(s, vbNarrow)          ' 全角→半角
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
                   
                       
             
        
        ' ========= バックアップ（最初に必ずコピー） =========
        Private Sub BackupSheet(ByVal shtName As String)
            If Not SheetExists(shtName) Then Exit Sub
            Dim src As Worksheet: Set src = Worksheets(shtName)
            Dim nm As String: nm = shtName & "_backup_" & Format(Now, "yyyymmdd_HHMMSS")
            src.Copy After:=src
            ActiveSheet.Name = nm
        End Sub
        
        ' ========= 小物関数 =========
        Private Function SheetExists(ByVal nm As String) As Boolean
            Dim ws As Worksheet
            On Error Resume Next
            Set ws = Worksheets(nm)
            SheetExists = Not ws Is Nothing
            On Error GoTo 0
        End Function
        
        Private Function FindCol(ws As Worksheet, headerRow As Long, names As Variant, Optional allowMissing As Boolean = False) As Long
            Dim lastC&, c&, h$, i&
            lastC = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
            For c = 1 To lastC
                h = NormalizeStr(ws.Cells(headerRow, c).value)
                For i = LBound(names) To UBound(names)
                    If h Like "*" & NormalizeStr(CStr(names(i))) & "*" Then
                        FindCol = c: Exit Function
                    End If
                Next
            Next
            If allowMissing Then FindCol = 0 Else FindCol = 0
        End Function
        
        Private Function NormalizeStr(ByVal s As String) As String
            s = Trim$(LCase$(Replace(Replace(Replace(Replace(CStr(s), vbCr, " "), vbLf, " "), vbTab, " "), "　", " ")))
            s = Replace(s, " ", "")
            NormalizeStr = s
        End Function
        
        Private Function NormalizeId(v) As String
            Dim s$, i&, ch$, out$
            s = CStr(v)
            For i = 1 To Len(s)
                ch = Mid$(s, i, 1)
                If ch >= "0" And ch <= "9" Then out = out & ch
            Next
            NormalizeId = out
        End Function
        
        Private Function NormalizeName(v) As String
            Dim s$
            s = CStr(v)
            s = Replace(s, "　", "")
            s = Replace(s, " ", "")
            NormalizeName = LCase$(s)
        End Function
        
        Private Function BuildEmpKey(ByVal empNo As String, ByVal empNm As String) As String
            If empNo <> "" Then BuildEmpKey = "ID:" & empNo & "|NM:" & empNm Else BuildEmpKey = "NM:" & empNm
        End Function
        
        Private Function ParseAmount(v) As Double
            Dim s$: s = CStr(v)
            s = Replace(s, ",", "")
            s = Replace(s, "￥", "")
            s = Replace(s, "円", "")
            s = Replace(s, "(", "-")
            s = Replace(s, ")", "")
            If IsNumeric(s) Then ParseAmount = CDbl(s) Else ParseAmount = 0#
        End Function
        
        Private Function HitAny(ByVal s As String, kw As Variant) As Boolean
            Dim i&
            For i = LBound(kw) To UBound(kw)
                If InStr(1, s, NormalizeStr(CStr(kw(i))), vbTextCompare) > 0 Then HitAny = True: Exit Function
            Next
        End Function
        
        Private Sub DateMax(dict As Object, key As String, ByVal d As Double)
            If Not dict.Exists(key) Then dict.Add key, d Else If d > dict(key) Then dict(key) = d
        End Sub
        
        Private Function Nz(v) As Double
            If IsNumeric(v) Then Nz = CDbl(v) Else Nz = 0#
        End Function
        
        Private Function NzDate(dict As Object, key As String) As Double
            If dict.Exists(key) Then NzDate = dict(key) Else NzDate = 0#
        End Function
        
        Private Function TryParseDate(v) As Double
            If IsDate(v) Then TryParseDate = CDbl(CDate(v)): Exit Function
            Dim s$: s = CStr(v)
            s = Replace(s, "年", "/"): s = Replace(s, "月", "/"): s = Replace(s, "日", "")
            s = Replace(s, "-", "/")
            If IsDate(s) Then TryParseDate = CDbl(CDate(s)) Else TryParseDate = 0#
        End Function
        
        
' （必要なら）名前側を取りたいときに使う補助。今は必須じゃないけど置いておくと便利。
Public Function ParseEmpName(ByVal key As String) As String
    Dim p As Long
    p = InStr(1, key, "|NM:")
    If p > 0 Then
        ParseEmpName = Mid$(key, p + 4)
    Else
        ParseEmpName = ""
    End If
End Function

Private Function ToDbl(ByVal v As Variant) As Double
    If IsError(v) Or isEmpty(v) Or v = "" Then Exit Function
    On Error Resume Next
    ToDbl = CDbl(v) ' 失敗したら0のまま
    On Error GoTo 0
End Function
'=== 標準モジュールの末尾に追加（どのSubの外！） ===
Private Function ColByHeader(ws As Worksheet, ByVal key As String, ByVal fallback As Long) As Long
    Dim c As Long: c = FindColLike(ws, key)
    If c > 0 Then
        ColByHeader = c
    Else
        ColByHeader = fallback   ' 見つからなければ既定列にフォールバック
    End If
End Function

' 見出しの部分一致で列番号を返す（1行目を探索）
Private Function FindColLikeA(ws As Worksheet, ByVal key As String) As Long
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim i As Long, h As String
    For i = 1 To lastCol
        h = CStr(ws.Cells(1, i).value)
        If Len(h) > 0 Then
            If InStr(h, key) > 0 Then FindColLikeA = i: Exit Function
        End If
    Next
End Function

' 金額/通貨表記を数値化（未定義なら入れて）
Private Function ValJPA(ByVal v As Variant) As Double
    If IsError(v) Or isEmpty(v) Or v = "" Then Exit Function
    Dim s As String: s = CStr(v)
    s = StrConv(s, vbNarrow)  ' 全角→半角
    s = Replace(s, "\", "")
    s = Replace(s, "円", "")
    s = Replace(s, ",", "")
    s = Replace(s, "　", "")
    s = Trim$(s)
    If Len(s) >= 2 And Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        s = "-" & Mid$(s, 2, Len(s) - 2)
    End If
    If IsNumeric(s) Then ValJPA = CDbl(s)
End Function

  Sub Distribute_To_RY_Columns()
    Dim ws As Worksheet
    Set ws = Worksheets(SH_SUM)
    
    Dim lastR As Long
    lastR = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastR
        ' A列が空行ならスキップ
        If Trim$(CStr(ws.Cells(r, 1).value)) <> "" Then
            Dim vF As Double, vJ As Double, vH As Double, vG As Double, vI As Double
            vF = ValJP(ws.Cells(r, 6).Value2) ' F: 手当2（顧客+RINK）
            vJ = ValJP(ws.Cells(r, 10).Value2) ' J: テレワーク手当
            vH = ValJP(ws.Cells(r, 8).Value2) ' H: 交通費
            vG = ValJP(ws.Cells(r, 7).Value2) ' G: 顧客請求分
            vI = ValJP(ws.Cells(r, 9).Value2) ' I: その他
            
            ' R/S: 手当2(F列)に値が入ってたら「夜間当番手当」＋金額
            If vF <> 0 Then
                ws.Cells(r, 18).value = "夜間当番手当"  ' R列
                ws.Cells(r, 19).value = vF                  ' S列
            End If
            
            ' T/U: テレワーク手当(J列)に値が入ってたら「テレワーク手当」＋金額
            If vJ <> 0 Then
                ws.Cells(r, 20).value = "テレワーク手当"   ' T列
                ws.Cells(r, 21).value = vJ                  ' U列
            End If
            
            ' V: 交通費(H列)に値が入ってたらその値を記載
            If vH <> 0 Then
                ws.Cells(r, 22).value = vH                  ' V列
            End If
            
            ' W: 顧客請求分(G列)に値が入ってたらその値を記載
            If vG <> 0 Then
                ws.Cells(r, 23).value = vG                  ' W列
            End If
            
            ' X: 非課税精算(立替金) は記載なし
            ' ws.Cells(r, 24).value = ""                    ' X列（何もしない）
            
            ' Y: その他(I列)に値が入ってたらその値を記載
            If vI <> 0 Then
                ws.Cells(r, 25).value = vI                  ' Y列
            End If
        End If
    Next r
    
    MsgBox "R～Y列への振り分けが完了しました。", vbInformation
End Sub

