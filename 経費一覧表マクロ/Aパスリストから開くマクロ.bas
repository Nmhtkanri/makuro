Attribute VB_Name = "Aパスリストから開くマクロ"
Option Explicit

' ====== 64/32bit 両対応の WNetGetConnection 宣言 ======
#If VBA7 Then
    Private Declare PtrSafe Function WNetGetConnectionA Lib "mpr.dll" ( _
        ByVal lpszLocalName As String, _
        ByVal lpszRemoteName As String, _
        ByRef lpnLength As Long) As Long
#Else
    Private Declare Function WNetGetConnectionA Lib "mpr.dll" ( _
        ByVal lpszLocalName As String, _
        ByVal lpszRemoteName As String, _
        ByRef lpnLength As Long) As Long
#End If

' ========== 公開エントリ ==========
' アクティブ行の C列パスを開く
Public Sub OpenReport_CurrentRow()
    Dim r As Long
    r = ActiveCell.Row
    OpenReport_FromRow r
End Sub

' 社員番号（A列）で行を探して C列パスを開く
Public Sub OpenReport_ByEmpNo(ByVal empNo As Variant)
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If ws.Cells(r, 1).value = empNo Then
            OpenReport_FromRow r
            Exit Sub
        End If
    Next
    MsgBox "社員番号 " & empNo & " の行が見つかりませんでした。", vbExclamation
End Sub

' ========== 基本処理 ==========
' 指定行の C列にあるパスを開く（ファイルを想定）
Public Sub OpenReport_FromRow(ByVal r As Long)
    On Error GoTo ErrHandler
    Dim ws As Worksheet, rawPath As String, norm As String, unc As String
    Dim target As String

    Set ws = ActiveSheet
    rawPath = CStr(ws.Cells(r, 3).value)  ' C列
    If Len(Trim$(rawPath)) = 0 Then
        MsgBox "C列にパスがありません（行 " & r & "）。", vbExclamation
        Exit Sub
    End If

    norm = CleanPath(rawPath)
    ' 既にUNCならそのまま、ドライブならUNC化を試す
    If IsUNCPath(norm) Then
        target = norm
    Else
        unc = DrivePathToUNC(norm)
        target = IIf(Len(unc) > 0, unc, norm)
    End If

    ' 存在チェック（ファイル想定）
    If Not FileExistsStrong(target) Then
        ' 代替：親フォルダを開く
        Dim parentFolder As String
        parentFolder = GetParentFolder(target)
        MsgBox "ファイルにアクセスできませんでした。" & vbCrLf & _
               "元: " & rawPath & vbCrLf & _
               "正規化: " & norm & vbCrLf & _
               "解決パス: " & target & vbCrLf & vbCrLf & _
               "親フォルダを開きますので、手動で確認してください。", vbExclamation
        If Len(parentFolder) > 0 Then
            Shell "explorer.exe " & """" & parentFolder & """", vbNormalFocus
        End If
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Workbooks.Open fileName:=target, ReadOnly:=False, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "オープン時エラー: " & Err.Number & " - " & Err.Description & vbCrLf & "パス: " & target, vbCritical
End Sub

' ========== ユーティリティ ==========
Private Function CleanPath(ByVal p As String) As String
    Dim s As String
    s = Trim$(p)
    ' 全角￥を半角\へ、全角スペースを半角へ（最低限）
    s = Replace(s, "￥", "\")
    s = Replace(s, ChrW(&H3000), " ")
    ' 連続ダブルクォート除去
    If Left$(s, 1) = """" And Right$(s, 1) = """" Then
        s = Mid$(s, 2, Len(s) - 2)
    End If
    CleanPath = s
End Function

Private Function IsUNCPath(ByVal p As String) As Boolean
    IsUNCPath = (Len(p) >= 2 And Left$(p, 2) = "\\")
End Function

' U:\folder\file → \\server\share\folder\file に解決（失敗時は ""）
Private Function DrivePathToUNC(ByVal p As String) As String
    On Error GoTo EH
    Dim drv As String, rest As String
    Dim remote As String * 260
    Dim n As Long, rc As Long

    If Len(p) < 2 Or Mid$(p, 2, 1) <> ":" Then
        DrivePathToUNC = ""
        Exit Function
    End If

    drv = Left$(p, 2)               ' "U:"
    rest = Mid$(p, 3)               ' "\NetMarks\FY2025\..."
    n = 260
    rc = WNetGetConnectionA(drv, remote, n)
    If rc = 0 Then
        ' remote に "\\server\share" が返る
        DrivePathToUNC = TrimNull(remote) & rest
    Else
        DrivePathToUNC = ""         ' 変換失敗
    End If
    Exit Function
EH:
    DrivePathToUNC = ""
End Function

Private Function TrimNull(ByVal s As String) As String
    Dim Z As Long
    Z = InStr(1, s, Chr$(0))
    If Z > 0 Then
        TrimNull = Left$(s, Z - 1)
    Else
        TrimNull = s
    End If
End Function

Private Function FileExistsStrong(ByVal filePath As String) As Boolean
    On Error Resume Next
    If Len(filePath) = 0 Then Exit Function
    FileExistsStrong = (Dir$(filePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem) <> "")
End Function

Private Function GetParentFolder(ByVal p As String) As String
    Dim i As Long
    For i = Len(p) To 1 Step -1
        If Mid$(p, i, 1) = "\" Or Mid$(p, i, 1) = "/" Then
            GetParentFolder = Left$(p, i - 1)
            Exit Function
        End If
    Next
    GetParentFolder = ""
End Function
' ===== フォルダをエクスプローラで開く系 =====

' アクティブ行のC列パスのフォルダを開く
Public Sub OpenFolder_CurrentRow()
    Dim r As Long
    r = ActiveCell.Row
    OpenFolder_FromRow r
End Sub

' A列=社員番号で行を探して、その行のC列パスのフォルダを開く
Public Sub OpenFolder_ByEmpNo(ByVal empNo As Variant)
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If ws.Cells(r, 1).value = empNo Then
            OpenFolder_FromRow r
            Exit Sub
        End If
    Next
    MsgBox "社員番号 " & empNo & " の行が見つかりませんでした。", vbExclamation
End Sub

' 指定行のC列にあるパスの「フォルダ」を開く
Public Sub OpenFolder_FromRow(ByVal r As Long)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim rawPath As String
    rawPath = CStr(ws.Cells(r, 3).value) ' C列
    If Len(Trim$(rawPath)) = 0 Then
        MsgBox "C列にパスがありません（行 " & r & "）。", vbExclamation
        Exit Sub
    End If
    OpenFolder_FromPath rawPath
End Sub

' 文字列で渡したパスの「フォルダ」を開く
' ・フォルダパスならそのまま開く
' ・ファイルパスなら親フォルダを開く
Public Sub OpenFolder_FromPath(ByVal anyPath As String)
    On Error GoTo ErrHandler
    Dim norm As String, target As String, unc As String, folder As String

    norm = CleanPath(anyPath)
    ' ファイルが来たら親フォルダに
    folder = EnsureFolderPath(norm)

    ' 既にUNCならそのまま、ドライブならUNC変換
    If IsUNCPath(folder) Then
        target = folder
    Else
        unc = DrivePathToUNC(folder)
        target = IIf(Len(unc) > 0, unc, folder)
    End If

    ' フォルダ存在チェック
    If Not FolderExistsStrong(target) Then
        MsgBox "フォルダにアクセスできませんでした。" & vbCrLf & _
               "元: " & anyPath & vbCrLf & _
               "正規化: " & norm & vbCrLf & _
               "解決フォルダ: " & target, vbExclamation
        Exit Sub
    End If

    Shell "explorer.exe " & """" & target & """", vbNormalFocus
    Exit Sub

ErrHandler:
    MsgBox "フォルダオープン時エラー: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "入力パス: " & anyPath, vbCritical
End Sub

' ===== ヘルパー（前のモジュールに無ければ追加して） =====

' フォルダの実在確認（UNC/ローカル両対応）
Private Function FolderExistsStrong(ByVal folderPath As String) As Boolean
    On Error GoTo EH
    Dim p As String
    p = RTrim$(folderPath)
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Len(p) = 0 Then Exit Function

    ' Dirで存在チェック
    Dim ret As String
    ret = Dir$(p, vbDirectory)
    If ret = "" Then
        FolderExistsStrong = False
    Else
        ' 属性でフォルダ判定
        FolderExistsStrong = ((GetAttr(p) And vbDirectory) = vbDirectory)
    End If
    Exit Function
EH:
    FolderExistsStrong = False
End Function

' 渡された文字列がファイルっぽかったら親フォルダに変換
Private Function EnsureFolderPath(ByVal p As String) As String
    Dim s As String: s = Trim$(p)
    If Len(s) = 0 Then Exit Function

    ' 末尾が \ で終わっていればフォルダとみなす
    If Right$(s, 1) = "\" Then
        EnsureFolderPath = Left$(s, Len(s) - 1)
        Exit Function
    End If

    ' 拡張子らしきもの or 実在するファイルなら親フォルダへ
    Dim looksLikeFile As Boolean
    looksLikeFile = (InStrRev(s, ".") > InStrRev(s, "\")) ' 最後の\より後に.がある
    If looksLikeFile Or FileExistsStrong(s) Then
        EnsureFolderPath = GetParentFolder(s)
    Else
        EnsureFolderPath = s
    End If
End Function


