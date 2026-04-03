Attribute VB_Name = "Module1"
Sub SendMassMail_WithCC()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim logWs As Worksheet
    Dim i As Long
    Dim subjectText As String
    Dim baseBody As String
    Dim empName As String
    Dim mailBody As String
    Dim ccAddress As String
    Dim lastRow As Long
    Dim logRow As Long
    Dim useBCCCheck As Boolean
    Dim bccAddress As String
    Dim mailCount As Long
    
    Set ws = ThisWorkbook.Sheets("メール送信")
    Set OutApp = CreateObject("Outlook.Application")
    
    subjectText = ws.Range("B1").Value
    baseBody = ws.Range("B2").Value
    ccAddress = ws.Range("D1").Value
    useBCCCheck = (ws.Range("F3").Value = True)
    
    On Error Resume Next
    Set logWs = ThisWorkbook.Sheets("MailLog")
    If logWs Is Nothing Then
        Set logWs = ThisWorkbook.Sheets.Add(After:=ws)
        logWs.Name = "MailLog"
        logWs.Range("A1:C1").Value = Array("送信日時", "氏名", "メールアドレス")
    End If
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    mailCount = 0
    For i = 4 To lastRow
        If ws.Cells(i, "A").Value = True Or ws.Cells(i, "A").Value = "TRUE" Then
            empName = ws.Cells(i, "C").Value
            
            mailBody = Replace(baseBody, "[対象者名]", empName)
            mailBody = empName & "さん" & vbCrLf & vbCrLf & mailBody
            
            ' BCC
            bccAddress = ""
            If useBCCCheck Then
                If Trim(ws.Cells(i, "E").Value) <> "" Then
                    bccAddress = ws.Cells(i, "E").Value
                End If
                If Trim(ws.Cells(i, "F").Value) <> "" Then
                    If bccAddress <> "" Then bccAddress = bccAddress & "; "
                    bccAddress = bccAddress & ws.Cells(i, "F").Value
                End If
            End If
            
            mailCount = mailCount + 1
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = ws.Cells(i, "D").Value
                .CC = ccAddress
                If bccAddress <> "" Then .BCC = bccAddress
                .Subject = subjectText
                .Body = mailBody
                .Importance = 2
                If useBCCCheck And mailCount > 1 Then
                    .Send
                Else
                    .Display
                End If
            End With
            
            logRow = logWs.Cells(logWs.Rows.Count, "A").End(xlUp).Row + 1
            logWs.Cells(logRow, "A").Value = Now
            logWs.Cells(logRow, "B").Value = empName
            logWs.Cells(logRow, "C").Value = ws.Cells(i, "D").Value
        End If
    Next i
    
    MsgBox "メール作成（CC付き）＆ログ記録が完了しました。", vbInformation
End Sub

Sub SelectTemplate()
    Dim wsMail As Worksheet
    Dim wsTemplate As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim templateList As String
    Dim templateCount As Long
    Dim userInput As String
    Dim selectedCol As Long
    
    Set wsMail = ThisWorkbook.Sheets("メール送信")
    Set wsTemplate = ThisWorkbook.Sheets("メールテンプレート")
    
    lastCol = wsTemplate.Cells(1, wsTemplate.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then
        MsgBox "該当するテンプレートがありません。", vbExclamation
        Exit Sub
    End If
    
    templateList = ""
    templateCount = 0
    For col = 2 To lastCol
        If Trim(wsTemplate.Cells(1, col).Value) <> "" Then
            templateCount = templateCount + 1
            templateList = templateList & templateCount & ": " & wsTemplate.Cells(1, col).Value & vbCrLf
        End If
    Next col
    
    If templateCount = 0 Then
        MsgBox "該当するテンプレートがありません。", vbExclamation
        Exit Sub
    End If
    
    userInput = InputBox(templateList & vbCrLf & "番号で選択してください", "テンプレート選択")
    If userInput = "" Then Exit Sub
    
    If Not IsNumeric(userInput) Then Exit Sub
    If CLng(userInput) < 1 Or CLng(userInput) > templateCount Then Exit Sub
    
    selectedCol = CLng(userInput) + 1
    wsMail.Range("B1").Value = wsTemplate.Cells(2, selectedCol).Value
    wsMail.Range("B2").Value = wsTemplate.Cells(3, selectedCol).Value
    
    MsgBox "件名と本文をセットしました。", vbInformation
End Sub
