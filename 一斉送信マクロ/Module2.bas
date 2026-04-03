Attribute VB_Name = "Module2"
Sub AutoCheck_And_SendMail()
    Dim wsMail As Worksheet
    Dim wsCSV As Worksheet
    Dim OutApp As Object
    Dim OutMail As Object
    Dim logWs As Worksheet
    Dim lastMailRow As Long, lastCSVRow As Long
    Dim i As Long, j As Long
    Dim empNoMail As String, empNoCSV As String
    Dim subjectText As String, baseBody As String, empName As String, mailBody As String
    Dim ccAddress As String
    Dim additionalCC As String
    Dim finalCC As String
    Dim foundMatch As Boolean
    Dim logRow As Long
    Dim useE3Check As Boolean
    Dim skipCSVCheck As Boolean
    Dim seisanAmount As String
    Dim useBCCCheck As Boolean
    Dim bccAddress As String
    Dim mailCount As Long
    
    Set wsMail = ThisWorkbook.Sheets("メール送信")
    Set wsCSV = ThisWorkbook.Sheets("一斉送信LOG")
    Set OutApp = CreateObject("Outlook.Application")
    
    subjectText = wsMail.Range("B1").Value
    baseBody = wsMail.Range("B2").Value
    ccAddress = wsMail.Range("D1").Value
    
    useE3Check = (wsMail.Range("E3").Value = True)
    useBCCCheck = (wsMail.Range("F3").Value = True)
    
    If Trim(wsCSV.Range("A1").Value) = "" And Trim(wsCSV.Range("B1").Value) = "" Then
        skipCSVCheck = True
    Else
        skipCSVCheck = False
        wsMail.Range("A4:A" & wsMail.Cells(wsMail.Rows.Count, "B").End(xlUp).Row).ClearContents
        
        lastMailRow = wsMail.Cells(wsMail.Rows.Count, "B").End(xlUp).Row
        lastCSVRow = wsCSV.Cells(wsCSV.Rows.Count, "A").End(xlUp).Row
        
        For i = 4 To lastMailRow
            empNoMail = Trim(wsMail.Cells(i, "B").Value)
            For j = 2 To lastCSVRow
                empNoCSV = Trim(wsCSV.Cells(j, "A").Value)
                If empNoMail <> "" And empNoMail = empNoCSV Then
                    wsMail.Cells(i, "A").Value = True
                    Exit For
                End If
            Next j
        Next i
    End If
    
    On Error Resume Next
    Set logWs = ThisWorkbook.Sheets("MailLog")
    If logWs Is Nothing Then
        Set logWs = ThisWorkbook.Sheets.Add(After:=wsMail)
        logWs.Name = "MailLog"
        logWs.Range("A1:C1").Value = Array("送信日時", "氏名", "メールアドレス")
    End If
    On Error GoTo 0
    
    lastMailRow = wsMail.Cells(wsMail.Rows.Count, "B").End(xlUp).Row
    If Not skipCSVCheck Then
        lastCSVRow = wsCSV.Cells(wsCSV.Rows.Count, "A").End(xlUp).Row
    End If
    
    mailCount = 0
    For i = 4 To lastMailRow
        If wsMail.Cells(i, "A").Value = True Then
            empName = wsMail.Cells(i, "C").Value
            
            ' 精算額 lookup from LOG
            seisanAmount = ""
            If Not skipCSVCheck Then
                empNoMail = Trim(CStr(wsMail.Cells(i, "B").Value))
                For j = 2 To lastCSVRow
                    If Trim(CStr(wsCSV.Cells(j, "A").Value)) = empNoMail Then
                        seisanAmount = CStr(wsCSV.Cells(j, "C").Value)
                        Exit For
                    End If
                Next j
            End If
            
            mailBody = Replace(baseBody, "[対象者名]", empName)
            mailBody = Replace(mailBody, "[精算額]", seisanAmount)
            mailBody = empName & "さん" & vbCrLf & vbCrLf & mailBody
            
            ' CC
            If useE3Check Then
                additionalCC = Trim(wsMail.Cells(i, "E").Value)
                If additionalCC <> "" Then
                    If ccAddress <> "" Then
                        finalCC = ccAddress & "; " & additionalCC
                    Else
                        finalCC = additionalCC
                    End If
                Else
                    finalCC = ccAddress
                End If
            Else
                finalCC = ccAddress
            End If
            
            ' BCC
            bccAddress = ""
            If useBCCCheck Then
                If Trim(wsMail.Cells(i, "E").Value) <> "" Then
                    bccAddress = wsMail.Cells(i, "E").Value
                End If
                If Trim(wsMail.Cells(i, "F").Value) <> "" Then
                    If bccAddress <> "" Then bccAddress = bccAddress & "; "
                    bccAddress = bccAddress & wsMail.Cells(i, "F").Value
                End If
            End If
            
            mailCount = mailCount + 1
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = wsMail.Cells(i, "D").Value
                .CC = finalCC
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
            logWs.Cells(logRow, "C").Value = wsMail.Cells(i, "D").Value
        End If
    Next i
    
    If skipCSVCheck Then
        MsgBox "手動チェック時のメール作成を完了しました。", vbInformation
    Else
        MsgBox "対象者を自動チェックし、メール作成を完了しました。", vbInformation
    End If
End Sub
