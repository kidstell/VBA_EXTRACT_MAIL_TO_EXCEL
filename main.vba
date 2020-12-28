Const olFolderInbox As Integer = 6
Dim strFilter As String
Dim ccc As Integer


Sub UpdateCapacityReportViaEmail()
    Dim dataSheetName
    Dim entrycount As Integer, remdatelength As Integer, idx As Integer, nIdx As Integer, iMaxTps As Integer
    Dim lastrowdate As Date, lastcomputedate As Date, lastdate As Date, finaldate As Date, latestmaildate As Date, nextdatetarget As Date, datesFields() As Date, date2date As Date
    Dim valueFields() As Integer, dayscount As Integer
    Dim finaldatelng As Long, latestmaildatelng As Long, playpos As Long
    Dim cell As Range
    Dim d7ave() As Variant, d7aveTrap() As Variant, usageratio() As Variant
    
    Dim dateset() As Variant, scoreset() As Variant
    
    dataSheetName = "SERVER TPS"
    Sheets(dataSheetName).Select
    sSubFolder = Sheets(dataSheetName).Range("G5")
    iMaxTps = Sheets(dataSheetName).Range("G3")
    
    If (iMaxTps = 0) Then iMaxTps = 1

    Dim oOlAp As Object, oOlns As Object, oOlInb As Object
    Dim oOlItm As Object
    Dim oOlMails As Object, OutAttch As Object, datafile As Object

    '~~> Outlook Variables for email
    Dim eSender As String, dtRecvd As String, dtSent As String
    Dim sSubj As String, sMsg As String
    
    Sheets(dataSheetName).Columns(1).Range("M1") = "Status: Running"

    '~~> Get Outlook instance
    Set oOlAp = GetObject(, "Outlook.application")
    Set oOlns = oOlAp.GetNamespace("MAPI")
    Set oOlInb = oOlns.GetDefaultFolder(olFolderInbox)
    
    '~~> Path for the attachment
    Dim OutputFolder As String
    OutputFolder = "D:\"
    'OutputFolder = ""

    '~~> New File Name for the attachment
    Dim NewFileName As String, datafilename As String
    NewFileName = OutputFolder & Format(Date, "DD-MM-YYYY") & "-"
    
    
    'if(cell(a5) is notempty then msgbox "you have specified the subfolder(cella5) at cella5" else msgbox "your inbox root will be search because you did not specify a subfolder at cella5"
    If (IsEmpty(sSubFolder) Or sSubFolder = "") Then
        MsgBox "your inbox root will be searched because you did not specify a subfolder at cell 'G5'"
        Set oOlInb = oOlns.GetDefaultFolder(olFolderInbox)
    Else
        MsgBox "you have specified the subfolder(" & sSubFolder & ") at cell 'G5'"
        Set oOlInb = oOlns.GetDefaultFolder(olFolderInbox).Folders(sSubFolder)
    End If
    
    
    entrycount = Sheets(dataSheetName).Cells(Sheets(dataSheetName).Rows.Count, "A").End(xlUp).Row
    lastrowdate = Sheets(dataSheetName).Columns(1).Cells(entrycount, 1).Value
    lastcomputedate = Sheets(dataSheetName).Range("G6")
    
    If (Not (IsEmpty(lastcomputedate)) And lastcomputedate > 0) Then
        lastdate = lastcomputedate
    Else
        lastdate = lastrowdate
    End If
    
    'strFilter = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " like '%HSDP PLATFORM MONITORING%'"
    strFilter = "[Subject] = ""[SERVER DEAMON] Routine Stats Report SERVER_1"" And [ReceivedTime] > '" & Format(lastdate, "dd/mm/yyyy") & "'"
    
    Set oOlMails = oOlInb.Items.Restrict(strFilter)
    oOlMails.Sort "[ReceivedTime]", True
    
    finaldate = Date 'Now()
    finaldatelng = CLng(Int(finaldate))
    
    dayscount = finaldatelng - CLng(Int(lastdate))
    ReDim valueFields(0 To (dayscount - 1))
    ReDim datesFields(0 To (dayscount - 1))
    ReDim usageratio(0 To (dayscount - 1))
    ReDim d7ave(0 To (dayscount - 1))
    latestmaildatelng = 0
    nextdatetarget = 0
    
    For Each oOlItm In oOlMails
        'eSender = oOlItm.SenderEmailAddress
        dtRecvd = oOlItm.ReceivedTime
        'dtSent = oOlItm.CreationTime
        'sSubj = oOlItm.Subject
        'sMsg = oOlItm.Body
        
        If (oOlItm.Attachments.Count <> 0) Then
            If (latestmaildatelng = 0) Then
                latestmaildate = CDate(dtRecvd)
                latestmaildatelng = CLng(Int(latestmaildate))
                nextdatetarget = latestmaildate 'CDate(dtRecvd)
            End If
            
            If (nextdatetarget > CDate(dtRecvd)) Then nextdatetarget = CDate(dtRecvd)
            
            If (nextdatetarget = CDate(dtRecvd)) Then
                'save this attachment and skip 7 days backwards if need be
                NewFileName = OutputFolder & Format(nextdatetarget, "DD-MM-YYYY") & "-"
                Set OutAttch = oOlItm.Attachments(1)
                datafilename = NewFileName & OutAttch.Filename
                If (Dir(datafilename) = "") Then OutAttch.SaveAsFile datafilename
                'open it and read it
                Application.ScreenUpdating = False
                Set datafile = Workbooks.Open(datafilename, True, True)
                dateset = datafile.Worksheets("sheet1_1").Columns(2).Range(Cells(27, 1), Cells(33, 1)).Value
                scoreset = datafile.Worksheets("sheet1_1").Columns(9).Range(Cells(27, 1), Cells(33, 1)).Value
                datafile.Close False
                Kill datafilename
                For idx = 1 To 7
                    date2date = Format(dateset(idx, 1), "DD/MM/YYYY")
                    If (date2date >= lastdate) Then
                        playpos = CLng(Int(date2date)) - lastdate
                        valueFields(playpos) = scoreset(idx, 1)
                        datesFields(playpos) = date2date
                    End If
                Next idx
                Application.ScreenUpdating = True
                nextdatetarget = nextdatetarget - 5
            End If
            If (nextdatetarget < lastdate) Then Exit For
        End If
    Next

    Set cell = Sheets(dataSheetName).Columns(1).Find(What:=lastdate)
    If (cell Is Nothing) Then
        nIdx = entrycount + 1
    Else
        nIdx = cell.Row
    End If
    
    usageratio = Sheets(dataSheetName).Columns(3).Range(Cells(1, 1), Cells(nIdx + dayscount, 1)).Value
    d7ave = Sheets(dataSheetName).Columns(4).Range(Cells(1, 1), Cells(nIdx + dayscount, 1)).Value
    For idx = 0 To (dayscount - 1)
        If (datesFields(idx) <> 0) Then
            usageratio(nIdx + idx, 1) = (valueFields(idx) / iMaxTps) * 100
            d7ave(nIdx + idx, 1) = Application.WorksheetFunction.Average(usageratio(nIdx + idx, 1), usageratio(nIdx + idx - 1, 1), usageratio(nIdx + idx - 2, 1), usageratio(nIdx + idx - 3, 1), usageratio(nIdx + idx - 4, 1), usageratio(nIdx + idx - 5, 1), usageratio(nIdx + idx - 6, 1))
            
            Sheets(dataSheetName).Columns(1).Range(Cells(nIdx + idx, 1), Cells(nIdx + idx, 1)).Value = datesFields(idx)
            Sheets(dataSheetName).Columns(2).Range(Cells(nIdx + idx, 1), Cells(nIdx + idx, 1)).Value = valueFields(idx)
            Sheets(dataSheetName).Columns(3).Range(Cells(nIdx + idx, 1), Cells(nIdx + idx, 1)).Value = Format(usageratio(nIdx + idx, 1), "#.00")
            Sheets(dataSheetName).Columns(4).Range(Cells(nIdx + idx, 1), Cells(nIdx + idx, 1)).Value = Format(d7ave(nIdx + idx, 1), "#.##")
        End If
    Next idx
    
    Sheets(dataSheetName).Columns(1).Range("M1") = "Status: Done"
    
    If (finaldate - datesFields(dayscount - 1) > 1) Then
        MsgBox "Your mail box has not received TPS for " & finaldate & ". Ensure your Mailbox is up-to-date"
    End If
End Sub
