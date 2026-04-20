Attribute VB_Name = "SendMailSzobeliMeghivo"
Option Explicit

' (Ujravágott) Feldolgozás batch-ekben, használja a ListRow.Index-et a pontos sorazonosításhoz.
' A további mûködés: iktsz feltöltés, mentés, batch feldolgozás, helyi sablonmásolat, retry, StepLog, MailErrors.

Const DEFAULT_BATCH_SIZE As Long = 20
Const MAX_RETRIES As Long = 3
Const RETRY_WAIT_SECONDS As Double = 1
Const MAX_DISPLAY_TEST As Long = 3

Private gErrorSheet As Worksheet
Private gErrorRow As Long
Private gStepsSheet As Worksheet
Private gStepsRow As Long

Sub PrepareIktatoAndSendBatch()
    Dim wsName As String: wsName = "diakadat"
    Dim loName As String: loName = "diakadat"
    Dim templatePath As String
    templatePath = "\\NS2\Felvételi\outlooksablon\szobeli-behivo.oft" ' <-- állítsd be
    
    Dim sendAutomatically As Boolean: sendAutomatically = True
    Dim batchSize As Long: batchSize = DEFAULT_BATCH_SIZE
    
    Dim startInput As String
    Dim startIktato As Long
    startInput = InputBox("Kérem a kezdõ iktatószámot (egész szám). Üres: folytatja a meglévõ iktsz-ek alapján:", "Kezdõ iktatószám", "")
    If Trim(startInput) = "" Then
        startIktato = 0
    ElseIf IsNumeric(startInput) Then
        startIktato = CLng(startInput)
    Else
        MsgBox "A megadott érték nem szám. Mûvelet megszakítva.", vbExclamation
        Exit Sub
    End If
    
    InitErrorLog
    InitStepLog
    
    Call FillIktszColumn(wsName, loName, startIktato)
    
    On Error Resume Next
    ThisWorkbook.Save
    DoEvents
    Application.Calculate
    On Error GoTo 0
    
    Call ProcessNextBatch(wsName, loName, templatePath, batchSize, sendAutomatically)
    
    MsgBox "Kész. Nézd meg a MailErrors és StepLog lapokat.", vbInformation
End Sub

Private Sub FillIktszColumn(wsName As String, loName As String, startIktato As Long)
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(wsName)
    Dim lo As ListObject: Set lo = ws.ListObjects(loName)
    If lo Is Nothing Then MsgBox "Tábla nem található.", vbExclamation: Exit Sub
    
    Dim hdrBizottsag As String: hdrBizottsag = "bizottsag"
    Dim hdrDatum As String: hdrDatum = "datum_nap"
    Dim hdrMail As String: hdrMail = "mail"
    Dim hdrIdopontKiadva As String: hdrIdopontKiadva = "idopont_kiadva"
    Dim hdrIktato As String: hdrIktato = "iktsz"
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, h As String
    For i = 1 To lo.HeaderRowRange.Columns.Count
        h = Trim(CStr(lo.HeaderRowRange.Cells(1, i).value & ""))
        If h <> "" Then dict(LCase(h)) = i
    Next i
    
    If Not dict.Exists(LCase(hdrIktato)) Then
        Dim newCol As ListColumn
        Set newCol = lo.ListColumns.add
        newCol.Name = hdrIktato
        dict(LCase(hdrIktato)) = newCol.Index
        LogStep 0, "", "", "FillIktsz", "Added iktsz column"
    End If
    
    Dim missing As String: missing = ""
    If Not dict.Exists(LCase(hdrBizottsag)) Then missing = missing & hdrBizottsag & vbCrLf
    If Not dict.Exists(LCase(hdrDatum)) Then missing = missing & hdrDatum & vbCrLf
    If Not dict.Exists(LCase(hdrMail)) Then missing = missing & hdrMail & vbCrLf
    If Not dict.Exists(LCase(hdrIdopontKiadva)) Then missing = missing & hdrIdopontKiadva & vbCrLf
    If missing <> "" Then MsgBox "Hiányzó fejléc(ek):" & vbCrLf & missing, vbExclamation: Exit Sub
    
    Dim currentNum As Long
    If startIktato > 0 Then
        currentNum = startIktato
    Else
        currentNum = 0
        Dim lr As ListRow
        Dim val As Variant
        For Each lr In lo.ListRows
            val = NzString(lr.Range.Cells(1, dict(LCase(hdrIktato))).value)
            If IsNumeric(val) Then If CLng(val) > currentNum Then currentNum = CLng(val)
        Next lr
        If currentNum = 0 Then currentNum = 1 Else currentNum = currentNum + 1
    End If
    
    Dim filled As Long: filled = 0
    Dim lr2 As ListRow
    For Each lr2 In lo.ListRows
        Dim bizottsagVal As String, rawDatum As Variant, mailVal As String, givenVal As String
        bizottsagVal = NzString(lr2.Range.Cells(1, dict(LCase(hdrBizottsag))).value)
        rawDatum = lr2.Range.Cells(1, dict(LCase(hdrDatum))).value
        mailVal = NzString(lr2.Range.Cells(1, dict(LCase(hdrMail))).value)
        givenVal = NzString(lr2.Range.Cells(1, dict(LCase(hdrIdopontKiadva))).value)
        
        If Trim(bizottsagVal) <> "" And Trim(NzString(rawDatum)) <> "" And LCase(Trim(givenVal)) <> "x" And Trim(mailVal) <> "" Then
            Dim curIktCell As Range
            Set curIktCell = lr2.Range.Cells(1, dict(LCase(hdrIktato)))
            If Trim(NzString(curIktCell.value)) = "" Then
                curIktCell.value = CStr(currentNum)
                LogStep lr2.Range.Row, mailVal, CStr(currentNum), "FillIktsz", "Assigned"
                currentNum = currentNum + 1
                filled = filled + 1
            End If
        End If
    Next lr2
    
    MsgBox "Iktsz kitöltés kész. Új iktsz: " & filled, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Hiba a FillIktszColumn során: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub ProcessNextBatch(wsName As String, loName As String, templatePath As String, batchSize As Long, sendAutomatically As Boolean)
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(wsName)
    Dim lo As ListObject: Set lo = ws.ListObjects(loName)
    
    Dim hdrIktato As String: hdrIktato = "iktsz"
    Dim hdrBizottsag As String: hdrBizottsag = "bizottsag"
    Dim hdrDatum As String: hdrDatum = "datum_nap"
    Dim hdrFnev As String: hdrFnev = "f_nev"
    Dim hdrMail As String: hdrMail = "mail"
    Dim hdrIdopontKiadva As String: hdrIdopontKiadva = "idopont_kiadva"
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, h As String
    For i = 1 To lo.HeaderRowRange.Columns.Count
        h = Trim(CStr(lo.HeaderRowRange.Cells(1, i).value & ""))
        If h <> "" Then dict(LCase(h)) = i
    Next i
    If Not dict.Exists(LCase(hdrIktato)) Then MsgBox "Nincs iktsz oszlop.", vbExclamation: Exit Sub
    
    Dim rowsToProcess As Collection: Set rowsToProcess = New Collection
    Dim lr As ListRow
    For Each lr In lo.ListRows
        Dim iktValTmp As String
        iktValTmp = NzString(lr.Range.Cells(1, dict(LCase(hdrIktato))).value)
        If Trim(iktValTmp) <> "" Then
            Dim givenVal As String, mailVal As String, rawDatum As Variant, bizottsagVal As String
            givenVal = NzString(lr.Range.Cells(1, dict(LCase(hdrIdopontKiadva))).value)
            mailVal = NzString(lr.Range.Cells(1, dict(LCase(hdrMail))).value)
            rawDatum = lr.Range.Cells(1, dict(LCase(hdrDatum))).value
            bizottsagVal = NzString(lr.Range.Cells(1, dict(LCase(hdrBizottsag))).value)
            
            If Trim(bizottsagVal) <> "" And Trim(NzString(rawDatum)) <> "" And LCase(Trim(givenVal)) <> "x" And Trim(mailVal) <> "" Then
                Dim key As Long
                If IsNumeric(iktValTmp) Then key = CLng(iktValTmp) Else key = 2147483640
                ' STORE the ListRow.Index (reliable inside the table) instead of worksheet row
                rowsToProcess.add CStr(key) & "|" & CStr(lr.Index)
            End If
        End If
    Next lr
    
    If rowsToProcess.Count = 0 Then MsgBox "Nincs feldolgozható sor.", vbInformation: Exit Sub
    
    Dim arr() As String
    ReDim arr(1 To rowsToProcess.Count)
    Dim idx As Long
    For idx = 1 To rowsToProcess.Count
        arr(idx) = rowsToProcess(idx)
    Next idx
    Call QuickSortStrings(arr, LBound(arr), UBound(arr))
    
    Dim toHandle As Collection: Set toHandle = New Collection
    Dim take As Long: take = Application.Min(batchSize, UBound(arr) - LBound(arr) + 1)
    Dim j As Long
    For j = LBound(arr) To LBound(arr) + take - 1
        toHandle.add arr(j)
    Next j
    
    Dim localTemplate As String: localTemplate = ""
    On Error Resume Next
    If Len(Trim(templatePath & "")) > 0 Then
        localTemplate = Environ("Temp") & "\temp_template.oft"
        On Error Resume Next: Kill localTemplate: On Error GoTo 0
        On Error Resume Next: FileCopy templatePath, localTemplate
        If Err.Number <> 0 Then localTemplate = templatePath: Err.clear: LogStep 0, "", "", "TemplateCopy", "Could not copy to Temp; using network path" Else LogStep 0, "", "", "TemplateCopy", "Copied template to " & localTemplate
        On Error GoTo 0
    End If
    
    Dim OutApp As Object
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    If OutApp Is Nothing Then Set OutApp = CreateObject("Outlook.Application")
    On Error GoTo ErrHandler
    
    Dim displayedCount As Long: displayedCount = 0
    Dim sentCount As Long: sentCount = 0
    Dim errCount As Long: errCount = 0
    
    Dim arrIndex As Variant
    For Each arrIndex In toHandle
        Dim parts() As String: parts = Split(arrIndex, "|")
        Dim listRowIndex As Long: listRowIndex = CLng(parts(1)) ' THIS IS ListRow.Index now
        Set lr = lo.ListRows(listRowIndex)
        Dim rowRange As Range: Set rowRange = lr.Range
        
        ' Változók a sorhoz
        Dim iktVal As String, bizottsagValRow As String, rawDatumRow As Variant
        Dim fnevVal As String, toAddr As String, alreadyGiven As String
        Dim datumVal As String, bizottsagLabel As String, iktatoStr As String
        Dim worksheetRowNum As Long: worksheetRowNum = rowRange.Row ' useful for logs
        
        iktVal = NzString(rowRange.Cells(1, dict(LCase(hdrIktato))).value)
        bizottsagValRow = NzString(rowRange.Cells(1, dict(LCase(hdrBizottsag))).value)
        rawDatumRow = rowRange.Cells(1, dict(LCase(hdrDatum))).value
        fnevVal = NzString(rowRange.Cells(1, dict(LCase(hdrFnev))).value)
        toAddr = NzString(rowRange.Cells(1, dict(LCase(hdrMail))).value)
        alreadyGiven = NzString(rowRange.Cells(1, dict(LCase(hdrIdopontKiadva))).value)
        
        If IsDate(rawDatumRow) Then
            datumVal = Format(CDate(rawDatumRow), "yyyy-mm-dd HH:nn")
        Else
            datumVal = NzString(rawDatumRow)
        End If
        
        If IsNumeric(bizottsagValRow) Then
            On Error Resume Next
            bizottsagLabel = Toldalek2(CInt(bizottsagValRow))
            If Err.Number <> 0 Then Err.clear
            On Error GoTo ErrHandler
        Else
            bizottsagLabel = bizottsagValRow
        End If
        
        iktatoStr = iktVal
        
        ' Debug: tároljuk mind a ListRow.Index-et és a Worksheet-row-t
        LogStep worksheetRowNum, toAddr, iktatoStr, "DebugValues", "ListRowIndex=" & listRowIndex & "; wsRow=" & worksheetRowNum & "; fnev=" & fnevVal & "; iktsz=" & iktVal & "; datum=" & datumVal
        
        ' Create item (with retries)
        Dim OutMail As Object
        Dim attempt As Long, createdOK As Boolean: createdOK = False
        Dim templateToUse As String: templateToUse = IIf(localTemplate <> "", localTemplate, templatePath)
        For attempt = 1 To MAX_RETRIES
            On Error Resume Next
            Set OutMail = OutApp.CreateItemFromTemplate(templateToUse)
            If Err.Number = 0 And Not OutMail Is Nothing Then
                createdOK = True
                LogStep worksheetRowNum, toAddr, iktatoStr, "CreateItem", "OK attempt " & attempt
                Exit For
            Else
                LogStep worksheetRowNum, toAddr, iktatoStr, "CreateItem", "FAILED attempt " & attempt & " Err: " & Err.Number & " - " & Err.Description
                Err.clear
                DoEvents
                Application.Wait Now + TimeValue("0:00:" & CStr(RETRY_WAIT_SECONDS))
            End If
            On Error GoTo ErrHandler
        Next attempt
        
        If Not createdOK Then LogMailError worksheetRowNum, toAddr, iktatoStr, -1, "CreateItem failed": errCount = errCount + 1: GoTo NextInBatch
        
        ' Body read / fallback
        Dim bodyHTML As String, plainBody As String
        On Error Resume Next
        bodyHTML = OutMail.HTMLBody
        On Error GoTo ErrHandler
        If Len(Trim(bodyHTML & "")) = 0 Then
            On Error Resume Next
            plainBody = OutMail.body
            On Error GoTo ErrHandler
            If Len(Trim(plainBody & "")) > 0 Then
                bodyHTML = "<html><body><div style='font-family:Arial,Helvetica,sans-serif;'>" & Replace(Replace(plainBody, vbCrLf, "<br/>"), vbTab, "&nbsp;&nbsp;") & "</div></body></html>"
                LogStep worksheetRowNum, toAddr, iktatoStr, "Body", "Used plain Body fallback"
            Else
                bodyHTML = "<html><body><div style='font-family:Arial,Helvetica,sans-serif;'><p>Kedves {{F_NEV}},</p><p>Értesítés: {{BIZOTTSAG}} - {{DATUM_NAP}}</p><p>Üdvözlettel,</p></div></body></html>"
                LogStep worksheetRowNum, toAddr, iktatoStr, "Body", "Used default fallback"
            End If
        Else
            LogStep worksheetRowNum, toAddr, iktatoStr, "Body", "HTMLBody len=" & Len(bodyHTML)
        End If
        
        bodyHTML = Replace(bodyHTML, "{{IKTATOSZAM}}", iktatoStr)
        bodyHTML = Replace(bodyHTML, "{{BIZOTTSAG}}", bizottsagLabel)
        bodyHTML = Replace(bodyHTML, "{{DATUM_NAP}}", datumVal)
        bodyHTML = Replace(bodyHTML, "{{F_NEV}}", fnevVal)
        
        Dim stamp As String
        stamp = "<div style='font-size:10px;color:#666;margin-top:12px;'>Küldve: " & Format(Now, "yyyy-mm-dd HH:nn") & "</div>"
        If InStr(1, LCase(bodyHTML), "</body>") > 0 Then
            bodyHTML = Replace(bodyHTML, "</body>", stamp & "</body>")
        Else
            bodyHTML = bodyHTML & stamp
        End If
        
        createdOK = False
        For attempt = 1 To MAX_RETRIES
            On Error Resume Next
            OutMail.HTMLBody = bodyHTML
            If Err.Number = 0 Then
                createdOK = True
                LogStep worksheetRowNum, toAddr, iktatoStr, "SetBody", "OK attempt " & attempt
                Exit For
            Else
                LogStep worksheetRowNum, toAddr, iktatoStr, "SetBody", "FAILED attempt " & attempt & " Err: " & Err.Number & " - " & Err.Description
                Err.clear
                DoEvents
                Application.Wait Now + TimeValue("0:00:" & CStr(RETRY_WAIT_SECONDS))
            End If
            On Error GoTo ErrHandler
        Next attempt
        
        If Not createdOK Then LogMailError worksheetRowNum, toAddr, iktatoStr, -2, "Set HTMLBody failed": errCount = errCount + 1: GoTo NextInBatch
        
        On Error Resume Next
        OutMail.Subject = "Értesítés szóbeli idõpontról - " & datumVal
        If Err.Number <> 0 Then LogStep worksheetRowNum, toAddr, iktatoStr, "Subject", "Err: " & Err.Number & " - " & Err.Description: Err.clear Else LogStep worksheetRowNum, toAddr, iktatoStr, "Subject", "OK"
        On Error GoTo ErrHandler
        
        On Error Resume Next
        OutMail.To = toAddr
        If Err.Number <> 0 Then LogStep worksheetRowNum, toAddr, iktatoStr, "SetTo", "Err: " & Err.Number & " - " & Err.Description: Err.clear Else LogStep worksheetRowNum, toAddr, iktatoStr, "SetTo", "OK"
        On Error GoTo ErrHandler
        
        DoEvents
        Application.Wait Now + TimeValue("0:00:" & CStr(RETRY_WAIT_SECONDS))
        
        If sendAutomatically Then
            Dim sendOK As Boolean: sendOK = False
            Dim sendAttempt As Long
            For sendAttempt = 1 To MAX_RETRIES
                On Error Resume Next
                OutMail.send
                If Err.Number = 0 Then sendOK = True: LogStep worksheetRowNum, toAddr, iktatoStr, "Send", "OK attempt " & sendAttempt: Exit For Else LogStep worksheetRowNum, toAddr, iktatoStr, "Send", "FAILED attempt " & sendAttempt & " Err: " & Err.Number & " - " & Err.Description: Err.clear: DoEvents: Application.Wait Now + TimeValue("0:00:" & CStr(RETRY_WAIT_SECONDS))
                On Error GoTo ErrHandler
            Next sendAttempt
            
            If sendOK Then
                rowRange.Cells(1, dict(LCase(hdrIdopontKiadva))).value = "x"
                sentCount = sentCount + 1
            Else
                LogMailError worksheetRowNum, toAddr, iktatoStr, -3, "Send failed after retries"
                errCount = errCount + 1
            End If
        Else
            If displayedCount < MAX_DISPLAY_TEST Then
                On Error Resume Next: OutMail.Display
                If Err.Number = 0 Then LogStep worksheetRowNum, toAddr, iktatoStr, "Display", "Displayed" Else LogStep worksheetRowNum, toAddr, iktatoStr, "Display", "Err: " & Err.Number & " - " & Err.Description: Err.clear
                On Error GoTo ErrHandler
                displayedCount = displayedCount + 1
            Else
                On Error Resume Next: OutMail.Save
                If Err.Number = 0 Then LogStep worksheetRowNum, toAddr, iktatoStr, "Save", "Saved Draft" Else LogStep worksheetRowNum, toAddr, iktatoStr, "Save", "Err: " & Err.Number & " - " & Err.Description: Err.clear
                On Error GoTo ErrHandler
            End If
        End If
        
        On Error Resume Next: Set OutMail = Nothing: On Error GoTo ErrHandler
NextInBatch:
    Next arrIndex
    
    On Error Resume Next
    If Len(localTemplate) > 0 And InStr(1, localTemplate, Environ("Temp"), vbTextCompare) > 0 Then Kill localTemplate
    On Error GoTo ErrHandler
    
    MsgBox "Batch feldolgozás kész. Sikeres: " & sentCount & ", Hibák: " & errCount, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Váratlan hiba: " & Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    LogMailError 0, "", "", Err.Number, Err.Description
End Sub

' -----------------------
' Naplózó, segédfüggvények
' -----------------------
Private Sub InitErrorLog()
    On Error Resume Next
    Set gErrorSheet = ThisWorkbook.Worksheets("MailErrors")
    If gErrorSheet Is Nothing Then Set gErrorSheet = ThisWorkbook.Worksheets.add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)): gErrorSheet.Name = "MailErrors" Else gErrorSheet.Cells.clear
    On Error GoTo 0
    gErrorSheet.Range("A1:F1").value = Array("SheetRow", "MailTo", "Iktsz", "ErrNumber", "ErrDescription", "Time")
    gErrorRow = 2
End Sub

Private Sub LogMailError(rowNum As Long, mailTo As String, iktato As String, errNum As Long, errDesc As String)
    On Error Resume Next
    If gErrorSheet Is Nothing Then InitErrorLog
    gErrorSheet.Cells(gErrorRow, 1).value = rowNum
    gErrorSheet.Cells(gErrorRow, 2).value = mailTo
    gErrorSheet.Cells(gErrorRow, 3).value = iktato
    gErrorSheet.Cells(gErrorRow, 4).value = errNum
    gErrorSheet.Cells(gErrorRow, 5).value = errDesc
    gErrorSheet.Cells(gErrorRow, 6).value = Now
    gErrorRow = gErrorRow + 1
    On Error GoTo 0
End Sub

Private Sub InitStepLog()
    On Error Resume Next
    Set gStepsSheet = ThisWorkbook.Worksheets("StepLog")
    If gStepsSheet Is Nothing Then Set gStepsSheet = ThisWorkbook.Worksheets.add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)): gStepsSheet.Name = "StepLog" Else gStepsSheet.Cells.clear
    On Error GoTo 0
    gStepsSheet.Range("A1:G1").value = Array("SheetRow", "MailTo", "Iktsz", "Step", "Message", "Time", "User")
    gStepsRow = 2
End Sub

Private Sub LogStep(rowNum As Long, mailTo As String, iktato As String, stepName As String, msg As String)
    On Error Resume Next
    If gStepsSheet Is Nothing Then InitStepLog
    gStepsSheet.Cells(gStepsRow, 1).value = rowNum
    gStepsSheet.Cells(gStepsRow, 2).value = mailTo
    gStepsSheet.Cells(gStepsRow, 3).value = iktato
    gStepsSheet.Cells(gStepsRow, 4).value = stepName
    gStepsSheet.Cells(gStepsRow, 5).value = msg
    gStepsSheet.Cells(gStepsRow, 6).value = Now
    gStepsSheet.Cells(gStepsRow, 7).value = Environ("USERNAME")
    gStepsRow = gStepsRow + 1
    On Error GoTo 0
End Sub

Private Function NzString(v As Variant) As String
    ' Visszaad egy üres stringet, ha hibás vagy Null az érték, különben trimelt string
    On Error Resume Next
    If IsError(v) Then
        NzString = ""
    ElseIf IsNull(v) Then
        NzString = ""
    Else
        NzString = Trim(CStr(v & ""))
    End If
    On Error GoTo 0
End Function

Function Toldalek2(szam As Integer) As String
    Select Case szam
        Case 1, 2, 4, 7, 9, 10: Toldalek2 = szam & "-es"
        Case 8, 3: Toldalek2 = szam & "-as"
        Case 5: Toldalek2 = szam & "-ös"
        Case 6: Toldalek2 = szam & "-os"
        Case Else: Toldalek2 = szam & "-"
    End Select
End Function

Private Sub QuickSortStrings(arr() As String, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long, mid As String, tmp As String
    low = first: high = last: mid = arr((first + last) \ 2)
    Do While low <= high
        Do While CLng(Split(arr(low), "|")(0)) < CLng(Split(mid, "|")(0)): low = low + 1: Loop
        Do While CLng(Split(arr(high), "|")(0)) > CLng(Split(mid, "|")(0)): high = high - 1: Loop
        If low <= high Then tmp = arr(low): arr(low) = arr(high): arr(high) = tmp: low = low + 1: high = high - 1
    Loop
    If first < high Then QuickSortStrings arr, first, high
    If low < last Then QuickSortStrings arr, low, last
End Sub
