Attribute VB_Name = "ExportPontozokSingleDoc"
Option Explicit

' Teljes modul — frissítve: az összesítõ dokumentumok fejlécébe is beírjuk
' a bizottság nevét és az idõpontot (ahogy a pontozó dokumentumnál).
' Állítsd be a sablonok útvonalait a modul elején, mentsd .xlsm, majd futtasd az ExportPontozok_Full_WithSummary eljárást.

' Word late-binding konstansok
Private Const wdPageBreak As Long = 7
Private Const wdReplaceAll As Long = 2
Private Const wdFindStop As Long = 0
Private Const wdFormatXMLDocument As Long = 16
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdAlignParagraphCenter As Long = 1

' === Konfiguráció - állítsd be az útvonalakat ===
Private Const MAIN_TEMPLATE_PATH As String = "\\NS2\Felvételi\Data\PontozolapTemplate.docx"
Private Const SUMMARY_TEMPLATE_PATH As String = "\\NS2\Felvételi\Data\OsszesitolapTemplate.docx"
Private Const OUTPUT_ROOT As String = "\\NS2\Felvételi\Data\Pontozo\"    ' gyökér mappa az eredményeknek
' ==================================================

' Fõ eljárás - integrált mûködés
Sub ExportPontozok_Full_WithSummary()
    On Error GoTo ErrMain
    Const N_PER_PAGE As Long = 4
    Const placeholder As String = "{{DATA_START}}"
    Const EXPORTED_COL_NAME As String = "exported" ' oszlopnév, ami jelöli a feldolgozott sorokat
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ellenõrizzük a mappát
    If Not fso.FolderExists(OUTPUT_ROOT) Then
        On Error Resume Next
        fso.CreateFolder OUTPUT_ROOT
        On Error GoTo 0
    End If
    
    ' Kérdések a futtatónak
    Dim markExported As Boolean, mergePerCommittee As Boolean, createSummaryDocs As Boolean
    If MsgBox("Jelöljem az exportált sorokat az Excel táblában (exported)?", vbYesNo + vbQuestion, "Jelölés") = vbYes Then
        markExported = True
    Else
        markExported = False
    End If
    If MsgBox("Összefûzzem a bizottsági .docx fájlokat egy per-bizottság fájlba?", vbYesNo + vbQuestion, "Merge") = vbYes Then
        mergePerCommittee = True
    Else
        mergePerCommittee = False
    End If
    If MsgBox("Készítsek per-bizottság Word összesítõ dokumentumot a sablon alapján (" & _
              fso.GetFileName(SUMMARY_TEMPLATE_PATH) & ")?", vbYesNo + vbQuestion, "Összesítõ") = vbYes Then
        createSummaryDocs = True
    Else
        createSummaryDocs = False
    End If
    
    ' Inicializáljuk a naplót
    InitExportLog
    
    ' Biztonságos sablonmásolat a Temp-be (fõ sablon)
    Dim localMainTemplate As String: localMainTemplate = Environ("Temp") & "\temp_pontozolap_template.docx"
    Dim mainTemplateUsed As String: mainTemplateUsed = ""
    If Not fso.FileExists(MAIN_TEMPLATE_PATH) Then
        MsgBox "A fõ sablon nem található: " & MAIN_TEMPLATE_PATH, vbCritical
        Exit Sub
    End If
    On Error Resume Next
    If fso.FileExists(localMainTemplate) Then fso.DeleteFile localMainTemplate, True
    Err.clear
    fso.CopyFile MAIN_TEMPLATE_PATH, localMainTemplate
    If Err.Number = 0 And fso.FileExists(localMainTemplate) Then
        mainTemplateUsed = localMainTemplate
    Else
        mainTemplateUsed = MAIN_TEMPLATE_PATH
    End If
    On Error GoTo 0
    
    ' Summary sablon ellenõrzése (ha kértük)
    Dim summaryTemplateUsed As String: summaryTemplateUsed = ""
    If createSummaryDocs Then
        If Not fso.FileExists(SUMMARY_TEMPLATE_PATH) Then
            MsgBox "Az összesítõ sablon nem található: " & SUMMARY_TEMPLATE_PATH & vbCrLf & "A summary dokumentumok nem készülnek.", vbExclamation
            createSummaryDocs = False
        Else
            ' másoljuk tempbe
            Dim localSummaryTemplate As String: localSummaryTemplate = Environ("Temp") & "\temp_osszesito_template.docx"
            On Error Resume Next
            If fso.FileExists(localSummaryTemplate) Then fso.DeleteFile localSummaryTemplate, True
            Err.clear
            fso.CopyFile SUMMARY_TEMPLATE_PATH, localSummaryTemplate
            If Err.Number = 0 And fso.FileExists(localSummaryTemplate) Then
                summaryTemplateUsed = localSummaryTemplate
            Else
                summaryTemplateUsed = SUMMARY_TEMPLATE_PATH
            End If
            On Error GoTo 0
        End If
    End If
    
    ' Beolvassuk az Excel tábla "diakadat"
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("diakadat")
    If ws Is Nothing Then MsgBox "Nincs 'diakadat' munkalap.", vbExclamation: Exit Sub
    Set lo = ws.ListObjects("diakadat")
    If lo Is Nothing Then MsgBox "Nincs 'diakadat' tábla (ListObject).", vbExclamation: Exit Sub
    On Error GoTo 0
    
    ' Oszlopok indexei
    Dim dictCols As Object: Set dictCols = CreateObject("Scripting.Dictionary")
    Dim c As Long, hdr As String
    For c = 1 To lo.HeaderRowRange.Columns.Count
        hdr = Trim(LCase(CStr(lo.HeaderRowRange.Cells(1, c).value & "")))
        If hdr <> "" Then dictCols(hdr) = c
    Next c
    If Not dictCols.Exists("f_nev") Or Not dictCols.Exists("bizottsag") Or Not dictCols.Exists("datum_nap") Then
        MsgBox "A 'diakadat' tábla nem tartalmazza a szükséges oszlopokat: f_nev, bizottsag, datum_nap", vbExclamation
        Exit Sub
    End If
    
    ' Ha jelölés kért, biztosítjuk az exported oszlopot
    If markExported Then
        If Not dictCols.Exists(EXPORTED_COL_NAME) Then
            On Error Resume Next
            lo.ListColumns.add
            lo.HeaderRowRange.Cells(1, lo.ListColumns.Count).value = EXPORTED_COL_NAME
            On Error GoTo 0
            ' frissítjük a dictCols
            dictCols.RemoveAll
            For c = 1 To lo.HeaderRowRange.Columns.Count
                hdr = Trim(LCase(CStr(lo.HeaderRowRange.Cells(1, c).value & "")))
                If hdr <> "" Then dictCols(hdr) = c
            Next c
        End If
    End If
    
    ' Csoportosítás: kulcs = bizottsag||datumStr ; tároljuk a ListRow objektumokat
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")
    Dim lr As ListRow
    Dim nm As String, biz As String, dt As Variant, key As String, dtStr As String
    For Each lr In lo.ListRows
        ' ha markExported és már jelölt, kihagyjuk
        If markExported Then
            Dim valExp As String
            valExp = Trim(CStr(lr.Range.Cells(1, dictCols(EXPORTED_COL_NAME)).value & ""))
            If valExp <> "" Then GoTo NextRowMain
        End If
        
        nm = Trim(CStr(lr.Range.Cells(1, dictCols("f_nev")).value & ""))
        If Len(nm) = 0 Then GoTo NextRowMain
        biz = Trim(CStr(lr.Range.Cells(1, dictCols("bizottsag")).value & ""))
        dt = lr.Range.Cells(1, dictCols("datum_nap")).value
        If IsDate(dt) Then
            dtStr = Format(CDate(dt), "yyyy-mm-dd_HHnn")
        Else
            dtStr = Trim(CStr(dt & ""))
            If dtStr = "" Then dtStr = "no_date"
        End If
        key = biz & "||" & dtStr
        If Not groups.Exists(key) Then
            Dim collRows As Collection: Set collRows = New Collection
            groups.add key, collRows
        End If
        groups(key).add lr  ' tároljuk a ListRow-t
NextRowMain:
    Next lr
    
    If groups.Count = 0 Then
        MsgBox "Nincs feldolgozható (új) adat.", vbInformation
        Exit Sub
    End If
    
    ' Indítjuk a Word-öt (late binding)
    Dim wdApp As Object
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    wdApp.Visible = False
    
    ' Megnyitjuk a main template (readonly)
    Dim templateDoc As Object
    On Error Resume Next
    Set templateDoc = wdApp.Documents.Open(mainTemplateUsed, ReadOnly:=True)
    If Err.Number <> 0 Or templateDoc Is Nothing Then
        MsgBox "Nem sikerült megnyitni a fõ sablont: " & Err.Number & " - " & Err.Description, vbCritical
        GoTo CleanupMain
    End If
    On Error GoTo 0
    Dim templateContent As Object: Set templateContent = templateDoc.Content
    
    ' Feldolgozás: minden csoportból készítünk egy docx fájlt
    Dim g As Variant, rowsColl As Collection, names As Collection
    Dim totalFiles As Long: totalFiles = 0
    For Each g In groups.keys
        Set rowsColl = groups(g)
        ' készítsük el a névlistát
        Set names = New Collection
        For Each lr In rowsColl
            names.add Trim(CStr(lr.Range.Cells(1, dictCols("f_nev")).value & ""))
        Next lr
        
        Dim parts() As String: parts = Split(g, "||")
        Dim bizLabel As String: bizLabel = parts(0)
        Dim dateLabel As String: dateLabel = parts(1)
        If bizLabel = "" Then bizLabel = "NoCommittee"
        If dateLabel = "" Then dateLabel = "no_date"
        
        Dim committeeFolder As String
        committeeFolder = fso.BuildPath(OUTPUT_ROOT, SafeFileName(bizLabel))
        If Not fso.FolderExists(committeeFolder) Then
            On Error Resume Next
            fso.CreateFolder committeeFolder
            On Error GoTo 0
        End If
        
        Dim pagesCreated As Long: pagesCreated = 0
        Dim outPath As String: outPath = ""
        Dim status As String: status = "OK"
        Dim message As String: message = ""
        
        On Error GoTo GroupErr
        ' Új doc létrehozása sablon alapján
        Dim newDoc As Object
        Set newDoc = wdApp.Documents.add(Template:=mainTemplateUsed, NewTemplate:=False)
        If newDoc Is Nothing Then Err.Raise vbObjectError + 1000, , "Nem sikerült dokumentumot létrehozni."
        
        ' Fejléc minden szekcióban (pontozó dokumentum)
        Dim sec As Object
        For Each sec In newDoc.Sections
            With sec.headers(wdHeaderFooterPrimary).Range
                .Text = bizLabel & "    " & dateLabel
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Font.Name = "Calibri"
                .Font.Size = 11
                .Font.Bold = True
            End With
        Next sec
        
        ' Feltöltjük oldalanként a neveket
        Dim i As Long, batchIndex As Long
        batchIndex = 0
        For i = 1 To names.Count Step N_PER_PAGE
            batchIndex = batchIndex + 1
            pagesCreated = pagesCreated + 1
            If batchIndex = 1 Then
                ' elsõ oldal: sablon már bent van
                If newDoc.Tables.Count >= 1 Then
                    Dim tbl0 As Object: Set tbl0 = newDoc.Tables(1)
                    Dim phRow As Long, phCol As Long
                    phRow = 0: phCol = 0
                    Call FindPlaceholderInTable(tbl0, placeholder, phRow, phCol)
                    If phRow > 0 Then
                        FillTableFromCell tbl0, names, i, N_PER_PAGE, phRow, phCol
                    Else
                        FillTableFirstColumn tbl0, names, i, N_PER_PAGE
                    End If
                Else
                    Dim rng0 As Object: Set rng0 = newDoc.Content
                    rng0.collapse Direction:=0
                    InsertSimpleList rng0, names, i, N_PER_PAGE
                End If
                ' globális helyõrzõk cseréje (ha léteznek)
                If RangeContainsText(newDoc.Range(0, newDoc.Content.End), "{{COMMITTEE}}") Then
                    ReplaceInDocumentRange newDoc.Range(0, newDoc.Content.End), "{{COMMITTEE}}", bizLabel
                End If
                If RangeContainsText(newDoc.Range(0, newDoc.Content.End), "{{DATE}}") Then
                    ReplaceInDocumentRange newDoc.Range(0, newDoc.Content.End), "{{DATE}}", dateLabel
                End If
            Else
                ' új oldal: beillesztjük a sablon tartalmát
                Dim pasteStart As Long: pasteStart = newDoc.Content.End
                newDoc.Range(newDoc.Content.End).InsertBreak Type:=wdPageBreak
                templateContent.Copy
                newDoc.Range(newDoc.Content.End).Paste
                Dim rngInserted As Object: Set rngInserted = newDoc.Range(pasteStart, newDoc.Content.End)
                If rngInserted.Tables.Count >= 1 Then
                    Dim insertedTbl As Object: Set insertedTbl = rngInserted.Tables(1)
                    Dim phR As Long, phC As Long
                    phR = 0: phC = 0
                    Call FindPlaceholderInTable(insertedTbl, placeholder, phR, phC)
                    If phR > 0 Then
                        FillTableFromCell insertedTbl, names, i, N_PER_PAGE, phR, phC
                    Else
                        FillTableFirstColumn insertedTbl, names, i, N_PER_PAGE
                    End If
                Else
                    rngInserted.collapse Direction:=0
                    InsertSimpleList rngInserted, names, i, N_PER_PAGE
                End If
                If RangeContainsText(rngInserted, "{{COMMITTEE}}") Then
                    ReplaceInDocumentRange rngInserted, "{{COMMITTEE}}", bizLabel
                End If
                If RangeContainsText(rngInserted, "{{DATE}}") Then
                    ReplaceInDocumentRange rngInserted, "{{DATE}}", dateLabel
                End If
            End If
        Next i
        
        ' Mentés (pontozó dokumentum)
        Dim outName As String
        outName = SafeFileName(bizLabel & "_" & dateLabel & ".docx")
        outPath = fso.BuildPath(committeeFolder, outName)
        newDoc.SaveAs2 fileName:=outPath, FileFormat:=wdFormatXMLDocument
        newDoc.Close SaveChanges:=False
        totalFiles = totalFiles + 1
        
        ' Azonnali summary dokumentum készítése ugyanabból a névlistából (ha kértük)
        If createSummaryDocs And summaryTemplateUsed <> "" Then
            On Error Resume Next
            Dim summaryDoc As Object
            Set summaryDoc = wdApp.Documents.add(Template:=summaryTemplateUsed, NewTemplate:=False)
            On Error GoTo 0
            If Not summaryDoc Is Nothing Then
                ' --- FEJLÉC beállítása a summary dokumentumnál is ---
                Dim ssec As Object
                For Each ssec In summaryDoc.Sections
                    With ssec.headers(wdHeaderFooterPrimary).Range
                        .Text = bizLabel & "    " & dateLabel
                        .ParagraphFormat.Alignment = wdAlignParagraphCenter
                        .Font.Name = "Calibri"
                        .Font.Size = 11
                        .Font.Bold = True
                    End With
                Next ssec
                ' -------------------------------------------------------
                
                ' Replace committee/date if placeholders exist
                If RangeContainsText(summaryDoc.Range(0, summaryDoc.Content.End), "{{COMMITTEE}}") Then
                    ReplaceInDocumentRange summaryDoc.Range(0, summaryDoc.Content.End), "{{COMMITTEE}}", bizLabel
                End If
                If RangeContainsText(summaryDoc.Range(0, summaryDoc.Content.End), "{{DATE}}") Then
                    ReplaceInDocumentRange summaryDoc.Range(0, summaryDoc.Content.End), "{{DATE}}", dateLabel
                End If
                
                ' Try to find PLACEHOLDER in first table
                Dim sPhR As Long, sPhC As Long
                sPhR = 0: sPhC = 0
                If summaryDoc.Tables.Count >= 1 Then
                    Call FindPlaceholderInTable(summaryDoc.Tables(1), placeholder, sPhR, sPhC)
                End If
                
                If sPhR > 0 Then
                    FillTableFromCell summaryDoc.Tables(1), names, 1, names.Count, sPhR, sPhC
                Else
                    ' fallback: try each table, else append simple list
                    Dim used As Boolean: used = False
                    Dim tblTry As Object
                    For Each tblTry In summaryDoc.Tables
                        sPhR = 0: sPhC = 0
                        Call FindPlaceholderInTable(tblTry, placeholder, sPhR, sPhC)
                        If sPhR > 0 Then
                            FillTableFromCell tblTry, names, 1, names.Count, sPhR, sPhC
                            used = True: Exit For
                        End If
                    Next tblTry
                    If Not used Then
                        If summaryDoc.Tables.Count >= 1 Then
                            FillTableFirstColumn summaryDoc.Tables(1), names, 1, names.Count
                        Else
                            Dim rngS As Object: Set rngS = summaryDoc.Content
                            rngS.collapse Direction:=0
                            rngS.InsertAfter vbCrLf & "Összesített nevek:" & vbCrLf
                            Dim ii As Long
                            For ii = 1 To names.Count
                                rngS.InsertAfter CStr(ii) & ". " & names(ii) & vbCrLf
                            Next ii
                        End If
                    End If
                End If
                
                ' Mentés: Osszesito_<bizLabel>_<dateLabel>.docx a committeeFolder-be
                Dim sumName As String: sumName = "Osszesito_" & SafeFileName(bizLabel & "_" & dateLabel) & ".docx"
                Dim sumPath As String: sumPath = fso.BuildPath(committeeFolder, sumName)
                summaryDoc.SaveAs2 fileName:=sumPath, FileFormat:=wdFormatXMLDocument
                summaryDoc.Close SaveChanges:=False
            End If
            Set summaryDoc = Nothing
        End If
        
        ' Jelölés Excelben (ha kértük) - timestampot írunk az exported oszlopba
        If markExported Then
            Dim rowItem As ListRow
            For Each rowItem In rowsColl
                On Error Resume Next
                rowItem.Range.Cells(1, dictCols(EXPORTED_COL_NAME)).value = Now
                On Error GoTo 0
            Next rowItem
        End If
        
        ' Log és Excel összesítõ frissítése
        LogEntry Now, bizLabel, dateLabel, pagesCreated, outPath, status, message
        ' Frissítjük a summary sheet (AllNames + Grader1..3)
        UpdateSummarySheet bizLabel, dateLabel, names, pagesCreated, outPath, status, 3
        
        Set newDoc = Nothing
        GoTo NextGroupMain
        
GroupErr:
        status = "ERROR"
        message = "Err#: " & Err.Number & " - " & Err.Description
        If Not newDoc Is Nothing Then
            On Error Resume Next
            newDoc.Close SaveChanges:=False
        End If
        LogEntry Now, bizLabel, dateLabel, pagesCreated, outPath, status, message
        Err.clear
        Resume NextGroupMain
        
NextGroupMain:
    Next g
    
    ' Bezárjuk a sablont
    On Error Resume Next
    If Not templateDoc Is Nothing Then templateDoc.Close SaveChanges:=False
    On Error GoTo 0
    
    ' Merge per committee (ha kértük)
    If mergePerCommittee Then
        MergeAllCommitteesInFolder OUTPUT_ROOT, wdApp, fso
    End If
    
    wdApp.Visible = True
    wdApp.Activate
    MsgBox "Kész. Létrehozott dokumentumok száma: " & totalFiles & vbCrLf & "Gyökér mappa: " & OUTPUT_ROOT, vbInformation
    Exit Sub

ErrMain:
    MsgBox "Váratlan hiba: " & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanupMain

CleanupMain:
    On Error Resume Next
    If Not templateDoc Is Nothing Then templateDoc.Close SaveChanges:=False
    Set templateDoc = Nothing
    Set wdApp = Nothing
    Set fso = Nothing
End Sub

' -------------------------
' Merge helper: összefûzi minden bizottság mappájában a .docx fájlokat egy ALL_<bizottsag>.docx fájlba
' Frissítve: a merged dokumentum fejlécébe is beírjuk a bizottság nevét és "various" dátumot.
' -------------------------
Private Sub MergeAllCommitteesInFolder(rootPath As String, wdApp As Object, fso As Object)
    On Error Resume Next
    Dim fld As Object, subfld As Object
    Set fld = fso.GetFolder(rootPath)
    For Each subfld In fld.SubFolders
        Dim mergedName As String
        mergedName = fso.BuildPath(subfld.path, "ALL_" & SafeFileName(subfld.Name) & ".docx")
        ' ha nincs benne .docx fájl, kihagyjuk
        Dim hasDocx As Boolean: hasDocx = False
        Dim f As Object
        For Each f In subfld.Files
            If LCase(fso.GetExtensionName(f.Name)) = "docx" Then hasDocx = True: Exit For
        Next f
        If Not hasDocx Then GoTo NextSub
        Dim mergedDoc As Object
        Set mergedDoc = wdApp.Documents.add(BlankTemplate:=False)
        ' --- FEJLÉC beállítása mergedDoc-on is ---
        Dim ssec As Object
        For Each ssec In mergedDoc.Sections
            With ssec.headers(wdHeaderFooterPrimary).Range
                .Text = subfld.Name & "    various"
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Font.Name = "Calibri"
                .Font.Size = 11
                .Font.Bold = True
            End With
        Next ssec
        ' --------------------------------------------
        For Each f In subfld.Files
            If LCase(fso.GetExtensionName(f.Name)) = "docx" Then
                mergedDoc.Range(mergedDoc.Content.End - 1, mergedDoc.Content.End - 1).InsertFile f.path
                mergedDoc.Range(mergedDoc.Content.End).InsertBreak Type:=wdPageBreak
            End If
        Next f
        mergedDoc.SaveAs2 fileName:=mergedName, FileFormat:=wdFormatXMLDocument
        mergedDoc.Close SaveChanges:=False
NextSub:
    Next subfld
    On Error GoTo 0
End Sub

' -------------------------
' Segédfüggvények: placeholder keresés és feltöltés
' -------------------------
Private Sub FindPlaceholderInTable(tbl As Object, placeholder As String, ByRef outRow As Long, ByRef outCol As Long)
    Dim r As Long, co As Long, plain As String
    outRow = 0: outCol = 0
    On Error Resume Next
    For r = 1 To tbl.rows.Count
        For co = 1 To tbl.Columns.Count
            plain = CellTextPlain(tbl.cell(r, co).Range.Text)
            If plain = placeholder Then
                outRow = r: outCol = co
                Exit Sub
            End If
        Next co
    Next r
    On Error GoTo 0
End Sub

Private Sub FillTableFromCell(tbl As Object, names As Collection, srcOffset As Long, n As Long, startRow As Long, startCol As Long)
    On Error GoTo ErrHandler
    Dim needRows As Long: needRows = startRow + n - 1
    EnsureTableRows tbl, needRows
    EnsureTableColumns tbl, startCol
    Dim k As Long, srcIdx As Long, targetRow As Long
    For k = 1 To n
        srcIdx = srcOffset + (k - 1)
        targetRow = startRow + (k - 1)
        If srcIdx >= 1 And srcIdx <= names.Count Then
            On Error Resume Next
            tbl.cell(targetRow, startCol).Range.Text = names(srcIdx)
            If Err.Number <> 0 Then
                Err.clear
                EnsureTableRows tbl, targetRow
                On Error Resume Next
                tbl.cell(targetRow, startCol).Range.Text = names(srcIdx)
            End If
            On Error GoTo ErrHandler
        Else
            On Error Resume Next
            tbl.cell(targetRow, startCol).Range.Text = ""
            On Error GoTo ErrHandler
        End If
    Next k
    Exit Sub
ErrHandler:
    Resume Next
End Sub

Private Sub FillTableFirstColumn(tbl As Object, names As Collection, srcOffset As Long, n As Long)
    On Error GoTo ErrHandler
    Dim firstRow As Long: firstRow = GetFirstDataRow(tbl)
    Dim needRows As Long: needRows = firstRow + n - 1
    EnsureTableRows tbl, needRows
    Dim k As Long, srcIdx As Long, targetRow As Long
    For k = 1 To n
        srcIdx = srcOffset + (k - 1)
        targetRow = firstRow + (k - 1)
        If srcIdx >= 1 And srcIdx <= names.Count Then
            On Error Resume Next
            tbl.cell(targetRow, 1).Range.Text = names(srcIdx)
            If Err.Number <> 0 Then
                Err.clear
                EnsureTableRows tbl, targetRow
                On Error Resume Next
                tbl.cell(targetRow, 1).Range.Text = names(srcIdx)
            End If
            On Error GoTo ErrHandler
        Else
            On Error Resume Next
            tbl.cell(targetRow, 1).Range.Text = ""
            On Error GoTo ErrHandler
        End If
    Next k
    Exit Sub
ErrHandler:
    Resume Next
End Sub

Private Sub EnsureTableRows(tbl As Object, n As Long)
    On Error Resume Next
    Dim cur As Long: cur = tbl.rows.Count
    Dim add As Long, r As Long
    If cur < n Then
        add = n - cur
        For r = 1 To add
            tbl.rows.add
        Next r
    End If
    On Error GoTo 0
End Sub

Private Sub EnsureTableColumns(tbl As Object, nCol As Long)
    On Error Resume Next
    Dim cur As Long: cur = tbl.Columns.Count
    Dim add As Long, c As Long
    If cur < nCol Then
        add = nCol - cur
        For c = 1 To add
            tbl.Columns.add
        Next c
    End If
    On Error GoTo 0
End Sub

Private Function CellTextPlain(Txt As String) As String
    Txt = Replace(Txt, vbCr, "")
    Txt = Replace(Txt, Chr(7), "")
    CellTextPlain = Trim(Txt)
End Function

Private Function GetFirstDataRow(tbl As Object) As Long
    On Error Resume Next
    Dim r As Long, s As String
    For r = 1 To tbl.rows.Count
        s = LCase(CellTextPlain(tbl.cell(r, 1).Range.Text))
        If s = "név" Or InStr(s, "név") > 0 Then
            GetFirstDataRow = r + 1
            Exit Function
        End If
    Next r
    If tbl.rows.Count >= 1 Then
        s = CellTextPlain(tbl.cell(1, 1).Range.Text)
        If Len(s) > 30 Then
            GetFirstDataRow = 2
            Exit Function
        End If
    End If
    GetFirstDataRow = 1
End Function

' -------------------------
' Public helper: Replace and Range check
' -------------------------
Public Sub ReplaceInDocumentRange(rng As Object, findText As String, replaceText As String)
    On Error Resume Next
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With
    On Error GoTo 0
End Sub

Public Function RangeContainsText(rng As Object, findText As String) As Boolean
    On Error GoTo ErrHandler
    With rng.Find
        .ClearFormatting
        .Text = findText
        .Forward = True
        .wrap = wdFindStop
        .Format = False
        .MatchCase = False
    End With
    RangeContainsText = rng.Find.Execute
    Exit Function
ErrHandler:
    RangeContainsText = False
    Err.clear
End Function

' -------------------------
' Logging (ExportLog)
' -------------------------
Private Sub InitExportLog()
    On Error Resume Next
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("ExportLog")
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sh.Name = "ExportLog"
    Else
        sh.Cells.clear
    End If
    sh.Range("A1:G1").value = Array("Timestamp", "Committee", "DateLabel", "Pages", "FilePath", "Status", "Message")
    sh.Columns("A:G").AutoFit
    On Error GoTo 0
End Sub

Private Sub LogEntry(ts As Date, committee As String, dateLabel As String, pages As Long, filePath As String, status As String, msg As String)
    On Error Resume Next
    Dim sh As Worksheet: Set sh = ThisWorkbook.Worksheets("ExportLog")
    Dim NextRow As Long: NextRow = sh.Cells(sh.rows.Count, "A").End(xlUp).Row + 1
    If NextRow < 2 Then NextRow = 2
    sh.Cells(NextRow, 1).value = ts
    sh.Cells(NextRow, 2).value = committee
    sh.Cells(NextRow, 3).value = dateLabel
    sh.Cells(NextRow, 4).value = pages
    sh.Cells(NextRow, 5).value = filePath
    sh.Cells(NextRow, 6).value = status
    sh.Cells(NextRow, 7).value = msg
    On Error GoTo 0
End Sub

' -------------------------
' Summary sheet update
' -------------------------
Public Sub UpdateSummarySheet(bizLabel As String, dateLabel As String, names As Collection, pages As Long, filePath As String, status As String, Optional numGraders As Long = 0)
    On Error GoTo ErrHandler
    Dim sh As Worksheet
    Dim wb As Workbook: Set wb = ThisWorkbook
    Const SUM_SHEET_NAME As String = "Osszesito"
    On Error Resume Next
    Set sh = wb.Worksheets(SUM_SHEET_NAME)
    On Error GoTo ErrHandler
    If sh Is Nothing Then
        Set sh = wb.Worksheets.add(After:=wb.Worksheets(wb.Worksheets.Count))
        sh.Name = SUM_SHEET_NAME
        sh.Range("A1:G1").value = Array("Timestamp", "Committee", "DateLabel", "AllNames", "Pages", "FilePath", "Status")
    End If
    Dim i As Long, baseCols As Long: baseCols = 7
    If numGraders >= 2 Then
        For i = 1 To numGraders
            sh.Cells(1, baseCols + i).value = "Grader" & i
        Next i
    End If
    Dim NextRow As Long: NextRow = sh.Cells(sh.rows.Count, "A").End(xlUp).Row + 1
    If NextRow < 2 Then NextRow = 2
    Dim allNames As String: allNames = ""
    Dim idx As Long
    For idx = 1 To names.Count
        If allNames <> "" Then allNames = allNames & "; "
        allNames = allNames & names(idx)
    Next idx
    sh.Cells(NextRow, 1).value = Now
    sh.Cells(NextRow, 2).value = bizLabel
    sh.Cells(NextRow, 3).value = dateLabel
    sh.Cells(NextRow, 4).value = allNames
    sh.Cells(NextRow, 5).value = pages
    sh.Cells(NextRow, 6).value = filePath
    sh.Cells(NextRow, 7).value = status
    If numGraders >= 2 Then
        Dim graders() As String
        ReDim graders(1 To numGraders)
        For idx = 1 To names.Count
            Dim g As Long: g = ((idx - 1) Mod numGraders) + 1
            If graders(g) = "" Then
                graders(g) = names(idx)
            Else
                graders(g) = graders(g) & "; " & names(idx)
            End If
        Next idx
        For i = 1 To numGraders
            sh.Cells(NextRow, baseCols + i).value = graders(i)
        Next i
    End If
    sh.Columns("A:" & ColumnLetter(baseCols + Application.Max(0, numGraders))).AutoFit
    Exit Sub
ErrHandler:
    Debug.Print "UpdateSummarySheet hiba: " & Err.Number & " - " & Err.Description
    Resume Next
End Sub

Private Function ColumnLetter(colNum As Long) As String
    Dim result As String: result = ""
    Dim n As Long: n = colNum
    Do While n > 0
        Dim rmd As Long
        rmd = (n - 1) Mod 26
        result = Chr(65 + rmd) & result
        n = (n - 1) \ 26
    Loop
    ColumnLetter = result
End Function

' -------------------------
' Utility és egyszerû lista beillesztés
' -------------------------
Private Function SafeFileName(s As String) As String
    Dim bad As Variant: bad = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next i
    SafeFileName = Trim(s)
End Function

Private Sub InsertSimpleList(rng As Object, names As Collection, srcOffset As Long, n As Long)
    Dim k As Long, srcIdx As Long
    For k = 1 To n
        srcIdx = srcOffset + (k - 1)
        If srcIdx >= 1 And srcIdx <= names.Count Then
            rng.InsertAfter CStr(k) & ". " & names(srcIdx) & vbCrLf
        Else
            rng.InsertAfter CStr(k) & ". " & vbCrLf
        End If
    Next k
End Sub

