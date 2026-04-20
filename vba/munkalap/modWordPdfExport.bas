Attribute VB_Name = "modWordPdfExport"
Option Explicit

' Excelbõl fut: Word dokumentumot megnyitja / csatlakozik és PDF-ekre bontja
Public Sub SavePagesGroupedByRowChange_Excel(Optional control As IRibbonControl)
    On Error GoTo EH

    ' --- Word konstansok (late binding) ---
    Const wdGoToPage As Long = 1
    Const wdGoToLine As Long = 3
    Const wdGoToAbsolute As Long = 1
    Const wdLine As Long = 5

    Const wdExportFormatPDF As Long = 17
    Const wdExportOptimizeForPrint As Long = 0
    Const wdExportFromTo As Long = 3
    Const wdExportDocumentContent As Long = 0
    Const wdExportCreateHeadingBookmarks As Long = 1

    Dim docPath As String
    docPath = PickWordFile("Válaszd ki a Word dokumentumot")
    If docPath = "" Then Exit Sub

    Dim outFolder As String
    outFolder = PickFolder("Válaszd ki a kimeneti mappát (PDF-ek)")
    If outFolder = "" Then Exit Sub

    Dim xRow As Long, xStart As Long, xEnd As Long
    xRow = AskLong("Hányadik sor alapján történjen a csoportosítás?", "PDF bontás", 1)
    If xRow <= 0 Then Exit Sub

    xStart = AskLong("Kezdõ oldal?", "PDF bontás", 1)
    If xStart <= 0 Then Exit Sub

    xEnd = AskLong("Utolsó oldal?", "PDF bontás", xStart)
    If xEnd < xStart Then Exit Sub

    ' --- Word példány ---
    Dim wdApp As Object
    Set wdApp = GetOrCreateWordApp(True)

    ' --- Doc megnyitása (ha már nyitva van, azt használjuk) ---
    Dim doc As Object
    Set doc = GetOrOpenWordDoc(wdApp, docPath)
    If doc Is Nothing Then
        MsgBox "Nem sikerült megnyitni a dokumentumot.", vbExclamation
        Exit Sub
    End If

    ' --- Feldolgozás ---
    Dim i As Long
    Dim currentKey As String, previousKey As String
    Dim groupStart As Long
    groupStart = xStart
    previousKey = vbNullString

    Application.ScreenUpdating = False

    For i = xStart To xEnd
        ' Ugrás az adott oldal adott sorára
        wdApp.Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Name:=CStr(i)
        wdApp.Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Name:=CStr(xRow - 1)
        wdApp.Selection.Extend
        wdApp.Selection.EndKey Unit:=wdLine
        wdApp.Selection.EscapeKey

        currentKey = wdApp.Selection.Range.Text
        currentKey = CleanFileKey(currentKey, i)

        ' Ha megváltozott a kulcs, mentsük az elõzõ csoportot
        If previousKey <> "" And currentKey <> previousKey Then
            ExportDocPagesToPDF doc, outFolder & "\" & previousKey & ".pdf", groupStart, i - 1, _
                                wdExportFormatPDF, wdExportOptimizeForPrint, wdExportFromTo, _
                                wdExportDocumentContent, wdExportCreateHeadingBookmarks
            groupStart = i
        End If

        previousKey = currentKey
    Next i

    ' Utolsó csoport mentése
    If previousKey <> "" Then
        ExportDocPagesToPDF doc, outFolder & "\" & previousKey & ".pdf", groupStart, xEnd, _
                            wdExportFormatPDF, wdExportOptimizeForPrint, wdExportFromTo, _
                            wdExportDocumentContent, wdExportCreateHeadingBookmarks
    End If

    Application.ScreenUpdating = True
    MsgBox "Kész: PDF-ek elkészültek.", vbInformation
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "Hiba: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub


' ==============================
' Word app + doc kezelõk
' ==============================
Private Function GetOrCreateWordApp(ByVal makeVisible As Boolean) As Object
    Dim wdApp As Object
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    On Error GoTo 0

    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If

    wdApp.Visible = makeVisible
    Set GetOrCreateWordApp = wdApp
End Function

Private Function GetOrOpenWordDoc(ByVal wdApp As Object, ByVal fullPath As String) As Object
    On Error GoTo EH

    Dim d As Object
    For Each d In wdApp.Documents
        If LCase$(d.fullName) = LCase$(fullPath) Then
            Set GetOrOpenWordDoc = d
            Exit Function
        End If
    Next d

    Set GetOrOpenWordDoc = wdApp.Documents.Open(fullPath, ReadOnly:=True)
    Exit Function

EH:
    Set GetOrOpenWordDoc = Nothing
End Function


' ==============================
' PDF export wrapper
' ==============================
Private Sub ExportDocPagesToPDF(ByVal doc As Object, ByVal outPath As String, ByVal pFrom As Long, ByVal pTo As Long, _
                               ByVal wdExportFormatPDF As Long, ByVal wdExportOptimizeForPrint As Long, ByVal wdExportFromTo As Long, _
                               ByVal wdExportDocumentContent As Long, ByVal wdExportCreateHeadingBookmarks As Long)
    doc.ExportAsFixedFormat OutputFileName:=outPath, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportFromTo, _
        From:=pFrom, _
        To:=pTo, _
        Item:=wdExportDocumentContent, _
        IncludeDocProps:=False, _
        KeepIRM:=False, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, _
        DocStructureTags:=True, _
        BitmapMissingFonts:=False, _
        UseISO19005_1:=False
End Sub


' ==============================
' UI segédek (Excel)
' ==============================
Private Function PickWordFile(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = title
        .Filters.clear
        .Filters.add "Word dokumentumok", "*.doc;*.docx;*.docm"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            PickWordFile = ""
        Else
            PickWordFile = .SelectedItems(1)
        End If
    End With
End Function

Private Function PickFolder(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = title
        If .Show <> -1 Then
            PickFolder = ""
        Else
            PickFolder = .SelectedItems(1)
        End If
    End With
End Function

Private Function AskLong(ByVal prompt As String, ByVal title As String, ByVal defValue As Long) As Long
    Dim s As String
    s = InputBox(prompt, title, CStr(defValue))
    s = Trim$(s)
    If s = "" Or Not IsNumeric(s) Then
        AskLong = 0
    Else
        AskLong = CLng(s)
    End If
End Function


' ==============================
' Fájlnév tisztítás
' ==============================
Private Function CleanFileKey(ByVal s As String, ByVal pageNum As Long) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, Chr(10), "")
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(9), "")

    s = Replace(s, "\", "")
    s = Replace(s, "/", "")
    s = Replace(s, ":", "")
    s = Replace(s, "*", "")
    s = Replace(s, "?", "")
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Replace(s, "|", "")
    s = Replace(s, """", "")

    s = Replace(s, " ", "")
    s = Trim$(s)

    If s = "" Then s = "Oldal_" & pageNum
    CleanFileKey = s
End Function

