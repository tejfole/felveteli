Attribute VB_Name = "modKozpontiIrasbeliPontokImport"
Option Explicit

Public Sub Import_KozpontiFelveteli_Pontszamok(Optional control As IRibbonControl)
    Dim srcPath As String
    srcPath = PickExcelFile("Válaszd ki a PONTSZÁMOS (forrás) Excel fájlt")
    If srcPath = "" Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim wbD As Workbook: Set wbD = ThisWorkbook
    Dim wsD As Worksheet: Set wsD = wbD.Worksheets("diakadat")
    Dim loD As ListObject: Set loD = wsD.ListObjects("diakadat")

    Dim wbS As Workbook, wsS As Worksheet
    On Error GoTo EH

    Set wbS = Workbooks.Open(srcPath, ReadOnly:=True)
    On Error Resume Next
    Set wsS = wbS.Worksheets("Export")
    On Error GoTo EH
    If wsS Is Nothing Then Set wsS = wbS.Worksheets(1)

    ' --- Cél oszlopindexek (táblában) ---
    Dim mapD As Object: Set mapD = BuildListObjectHeaderMapNorm(loD)
    If Not mapD.Exists(NKey("oktazon")) Then Err.Raise 1001, , "A cél táblában nincs: oktazon"
    If Not mapD.Exists(NKey("p_magyar")) Then Err.Raise 1002, , "A cél táblában nincs: p_magyar"
    If Not mapD.Exists(NKey("p_matek")) Then Err.Raise 1003, , "A cél táblában nincs: p_matek"

    Dim colKeyD As Long: colKeyD = mapD(NKey("oktazon"))
    Dim colPMag As Long: colPMag = mapD(NKey("p_magyar"))
    Dim colPMat As Long: colPMat = mapD(NKey("p_matek"))

    ' --- Cél index: oktazon -> ListRow.Index ---
    Dim idxD As Object: Set idxD = CreateObject("Scripting.Dictionary")
    BuildDestIndex loD, colKeyD, idxD

    ' --- Forrás: két fejléc sor ---
    Dim headerRowGroup As Long: headerRowGroup = 1
    Dim headerRowSub As Long: headerRowSub = 2

    ' Oktazon fejléc neve bekérés (a 2. sorban keressük)
    Dim srcKeyHeader As String
    srcKeyHeader = InputBox("Forrás kulcs oszlop FEJLÉCE (2. fejléc sor):", "Kulcs kiválasztás", "Oktatási azonosító")
    If Trim$(srcKeyHeader) = "" Then GoTo CleanExit

    ' Forrás oszlopok felderítése (csoport+alfejléc)
    Dim colKeyS As Long
    colKeyS = FindSubHeaderCol(wsS, headerRowSub, srcKeyHeader)
    If colKeyS = 0 Then Err.Raise 2001, , "Nem találom a forrás kulcs oszlopot (2. sor): " & srcKeyHeader

    Dim groupName As String: groupName = "Központi felvételi eredmények"
    Dim colMagS As Long, colMatS As Long
    colMagS = FindGroupedCol(wsS, headerRowGroup, headerRowSub, groupName, "Magyar nyelv elért pontszám")
    colMatS = FindGroupedCol(wsS, headerRowGroup, headerRowSub, groupName, "Matematika elért pontszám")

    If colMagS = 0 Then Err.Raise 2002, , "Nem találom: [" & groupName & "] / [Magyar nyelv elért pontszám]"
    If colMatS = 0 Then Err.Raise 2003, , "Nem találom: [" & groupName & "] / [Matematika elért pontszám]"

    ' --- Forrás bejárás ---
    Dim lastRow As Long
    lastRow = wsS.Cells(wsS.rows.Count, colKeyS).End(xlUp).Row
    If lastRow < headerRowSub + 1 Then GoTo CleanExit

    Dim seenS As Object: Set seenS = CreateObject("Scripting.Dictionary") ' forrás duplák
    Dim dupReport As String: dupReport = ""
    Dim missingReport As String: missingReport = ""

    Dim updCount As Long: updCount = 0
    Dim missCount As Long: missCount = 0
    Dim dupCount As Long: dupCount = 0

    Dim r As Long
    For r = headerRowSub + 1 To lastRow
        Dim k As String
        k = Trim$(CStr(wsS.Cells(r, colKeyS).value))
        If k = "" Then GoTo NextR

        ' Forrás duplák
        If seenS.Exists(k) Then
            dupCount = dupCount + 1
            If dupCount <= 30 Then dupReport = dupReport & "• " & k & " (sor " & r & ", már volt: sor " & seenS(k) & ")" & vbCrLf
            GoTo NextR ' preferáljuk az elsõt
        Else
            seenS(k) = r
        End If

        ' Célban megvan?
        If Not idxD.Exists(k) Then
            missCount = missCount + 1
            If missCount <= 30 Then missingReport = missingReport & "• " & k & " (forrás sor " & r & ")" & vbCrLf
            GoTo NextR
        End If

        Dim lr As ListRow
        Set lr = loD.ListRows(idxD(k))

        ' értékek (pontszámok)
        Dim vMag As Variant, vMat As Variant
        vMag = wsS.Cells(r, colMagS).value
        vMat = wsS.Cells(r, colMatS).value

        ' ha üres, hagyjuk békén; ha van szám, írjuk be
        If Trim$(CStr(vMag)) <> "" Then lr.Range.Cells(1, colPMag).value = vMag
        If Trim$(CStr(vMat)) <> "" Then lr.Range.Cells(1, colPMat).value = vMat

        updCount = updCount + 1

NextR:
    Next r

    wbD.Save

    Dim msg As String
    msg = "Pontszám import kész." & vbCrLf & _
          "Frissített rekordok: " & updCount & vbCrLf & _
          "Forrás duplák: " & dupCount & vbCrLf & _
          "Célban nem található oktazon: " & missCount

    If dupReport <> "" Then msg = msg & vbCrLf & vbCrLf & "Duplikált oktazonok a forrásban (elsõt vettük):" & vbCrLf & dupReport
    If missingReport <> "" Then msg = msg & vbCrLf & vbCrLf & "Forrásban van, célban nincs:" & vbCrLf & missingReport

    MsgBox msg, vbInformation

    ' >>> AUTOMATIKUS ÚJRASZÁMOLÁS (írásbeli pont import után) <<<
    If updCount > 0 Then
        RecalcPontok_Automatikus
    End If
    
CleanExit:
    On Error Resume Next
    wbS.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

EH:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Hiba: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume CleanExit
    
End Sub

' =========================
' FORRÁS FEJLÉC KERESÕK (2 sor + merged group)
' =========================

' 2. fejléc sorban keres "pontos" (normalizált) egyezéssel
Private Function FindSubHeaderCol(ws As Worksheet, headerRowSub As Long, headerText As String) As Long
    Dim lastCol As Long: lastCol = ws.Cells(headerRowSub, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If NKey(CStr(ws.Cells(headerRowSub, c).value)) = NKey(headerText) Then
            FindSubHeaderCol = c
            Exit Function
        End If
    Next c
    FindSubHeaderCol = 0
End Function

' 1. sor (group) + 2. sor (sub) alapján keres, a group cella lehet merged
Private Function FindGroupedCol(ws As Worksheet, headerRowGroup As Long, headerRowSub As Long, groupHeader As String, subHeader As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(headerRowSub, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        Dim g As String, s As String
        g = GroupHeaderText(ws, headerRowGroup, c)
        s = CStr(ws.Cells(headerRowSub, c).value)

        If NKey(g) = NKey(groupHeader) And NKey(s) = NKey(subHeader) Then
            FindGroupedCol = c
            Exit Function
        End If
    Next c

    FindGroupedCol = 0
End Function

' Visszaadja az 1. fejléc sor csoportszövegét; ha merged, a merge area bal-felsõ celláját
Private Function GroupHeaderText(ws As Worksheet, headerRowGroup As Long, col As Long) As String
    Dim cell As Range: Set cell = ws.Cells(headerRowGroup, col)
    If cell.MergeCells Then
        GroupHeaderText = CStr(cell.MergeArea.Cells(1, 1).value)
    Else
        GroupHeaderText = CStr(cell.value)
    End If
End Function

' =========================
' CÉL TÁBLA SEGÉDEK
' =========================

Private Function BuildListObjectHeaderMapNorm(lo As ListObject) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        d(NKey(lo.ListColumns(i).Name)) = i
    Next i
    Set BuildListObjectHeaderMapNorm = d
End Function

Private Sub BuildDestIndex(lo As ListObject, keyColIndex As Long, idx As Object)
    idx.RemoveAll
    If lo.ListRows.Count = 0 Then Exit Sub

    Dim i As Long, k As String
    For i = 1 To lo.ListRows.Count
        k = Trim$(CStr(lo.DataBodyRange.Cells(i, keyColIndex).value))
        If k <> "" Then idx(k) = i
    Next i
End Sub

' =========================
' NORMALIZÁLÓ
' =========================
Private Function NKey(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace(t, ChrW(160), " ")
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    t = Replace(t, "-", " ")
    t = Replace(t, "—", " ")
    t = Replace(t, "–", " ")

    t = Replace(t, "á", "a")
    t = Replace(t, "é", "e")
    t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o")
    t = Replace(t, "ö", "o")
    t = Replace(t, "õ", "o")
    t = Replace(t, "ú", "u")
    t = Replace(t, "ü", "u")
    t = Replace(t, "û", "u")

    NKey = t
End Function

' =========================
' FILE PICKER
' =========================
Private Function PickExcelFile(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = title
        .Filters.clear
        .Filters.add "Excel fájlok", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        If .Show <> -1 Then PickExcelFile = "" Else PickExcelFile = .SelectedItems(1)
    End With
End Function


