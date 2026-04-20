Attribute VB_Name = "jelszogeneralasCsvfilekeszites"
Option Explicit

Sub GeneratePasswordsFromTableAndExportCSV_Final_UniqueWithLogClean(Optional control As IRibbonControl)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As ListRow
    Dim aNev As String
    Dim szulI As String
    Dim jelszo As String
    Dim fileNo As Integer
    Dim logNo As Integer
    Dim folderPath As String
    Dim csvFile As String
    Dim logFile As String
    Dim lineText As String
    Dim fd As FileDialog
    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")

    ' FIX: diakadat tábla használata
    Set tbl = GetTableByName("diakadat")
    If tbl Is Nothing Then
        MsgBox "Nem található a 'diakadat' nevû tábla a munkafüzetben!", vbCritical
        Exit Sub
    End If
    Set ws = tbl.parent

    On Error GoTo HianyzoOszlop
    Dim aNevCol As ListColumn
    Dim szulICol As ListColumn
    Dim jelszoCol As ListColumn
    Dim oktazonCol As ListColumn

    Set aNevCol = tbl.ListColumns("f_a_nev")
    Set szulICol = tbl.ListColumns("f_szul_ido")
    Set jelszoCol = tbl.ListColumns("jelszo")
    Set oktazonCol = tbl.ListColumns("oktazon")
    On Error GoTo 0

    ' Jelszavak generálása a DIAKADAT táblában
    For Each r In tbl.ListRows
        Dim oktazonValue As String
        oktazonValue = Trim(CStr(r.Range(1, oktazonCol.Index).value))

        If oktazonValue = "" Then
            r.Range(1, jelszoCol.Index).value = ""
        Else
            aNev = CleanName(CStr(r.Range(1, aNevCol.Index).value))
            szulI = CStr(r.Range(1, szulICol.Index).value)

            If IsDate(szulI) Then
                szulI = Format(CDate(szulI), "yyyymmdd")
            Else
                szulI = DigitsOnly(szulI)
            End If

            If Len(aNev) >= 3 And Len(szulI) = 8 Then
                jelszo = Left$(aNev, 3) & szulI
                r.Range(1, jelszoCol.Index).value = jelszo
            Else
                r.Range(1, jelszoCol.Index).value = "HIBA"
            End If
        End If
    Next r

    ' Mappa választás
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = "Válassz mappát a CSV és log fájl mentéséhez"
    If fd.Show = -1 Then
        folderPath = fd.SelectedItems(1)
    Else
        MsgBox "Mappaválasztás megszakítva!", vbExclamation
        Exit Sub
    End If

    csvFile = folderPath & "\jelszavak.csv"
    logFile = folderPath & "\hibas_sorok_log.txt"

    fileNo = FreeFile
    Open csvFile For Output As #fileNo

    logNo = FreeFile
    Open logFile For Output As #logNo

    Print #fileNo, "fajlnev;jelszo"

    ' Export a DIAKADAT táblából
    For Each r In tbl.ListRows
        Dim oktazonExport As String
        Dim jelszoExport As String

        oktazonExport = Trim(CStr(r.Range(1, oktazonCol.Index).value))
        jelszoExport = Trim(CStr(r.Range(1, jelszoCol.Index).value))

        If oktazonExport = "" Then GoTo SkipNext
        If dict.Exists(oktazonExport) Then GoTo SkipNext

        If jelszoExport = "" Or UCase$(jelszoExport) = "HIBA" Then
            Print #logNo, "Hibás sor - Oktazon: " & oktazonExport & ", Jelszó: " & jelszoExport
            GoTo SkipNext
        End If

        lineText = oktazonExport & ";" & jelszoExport
        Print #fileNo, lineText
        dict.add oktazonExport, True

SkipNext:
    Next r

    Close #fileNo
    Close #logNo

    MsgBox "CSV és hibás sorok log sikeresen elkészült! (forrás: diakadat)", vbInformation
    Exit Sub

HianyzoOszlop:
    MsgBox "Hiányzik valamelyik szükséges oszlop a 'diakadat' táblából: 'f_a_nev', 'f_szul_ido', 'jelszo', 'oktazon'!", vbCritical
End Sub

Private Function GetTableByName(ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase$(lo.Name) = LCase$(tableName) Then
                Set GetTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set GetTableByName = Nothing
End Function

Private Function CleanName(ByVal szoveg As String) As String
    szoveg = LCase$(Trim$(szoveg))

    If Left$(szoveg, 3) = "dr." Then
        szoveg = Trim$(mid$(szoveg, 4))
    End If

    szoveg = Replace(szoveg, "á", "a")
    szoveg = Replace(szoveg, "é", "e")
    szoveg = Replace(szoveg, "í", "i")
    szoveg = Replace(szoveg, "ó", "o")
    szoveg = Replace(szoveg, "ö", "o")
    szoveg = Replace(szoveg, "õ", "o")
    szoveg = Replace(szoveg, "ú", "u")
    szoveg = Replace(szoveg, "ü", "u")
    szoveg = Replace(szoveg, "û", "u")

    szoveg = Replace(szoveg, " ", "")
    szoveg = Replace(szoveg, "-", "")
    szoveg = Replace(szoveg, ".", "")
    szoveg = Replace(szoveg, "'", "")

    CleanName = szoveg
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, t As String
    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        If ch Like "#" Then t = t & ch
    Next i
    DigitsOnly = t
End Function
