Attribute VB_Name = "RangsorFeltoltesDiakadatbol"
Sub MasolasDiakadatbolRangsorba(Optional control As IRibbonControl)
    Dim wb As Workbook
    Dim diakTbl As ListObject, rangsorTbl As ListObject, tbl As ListObject
    Dim i As Long, j As Long
    Dim forrasAdatok As Variant
    Dim atnevezes As Object
    Dim ws As Worksheet
    Dim rangsorOktazonIndex As Long, diakOktazonIndex As Long
    Dim rangsorData As Variant
    Dim oktazonDict As Object
    Dim logSorok As String
    Dim logFSO As Object, logFile As Object, logUt As String

    Set wb = ThisWorkbook

    ' Tábla beazonosítás
    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = "diakadat" Then Set diakTbl = tbl
            If tbl.Name = "rangsor" Then Set rangsorTbl = tbl
        Next tbl
    Next ws

    If diakTbl Is Nothing Or rangsorTbl Is Nothing Then
        MsgBox "Nem található valamelyik tábla (diakadat vagy rangsor)!", vbCritical
        Exit Sub
    End If

    ' Átmásolható oszlopok (forrás › cél)
    Set atnevezes = CreateObject("Scripting.Dictionary")
    atnevezes.add "f_nev", "nev"
    atnevezes.add "oktazon", "oktazon"
    atnevezes.add "irasbeliossz", "irasbeliossz"
    atnevezes.add "p_mindossz", "p_mindossz"
    atnevezes.add "j_1000", "j_1000"
    atnevezes.add "j_2000", "j_2000"
    atnevezes.add "j_3000", "j_3000"
    atnevezes.add "j_4000", "j_4000"

    ' Oktazon azonosítók elõkészítése rangsor táblából
    Set oktazonDict = CreateObject("Scripting.Dictionary")
    rangsorData = rangsorTbl.DataBodyRange.value
    rangsorOktazonIndex = rangsorTbl.ListColumns("oktazon").Index

    For i = 1 To UBound(rangsorData, 1)
        Dim oktazonKey As String
        oktazonKey = Trim(CStr(rangsorData(i, rangsorOktazonIndex)))
        If oktazonKey <> "" Then
            oktazonDict(oktazonKey) = i ' sorindex mentése
        End If
    Next i

    ' Másolás soronként
    forrasAdatok = diakTbl.DataBodyRange.value
    diakOktazonIndex = diakTbl.ListColumns("oktazon").Index

    For i = 1 To UBound(forrasAdatok, 1)
        Dim diakOktazon As String, jelolSarga As Boolean
        diakOktazon = Trim(CStr(forrasAdatok(i, diakOktazonIndex)))
        jelolSarga = False

        Dim celSor As Range

        If diakOktazon = "" Then
            ' Hiányzó oktazon: új sor, sárga jelölés + log
            Set celSor = rangsorTbl.ListRows.add.Range
            jelolSarga = True
            logSorok = logSorok & "Hiányzó oktazon a diakadat sor " & i & vbCrLf
        ElseIf oktazonDict.Exists(diakOktazon) Then
            ' létezõ sor › frissítés
            Set celSor = rangsorTbl.DataBodyRange.rows(oktazonDict(diakOktazon))
        Else
            ' új sor, oktazon alapján
            Set celSor = rangsorTbl.ListRows.add.Range
            oktazonDict.add diakOktazon, celSor.Row - rangsorTbl.HeaderRowRange.Row
        End If

        ' Elõször visszaállítjuk az alap színt
        celSor.Interior.ColorIndex = xlNone

        ' Csak a megengedett oszlopokat írjuk
        Dim kulcs As Variant
        For Each kulcs In atnevezes.keys
            Dim forrasCol As Long, celCol As Long
            On Error Resume Next
            forrasCol = diakTbl.ListColumns(kulcs).Index
            celCol = rangsorTbl.ListColumns(atnevezes(kulcs)).Index
            On Error GoTo 0
            If forrasCol > 0 And celCol > 0 Then
                celSor.Cells(1, celCol).value = forrasAdatok(i, forrasCol)
            End If
        Next kulcs

        ' Sárga háttér, ha szükséges
        If jelolSarga Then
            celSor.Interior.color = RGB(255, 255, 0)
        End If
    Next i

    ' Log fájl mentése, ha volt hiba
    If logSorok <> "" Then
        logUt = wb.path & "\rangsor_masolas_log.txt"
        Set logFSO = CreateObject("Scripting.FileSystemObject")
        Set logFile = logFSO.CreateTextFile(logUt, True, True)
        logFile.Write "Figyelmeztetések - " & Format(Now, "yyyy-mm-dd HH:MM:SS") & vbCrLf & logSorok
        logFile.Close
        MsgBox "Másolás kész. Figyelmeztetések naplózva: " & vbCrLf & logUt, vbExclamation
    Else
        MsgBox "Másolás kész. Nem volt hiba!", vbInformation
    End If
End Sub

