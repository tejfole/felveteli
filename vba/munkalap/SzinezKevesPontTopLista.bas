Attribute VB_Name = "SzinezKevesPontTopLista"
Sub SzinezzTopEsKevesPontokatRangsorban(Optional control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Dim tbl As ListObject
    Dim ws As Worksheet, wsAdatok As Worksheet
    Dim i As Long, n As Long
    Dim colMagyar As Long, colMatek As Long, colPont As Long
    Dim colRangsor As Long, colSzobeli As Long
    Dim ponthatar As Double, ponthatarInput As String
    Dim vanSzobeli As Boolean
    Dim dbKevesPont As Long, dbTopPont As Long
    Dim adatTomb As Variant
    Dim rangsorRange As Range
    Dim kevesPontSorok As String, topPontSorok As String
    Dim irasbeliPont As Double, mindPont As Double
    Dim startTime As Double
    
    startTime = Timer
    vanSzobeli = False
    dbKevesPont = 0
    dbTopPont = 0
    kevesPontSorok = ""
    topPontSorok = ""

    ' Képernyõfrissítés kikapcsolása
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- Munkalap elérés ---
    On Error Resume Next
    Set wsAdatok = ThisWorkbook.Sheets("adatok")
    On Error GoTo ErrorHandler
    
    If wsAdatok Is Nothing Then
        MsgBox "?? A 'adatok' nevû munkalap nem található!", vbCritical
        GoTo Cleanup
    End If

    ' --- Ponthatár kezelése ---
    If IsEmpty(wsAdatok.Range("A14").value) Or Not IsNumeric(wsAdatok.Range("A14").value) Then
        ponthatarInput = InputBox("Add meg a ponthatárt, amely felett zölddel jelölje a tanulókat:", "Top lista ponthatár", "160")
        If ponthatarInput = "" Or Not IsNumeric(ponthatarInput) Then
            MsgBox "?? Érvénytelen ponthatár!", vbExclamation
            GoTo Cleanup
        End If
        ponthatar = val(ponthatarInput)
        wsAdatok.Range("A14").value = ponthatar
    Else
        ponthatar = val(wsAdatok.Range("A14").value)
    End If

    ' --- diakadat tábla keresése ---
    Set tbl = FindTable("diakadat")
    If tbl Is Nothing Then
        MsgBox "?? A 'diakadat' nevû tábla nem található!", vbCritical
        GoTo Cleanup
    End If
    
    If tbl.ListRows.Count = 0 Then
        MsgBox "?? A táblázat üres!", vbExclamation
        GoTo Cleanup
    End If

    ' Oszlop indexek
    On Error Resume Next
    colMagyar = tbl.ListColumns("p_magyar").Index
    colMatek = tbl.ListColumns("p_matek").Index
    colPont = tbl.ListColumns("p_mindossz").Index
    colRangsor = tbl.ListColumns("rangsor").Index
    colSzobeli = tbl.ListColumns("szobeli").Index
    
    If Err.Number <> 0 Then
        MsgBox "? Hiányzó oszlop(ok) a táblázatban!" & vbCrLf & Err.Description, vbCritical
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler

    ' --- Beolvasás tömbbe ---
    adatTomb = tbl.DataBodyRange.value
    n = UBound(adatTomb, 1)

    ' --- Ellenõrzés: van-e szóbeli pont ---
    For i = 1 To n
        If IsNumeric(adatTomb(i, colSzobeli)) And val(adatTomb(i, colSzobeli)) > 0 Then
            vanSzobeli = True
            Exit For
        End If
    Next i

    ' --- ELSÕ MENET: Összes formázás törlése (GYORS!) ---
    Set rangsorRange = tbl.ListColumns("rangsor").DataBodyRange
    rangsorRange.Interior.ColorIndex = xlNone

    ' --- MÁSODIK MENET: Kategorizálás (mely sorok legyenek színezve) ---
    For i = 1 To n
        irasbeliPont = SafeVal(adatTomb(i, colMagyar)) + SafeVal(adatTomb(i, colMatek))
        mindPont = SafeVal(adatTomb(i, colPont))

        If irasbeliPont < KEVES_IRASBELI_KUSZOB Then
            ' Kevés pont: piros
            If kevesPontSorok <> "" Then kevesPontSorok = kevesPontSorok & ","
            kevesPontSorok = kevesPontSorok & i
            dbKevesPont = dbKevesPont + 1
            
        ElseIf vanSzobeli And mindPont >= ponthatar Then
            ' Top pont: zöld
            If topPontSorok <> "" Then topPontSorok = topPontSorok & ","
            topPontSorok = topPontSorok & i
            dbTopPont = dbTopPont + 1
        End If
    Next i

    ' --- HARMADIK MENET: Tömb-alapú színezés (EGY LÉPÉSBEN!) ---
    ' Kevés pont (piros)
    If kevesPontSorok <> "" Then
        ColorRowsByList rangsorRange, kevesPontSorok, RGB(255, 200, 200)
    End If
    
    ' Top pont (zöld)
    If topPontSorok <> "" Then
        ColorRowsByList rangsorRange, topPontSorok, RGB(200, 255, 200)
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    If Err.Number = 0 Then
        Dim elapsed As Double
        elapsed = Round(Timer - startTime, 3)
        
        MsgBox "? Színezés kész!" & vbCrLf & vbCrLf & _
               "?? Összesítés:" & vbCrLf & _
               "• Írásbeli < 55 pont: " & dbKevesPont & " fõ" & vbCrLf & _
               IIf(vanSzobeli, "• Elérte a ponthatárt (" & ponthatar & "): " & dbTopPont & " fõ", _
               "• Szóbeli nem szerepelt, top lista kihagyva.") & vbCrLf & vbCrLf & _
               "?? Futási idõ: " & elapsed & " mp", vbInformation
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "? Hiba történt: " & Err.Description & vbCrLf & _
           "Hibakód: " & Err.Number, vbCritical
End Sub

' ========== SEGÉDFÜGGVÉNYEK ==========

Private Function FindTable(tableName As String) As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = tableName Then
                Set FindTable = tbl
                Exit Function
            End If
        Next tbl
    Next ws
    
    Set FindTable = Nothing
End Function

Private Function SafeVal(ByVal value As Variant) As Double
    If IsEmpty(value) Then
        SafeVal = 0
    ElseIf IsError(value) Then
        SafeVal = 0
    ElseIf IsNumeric(value) Then
        SafeVal = CDbl(value)
    Else
        SafeVal = val(CStr(value))
    End If
End Function

Private Sub ColorRowsByList(ByVal targetRange As Range, ByVal rowList As String, ByVal color As Long)
    ' Tömb-alapú színezés: vesszõvel elválasztott sorlista alapján
    ' Pl: "1,3,5,7" › 1., 3., 5., 7. sor színezése
    
    On Error Resume Next
    Dim rows() As String
    Dim i As Long
    Dim rowNum As Long
    Dim cellsToColor As Range
    
    ' Sorlista szétválasztása
    rows = Split(rowList, ",")
    
    ' Union használata: összegyûjtjük a cellákat, majd EGY LÉPÉSBEN színezzük
    For i = LBound(rows) To UBound(rows)
        rowNum = CLng(Trim(rows(i)))
        
        If cellsToColor Is Nothing Then
            Set cellsToColor = targetRange.Cells(rowNum, 1)
        Else
            Set cellsToColor = Union(cellsToColor, targetRange.Cells(rowNum, 1))
        End If
        
        ' Excel Union limitet: max 500-1000 tartomány egyszerre
        ' Ha sok sor van, színezzük részletekben
        If (i - LBound(rows) + 1) Mod 500 = 0 Or i = UBound(rows) Then
            If Not cellsToColor Is Nothing Then
                cellsToColor.Interior.color = color
                Set cellsToColor = Nothing
            End If
        End If
    Next i
    
    ' Utolsó csoport színezése
    If Not cellsToColor Is Nothing Then
        cellsToColor.Interior.color = color
    End If
    
    On Error GoTo 0
End Sub

