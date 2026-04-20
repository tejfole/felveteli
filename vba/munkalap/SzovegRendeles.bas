Attribute VB_Name = "SzovegRendeles"
Option Explicit

Private Function NzString(v As Variant) As String
    If IsError(v) Then
        NzString = ""
    ElseIf IsNull(v) Then
        NzString = ""
    Else
        NzString = Trim$(CStr(v & ""))
    End If
End Function

Private Function AddEsIfNumeric(ByVal txt As String) As String
    Dim s As String
    s = Trim$(CStr(txt))

    If Len(s) = 0 Then
        AddEsIfNumeric = ""
        Exit Function
    End If

    If LCase$(Right$(s, 3)) = "-es" Then
        AddEsIfNumeric = s
        Exit Function
    End If

    If IsNumeric(s) Then
        AddEsIfNumeric = s & "-es"
    Else
        AddEsIfNumeric = s
    End If
End Function

Sub SzovegRendelesStrukturaltTablakkal_WriteIfChanged()
    Dim listaT As ListObject, rangsorT As ListObject, szovegekT As ListObject
    Dim wsLista As Worksheet, wsRangsor As Worksheet, wsAdatok As Worksheet
    Dim dictLista As Object, dictRangsor As Object, dictSzovegek As Object
    Dim lr As ListRow, sr As ListRow
    Dim kategoria As String, szoveg As String, indok As String, hatarozat As String
    Dim megszolitInput As String, orommelInput As String, gratulaInput As String
    Dim irasbeli As Double
    Dim tagozat As String, nevelo As String
    Dim nyelv1 As String, nyelv2 As String, nyelvossz As String
    Dim startScreenUpdating As Boolean, startEnableEvents As Boolean
    Dim startCalcMode As XlCalculation

    On Error GoTo ErrHandler

    startScreenUpdating = Application.ScreenUpdating
    startEnableEvents = Application.EnableEvents
    startCalcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsLista = ThisWorkbook.Sheets("lista")
    Set wsRangsor = ThisWorkbook.Sheets("rangsor")
    Set wsAdatok = ThisWorkbook.Sheets("adatok")

    Set listaT = wsLista.ListObjects("lista")
    Set rangsorT = wsRangsor.ListObjects("rangsor")
    Set szovegekT = wsAdatok.ListObjects("szovegek")

    If listaT Is Nothing Or rangsorT Is Nothing Or szovegekT Is Nothing Then
        MsgBox "Nem található valamelyik várt tábla (lista, rangsor vagy szovegek).", vbExclamation
        GoTo Cleanup
    End If

    Set dictLista = CreateObject("Scripting.Dictionary")
    Set dictRangsor = CreateObject("Scripting.Dictionary")
    Set dictSzovegek = CreateObject("Scripting.Dictionary")

    Dim i As Long, h As String
    For i = 1 To listaT.HeaderRowRange.Columns.Count
        h = Trim$(LCase$(CStr(listaT.HeaderRowRange.Cells(1, i).value & "")))
        If Len(h) > 0 Then dictLista(h) = i
    Next i
    For i = 1 To rangsorT.HeaderRowRange.Columns.Count
        h = Trim$(LCase$(CStr(rangsorT.HeaderRowRange.Cells(1, i).value & "")))
        If Len(h) > 0 Then dictRangsor(h) = i
    Next i
    For i = 1 To szovegekT.HeaderRowRange.Columns.Count
        h = Trim$(LCase$(CStr(szovegekT.HeaderRowRange.Cells(1, i).value & "")))
        If Len(h) > 0 Then dictSzovegek(h) = i
    Next i

    Dim mustHave As Variant
    mustHave = Array("nev", "tagozat", "ny_1_nagy", "ny_2", "ny_osszefuz", "ok")
    For i = LBound(mustHave) To UBound(mustHave)
        If Not dictLista.Exists(mustHave(i)) Then
            MsgBox "Hiányzó oszlop a 'lista' táblában: " & mustHave(i), vbExclamation
            GoTo Cleanup
        End If
    Next i

    mustHave = Array("nev", "irasbeliossz", "felvesz", "mastvalaszt", "elut")
    For i = LBound(mustHave) To UBound(mustHave)
        If Not dictRangsor.Exists(mustHave(i)) Then
            MsgBox "Hiányzó oszlop a 'rangsor' táblában: " & mustHave(i), vbExclamation
            GoTo Cleanup
        End If
    Next i

    If Not dictSzovegek.Exists("kategoria") Then
        MsgBox "Hiányzó 'kategoria' oszlop a 'szovegek' táblában.", vbExclamation
        GoTo Cleanup
    End If

    Dim ensureCols As Variant, colName As Variant
    ensureCols = Array("szoveg", "indok", "megszolit", "hatarozat", "orommel", "gratula")
    For Each colName In ensureCols
        If Not dictLista.Exists(LCase$(CStr(colName))) Then
            listaT.ListColumns.add.Name = CStr(colName)
            dictLista(LCase$(CStr(colName))) = listaT.ListColumns(CStr(colName)).Index
        End If
    Next colName

    Dim listaNevCol As Long: listaNevCol = dictLista("nev")
    Dim listaTagCol As Long: listaTagCol = dictLista("tagozat")
    Dim listaNy1Col As Long: listaNy1Col = dictLista("ny_1_nagy")
    Dim listaNy2Col As Long: listaNy2Col = dictLista("ny_2")
    Dim listaNyOsszCol As Long: listaNyOsszCol = dictLista("ny_osszefuz")
    Dim listaOkCol As Long: listaOkCol = dictLista("ok")
    Dim listaSzovegCol As Long: listaSzovegCol = dictLista("szoveg")
    Dim listaIndokCol As Long: listaIndokCol = dictLista("indok")
    Dim listaMegszCol As Long: listaMegszCol = dictLista("megszolit")
    Dim listaHatarCol As Long: listaHatarCol = dictLista("hatarozat")
    Dim listaOromCol As Long: listaOromCol = dictLista("orommel")
    Dim listaGratCol As Long: listaGratCol = dictLista("gratula")

    Dim rangNevCol As Long: rangNevCol = dictRangsor("nev")
    Dim rangIrasCol As Long: rangIrasCol = dictRangsor("irasbeliossz")
    Dim rangFelvCol As Long: rangFelvCol = dictRangsor("felvesz")
    Dim rangMastCol As Long: rangMastCol = dictRangsor("mastvalaszt")
    Dim rangElutCol As Long: rangElutCol = dictRangsor("elut")

    Dim changedCount As Long: changedCount = 0
    Dim changedRows As Long: changedRows = 0

    For Each lr In listaT.ListRows
        Dim nev As String: nev = NzString(lr.Range.Cells(1, listaNevCol).value)
        Dim nevKereso As String: nevKereso = Trim$(LCase$(nev))
        kategoria = ""
        szoveg = "Nincs adat"
        indok = ""
        hatarozat = ""
        megszolitInput = ""
        orommelInput = ""
        gratulaInput = ""
        tagozat = NzString(lr.Range.Cells(1, listaTagCol).value)
        nyelv1 = NzString(lr.Range.Cells(1, listaNy1Col).value)
        nyelv2 = NzString(lr.Range.Cells(1, listaNy2Col).value)
        nyelvossz = NzString(lr.Range.Cells(1, listaNyOsszCol).value)

        For Each sr In rangsorT.ListRows
            If Trim$(LCase$(NzString(sr.Range.Cells(1, rangNevCol).value))) = nevKereso Then
                irasbeli = 0
                If IsNumeric(sr.Range.Cells(1, rangIrasCol).value) Then irasbeli = CDbl(sr.Range.Cells(1, rangIrasCol).value)

                If irasbeli < 70 Then
                    kategoria = "elegtelen"
                Else
                    If LCase$(NzString(sr.Range.Cells(1, rangFelvCol).value)) = "x" Then
                        kategoria = "felvesz"
                    ElseIf LCase$(NzString(sr.Range.Cells(1, rangMastCol).value)) = "x" Then
                        kategoria = "mastvalasz"
                    ElseIf LCase$(NzString(sr.Range.Cells(1, rangElutCol).value)) = "x" Then
                        kategoria = "elut"
                    End If
                End If
                Exit For
            End If
        Next sr

        If kategoria <> "" Then
            If kategoria = "felvesz" Then
                Dim resz1 As String, resz2 As String
                Dim indok1 As String, indok2 As String
                Dim hat1 As String, hat2 As String, hat3 As String

                For Each sr In szovegekT.ListRows
                    If Trim$(LCase$(NzString(sr.Range.Cells(1, dictSzovegek("kategoria")).value))) = "felvesz" Then
                        resz1 = NzString(sr.Range.Cells(1, dictSzovegek("resz1")).value)
                        resz2 = NzString(sr.Range.Cells(1, dictSzovegek("resz2")).value)
                        indok1 = NzString(sr.Range.Cells(1, dictSzovegek("indok1")).value)
                        indok2 = NzString(sr.Range.Cells(1, dictSzovegek("indok2")).value)
                        hat1 = NzString(sr.Range.Cells(1, dictSzovegek("hatarozat1")).value)
                        hat2 = NzString(sr.Range.Cells(1, dictSzovegek("hatarozat2")).value)
                        hat3 = NzString(sr.Range.Cells(1, dictSzovegek("hatarozat3")).value)
                        megszolitInput = NzString(sr.Range.Cells(1, dictSzovegek("megszolit")).value)
                        orommelInput = NzString(sr.Range.Cells(1, dictSzovegek("orommel")).value)
                        gratulaInput = NzString(sr.Range.Cells(1, dictSzovegek("gratula")).value)
                        Exit For
                    End If
                Next sr

                If Trim$(tagozat) = "1000" Then nevelo = "az" Else nevelo = "a"

                szoveg = Trim$(nyelv1 & " " & resz1 & " " & nyelv2 & " " & resz2)
                indok = Trim$(indok1 & " " & nyelvossz & " " & indok2)
                hatarozat = Trim$(nev & " " & hat1 & " " & nevelo & " " & AddEsIfNumeric(tagozat) & " " & hat2 & " " & nyelvossz & " " & hat3)

            Else
                For Each sr In szovegekT.ListRows
                    If Trim$(LCase$(NzString(sr.Range.Cells(1, dictSzovegek("kategoria")).value))) = Trim$(LCase$(kategoria)) Then
                        If kategoria = "elut" And irasbeli >= 70 Then
                            Dim elutResz1 As String, elutResz2 As String, elutasitasOk As String, elutasitasOkEs As String
                            elutResz1 = NzString(sr.Range.Cells(1, dictSzovegek("resz1")).value)
                            elutResz2 = NzString(sr.Range.Cells(1, dictSzovegek("resz2")).value)
                            elutasitasOk = NzString(lr.Range.Cells(1, listaOkCol).value)

                            If InStr(elutasitasOk, "1000") > 0 Then nevelo = "az" Else nevelo = "a"

                            elutasitasOkEs = AddEsIfNumeric(elutasitasOk)
                            szoveg = Trim$(elutResz1 & " " & nevelo & " " & elutasitasOkEs & " " & elutResz2)
                        Else
                            szoveg = NzString(sr.Range.Cells(1, dictSzovegek("resz1")).value)
                        End If
                        Exit For
                    End If
                Next sr
            End If
        End If

        Dim anyChangeThisRow As Boolean: anyChangeThisRow = False

        If kategoria <> "" And szoveg <> "Nincs adat" Then
            If NzString(lr.Range.Cells(1, listaSzovegCol).value) <> szoveg Then lr.Range.Cells(1, listaSzovegCol).value = szoveg: changedCount = changedCount + 1: anyChangeThisRow = True
            If NzString(lr.Range.Cells(1, listaIndokCol).value) <> indok Then lr.Range.Cells(1, listaIndokCol).value = indok: changedCount = changedCount + 1: anyChangeThisRow = True
            If NzString(lr.Range.Cells(1, listaMegszCol).value) <> megszolitInput Then lr.Range.Cells(1, listaMegszCol).value = megszolitInput: changedCount = changedCount + 1: anyChangeThisRow = True
            If NzString(lr.Range.Cells(1, listaHatarCol).value) <> hatarozat Then lr.Range.Cells(1, listaHatarCol).value = hatarozat: changedCount = changedCount + 1: anyChangeThisRow = True
            If NzString(lr.Range.Cells(1, listaOromCol).value) <> orommelInput Then lr.Range.Cells(1, listaOromCol).value = orommelInput: changedCount = changedCount + 1: anyChangeThisRow = True
            If NzString(lr.Range.Cells(1, listaGratCol).value) <> gratulaInput Then lr.Range.Cells(1, listaGratCol).value = gratulaInput: changedCount = changedCount + 1: anyChangeThisRow = True
        Else
            If NzString(lr.Range.Cells(1, listaSzovegCol).value) <> "" Then lr.Range.Cells(1, listaSzovegCol).value = "": changedCount = changedCount + 1: anyChangeThisRow = True
            If NzString(lr.Range.Cells(1, listaIndokCol).value) <> "" Then lr.Range.Cells(1, listaIndokCol).value = "": changedCount = changedCount + 1: anyChangeThisRow = True
            If NzString(lr.Range.Cells(1, listaMegszCol).value) <> "" Then lr.Range.Cells(1, listaMegszCol).value = "": changedCount = changedCount + 1: anyChangeThisRow = True
            If NzString(lr.Range.Cells(1, listaHatarCol).value) <> "" Then lr.Range.Cells(1, listaHatarCol).value = "": changedCount = changedCount + 1: anyChangeThisRow = True
            If NzString(lr.Range.Cells(1, listaOromCol).value) <> "" Then lr.Range.Cells(1, listaOromCol).value = "": changedCount = changedCount + 1: anyChangeThisRow = True
            If NzString(lr.Range.Cells(1, listaGratCol).value) <> "" Then lr.Range.Cells(1, listaGratCol).value = "": changedCount = changedCount + 1: anyChangeThisRow = True
        End If

        If anyChangeThisRow Then changedRows = changedRows + 1
    Next lr

    On Error Resume Next
    ufrKesz.Show
    On Error GoTo 0

    MsgBox "Feldolgozás kész. Módosított cellák: " & changedCount & " (módosított sorok: " & changedRows & ")", vbInformation

Cleanup:
    Application.Calculation = startCalcMode
    Application.EnableEvents = startEnableEvents
    Application.ScreenUpdating = startScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Hiba a feldolgozás során: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub
