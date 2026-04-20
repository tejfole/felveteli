Attribute VB_Name = "IdeiglenesRangsorolas"
Option Explicit

Public Sub Publikalas_AzonositoPont_Sorszam_RangsorSzerint(Optional control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsInput As Worksheet: Set wsInput = wb.Worksheets("diakadat")
    Dim loInput As ListObject: Set loInput = wsInput.ListObjects("diakadat")

    ' --- opcionális név oszlop ---
    Dim includeName As Boolean
    includeName = (MsgBox("Kerüljön NÉV oszlop is a publikált listába?", vbQuestion + vbYesNo, "Publikálás beállítás") = vbYes)

    ' --- p_mindossz tizedesek ---
    Dim decStr As String
    decStr = Trim$(InputBox("Hány tizedesjeggyel jelenjen meg a p_mindossz a publikált lapon? (0-6)", _
                            "Formátum", "2"))

    Dim decimals As Long
    decimals = 2
    If decStr <> "" Then
        If IsNumeric(decStr) Then
            decimals = CLng(decStr)
            If decimals < 0 Then decimals = 0
            If decimals > 6 Then decimals = 6
        End If
    End If

    ' --- minimum pontszám szûrõ (opcionális) ---
    Dim minPontStr As String
    minPontStr = LCase$(Trim$(InputBox( _
        "Minimum p_mindossz, ami bekerül a listába (csak az elfogadottakra)." & vbCrLf & _
        "Hagyd üresen / írd be: mind / összes  › nincs szûrés.", _
        "Minimum pont", "mind")))

    Dim useMinPont As Boolean
    Dim minPont As Double
    useMinPont = False
    minPont = 0

    If minPontStr <> "" And minPontStr <> "mind" And minPontStr <> "osszes" And minPontStr <> "összes" Then
        minPontStr = Replace(minPontStr, ",", ".")
        If IsNumeric(minPontStr) Then
            minPont = CDbl(minPontStr)
            useMinPont = True
        Else
            MsgBox "A minimum pont nem szám: " & minPontStr, vbExclamation
            Exit Sub
        End If
    End If

    ' --- tagozat szûrés ---
    Dim szuresMezo As String
    szuresMezo = LCase$(Trim$(InputBox("Melyik tagozat? (j_1000/j_2000/j_3000/j_4000 vagy mind)", _
                                       "Publikálás – szûrés", "mind")))
    If szuresMezo = "" Then Exit Sub

    Dim ixSzures As Long
    If szuresMezo <> "mind" Then
        ixSzures = LoCol(loInput, szuresMezo)
        If ixSzures = 0 Then
            MsgBox "Nincs ilyen oszlop: " & szuresMezo, vbCritical
            Exit Sub
        End If
    End If

    ' --- cél lapnév ---
    Dim alapNev As String
    If szuresMezo = "mind" Then
        alapNev = "publik"
    Else
        alapNev = Replace(szuresMezo, "j_", "")
    End If

    Dim celLapNevRaw As String
    celLapNevRaw = Trim$(InputBox("Add meg az új munkalap nevét:", "Cél munkalap", alapNev))
    If celLapNevRaw = "" Then Exit Sub

    Dim celLapNev As String
    celLapNev = SanitizeSheetName(celLapNevRaw)

    ' --- kötelezõ oszlopok ---
    Dim ixOkt As Long, ixPont As Long, ixJelige As Long, ixRang As Long
    Dim ixSzobeli As Long, ixIrasbeli As Long, ixBiz As Long
    Dim ixNev As Long

    ixOkt = LoCol(loInput, "oktazon")
    ixPont = LoCol(loInput, "p_mindossz")
    ixJelige = LoCol(loInput, "f_jelige")
    ixRang = LoCol(loInput, "rangsor")

    ixSzobeli = LoCol(loInput, "szobeli")
    ixIrasbeli = LoCol(loInput, "irasbeliossz")
    ixBiz = LoCol(loInput, "p_bizonyitvany")

    ixNev = LoCol(loInput, "f_nev")
    If ixNev = 0 Then ixNev = LoCol(loInput, "a_nev")

    If ixOkt = 0 Or ixPont = 0 Or ixJelige = 0 Or ixRang = 0 Then
        MsgBox "Hiányzó oszlop a diakadat táblában (oktazon/p_mindossz/f_jelige/rangsor).", vbCritical
        Exit Sub
    End If
    If ixSzobeli = 0 Or ixIrasbeli = 0 Or ixBiz = 0 Then
        MsgBox "Hiányzó oszlop az elutasítás ellenõrzéséhez (szobeli/irasbeliossz/p_bizonyitvany).", vbCritical
        Exit Sub
    End If
    If includeName And ixNev = 0 Then
        MsgBox "A név oszlop bekapcsolva, de nincs 'f_nev' (vagy 'nev') oszlop a diakadat táblában.", vbCritical
        Exit Sub
    End If

    If loInput.ListRows.Count = 0 Then
        MsgBox "A diakadat tábla üres.", vbExclamation
        Exit Sub
    End If

    Dim arr As Variant
    arr = loInput.DataBodyRange.value

    Dim maxN As Long: maxN = UBound(arr, 1)

    ' accepted: azonosito, nev, pont, rangsor
    ' rejected: azonosito, nev, "Elutasítva"
    Dim accepted() As Variant, rejected() As Variant
    ReDim accepted(1 To maxN, 1 To 4)
    ReDim rejected(1 To maxN, 1 To 3)

    Dim nAcc As Long: nAcc = 0
    Dim nRej As Long: nRej = 0

    Dim r As Long
    For r = 1 To maxN
        Dim includeRow As Boolean: includeRow = True

        If szuresMezo <> "mind" Then
            includeRow = (LCase$(Trim$(CStr(arr(r, ixSzures)))) = "x")
        End If
        If Not includeRow Then GoTo NextRow

        Dim az As String, nev As String
        az = CleanStr(arr(r, ixJelige))
        If Len(az) = 0 Then az = CleanStr(arr(r, ixOkt))
        If ixNev > 0 Then nev = CleanStr(arr(r, ixNev)) Else nev = ""

        Dim elutasitva As Boolean
        elutasitva = (NzDbl(arr(r, ixSzobeli)) = 0) Or (NzDbl(arr(r, ixIrasbeli)) = 0) Or (NzDbl(arr(r, ixBiz)) = 0)

        If elutasitva Then
            nRej = nRej + 1
            rejected(nRej, 1) = az
            rejected(nRej, 2) = nev
            rejected(nRej, 3) = "Elutasítva"
            GoTo NextRow
        End If

        If useMinPont Then
            If NzDbl(arr(r, ixPont)) < minPont Then GoTo NextRow
        End If

        If Len(Trim$(CStr(arr(r, ixRang)))) = 0 Then GoTo NextRow

        nAcc = nAcc + 1
        accepted(nAcc, 1) = az
        accepted(nAcc, 2) = nev
        accepted(nAcc, 3) = arr(r, ixPont)
        accepted(nAcc, 4) = arr(r, ixRang)

NextRow:
    Next r

    If nAcc = 0 And nRej = 0 Then
        MsgBox "Nincs találat a szûrés/min pont alapján.", vbExclamation
        Exit Sub
    End If

    ' --- output lap csere ---
    Dim wsOutput As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(celLapNev).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set wsOutput = wb.Worksheets.add(After:=wsInput)
    wsOutput.Name = celLapNev

    If includeName Then
        ' A:sorszam B:azonosito C:nev D:p_mindossz E:rangsor(segéd)
        wsOutput.Range("A1:E1").value = Array("sorszam", "azonosito", "nev", "p_mindossz", "rangsor")
        wsOutput.Columns("B").NumberFormat = "@"

        If nAcc > 0 Then
            wsOutput.Range("B2").Resize(nAcc, 4).value = accepted

            With wsOutput.Sort
                .SortFields.clear
                .SortFields.add key:=wsOutput.Range("E2").Resize(nAcc, 1), Order:=xlAscending
                .SortFields.add key:=wsOutput.Range("B2").Resize(nAcc, 1), Order:=xlAscending
                .SetRange wsOutput.Range("A1:E1").Resize(nAcc + 1, 5)
                .Header = xlYes
                .Apply
            End With

            wsOutput.Range("A2").value = 1
            wsOutput.Range("A2").AutoFill Destination:=wsOutput.Range("A2").Resize(nAcc, 1), Type:=xlFillSeries
        End If

        If nRej > 0 Then
            Dim startRejN As Long
            startRejN = 2 + nAcc

            wsOutput.Range("B" & startRejN).Resize(nRej, 1).value = Slice2D(rejected, nRej, 1) ' azonosito
            wsOutput.Range("C" & startRejN).Resize(nRej, 1).value = Slice2D(rejected, nRej, 2) ' nev
            wsOutput.Range("D" & startRejN).Resize(nRej, 1).value = Slice2D(rejected, nRej, 3) ' Elutasítva

            wsOutput.Range("A" & startRejN).Resize(nRej, 1).value = ""
            wsOutput.Range("E" & startRejN).Resize(nRej, 1).value = ""
        End If

        wsOutput.Columns("A").NumberFormat = "0"
        If decimals = 0 Then
            wsOutput.Columns("D").NumberFormat = "0"
        Else
            wsOutput.Columns("D").NumberFormat = "0." & String$(decimals, "0")
        End If

        wsOutput.Columns("E").Delete ' segéd rangsor törlés
        wsOutput.rows(1).Font.Bold = True
        wsOutput.Columns("A:D").AutoFit

    Else
        ' A:sorszam B:azonosito C:p_mindossz D:rangsor(segéd)
        wsOutput.Range("A1:D1").value = Array("sorszam", "azonosito", "p_mindossz", "rangsor")
        wsOutput.Columns("B").NumberFormat = "@"

        If nAcc > 0 Then
            Dim accNoName() As Variant
            accNoName = BuildAcceptedNoName(accepted, nAcc)

            wsOutput.Range("B2").Resize(nAcc, 3).value = accNoName

            With wsOutput.Sort
                .SortFields.clear
                .SortFields.add key:=wsOutput.Range("D2").Resize(nAcc, 1), Order:=xlAscending
                .SortFields.add key:=wsOutput.Range("B2").Resize(nAcc, 1), Order:=xlAscending
                .SetRange wsOutput.Range("A1:D1").Resize(nAcc + 1, 4)
                .Header = xlYes
                .Apply
            End With

            wsOutput.Range("A2").value = 1
            wsOutput.Range("A2").AutoFill Destination:=wsOutput.Range("A2").Resize(nAcc, 1), Type:=xlFillSeries
        End If

        If nRej > 0 Then
            Dim startRej As Long
            startRej = 2 + nAcc

            wsOutput.Range("B" & startRej).Resize(nRej, 1).value = Slice2D(rejected, nRej, 1)
            wsOutput.Range("C" & startRej).Resize(nRej, 1).value = Slice2D(rejected, nRej, 3)

            wsOutput.Range("A" & startRej).Resize(nRej, 1).value = ""
            wsOutput.Range("D" & startRej).Resize(nRej, 1).value = ""
        End If

        wsOutput.Columns("A").NumberFormat = "0"
        If decimals = 0 Then
            wsOutput.Columns("C").NumberFormat = "0"
        Else
            wsOutput.Columns("C").NumberFormat = "0." & String$(decimals, "0")
        End If

        wsOutput.Columns("D").Delete
        wsOutput.rows(1).Font.Bold = True
        wsOutput.Columns("A:C").AutoFit
    End If

    wsOutput.Range("A2").Select
    ActiveWindow.FreezePanes = True

    Dim msg As String
    msg = "Publikálható lista elkészült. Lap: " & wsOutput.Name & vbCrLf & _
          "Rangsorolt (sorszámozott): " & nAcc & " fõ"
    If nRej > 0 Then msg = msg & vbCrLf & "Lista végén (Elutasítva, sorszám nélkül): " & nRej & " fõ"
    MsgBox msg, vbInformation
    Exit Sub

EH:
    MsgBox "Hiba: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Function BuildAcceptedNoName(ByRef accepted As Variant, ByVal nAcc As Long) As Variant
    Dim outArr() As Variant, i As Long
    ReDim outArr(1 To nAcc, 1 To 3)
    For i = 1 To nAcc
        outArr(i, 1) = accepted(i, 1) ' azonosito
        outArr(i, 2) = accepted(i, 3) ' p_mindossz
        outArr(i, 3) = accepted(i, 4) ' rangsor
    Next i
    BuildAcceptedNoName = outArr
End Function

Private Function Slice2D(ByRef arr As Variant, ByVal n As Long, ByVal ColIndex As Long) As Variant
    Dim outArr() As Variant, i As Long
    ReDim outArr(1 To n, 1 To 1)
    For i = 1 To n
        outArr(i, 1) = arr(i, ColIndex)
    Next i
    Slice2D = outArr
End Function

Private Function LoCol(ByVal lo As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    LoCol = lo.ListColumns(colName).Index
    On Error GoTo 0
End Function

Private Function CleanStr(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace$(s, ChrW$(160), " ")
    CleanStr = Trim$(s)
End Function

Private Function NzDbl(ByVal v As Variant) As Double
    If IsError(v) Then
        NzDbl = 0
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        NzDbl = 0
    ElseIf IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = 0
    End If
End Function

Private Function SanitizeSheetName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    s = Trim$(s)

    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "_")
    Next i

    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    Do While InStr(s, "__") > 0
        s = Replace$(s, "__", "_")
    Loop

    If Len(s) = 0 Then s = "publik"
    If Len(s) > 31 Then s = Left$(s, 31)

    SanitizeSheetName = s
End Function

