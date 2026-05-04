Attribute VB_Name = "modListaGeneralas"
Option Explicit

' ============================================================
' LISTA GENERALAS
' - SzuressNev
' - SzuressOktazon
' - SzuressOk
' - FrissitListaTablat
'
' Forras: rangsor tabla
' Cel:    lista tabla
'
' Fontos logika:
' - felvesz: marad felveteli lista
' - elut: azok a j_xxxx tagozatok kerulnek be, ahova jelentkezett,
'         de nem oda lett felveve
' - ha felvesz = x, akkor a rangsor[tagozat] mutatja a felvett tagozatot
' - visszalepett = x es nemteljesitett = x eseten a tanulo csak
'   visszalepett listaban szerepel, nem a nemteljesitett listaban
' - kevespont_hatar alatti tanulo nem szerepelhet a nemteljesitett listaban
' - kevespont / elutkevespont listaba csak az kerul,
'   akinek irasbeliossz < kevespont_hatar
' - kevespont_hatar csak 1..100 kozotti ertekkent ervenyes
'
' JAVITAS (2026-05):
' - Nem transzponalunk 1D tombot 2D-ve (Application.Transpose), mert az elcsuszasokat okoz.
' - A kevéspont szűrésnél a 0 pontot NEM tekintjük "kevéspont"-nak; inkább "nemteljesített" (nincs írásbeli).
' ============================================================

Public Function SzuressNev(Optional ByVal valasztas As String = "", _
                           Optional ByVal SorIndex As Long = 0, _
                           Optional ByVal KeresettErtek As Variant = "x") As Variant
    On Error GoTo Hibakezeles

    Dim arr As Variant
    arr = BuildListaByField("nev", valasztas, KeresettErtek)

    SzuressNev = ReturnArrayItemOrTranspose(arr, SorIndex)
    Exit Function

Hibakezeles:
    SzuressNev = ""
End Function

Public Function SzuressOktazon(Optional ByVal valasztas As String = "", _
                               Optional ByVal SorIndex As Long = 0, _
                               Optional ByVal KeresettErtek As Variant = "x") As Variant
    On Error GoTo Hibakezeles

    Dim arr As Variant
    arr = BuildListaByField("oktazon", valasztas, KeresettErtek)

    SzuressOktazon = ReturnArrayItemOrTranspose(arr, SorIndex)
    Exit Function

Hibakezeles:
    SzuressOktazon = ""
End Function

Public Function SzuressOk(Optional ByVal valasztas As String = "", _
                          Optional ByVal SorIndex As Long = 0, _
                          Optional ByVal KeresettErtek As Variant = "x") As String
    On Error GoTo Hibakezeles

    Dim okok As Variant
    okok = BuildOkLista(valasztas, KeresettErtek)

    If IsEmpty(okok) Then
        SzuressOk = ""
    ElseIf IsArray(okok) Then
        If SorIndex >= LBound(okok) And SorIndex <= UBound(okok) Then
            SzuressOk = CStr(okok(SorIndex))
        Else
            SzuressOk = ""
        End If
    ElseIf SorIndex <= 1 Then
        SzuressOk = CStr(okok)
    Else
        SzuressOk = ""
    End If

    Exit Function

Hibakezeles:
    SzuressOk = ""
End Function

Public Sub FrissitListaTablat(Optional control As IRibbonControl, _
                              Optional ByVal Csendes As Boolean = False)
    On Error GoTo Hibakezeles

    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim valasztas As String

    Set tbl = FindTableByName(ThisWorkbook, "lista")

    If tbl Is Nothing Then
        If Not Csendes Then MsgBox "Nem található a 'lista' nevű tábla.", vbCritical
        Exit Sub
    End If

    Set ws = tbl.Parent
    valasztas = Trim$(CStr(ws.Range("B1").Value))

    If valasztas = "" Then
        If Not Csendes Then MsgBox "A B1 cellában nincs kiválasztott szűrési feltétel.", vbExclamation
        Exit Sub
    End If

    Dim nevek As Variant
    Dim oktazonok As Variant
    Dim darab As Long

    nevek = SzuressNev(valasztas)
    oktazonok = SzuressOktazon(valasztas)

    darab = GetVectorCount(nevek)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ResizeListObjectRows tbl, darab

    If darab = 0 Then
        ClearListaInputColumns tbl

        If Not Csendes Then
            MsgBox "Nincs találat a kiválasztott szűrésre: " & valasztas, vbInformation
        End If

        GoTo CleanExit
    End If

    Dim cNev As Long
    Dim cOktazon As Long
    Dim cOk As Long

    cNev = GetRequiredColIndex(tbl, "nev")
    cOktazon = GetRequiredColIndex(tbl, "oktazon")
    cOk = GetRequiredColIndex(tbl, "ok")

    ClearListColumnContents tbl, cNev
    ClearListColumnContents tbl, cOktazon
    ClearListColumnContents tbl, cOk

    Dim i As Long

    For i = 1 To darab
        tbl.DataBodyRange.Cells(i, cNev).Value = GetVectorItem(nevek, i)
        tbl.DataBodyRange.Cells(i, cOktazon).Value = GetVectorItem(oktazonok, i)
        tbl.DataBodyRange.Cells(i, cOk).Value = SzuressOk(valasztas, i)
    Next i

    If Not Csendes Then
        MsgBox "Lista frissítve. Sorok száma: " & darab, vbInformation
    End If

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Hibakezeles:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Not Csendes Then
        MsgBox "Hiba a lista frissítése közben: " & Err.Number & vbCrLf & Err.Description, vbCritical
    End If
End Sub

Private Function ReturnArrayItemOrTranspose(ByVal arr As Variant, ByVal SorIndex As Long) As Variant
    ' FONTOS: a retArr/okArr 1D tömbként (1..n) készül a BuildSzuresArrays-ban.
    ' Itt NEM transzponálunk, mert az 2D tömböt hozhat létre, ami elcsúszást okozhat.
    On Error GoTo Hiba

    If IsEmpty(arr) Then
        ReturnArrayItemOrTranspose = ""
        Exit Function
    End If

    If IsArray(arr) Then
        Dim lb As Long
        Dim ub As Long

        lb = LBound(arr)
        ub = UBound(arr)

        If ub < lb Then
            ReturnArrayItemOrTranspose = ""
        ElseIf SorIndex <= 0 Then
            ReturnArrayItemOrTranspose = arr
        ElseIf SorIndex >= lb And SorIndex <= ub Then
            ReturnArrayItemOrTranspose = arr(SorIndex)
        Else
            ReturnArrayItemOrTranspose = ""
        End If
    Else
        If SorIndex <= 1 Then
            ReturnArrayItemOrTranspose = CStr(arr)
        Else
            ReturnArrayItemOrTranspose = ""
        End If
    End If

    Exit Function

Hiba:
    ReturnArrayItemOrTranspose = ""
End Function

Private Function BuildListaByField(ByVal ReturnField As String, _
                                   ByVal valasztas As String, _
                                   ByVal KeresettErtek As Variant) As Variant
    On Error GoTo Hibakezeles

    Dim retVals As Variant
    Dim okVals As Variant

    BuildSzuresArrays ReturnField, valasztas, KeresettErtek, retVals, okVals
    BuildListaByField = retVals
    Exit Function

Hibakezeles:
    BuildListaByField = Empty
End Function

Private Function BuildOkLista(ByVal valasztas As String, ByVal KeresettErtek As Variant) As Variant
    On Error GoTo Hibakezeles

    Dim retVals As Variant
    Dim okVals As Variant

    BuildSzuresArrays "nev", valasztas, KeresettErtek, retVals, okVals
    BuildOkLista = okVals
    Exit Function

Hibakezeles:
    BuildOkLista = Empty
End Function

Private Sub BuildSzuresArrays(ByVal ReturnField As String, _
                              ByVal valasztas As String, _
                              ByVal KeresettErtek As Variant, _
                              ByRef retOut As Variant, _
                              ByRef okOut As Variant)
    On Error GoTo Hibakezeles

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim d As Variant

    Set ws = ThisWorkbook.Worksheets("rangsor")
    Set tbl = ws.ListObjects("rangsor")

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    d = tbl.DataBodyRange.Value

    If Len(valasztas) = 0 Then
        On Error Resume Next
        valasztas = NormText(Application.Caller.Worksheet.Range("B1").Value)
        On Error GoTo Hibakezeles
    Else
        valasztas = NormText(valasztas)
    End If

    If IsNemTeljesitettValasztas(valasztas) Then
        valasztas = "nemteljesitett"
    End If

    Dim cRet As Long
    Dim cIras As Long
    Dim cElut As Long
    Dim cVissza As Long
    Dim cFelvesz As Long
    Dim cMastValaszt As Long
    Dim cNemTeljesitett As Long
    Dim cTagozat As Long

    cRet = GetRequiredColIndex(tbl, ReturnField)
    cIras = GetRequiredColIndex(tbl, "irasbeliossz")
    cElut = GetRequiredColIndex(tbl, "elut")
    cVissza = GetRequiredColIndex(tbl, "visszalepett")
    cFelvesz = GetRequiredColIndex(tbl, "felvesz")
    cMastValaszt = GetRequiredColIndex(tbl, "mastvalaszt")
    cNemTeljesitett = GetRequiredColIndex(tbl, "nemteljesitett")
    cTagozat = GetRequiredColIndex(tbl, "tagozat")

    Dim cJ(1 To 4) As Long
    cJ(1) = GetRequiredColIndex(tbl, "j_1000")
    cJ(2) = GetRequiredColIndex(tbl, "j_2000")
    cJ(3) = GetRequiredColIndex(tbl, "j_3000")
    cJ(4) = GetRequiredColIndex(tbl, "j_4000")

    Dim jKod(1 To 4) As String
    jKod(1) = "1000"
    jKod(2) = "2000"
    jKod(3) = "3000"
    jKod(4) = "4000"

    Dim kevesPontLimit As Double
    kevesPontLimit = GetKevespontHatarLocal(55)

    Dim retArr() As String
    Dim okArr() As String
    Dim n As Long
    Dim i As Long
    Dim j As Long

    ReDim retArr(1 To UBound(d, 1) * 8)
    ReDim okArr(1 To UBound(d, 1) * 8)

    Dim keresX As Boolean
    keresX = (NormText(KeresettErtek) = "x")

    For i = 1 To UBound(d, 1)
        Dim retVal As String

        retVal = Trim$(CStr(d(i, cRet)))
        If Len(retVal) = 0 Then GoTo NextI

        Select Case valasztas

            Case "elut"
                If IsX(d(i, cVissza)) Then GoTo NextI

                If IsX(d(i, cElut)) Or IsX(d(i, cFelvesz)) Then
                    Dim felvettTagozat As String
                    felvettTagozat = Trim$(CStr(d(i, cTagozat)))

                    For j = 1 To 4
                        If IsX(d(i, cJ(j))) Then
                            If Not (IsX(d(i, cFelvesz)) And jKod(j) = felvettTagozat) Then
                                n = n + 1
                                retArr(n) = retVal
                                okArr(n) = jKod(j)
                            End If
                        End If
                    Next j
                End If

            Case "elutkevespont"
                ' Ebbe a listába CSAK a valóban kevéspontos tanulók kerülhetnek.
                If IsX(d(i, cVissza)) Then GoTo NextI
                If IsX(d(i, cFelvesz)) Then GoTo NextI

                If IsKevesPontValue(d(i, cIras), kevesPontLimit) Then
                    For j = 1 To 4
                        If IsX(d(i, cJ(j))) Then
                            n = n + 1
                            retArr(n) = retVal
                            okArr(n) = jKod(j)
                        End If
                    Next j
                End If

            Case "kevespont"
                If IsX(d(i, cVissza)) Then GoTo NextI
                If IsX(d(i, cFelvesz)) Then GoTo NextI

                If IsKevesPontValue(d(i, cIras), kevesPontLimit) Then
                    For j = 1 To 4
                        If IsX(d(i, cJ(j))) Then
                            n = n + 1
                            retArr(n) = retVal
                            okArr(n) = jKod(j)
                        End If
                    Next j
                End If

            Case "felvesz"
                If IsX(d(i, cVissza)) Then GoTo NextI

                If keresX Then
                    If IsX(d(i, cFelvesz)) Then
                        n = n + 1
                        retArr(n) = retVal
                        okArr(n) = ""
                    End If
                Else
                    If NormText(d(i, cFelvesz)) = NormText(KeresettErtek) Then
                        n = n + 1
                        retArr(n) = retVal
                        okArr(n) = ""
                    End If
                End If

            Case "mastvalaszt"
                If IsX(d(i, cVissza)) Then GoTo NextI
                If IsX(d(i, cFelvesz)) Then GoTo NextI

                If keresX Then
                    If IsX(d(i, cMastValaszt)) Then
                        For j = 1 To 4
                            If IsX(d(i, cJ(j))) Then
                                n = n + 1
                                retArr(n) = retVal
                                okArr(n) = jKod(j)
                            End If
                        Next j
                    End If
                Else
                    If NormText(d(i, cMastValaszt)) = NormText(KeresettErtek) Then
                        For j = 1 To 4
                            If IsX(d(i, cJ(j))) Then
                                n = n + 1
                                retArr(n) = retVal
                                okArr(n) = jKod(j)
                            End If
                        Next j
                    End If
                End If

            Case "visszalep", "visszalepett"
                If keresX Then
                    If IsX(d(i, cVissza)) Then
                        For j = 1 To 4
                            If IsX(d(i, cJ(j))) Then
                                n = n + 1
                                retArr(n) = retVal
                                okArr(n) = jKod(j)
                            End If
                        Next j
                    End If
                Else
                    If NormText(d(i, cVissza)) = NormText(KeresettErtek) Then
                        For j = 1 To 4
                            If IsX(d(i, cJ(j))) Then
                                n = n + 1
                                retArr(n) = retVal
                                okArr(n) = jKod(j)
                            End If
                        Next j
                    End If
                End If

            Case "nemteljesitett"
                ' Visszalépett tanuló ne legyen nem teljesített listában.
                If IsX(d(i, cVissza)) Then GoTo NextI

                ' Kevéspontos tanuló se legyen nem teljesített listában.
                ' FONTOS: 0 pont NEM számít "kevéspont"-nak (az nálatok inkább "nem teljesített": nem adta meg az írásbelit).
                If IsKevesPontValue(d(i, cIras), kevesPontLimit) Then GoTo NextI

                If keresX Then
                    If IsX(d(i, cNemTeljesitett)) Then
                        For j = 1 To 4
                            If IsX(d(i, cJ(j))) Then
                                n = n + 1
                                retArr(n) = retVal
                                okArr(n) = jKod(j)
                            End If
                        Next j
                    End If
                Else
                    If NormText(d(i, cNemTeljesitett)) = NormText(KeresettErtek) Then
                        For j = 1 To 4
                            If IsX(d(i, cJ(j))) Then
                                n = n + 1
                                retArr(n) = retVal
                                okArr(n) = jKod(j)
                            End If
                        Next j
                    End If
                End If

        End Select

NextI:
    Next i

    If n = 0 Then Exit Sub

    ReDim Preserve retArr(1 To n)
    ReDim Preserve okArr(1 To n)

    retOut = retArr
    okOut = okArr
    Exit Sub

Hibakezeles:
    retOut = Empty
    okOut = Empty
End Sub

Private Function IsNemTeljesitettValasztas(ByVal s As String) As Boolean
    s = NormText(s)
    s = Replace$(s, "-", " ")
    s = Replace$(s, "_", " ")

    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop

    IsNemTeljesitettValasztas = _
        (s = "nem teljesitett") Or _
        (s = "nem teljesített") Or _
        (s = "nemteljesitett") Or _
        (s = "nem teljesítette") Or _
        (s = "nincs teljesitve") Or _
        (s = "nincs teljesítve") Or _
        (s = "hianyos") Or _
        (s = "hiányos")
End Function

Private Function IsKevesPontValue(ByVal v As Variant, ByVal limit As Double) As Boolean
    ' Szigorúbb konverzió: a hibás/üres pontszám ne számítson automatikusan kevés pontnak.
    ' Üzleti szabály: 0 pont NEM kevéspont, az inkább "nem teljesített".
    Dim s As String

    If IsError(v) Or IsNull(v) Then
        IsKevesPontValue = False
        Exit Function
    End If

    s = Trim$(CStr(v & ""))
    If s = "" Then
        IsKevesPontValue = False
        Exit Function
    End If

    s = Replace$(s, ",", ".")
    If Not IsNumeric(s) Then
        IsKevesPontValue = False
        Exit Function
    End If

    Dim num As Double
    num = CDbl(s)

    ' 0 vagy negatív: nálatok inkább "nem teljesített" / nincs írásbeli
    If num <= 0 Then
        IsKevesPontValue = False
        Exit Function
    End If

    IsKevesPontValue = (num < limit)
End Function

Private Function GetKevespontHatarLocal(Optional ByVal defaultValue As Double = 55) As Double
    On Error GoTo EH

    Dim tbl As ListObject
    Set tbl = FindTableByName(ThisWorkbook, "tbl_hatarozat_beallitas")

    If tbl Is Nothing Then
        GetKevespontHatarLocal = defaultValue
        Exit Function
    End If

    If tbl.DataBodyRange Is Nothing Then
        GetKevespontHatarLocal = defaultValue
        Exit Function
    End If

    Dim cKulcs As Long
    Dim cErtek As Long

    cKulcs = tbl.ListColumns("kulcs").Index
    cErtek = tbl.ListColumns("ertek").Index

    Dim r As ListRow
    Dim v As String
    Dim num As Double

    For Each r In tbl.ListRows
        If NormText(r.Range.Cells(1, cKulcs).Value) = "kevespont_hatar" Then
            v = Trim$(CStr(r.Range.Cells(1, cErtek).Value & ""))
            v = Replace$(v, ",", ".")

            If IsNumeric(v) Then
                num = CDbl(v)

                ' Az irasbeliossz írásbeli pont, ezért itt nem fogadunk el
                ' tagozati teljes ponthatárt, pl. 140/150/160.
                If num > 0 And num <= 100 Then
                    GetKevespontHatarLocal = num
                Else
                    GetKevespontHatarLocal = defaultValue
                End If
            Else
                GetKevespontHatarLocal = defaultValue
            End If

            Exit Function
        End If
    Next r

    GetKevespontHatarLocal = defaultValue
    Exit Function

EH:
    GetKevespontHatarLocal = defaultValue
End Function

Private Function FindTableByName(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If LCase$(lo.Name) = LCase$(tableName) Then
                Set FindTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws

    Set FindTableByName = Nothing
End Function

Private Function GetRequiredColIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    On Error GoTo NemTalalhato

    GetRequiredColIndex = tbl.ListColumns(colName).Index
    Exit Function

NemTalalhato:
    Err.Raise vbObjectError + 513, "modListaGeneralas", _
              "Hiányzó oszlop a(z) '" & tbl.Name & "' táblában: '" & colName & "'"
End Function

Private Function NormText(ByVal v As Variant) As String
    Dim s As String

    If IsError(v) Or IsNull(v) Then
        NormText = ""
        Exit Function
    End If

    s = CStr(v)
    s = Replace$(s, ChrW$(160), " ")
    s = Replace$(s, vbTab, " ")
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, ChrW$(8203), "")
    s = Replace$(s, ChrW$(65279), "")
    s = Trim$(s)

    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop

    NormText = LCase$(s)
End Function

Private Function IsX(ByVal v As Variant) As Boolean
    IsX = (NormText(v) = "x")
End Function

Private Sub ResizeListObjectRows(ByVal tbl As ListObject, ByVal targetRows As Long)
    Dim currentRows As Long
    Dim i As Long

    If targetRows < 0 Then targetRows = 0

    If tbl.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = tbl.ListRows.Count
    End If

    If currentRows > targetRows Then
        For i = currentRows To targetRows + 1 Step -1
            tbl.ListRows(i).Delete
        Next i

    ElseIf currentRows < targetRows Then
        For i = currentRows + 1 To targetRows
            tbl.ListRows.Add
        Next i
    End If
End Sub

Private Sub ClearListaInputColumns(ByVal tbl As ListObject)
    On Error Resume Next

    ClearListColumnContents tbl, tbl.ListColumns("nev").Index
    ClearListColumnContents tbl, tbl.ListColumns("oktazon").Index
    ClearListColumnContents tbl, tbl.ListColumns("ok").Index

    On Error GoTo 0
End Sub

Private Sub ClearListColumnContents(ByVal tbl As ListObject, ByVal colIndex As Long)
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    tbl.ListColumns(colIndex).DataBodyRange.ClearContents
End Sub

Private Function GetVectorCount(ByVal v As Variant) As Long
    On Error GoTo Nincs

    If IsEmpty(v) Then
        GetVectorCount = 0
        Exit Function
    End If

    If IsArray(v) Then
        GetVectorCount = UBound(v) - LBound(v) + 1
        Exit Function
    End If

    If Trim$(CStr(v)) = "" Then
        GetVectorCount = 0
    Else
        GetVectorCount = 1
    End If

    Exit Function

Nincs:
    GetVectorCount = 0
End Function

Private Function GetVectorItem(ByVal v As Variant, ByVal index1Based As Long) As String
    On Error GoTo Hiba

    If IsEmpty(v) Then
        GetVectorItem = ""
        Exit Function
    End If

    If IsArray(v) Then
        GetVectorItem = CStr(v(index1Based))
        Exit Function
    End If

    If index1Based = 1 Then
        GetVectorItem = CStr(v)
    Else
        GetVectorItem = ""
    End If

    Exit Function

Hiba:
    GetVectorItem = ""
End Function
