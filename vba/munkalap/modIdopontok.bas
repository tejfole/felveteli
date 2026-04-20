Attribute VB_Name = "modIdopontok"
Option Explicit

' Belépési pont a Sheet modulból
Public Sub AssignDatumNap_FromIdopontTabla( _
    ByVal loD As ListObject, _
    ByVal rowIdx As Long, _
    ByVal biz As Long, _
    ByVal kapacitas As Long _
)
    On Error GoTo Fail

    Dim wsT As Worksheet
    Dim loT As ListObject

    Set wsT = ThisWorkbook.Worksheets("idopontok")
    Set loT = wsT.ListObjects("tbl_idopontok")
    If loT Is Nothing Then Exit Sub

    If loT.ListRows.Count = 0 Or loT.DataBodyRange Is Nothing Then
        MsgBox "Nincs idõpont a tbl_idopontok táblában. Vegyél fel idõpontokat az Idõpontok menübõl.", vbExclamation
        Exit Sub
    End If

    If loD Is Nothing Then Exit Sub
    If loD.DataBodyRange Is Nothing Then Exit Sub
    If rowIdx < 1 Or rowIdx > loD.ListRows.Count Then Exit Sub

    ' ===== 0) KEVÉS ÍRÁSBELI (felülbírálható) =====
    Dim colIras As Long, irasPont As Double
    colIras = GetLoColIndex(loD, "irasbeliossz")

    If colIras > 0 Then
        irasPont = val(loD.DataBodyRange.Cells(rowIdx, colIras).value)

        If irasPont < KEVES_IRASBELI_KUSZOB Then
            Dim nev As String, okt As String
            nev = CStr(GetCellValueSafe(loD, rowIdx, "f_nev"))
            okt = CStr(GetCellValueSafe(loD, rowIdx, "oktazon"))

            Dim msg As String
            msg = "Kevés írásbeli pont (" & KEVES_IRASBELI_KUSZOB & " alatt) – alapból nem kap idõpontot." & vbCrLf & vbCrLf & _
                  "Név: " & nev & vbCrLf & _
                  "Oktazon: " & okt & vbCrLf & _
                  "Írásbeli: " & irasPont & vbCrLf & vbCrLf & _
                  "Felülbírálod és mégis kiosztod az idõpontot?"

            If MsgBox(msg, vbExclamation + vbYesNo, "Kevés pont") = vbNo Then
                Exit Sub
            Else
                Call SetCellValueSafe(loD, rowIdx, "megjegyzes", "Kevés írásbeli – felülbírálva")
            End If
        End If
    End If
    ' ===== /KEVÉS ÍRÁSBELI =====

    ' ===== 1) aktív idõpontok listája =====
    Dim arrT As Variant
    arrT = loT.DataBodyRange.value

    Dim iDtT As Long, iAkT As Long
    iDtT = GetLoColIndex(loT, "datum_nap")
    iAkT = GetLoColIndex(loT, "aktiv")

    If iDtT = 0 Or iAkT = 0 Then
        MsgBox "Hiányzó oszlop a tbl_idopontok táblában: datum_nap és/vagy aktiv.", vbExclamation
        Exit Sub
    End If

    ' ===== 2) foglaltság számolás a diakadatból =====
    Dim iBizD As Long, iDtD As Long
    iBizD = GetLoColIndex(loD, "bizottsag")
    iDtD = GetLoColIndex(loD, "datum_nap")
    If iDtD = 0 Then iDtD = GetLoColIndex(loD, "idopont_nap") ' fallback

    If iBizD = 0 Or iDtD = 0 Then
        MsgBox "Hiányzó oszlop a diakadat táblában: bizottsag és/vagy datum_nap (idopont_nap).", vbExclamation
        Exit Sub
    End If

    Dim arrD As Variant
    arrD = loD.DataBodyRange.value

    Dim activeDates As Collection
    Set activeDates = New Collection

    Dim r As Long
    For r = 1 To UBound(arrT, 1)
        If CLng(val(arrT(r, iAkT))) = 1 Then
            Dim dtTmp As Date
            If TryParseDateTimeAny(arrT(r, iDtT), dtTmp) Then
                activeDates.add dtTmp
            End If
        End If
    Next r

    If activeDates.Count = 0 Then
        MsgBox "Nincs AKTÍV idõpont a tbl_idopontok táblában.", vbExclamation
        Exit Sub
    End If

    ' ===== 3) lista: idõpont + szabad hely =====
    Dim items() As String, keys() As String
    ReDim items(1 To activeDates.Count)
    ReDim keys(1 To activeDates.Count)

    Dim i As Long
    For i = 1 To activeDates.Count
        Dim dt As Date
        dt = activeDates(i)

        Dim used As Long
        used = CountAssignedInArr(arrD, biz, dt, iBizD, iDtD)

        Dim free As Long
        free = kapacitas - used

        keys(i) = CStr(CDbl(dt))
        items(i) = Format$(dt, "yyyy.mm.dd hh:nn:ss") & "   (szabad: " & free & ")"
    Next i

    ' ===== 4) csak szabad idõpontok =====
    Dim items2() As String, keys2() As String, n As Long
    ReDim items2(1 To UBound(items))
    ReDim keys2(1 To UBound(items))

    For i = 1 To UBound(items)
        If ExtractFree(items(i)) > 0 Then
            n = n + 1
            items2(n) = items(i)
            keys2(n) = keys(i)
        End If
    Next i

    If n = 0 Then
        MsgBox "Nincs szabad hely egyik aktív idõpontban sem ennél a bizottságnál (" & biz & ").", vbExclamation
        Exit Sub
    End If

    ReDim Preserve items2(1 To n)
    ReDim Preserve keys2(1 To n)

    ' ===== 5) választás InputBox-szal =====
    Dim pick As Long
    pick = ChooseIndexFromList("Idõpont választás (Bizottság " & biz & ")", items2)
    If pick = 0 Then Exit Sub

    Dim dtChosen As Date
    dtChosen = CDate(CDbl(keys2(pick)))

    ' Újraellenõrzés
    If CountAssignedInArr(arrD, biz, dtChosen, iBizD, iDtD) >= kapacitas Then
        MsgBox "Közben betelt ez az idõpont. Válassz másikat.", vbExclamation
        Exit Sub
    End If

    ' ===== 6) beírás =====
    loD.DataBodyRange.Cells(rowIdx, iDtD).value = dtChosen
    loD.DataBodyRange.Cells(rowIdx, iDtD).NumberFormat = "yyyy.mm.dd hh:mm:ss"

    Exit Sub

Fail:
    MsgBox "Hiba az idõpont kiosztás közben: " & Err.Description, vbExclamation
End Sub

' ===== segédek =====

Private Function GetLoColIndex(ByVal lo As ListObject, ByVal colName As String) As Long
    On Error GoTo Fail
    GetLoColIndex = lo.ListColumns(colName).Index
    Exit Function
Fail:
    GetLoColIndex = 0
End Function

Private Function GetCellValueSafe(ByVal lo As ListObject, ByVal rowIdx As Long, ByVal colName As String) As Variant
    Dim ix As Long
    ix = GetLoColIndex(lo, colName)
    If ix = 0 Then Exit Function
    GetCellValueSafe = lo.DataBodyRange.Cells(rowIdx, ix).value
End Function

Private Sub SetCellValueSafe(ByVal lo As ListObject, ByVal rowIdx As Long, ByVal colName As String, ByVal v As Variant)
    Dim ix As Long
    ix = GetLoColIndex(lo, colName)
    If ix = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIdx, ix).value = v
End Sub

Private Function ExtractFree(ByVal s As String) As Long
    Dim p As Long: p = InStr(1, s, "szabad:", vbTextCompare)
    If p = 0 Then Exit Function
    Dim t As String: t = mid$(s, p + Len("szabad:"))
    t = Replace(t, ")", "")
    ExtractFree = CLng(val(Trim$(t)))
End Function

' FONTOS: legyen Public, hogy biztosan elérhetõ legyen minden modulból
Public Function ChooseIndexFromList(ByVal title As String, ByRef items() As String) As Long
    Dim lb As Long, ub As Long

    On Error Resume Next
    lb = LBound(items): ub = UBound(items)
    If Err.Number <> 0 Then
        Err.clear
        Exit Function
    End If
    On Error GoTo 0

    Dim msg As String, i As Long
    For i = lb To ub
        msg = msg & i & ". " & items(i) & vbCrLf
    Next i

    Dim ans As String
    ans = Trim$(InputBox(msg, title, CStr(lb)))
    If ans = "" Or Not IsNumeric(ans) Then Exit Function

    Dim n As Long: n = CLng(ans)
    If n < lb Or n > ub Then Exit Function

    ChooseIndexFromList = n
End Function

Private Function CountAssignedInArr(ByRef arrD As Variant, ByVal biz As Long, ByVal dt As Date, ByVal iBizD As Long, ByVal iDtD As Long) As Long
    Dim r As Long, cnt As Long, dtTmp As Date

    If IsEmpty(arrD) Then Exit Function

    For r = 1 To UBound(arrD, 1)
        If CLng(val(arrD(r, iBizD))) = biz Then
            If TryParseDateTimeAny(arrD(r, iDtD), dtTmp) Then
                If CDbl(dtTmp) = CDbl(dt) Then cnt = cnt + 1
            End If
        End If
    Next r

    CountAssignedInArr = cnt
End Function

Private Function TryParseDateTimeAny(ByVal v As Variant, ByRef dtOut As Date) As Boolean
    On Error GoTo Fail

    If IsDate(v) Then
        dtOut = CDate(v)
        TryParseDateTimeAny = True
        Exit Function
    End If

    Dim s As String: s = Trim$(CStr(v))
    If s = "" Then GoTo Fail

    Dim parts() As String, d() As String, t() As String
    parts = Split(s, " ")
    d = Split(parts(0), ".")
    If UBound(d) < 2 Then GoTo Fail

    Dim yyyy As Long, mm As Long, dd As Long
    yyyy = CLng(val(d(0))): mm = CLng(val(d(1))): dd = CLng(val(d(2)))
    If yyyy < 1900 Or mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then GoTo Fail

    Dim hh As Long, nn As Long, ss As Long
    hh = 0: nn = 0: ss = 0

    If UBound(parts) >= 1 Then
        t = Split(parts(1), ":")
        hh = CLng(val(t(0)))
        If UBound(t) >= 1 Then nn = CLng(val(t(1)))
        If UBound(t) >= 2 Then ss = CLng(val(t(2)))
    End If

    dtOut = DateSerial(yyyy, mm, dd) + TimeSerial(hh, nn, ss)
    TryParseDateTimeAny = True
    Exit Function

Fail:
    TryParseDateTimeAny = False
End Function

