Attribute VB_Name = "modIdopontMenu"
Option Explicit

' Ribbonból is hívható
Public Sub Idopontok_Menu(Optional control As IRibbonControl)
    On Error GoTo EH

    Dim lo As ListObject: Set lo = GetIdopontTabla_V2()
    If lo Is Nothing Then
        MsgBox "Nem találom az idõpont táblát: idopontok / tbl_idopontok", vbExclamation
        Exit Sub
    End If

    Dim choice As String
    Do
        choice = InputBox( _
            BuildMenuText(lo), _
            "Idõpontok menü", _
            "1")
        choice = Trim$(choice)
        If choice = "" Then Exit Sub

        Select Case choice
            Case "1": Idopontok_Listaz lo
            Case "2": Idopontok_UjFelvitel lo
            Case "3": Idopontok_AktivToggle lo
            Case "4": Idopontok_Torles lo
            Case "5": Idopontok_DuplikatumTakaritas lo
            Case "6": Idopontok_MindenInaktiv lo
            Case "7": If Idopontok_MindenTorles(lo) Then Exit Sub
            Case "8": Idopontok_TomegesOrankentiGeneralas lo
            Case "9": Idopontok_LejartakInaktivalasa lo
            Case Else
                MsgBox "Érvénytelen választás.", vbExclamation
        End Select
    Loop

    Exit Sub
EH:
    MsgBox "Idõpont menü hiba: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

Private Function BuildMenuText(lo As ListObject) As String
    BuildMenuText = _
        "Válassz mûveletet:" & vbCrLf & _
        "  1. Lista / Áttekintés" & vbCrLf & _
        "  2. Új idõpont(ok) felvétele (kézi, soronként)" & vbCrLf & _
        "  3. Aktív kapcsolás (toggle) index alapján" & vbCrLf & _
        "  4. Idõpont törlése index alapján" & vbCrLf & _
        "  5. Duplikátumok takarítása (ugyanaz a dátum -> 1 sor)" & vbCrLf & _
        "  6. Minden inaktiválása" & vbCrLf & _
        "  7. Minden törlése" & vbCrLf & _
        "  8. Tömeges generálás óránként (tartomány + idõsáv + szünet)" & vbCrLf & _
        "  9. Lejárt idõpontok inaktiválása (datum_nap < Now)" & vbCrLf & _
        vbCrLf & _
        "Tábla: " & lo.parent.Name & " / " & lo.Name & " (sorok: " & lo.ListRows.Count & ")" & vbCrLf & _
        vbCrLf & _
        "Kilépés: Mégse / üres"
End Function

' =========================
' 1) LISTA
' =========================
Private Sub Idopontok_Listaz(lo As ListObject)
    Dim iDt As Long: iDt = GetColIndex_Menu(lo, "datum_nap")
    Dim iAk As Long: iAk = GetColIndex_Menu(lo, "aktiv")
    If iDt = 0 Or iAk = 0 Then Exit Sub

    If lo.ListRows.Count = 0 Or lo.DataBodyRange Is Nothing Then
        MsgBox "Nincs idõpont.", vbInformation
        Exit Sub
    End If

    Dim arr As Variant: arr = lo.DataBodyRange.value

    Dim msg As String
    msg = "Idõpontok (index / dátum / aktív):" & vbCrLf & vbCrLf

    Dim r As Long
    For r = 1 To UBound(arr, 1)
        msg = msg & r & ". " & _
              FormatAnyDate(arr(r, iDt)) & _
              "   | aktiv=" & CStr(arr(r, iAk)) & vbCrLf
        If r = 40 Then
            msg = msg & "... (csak az elsõ 40 sor látszik)" & vbCrLf
            Exit For
        End If
    Next r

    MsgBox msg, vbInformation
End Sub

' =========================
' 2) ÚJ FELVITEL (kézi)
' =========================
Private Sub Idopontok_UjFelvitel(lo As ListObject)
    Dim iDt As Long: iDt = GetColIndex_Menu(lo, "datum_nap")
    Dim iAk As Long: iAk = GetColIndex_Menu(lo, "aktiv")
    If iDt = 0 Or iAk = 0 Then Exit Sub

    Dim s As String
    s = InputBox( _
        "Adj meg idõpontokat soronként (ENTER-rel új sor)." & vbCrLf & _
        "Formátum ajánlott: 2026.03.07 09:00:00 (a másodperc opcionális)" & vbCrLf & _
        "Példa:" & vbCrLf & _
        "2026.03.07 09:00" & vbCrLf & _
        "2026.03.07 10:00", _
        "Új idõpontok felvétele")
    s = Trim$(s)
    If s = "" Then Exit Sub

    Dim lines() As String
    lines = Split(Replace(s, vbCrLf, vbLf), vbLf)

    Dim addCount As Long: addCount = 0
    Dim badCount As Long: badCount = 0
    Dim badList As String: badList = ""

    Application.ScreenUpdating = False

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim dt As Date
        If TryParseHuDateTime_Menu(lines(i), dt) Then
            Dim lr As ListRow
            Set lr = lo.ListRows.add
            lr.Range.Cells(1, iDt).value = dt
            lr.Range.Cells(1, iDt).NumberFormat = "yyyy.mm.dd hh:mm:ss"
            lr.Range.Cells(1, iAk).value = 1 ' aktív
            addCount = addCount + 1
        Else
            If Trim$(lines(i)) <> "" Then
                badCount = badCount + 1
                If badCount <= 10 Then badList = badList & "• " & lines(i) & vbCrLf
            End If
        End If
    Next i

    Application.ScreenUpdating = True

    Dim msg As String
    msg = "Felvéve: " & addCount & " idõpont." & vbCrLf & _
          "Hibás sorok: " & badCount
    If badList <> "" Then msg = msg & vbCrLf & vbCrLf & "Példák hibás sorokra:" & vbCrLf & badList

    MsgBox msg, vbInformation
End Sub

' =========================
' 3) AKTÍV TOGGLE
' =========================
Private Sub Idopontok_AktivToggle(lo As ListObject)
    If lo.ListRows.Count = 0 Then
        MsgBox "Nincs idõpont.", vbExclamation
        Exit Sub
    End If

    Dim iAk As Long: iAk = GetColIndex_Menu(lo, "aktiv")
    If iAk = 0 Then Exit Sub

    Dim idx As Long
    idx = CLng(val(InputBox("Melyik INDEX-et kapcsoljam át? (1.." & lo.ListRows.Count & ")", "Aktív kapcsolás")))
    If idx < 1 Or idx > lo.ListRows.Count Then Exit Sub

    Dim v As Variant
    v = lo.DataBodyRange.Cells(idx, iAk).value

    Dim newVal As Long
    newVal = IIf(CLng(val(v)) = 1, 0, 1)

    lo.DataBodyRange.Cells(idx, iAk).value = newVal
    MsgBox "Kész. Index " & idx & " aktiv=" & newVal, vbInformation
End Sub

' =========================
' 4) TÖRLÉS
' =========================
Private Sub Idopontok_Torles(lo As ListObject)
    If lo.ListRows.Count = 0 Then
        MsgBox "Nincs idõpont.", vbExclamation
        Exit Sub
    End If

    Dim idx As Long
    idx = CLng(val(InputBox("Melyik INDEX-et töröljem? (1.." & lo.ListRows.Count & ")", "Idõpont törlése")))
    If idx < 1 Or idx > lo.ListRows.Count Then Exit Sub

    If MsgBox("Biztosan törlöd a(z) " & idx & ". sort az idõpont táblából?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    lo.ListRows(idx).Delete
    MsgBox "Törölve.", vbInformation
End Sub

' =========================
' 5) DUPLIKÁTUM TAKARÍTÁS
' =========================
Private Sub Idopontok_DuplikatumTakaritas(lo As ListObject)
    Dim iDt As Long: iDt = GetColIndex_Menu(lo, "datum_nap")
    Dim iAk As Long: iAk = GetColIndex_Menu(lo, "aktiv")
    If iDt = 0 Or iAk = 0 Then Exit Sub

    If lo.ListRows.Count = 0 Or lo.DataBodyRange Is Nothing Then
        MsgBox "Nincs idõpont.", vbInformation
        Exit Sub
    End If

    Dim arr As Variant: arr = lo.DataBodyRange.value
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long, dt As Date

    Dim delIdx() As Long
    ReDim delIdx(1 To lo.ListRows.Count)

    Dim delN As Long: delN = 0

    For r = 1 To UBound(arr, 1)
        If TryParseHuDateTime_Menu(arr(r, iDt), dt) Then
            Dim k As String: k = CStr(CDbl(dt))
            If dict.Exists(k) Then
                delN = delN + 1
                delIdx(delN) = r
            Else
                dict.add k, True
            End If
        End If
    Next r

    If delN = 0 Then
        MsgBox "Nincs duplikált idõpont.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim i As Long
    For i = delN To 1 Step -1
        lo.ListRows(delIdx(i)).Delete
    Next i
    Application.ScreenUpdating = True

    MsgBox "Duplikátum takarítás kész. Törölt sorok: " & delN, vbInformation
End Sub

' =========================
' 6) MINDEN INAKTIVÁLÁSA
' =========================
Private Sub Idopontok_MindenInaktiv(lo As ListObject)
    Dim iAk As Long: iAk = GetColIndex_Menu(lo, "aktiv")
    If iAk = 0 Then Exit Sub

    If lo.ListRows.Count = 0 Then
        MsgBox "Nincs idõpont.", vbInformation
        Exit Sub
    End If

    If MsgBox("Biztosan INAKTÍVRA állítod az összes idõpontot?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    lo.ListColumns(iAk).DataBodyRange.value = 0
    MsgBox "Kész: minden idõpont inaktív.", vbInformation
End Sub

' =========================
' 7) MINDEN TÖRLÉSE
' =========================
Private Function Idopontok_MindenTorles(lo As ListObject) As Boolean
    If lo.ListRows.Count = 0 Then
        MsgBox "Nincs idõpont.", vbInformation
        Idopontok_MindenTorles = True
        Exit Function
    End If

    If MsgBox("BIZTOSAN törlöd az ÖSSZES idõpontot? Ez nem visszavonható.", vbYesNo + vbCritical) <> vbYes Then Exit Function

    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete

    MsgBox "Minden idõpont törölve.", vbInformation
    Idopontok_MindenTorles = True
End Function

' =========================
' 8) TÖMEGES GENERÁLÁS ÓRÁNKÉNT (fix 60 perc)
' =========================
Private Sub Idopontok_TomegesOrankentiGeneralas(lo As ListObject)
    On Error GoTo EH

    Dim iDt As Long: iDt = GetColIndex_Menu(lo, "datum_nap")
    Dim iAk As Long: iAk = GetColIndex_Menu(lo, "aktiv")
    If iDt = 0 Or iAk = 0 Then Exit Sub

    Dim sRange As String
    sRange = Trim$(InputBox("Dátumtartomány (YYYY.MM.DD - YYYY.MM.DD)" & vbCrLf & _
                            "Példa: 2026.03.07 - 2026.03.10", "Tömeges generálás óránként", "2026.03.07 - 2026.03.10"))
    If sRange = "" Then Exit Sub

    Dim d1 As Date, d2 As Date
    If Not TryParseDateRangeHu(sRange, d1, d2) Then
        MsgBox "Hibás dátumtartomány. Formátum: 2026.03.07 - 2026.03.10", vbExclamation
        Exit Sub
    End If

    Dim sTime As String
    sTime = Trim$(InputBox("Idõsáv (HH:MM - HH:MM)" & vbCrLf & _
                           "Példa: 08:00 - 14:00", "Tömeges generálás óránként", "08:00 - 14:00"))
    If sTime = "" Then Exit Sub

    Dim t1 As Date, t2 As Date
    If Not TryParseTimeRange(sTime, t1, t2) Then
        MsgBox "Hibás idõsáv. Formátum: 08:00 - 14:00", vbExclamation
        Exit Sub
    End If

    Dim sBreaks As String
    sBreaks = Trim$(InputBox("Szünet(ek) kihagyása (opcionális)" & vbCrLf & _
                             "Formátum: HH:MM-HH:MM;HH:MM-HH:MM" & vbCrLf & _
                             "Példa: 12:00-12:30" & vbCrLf & _
                             "Üres = nincs szünet", _
                             "Tömeges generálás óránként", "12:00-12:30"))

    Dim breaks As Variant
    breaks = ParseBreakRangesOrEmpty(sBreaks)

    Dim sDays As String
    sDays = Trim$(InputBox("Mely napokon?" & vbCrLf & _
                           "1=hétfõ ... 7=vasárnap" & vbCrLf & _
                           "Példák: 1-5  |  6-7  |  1,3,5  |  1-7", _
                           "Tömeges generálás óránként", "1-5"))
    If sDays = "" Then Exit Sub

    Dim dayAllowed(1 To 7) As Boolean
    If Not TryParseDaySpec(sDays, dayAllowed) Then
        MsgBox "Hibás nap megadás. Példa: 1-5 vagy 1,3,5", vbExclamation
        Exit Sub
    End If

    ' meglévõk a duplikáció ellen
    Dim existing As Object
    Set existing = CreateObject("Scripting.Dictionary")
    existing.CompareMode = 0

    If lo.ListRows.Count > 0 And Not lo.DataBodyRange Is Nothing Then
        Dim arr As Variant: arr = lo.DataBodyRange.value
        Dim r As Long, dtE As Date
        For r = 1 To UBound(arr, 1)
            If TryParseHuDateTime_Menu(arr(r, iDt), dtE) Then
                existing(CStr(CDbl(dtE))) = True
            End If
        Next r
    End If

    Dim addCount As Long, skipCount As Long, breakSkip As Long
    addCount = 0: skipCount = 0: breakSkip = 0

    Application.ScreenUpdating = False

    Dim d As Date, cur As Date, dt As Date
    Dim oneHour As Double
    oneHour = 1# / 24# ' 60 perc

    For d = DateSerial(Year(d1), Month(d1), Day(d1)) To DateSerial(Year(d2), Month(d2), Day(d2))
        Dim dow As Long
        dow = Weekday(d, vbMonday) ' 1..7

        If dayAllowed(dow) Then
            cur = TimeSerial(Hour(t1), Minute(t1), 0)

            Do While cur <= TimeSerial(Hour(t2), Minute(t2), 0)

                If IsInAnyBreak(cur, breaks) Then
                    breakSkip = breakSkip + 1
                Else
                    dt = d + cur

                    Dim k As String
                    k = CStr(CDbl(dt))

                    If existing.Exists(k) Then
                        skipCount = skipCount + 1
                    Else
                        Dim lr As ListRow
                        Set lr = lo.ListRows.add
                        lr.Range.Cells(1, iDt).value = dt
                        lr.Range.Cells(1, iDt).NumberFormat = "yyyy.mm.dd hh:mm:ss"
                        lr.Range.Cells(1, iAk).value = 1 ' aktív

                        existing.add k, True
                        addCount = addCount + 1
                    End If
                End If

                cur = cur + oneHour
            Loop
        End If
    Next d

    Application.ScreenUpdating = True

    MsgBox "Tömeges óránkénti generálás kész." & vbCrLf & _
           "Felvéve (aktív=1): " & addCount & vbCrLf & _
           "Kihagyva (duplikát): " & skipCount & vbCrLf & _
           "Kihagyva (szünet): " & breakSkip, vbInformation
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "Tömeges generálás hiba: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

' =========================
' 9) LEJÁRTAK INAKTIVÁLÁSA (datum_nap < Now)
' =========================
Private Sub Idopontok_LejartakInaktivalasa(lo As ListObject)
    On Error GoTo EH

    Dim iDt As Long: iDt = GetColIndex_Menu(lo, "datum_nap")
    Dim iAk As Long: iAk = GetColIndex_Menu(lo, "aktiv")
    If iDt = 0 Or iAk = 0 Then Exit Sub

    If lo.ListRows.Count = 0 Or lo.DataBodyRange Is Nothing Then
        MsgBox "Nincs idõpont.", vbInformation
        Exit Sub
    End If

    If MsgBox("Inaktiváljam az összes LEJÁRT idõpontot? (datum_nap < most)", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    Dim arr As Variant: arr = lo.DataBodyRange.value
    Dim r As Long, dt As Date, changed As Long
    changed = 0

    Application.ScreenUpdating = False

    For r = 1 To UBound(arr, 1)
        If TryParseHuDateTime_Menu(arr(r, iDt), dt) Then
            If dt < Now Then
                If CLng(val(arr(r, iAk))) <> 0 Then
                    lo.DataBodyRange.Cells(r, iAk).value = 0
                    changed = changed + 1
                End If
            End If
        End If
    Next r

    Application.ScreenUpdating = True

    MsgBox "Kész. Inaktivált lejárt idõpontok: " & changed, vbInformation
    Exit Sub

EH:
    Application.ScreenUpdating = True
    MsgBox "Lejárt inaktiválás hiba: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

' =========================================================
' Segédek (csak ehhez a menühöz)
' =========================================================
Private Function GetColIndex_Menu(lo As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    GetColIndex_Menu = lo.ListColumns(colName).Index
    On Error GoTo 0
    If GetColIndex_Menu = 0 Then
        MsgBox "Hiányzó oszlop a táblában: " & colName, vbExclamation
    End If
End Function

Private Function FormatAnyDate(ByVal v As Variant) As String
    Dim dt As Date
    If TryParseHuDateTime_Menu(v, dt) Then
        FormatAnyDate = Format$(dt, "yyyy.mm.dd hh:nn:ss")
    Else
        FormatAnyDate = CStr(v)
    End If
End Function

Private Function TryParseDateRangeHu(ByVal s As String, ByRef d1 As Date, ByRef d2 As Date) As Boolean
    On Error GoTo Fail
    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")
    s = Replace(s, " ", "")

    Dim parts() As String
    parts = Split(s, "-")
    If UBound(parts) <> 1 Then GoTo Fail

    Dim a As Date, b As Date
    If Not TryParseHuDateTime_Menu(parts(0), a) Then GoTo Fail
    If Not TryParseHuDateTime_Menu(parts(1), b) Then GoTo Fail

    d1 = DateSerial(Year(a), Month(a), Day(a))
    d2 = DateSerial(Year(b), Month(b), Day(b))
    If d2 < d1 Then GoTo Fail

    TryParseDateRangeHu = True
    Exit Function
Fail:
    TryParseDateRangeHu = False
End Function

Private Function TryParseTimeRange(ByVal s As String, ByRef t1 As Date, ByRef t2 As Date) As Boolean
    On Error GoTo Fail
    s = Replace(s, "–", "-")
    s = Replace(s, "—", "-")
    s = Replace(s, " ", "")

    Dim parts() As String
    parts = Split(s, "-")
    If UBound(parts) <> 1 Then GoTo Fail

    Dim a As Date, b As Date
    If Not TryParseTimeOnly(parts(0), a) Then GoTo Fail
    If Not TryParseTimeOnly(parts(1), b) Then GoTo Fail

    t1 = a
    t2 = b
    If t2 < t1 Then GoTo Fail

    TryParseTimeRange = True
    Exit Function
Fail:
    TryParseTimeRange = False
End Function

Private Function TryParseTimeOnly(ByVal s As String, ByRef tOut As Date) As Boolean
    On Error GoTo Fail
    s = Trim$(s)
    If s = "" Then GoTo Fail

    Dim p() As String
    p = Split(s, ":")

    Dim hh As Long, nn As Long, ss As Long
    hh = 0: nn = 0: ss = 0

    hh = CLng(val(p(0)))
    If UBound(p) >= 1 Then nn = CLng(val(p(1)))
    If UBound(p) >= 2 Then ss = CLng(val(p(2)))

    If hh < 0 Or hh > 23 Then GoTo Fail
    If nn < 0 Or nn > 59 Then GoTo Fail
    If ss < 0 Or ss > 59 Then GoTo Fail

    tOut = TimeSerial(hh, nn, ss)
    TryParseTimeOnly = True
    Exit Function
Fail:
    TryParseTimeOnly = False
End Function

Private Function TryParseDaySpec(ByVal s As String, ByRef dayAllowed() As Boolean) As Boolean
    On Error GoTo Fail

    Dim i As Long
    For i = 1 To 7
        dayAllowed(i) = False
    Next i

    s = Replace(s, " ", "")

    If InStr(s, "-") > 0 Then
        Dim parts() As String
        parts = Split(s, "-")
        If UBound(parts) <> 1 Then GoTo Fail

        Dim a As Long, b As Long
        a = CLng(val(parts(0)))
        b = CLng(val(parts(1)))
        If a < 1 Or a > 7 Or b < 1 Or b > 7 Then GoTo Fail
        If b < a Then GoTo Fail

        For i = a To b
            dayAllowed(i) = True
        Next i

    ElseIf InStr(s, ",") > 0 Then
        Dim xs() As String
        xs = Split(s, ",")
        For i = LBound(xs) To UBound(xs)
            Dim n As Long
            n = CLng(val(xs(i)))
            If n < 1 Or n > 7 Then GoTo Fail
            dayAllowed(n) = True
        Next i

    Else
        Dim one As Long
        one = CLng(val(s))
        If one < 1 Or one > 7 Then GoTo Fail
        dayAllowed(one) = True
    End If

    TryParseDaySpec = True
    Exit Function
Fail:
    TryParseDaySpec = False
End Function

' ===== szünetek kezelése =====

Private Function ParseBreakRangesOrEmpty(ByVal s As String) As Variant
    s = Trim$(s)
    If s = "" Then
        ParseBreakRangesOrEmpty = Empty
        Exit Function
    End If

    Dim parts() As String
    parts = Split(s, ";")

    Dim out() As Variant
    ReDim out(0 To UBound(parts))

    Dim i As Long, n As Long
    n = -1

    For i = LBound(parts) To UBound(parts)
        Dim p As String
        p = Trim$(parts(i))
        If p <> "" Then
            Dim tStart As Date, tEnd As Date
            If TryParseTimeRange(p, tStart, tEnd) Then
                n = n + 1
                out(n) = Array(tStart, tEnd)
            Else
                Err.Raise vbObjectError + 513, , "Hibás szünet tartomány: " & p
            End If
        End If
    Next i

    If n < 0 Then
        ParseBreakRangesOrEmpty = Empty
    Else
        ReDim Preserve out(0 To n)
        ParseBreakRangesOrEmpty = out
    End If
End Function

Private Function IsInAnyBreak(ByVal t As Date, ByVal breaks As Variant) As Boolean
    On Error GoTo SafeNo
    If IsEmpty(breaks) Then Exit Function

    Dim i As Long
    For i = LBound(breaks) To UBound(breaks)
        Dim one As Variant
        one = breaks(i)

        Dim tStart As Date, tEnd As Date
        tStart = one(0): tEnd = one(1)

        ' [tStart, tEnd)
        If t >= tStart And t < tEnd Then
            IsInAnyBreak = True
            Exit Function
        End If
    Next i

SafeNo:
End Function

' ===== dátum+idõ parse (menühez) =====
Private Function TryParseHuDateTime_Menu(ByVal v As Variant, ByRef dtOut As Date) As Boolean
    On Error GoTo Fail

    If IsDate(v) Then
        dtOut = CDate(v)
        TryParseHuDateTime_Menu = True
        Exit Function
    End If

    Dim s As String: s = Trim$(CStr(v))
    If s = "" Then GoTo Fail

    s = Replace(s, ChrW(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    Dim datePart As String, timePart As String
    If InStr(s, " ") > 0 Then
        datePart = Split(s, " ")(0)
        timePart = Split(s, " ")(1)
    Else
        datePart = s
        timePart = "00:00:00"
    End If

    datePart = Replace(datePart, "-", ".")
    Dim d() As String: d = Split(datePart, ".")
    If UBound(d) <> 2 Then GoTo Fail

    Dim yyyy As Long, mm As Long, dd As Long
    yyyy = CLng(val(d(0)))
    mm = CLng(val(d(1)))
    dd = CLng(val(d(2)))

    Dim t() As String: t = Split(timePart, ":")
    Dim hh As Long, nn As Long, ss As Long
    hh = 0: nn = 0: ss = 0
    If UBound(t) >= 0 Then hh = CLng(val(t(0)))
    If UBound(t) >= 1 Then nn = CLng(val(t(1)))
    If UBound(t) >= 2 Then ss = CLng(val(t(2)))

    dtOut = DateSerial(yyyy, mm, dd) + TimeSerial(hh, nn, ss)
    TryParseHuDateTime_Menu = True
    Exit Function

Fail:
    TryParseHuDateTime_Menu = False
End Function

