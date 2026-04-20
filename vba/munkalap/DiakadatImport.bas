Attribute VB_Name = "DiakadatImport"
Option Explicit

' =========================
' FÕ MAKRO
' =========================
Public Sub Import_Export_Into_ThisWorkbook_Diakadat(Optional control As IRibbonControl)
    Dim srcPath As String
    srcPath = PickExcelFile("Válaszd ki a FORRÁS Excel fájlt")
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

    Dim headerRowS As Long: headerRowS = 1
    Dim mapS As Object: Set mapS = BuildHeaderMapNorm(wsS, headerRowS) ' normalizált
    Dim mapD As Object: Set mapD = BuildListObjectHeaderMapNorm(loD)   ' normalizált

    Dim srcKeyHeader As String
    srcKeyHeader = InputBox("Forrás kulcs oszlop fejléce:", "Kulcs kiválasztás", "Oktatási azonosító")
    If Trim$(srcKeyHeader) = "" Then GoTo CleanExit

    Dim dstKeyHeader As String: dstKeyHeader = "oktazon"

    If Not mapS.Exists(NKey(srcKeyHeader)) Then
        MsgBox "A forrásban nem találom ezt a kulcs fejléct: " & srcKeyHeader, vbExclamation
        GoTo CleanExit
    End If
    If Not mapD.Exists(NKey(dstKeyHeader)) Then
        MsgBox "A cél táblában nincs '" & dstKeyHeader & "' oszlop.", vbExclamation
        GoTo CleanExit
    End If
    If Not mapD.Exists(NKey("I_ker_irsz")) Then
        MsgBox "A cél táblában nincs 'I_ker_irsz' oszlop.", vbExclamation
        GoTo CleanExit
    End If

    Dim colKeyS As Long: colKeyS = mapS(NKey(srcKeyHeader))
    Dim colKeyD As Long: colKeyD = mapD(NKey(dstKeyHeader))
    Dim colIKer As Long: colIKer = mapD(NKey("I_ker_irsz"))

    Dim dupMode As Long
    dupMode = PickDupMode()
    If dupMode = 0 Then GoTo CleanExit

    ' --- Mapping: NKey(forrás fejléc) -> cél oszlopnév ---
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    m(NKey("Oktatási azonosító")) = "oktazon"
    m(NKey("Név")) = "f_nev"
    m(NKey("Születési hely")) = "f_szul_hely"
    m(NKey("Születési dátum")) = "f_szul_ido"
    m(NKey("Anyja születéskori neve")) = "f_a_nev"

    ' EMAIL aliasok -> mail
    m(NKey("Értesítési e-mail")) = "mail"
    m(NKey("Értesítési e-mail cím")) = "mail"
    m(NKey("Értesítési e-mail címek")) = "mail"
    m(NKey("Értesítési email")) = "mail"
    m(NKey("Értesítési email cím")) = "mail"
    m(NKey("Értesítési email címek")) = "mail"
    m(NKey("E-mail")) = "mail"
    m(NKey("Email")) = "mail"
    m(NKey("Kapcsolattartó e-mail")) = "mail"
    m(NKey("Kapcsolattartó email")) = "mail"

    m(NKey("Állandó lakcím")) = "a_cim"

    ' TEL aliasok -> tel
    m(NKey("Értesítési telefonszámok")) = "tel"
    m(NKey("Telefonszám")) = "tel"
    m(NKey("Telefon")) = "tel"
    m(NKey("Mobilszám")) = "tel"
    m(NKey("Mobil")) = "tel"
    m(NKey("Kapcsolattartó telefonszám")) = "tel"
    m(NKey("Kapcsolattartó telefon")) = "tel"

    m(NKey("Értesítési név")) = "ert_nev"
    m(NKey("Értesítési cím")) = "ert_cim"

    ' >>> Itt a változás: OM helyett iskolanév
    m(NKey("Általános iskola neve")) = "isknev"
    m(NKey("Általános iskola")) = "isknev"
    m(NKey("Iskola neve")) = "isknev"

    m(NKey("SNI")) = "f_SNI2"
    m(NKey("BTMN")) = "f_BTNN"
    m(NKey("Jelige")) = "f_jelige"
    m(NKey("001/1000")) = "j_1000"
    m(NKey("001/2000")) = "j_2000"
    m(NKey("001/3000")) = "j_3000"
    m(NKey("001/4000")) = "j_4000"
    m(NKey("Megjegyzés")) = "megjegyzes"

    ' Cél index
    Dim idxD As Object: Set idxD = CreateObject("Scripting.Dictionary")
    BuildDestIndex loD, colKeyD, idxD

    ' Forrás pick + duplariport
    Dim srcPick As Object: Set srcPick = CreateObject("Scripting.Dictionary")
    Dim srcDupReport As String
    srcDupReport = BuildSourcePickedRowIndex(wsS, headerRowS, colKeyS, dupMode, srcPick)
    If srcDupReport <> "" Then
        MsgBox "Forrás duplikáció riport:" & vbCrLf & vbCrLf & srcDupReport, vbExclamation
    End If

    Dim key As Variant, newCount As Long, updCount As Long

    For Each key In srcPick.keys
        Dim r As Long: r = CLng(srcPick(key))
        Dim k As String: k = CStr(key)
        If k = "" Then GoTo NextKey

        Dim lr As ListRow
        If idxD.Exists(k) Then
            Set lr = loD.ListRows(idxD(k))
            updCount = updCount + 1
        Else
            Set lr = loD.ListRows.add
            lr.Range.Cells(1, colKeyD).value = k
            idxD(k) = lr.Index
            newCount = newCount + 1
        End If

        Dim wroteMail As Boolean: wroteMail = False
        Dim wroteTel As Boolean: wroteTel = False

        Dim srcH As Variant, dstColName As String
        For Each srcH In m.keys
            dstColName = m(srcH)

            If NKey(dstColName) = NKey("oktazon") Then GoTo ContinueField
            If Not (mapS.Exists(srcH) And mapD.Exists(NKey(dstColName))) Then GoTo ContinueField

            Dim cS As Long, cD As Long, v As Variant
            cS = mapS(srcH)
            cD = mapD(NKey(dstColName))
            v = wsS.Cells(r, cS).value

            ' dátum
            If NKey(dstColName) = NKey("f_szul_ido") Then
                v = CoerceToDateOrKeep(v)
            End If

            ' SNI/BTMN: igen -> x, különben üres
            If NKey(dstColName) = NKey("f_sni2") Or NKey(dstColName) = NKey("f_btnn") Then
                v = YesToX(v)
            End If

            ' MAIL: preferált elsõ érvényes
            If NKey(dstColName) = NKey("mail") Then
                If Not wroteMail Then
                    Dim em As String
                    em = EmailFirstValid(CStr(v), 1)
                    If em <> "" Then
                        lr.Range.Cells(1, cD).value = em
                        wroteMail = True
                    End If
                End If
                GoTo ContinueField
            End If

            ' TEL: preferált elsõ érvényes
            If NKey(dstColName) = NKey("tel") Then
                If Not wroteTel Then
                    Dim tel1 As String
                    tel1 = PhoneFirstValid(CStr(v), 1)
                    If tel1 <> "" Then
                        lr.Range.Cells(1, cD).value = tel1
                        wroteTel = True
                    End If
                End If
                GoTo ContinueField
            End If

            ' minden más simán
            lr.Range.Cells(1, cD).value = v

ContinueField:
        Next srcH

        ' I_ker_irsz: 1010..1019 az a_cim-ben
        Dim addr As String
        addr = CStr(GetValueFromRowIfExistsNorm(wsS, r, mapS, "Állandó lakcím"))
        If IsBudapest101x(addr) Then
            lr.Range.Cells(1, colIKer).value = "x"
        Else
            lr.Range.Cells(1, colIKer).ClearContents
        End If

NextKey:
    Next key

    wbD.Save
    MsgBox "Import kész." & vbCrLf & "Új: " & newCount & " | Frissített: " & updCount, vbInformation

CleanExit:
    On Error Resume Next
    wbS.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

EH:
    MsgBox "Hiba: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' =========================
' EMAIL / TEL PARSER
' =========================
Private Function EmailFirstValid(ByVal szoveg As String, ByVal preferalt As Integer) As String
    Dim emails As Collection: Set emails = ExtractEmails(szoveg)
    If emails.Count = 0 Then EmailFirstValid = "": Exit Function
    If preferalt = 2 And emails.Count >= 2 Then EmailFirstValid = CStr(emails(2)) Else EmailFirstValid = CStr(emails(1))
End Function

Private Function PhoneFirstValid(ByVal szoveg As String, ByVal preferalt As Integer) As String
    Dim phones As Collection: Set phones = ExtractPhones(szoveg)
    If phones.Count = 0 Then PhoneFirstValid = "": Exit Function
    If preferalt = 2 And phones.Count >= 2 Then PhoneFirstValid = CStr(phones(2)) Else PhoneFirstValid = CStr(phones(1))
End Function

Private Function ExtractEmails(ByVal s As String) As Collection
    Dim col As New Collection
    s = Replace(s, ChrW(160), " ")

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})"

    Dim m As Object, ms As Object
    Set ms = re.Execute(s)

    For Each m In ms
        On Error Resume Next
        col.add LCase$(Trim$(m.value)), LCase$(Trim$(m.value))
        On Error GoTo 0
    Next m

    Set ExtractEmails = col
End Function

Private Function ExtractPhones(ByVal s As String) As Collection
    Dim col As New Collection
    s = Replace(s, ChrW(160), " ")

    Dim parts As Variant, p As Variant
    parts = Split(MultiReplace(s, Array(vbCrLf, vbCr, vbLf, ";"), ","), ",")

    For Each p In parts
        Dim cleaned As String
        cleaned = NormalizePhoneToken(CStr(p))
        cleaned = CanonicalHuPhone(cleaned)

        If cleaned <> "" Then
            On Error Resume Next
            col.add cleaned, cleaned
            On Error GoTo 0
        End If
    Next p

    Set ExtractPhones = col
End Function

Private Function NormalizePhoneToken(ByVal t As String) As String
    t = Trim$(t)
    If t = "" Then NormalizePhoneToken = "": Exit Function

    If InStr(t, ":") > 0 Then t = Trim$(mid$(t, InStrRev(t, ":") + 1))

    Dim i As Long, ch As String, out As String
    For i = 1 To Len(t)
        ch = mid$(t, i, 1)
        If ch Like "#" Then
            out = out & ch
        ElseIf ch = "+" Then
            If out = "" Then out = "+"
        End If
    Next i

    NormalizePhoneToken = out
End Function

Private Function CanonicalHuPhone(ByVal t As String) As String
    If t = "" Then CanonicalHuPhone = "": Exit Function

    Dim digits As String
    digits = t
    If Left$(digits, 1) = "+" Then digits = mid$(digits, 2)

    Dim i As Long, ch As String, d As String
    For i = 1 To Len(digits)
        ch = mid$(digits, i, 1)
        If ch Like "#" Then d = d & ch
    Next i

    If Left$(d, 2) = "06" Then d = "36" & mid$(d, 3)
    If Len(d) = 9 Then d = "36" & d

    If Len(d) <> 11 Then
        CanonicalHuPhone = ""
    ElseIf Left$(d, 2) <> "36" Then
        CanonicalHuPhone = ""
    Else
        CanonicalHuPhone = "+" & d
    End If
End Function

Private Function MultiReplace(ByVal s As String, ByVal findArr As Variant, ByVal repl As String) As String
    Dim i As Long
    For i = LBound(findArr) To UBound(findArr)
        s = Replace(s, CStr(findArr(i)), repl)
    Next i
    MultiReplace = s
End Function

' =========================
' NORMALIZÁLT HEADER MAP
' =========================
Private Function BuildHeaderMapNorm(ws As Worksheet, headerRow As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long: lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, h As String, nk As String
    For c = 1 To lastCol
        h = Trim$(CStr(ws.Cells(headerRow, c).value))
        If h <> "" Then
            nk = NKey(h)
            If Not d.Exists(nk) Then d(nk) = c
        End If
    Next c
    Set BuildHeaderMapNorm = d
End Function

Private Function BuildListObjectHeaderMapNorm(lo As ListObject) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        d(NKey(lo.ListColumns(i).Name)) = i
    Next i
    Set BuildListObjectHeaderMapNorm = d
End Function

Private Function GetValueFromRowIfExistsNorm(ws As Worksheet, ByVal r As Long, mapS As Object, ByVal headerName As String) As Variant
    Dim nk As String: nk = NKey(headerName)
    If mapS.Exists(nk) Then
        GetValueFromRowIfExistsNorm = ws.Cells(r, mapS(nk)).value
    Else
        GetValueFromRowIfExistsNorm = vbNullString
    End If
End Function

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
' DUPLIKÁCIÓ / PICK
' =========================
Private Function PickDupMode() As Long
    Dim inp As String
    inp = InputBox("Forrás duplakulcs esetén:" & vbCrLf & _
                   "1 = elsõ (ajánlott)" & vbCrLf & _
                   "2 = utolsó" & vbCrLf & _
                   "3 = kérdez", "Duplakulcs kezelés", "1")
    If Trim$(inp) = "" Then
        PickDupMode = 0
    ElseIf IsNumeric(inp) Then
        PickDupMode = CLng(inp)
        If PickDupMode < 1 Or PickDupMode > 3 Then PickDupMode = 1
    Else
        PickDupMode = 1
    End If
End Function

Private Function BuildSourcePickedRowIndex(ws As Worksheet, headerRow As Long, keyCol As Long, dupMode As Long, _
                                          ByRef picked As Object) As String
    picked.RemoveAll

    Dim dRows As Object: Set dRows = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, keyCol).End(xlUp).Row

    Dim r As Long, k As String
    For r = headerRow + 1 To lastRow
        k = Trim$(CStr(ws.Cells(r, keyCol).value))
        If k <> "" Then
            If Not dRows.Exists(k) Then
                Dim col As Collection: Set col = New Collection
                col.add r
                dRows.add k, col
            Else
                dRows(k).add r
            End If
        End If
    Next r

    Dim report As String: report = ""
    Dim key As Variant, shown As Long: shown = 0

    For Each key In dRows.keys
        Dim rowsCol As Collection: Set rowsCol = dRows(key)
        Dim chosenRow As Long

        If rowsCol.Count = 1 Then
            chosenRow = CLng(rowsCol(1))
        Else
            shown = shown + 1
            report = report & "• " & key & " : sorok = " & JoinCollection(rowsCol, ", ") & vbCrLf
            Select Case dupMode
                Case 1: chosenRow = CLng(rowsCol(1))
                Case 2: chosenRow = CLng(rowsCol(rowsCol.Count))
                Case 3
                    chosenRow = AskPickRowForKey(CStr(key), rowsCol)
                    If chosenRow = 0 Then chosenRow = CLng(rowsCol(1))
            End Select
            If shown >= 25 Then report = report & "… (továbbiak elrejtve)" & vbCrLf: Exit For
        End If

        picked(CStr(key)) = chosenRow
    Next key

    BuildSourcePickedRowIndex = report
End Function

Private Function AskPickRowForKey(ByVal key As String, rowsCol As Collection) As Long
    Dim inp As String
    inp = InputBox("Duplikált kulcs: " & key & vbCrLf & _
                   "Sorok: " & JoinCollection(rowsCol, ", ") & vbCrLf & _
                   "Írd be, melyik sort vegyem (üres = elsõ).", "Duplakulcs kiválasztása")
    If Trim$(inp) = "" Then AskPickRowForKey = 0 Else AskPickRowForKey = CLng(inp)
End Function

Private Function JoinCollection(col As Collection, ByVal sep As String) As String
    Dim i As Long, s As String
    For i = 1 To col.Count
        s = s & CStr(col(i))
        If i < col.Count Then s = s & sep
    Next i
    JoinCollection = s
End Function

' =========================
' CÉL INDEX
' =========================
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
' EGYÉB SZABÁLYOK
' =========================
Private Function CoerceToDateOrKeep(v As Variant) As Variant
    On Error GoTo Fail
    If IsDate(v) Then CoerceToDateOrKeep = CDate(v): Exit Function
    If IsNumeric(v) Then CoerceToDateOrKeep = DateSerial(1899, 12, 30) + CDbl(v): Exit Function
Fail:
    CoerceToDateOrKeep = v
End Function

Private Function IsBudapest101x(ByVal addressText As String) As Boolean
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(^|[^0-9])(101[0-9])([^0-9]|$)"
    re.Global = False
    re.IgnoreCase = True
    IsBudapest101x = re.Test(CStr(addressText))
End Function

Private Function YesToX(ByVal v As Variant) As Variant
    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    If s = "" Then
        YesToX = vbNullString
    ElseIf s = "igen" Or s = "i" Or s = "x" Or s = "1" Or s = "true" Or s = "yes" Then
        YesToX = "x"
    Else
        YesToX = vbNullString
    End If
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

