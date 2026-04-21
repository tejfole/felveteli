Attribute VB_Name = "IktszMenuCallbacks"
Option Explicit

Private Const IDOPONT_KIADVA_MARKER As String = "x"

' Iktsz menü callbackek:
' - intézményi: lista.iktsz kitöltés isk_nev csoportosítással
' - határozat: lista.iktsz egyedi, szekvenciális (csak nem üres hatarozat sorokra)
' - szóbeli: diakadat.iktsz feltételes, szekvenciális

Public Sub Iktsz_Menu_Intezmenyi(Optional control As IRibbonControl)
    FillIktsz_GroupByKey "lista", "isk_nev", "iktsz", 1
End Sub

Public Sub Iktsz_Menu_Hatarozat(Optional control As IRibbonControl)
    FillIktsz_Sequential_ListaHatarozat "lista", "hatarozat", "iktsz"
End Sub

Public Sub Iktsz_Menu_Szobeli(Optional control As IRibbonControl)
    FillIktsz_Sequential_Diakadat "diakadat", "iktsz", "bizottsag", "datum_nap", "mail", "idopont_kiadva"
End Sub

Private Sub FillIktsz_GroupByKey(ByVal tableName As String, ByVal keyColName As String, ByVal iktszColName As String, ByVal defaultStart As Long)
    Dim lo As ListObject
    Set lo = FindListObjectInWorkbook(tableName)
    If lo Is Nothing Then
        MsgBox "Nem található '" & tableName & "' nevű tábla.", vbCritical
        Exit Sub
    End If

    Dim keyCol As Long, iktszCol As Long
    keyCol = FindColumnIndex(lo, keyColName)
    iktszCol = FindColumnIndex(lo, iktszColName)
    If keyCol = 0 Or iktszCol = 0 Then
        MsgBox "Hiányzik a(z) '" & keyColName & "' vagy '" & iktszColName & "' oszlop.", vbCritical
        Exit Sub
    End If

    Dim startNum As Long, cancelled As Boolean
    startNum = PromptForStartNumber("Add meg a kezdő iktsz számot:", defaultStart, cancelled)
    If cancelled Then Exit Sub

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lr As ListRow, key As String
    For Each lr In lo.ListRows
        key = Trim$(CStr(lr.Range.Cells(1, keyCol).Value & ""))
        If key = "" Then
            lr.Range.Cells(1, iktszCol).Value = ""
        Else
            If Not dict.Exists(key) Then
                dict.Add key, startNum
                startNum = startNum + 1
            End If
            lr.Range.Cells(1, iktszCol).Value = dict(key)
        End If
    Next lr

    MsgBox "Kész: intézményi iktsz kitöltve (isk_nev csoportosítás).", vbInformation
End Sub

Private Sub FillIktsz_Sequential_ListaHatarozat(ByVal tableName As String, ByVal requiredColName As String, ByVal iktszColName As String)
    Dim lo As ListObject
    Set lo = FindListObjectInWorkbook(tableName)
    If lo Is Nothing Then
        MsgBox "Nem található '" & tableName & "' nevű tábla.", vbCritical
        Exit Sub
    End If

    Dim requiredCol As Long, iktszCol As Long
    requiredCol = FindColumnIndex(lo, requiredColName)
    iktszCol = FindColumnIndex(lo, iktszColName)
    If requiredCol = 0 Or iktszCol = 0 Then
        MsgBox "Hiányzik a(z) '" & requiredColName & "' vagy '" & iktszColName & "' oszlop.", vbCritical
        Exit Sub
    End If

    Dim startNum As Long
    startNum = PromptForStartOrContinue("Kezdő iktsz (üres = folytatás a jelenlegi max után):", lo, iktszCol)
    If startNum = -1 Then Exit Sub

    Dim lr As ListRow, requiredVal As String, currentIkt As String, filled As Long
    For Each lr In lo.ListRows
        requiredVal = Trim$(CStr(lr.Range.Cells(1, requiredCol).Value & ""))
        currentIkt = Trim$(CStr(lr.Range.Cells(1, iktszCol).Value & ""))

        If requiredVal <> "" And currentIkt = "" Then
            lr.Range.Cells(1, iktszCol).Value = startNum
            startNum = startNum + 1
            filled = filled + 1
        End If
    Next lr

    MsgBox "Kész: határozat iktsz kitöltve. Új iktsz: " & filled, vbInformation
End Sub

Private Sub FillIktsz_Sequential_Diakadat(ByVal tableName As String, ByVal iktszColName As String, ByVal bizottsagColName As String, ByVal datumColName As String, ByVal mailColName As String, ByVal idopontKiadvaColName As String)
    Dim lo As ListObject
    Set lo = FindListObjectInWorkbook(tableName)
    If lo Is Nothing Then
        MsgBox "Nem található '" & tableName & "' nevű tábla.", vbCritical
        Exit Sub
    End If

    Dim iktszCol As Long, bizottsagCol As Long, datumCol As Long, mailCol As Long, idopontKiadvaCol As Long
    iktszCol = FindColumnIndex(lo, iktszColName)
    bizottsagCol = FindColumnIndex(lo, bizottsagColName)
    datumCol = FindColumnIndex(lo, datumColName)
    mailCol = FindColumnIndex(lo, mailColName)
    idopontKiadvaCol = FindColumnIndex(lo, idopontKiadvaColName)

    If iktszCol = 0 Or bizottsagCol = 0 Or datumCol = 0 Or mailCol = 0 Or idopontKiadvaCol = 0 Then
        MsgBox "Hiányzó oszlop(ok): iktsz/bizottsag/datum_nap/mail/idopont_kiadva", vbCritical
        Exit Sub
    End If

    Dim startNum As Long
    startNum = PromptForStartOrContinue("Kezdő iktsz (üres = folytatás a jelenlegi max után):", lo, iktszCol)
    If startNum = -1 Then Exit Sub

    Dim lr As ListRow
    Dim bizottsagVal As String, datumVal As String, mailVal As String, idopontKiadvaVal As String, currentIkt As String
    Dim filled As Long

    For Each lr In lo.ListRows
        bizottsagVal = Trim$(CStr(lr.Range.Cells(1, bizottsagCol).Value & ""))
        datumVal = Trim$(CStr(lr.Range.Cells(1, datumCol).Value & ""))
        mailVal = Trim$(CStr(lr.Range.Cells(1, mailCol).Value & ""))
        idopontKiadvaVal = LCase$(Trim$(CStr(lr.Range.Cells(1, idopontKiadvaCol).Value & "")))
        currentIkt = Trim$(CStr(lr.Range.Cells(1, iktszCol).Value & ""))

        If bizottsagVal <> "" And datumVal <> "" And mailVal <> "" And idopontKiadvaVal <> IDOPONT_KIADVA_MARKER And currentIkt = "" Then
            lr.Range.Cells(1, iktszCol).Value = startNum
            startNum = startNum + 1
            filled = filled + 1
        End If
    Next lr

    MsgBox "Kész: szóbeli iktsz kitöltve. Új iktsz: " & filled, vbInformation
End Sub

Private Function PromptForStartNumber(ByVal prompt As String, ByVal defaultValue As Long, ByRef cancelled As Boolean) As Long
    Dim inputText As String
    inputText = Trim$(InputBox(prompt, "Kezdő iktsz", CStr(defaultValue)))

    If inputText = "" Then
        PromptForStartNumber = 0
        cancelled = True
        Exit Function
    End If

    If Not TryParseLongInput(inputText, PromptForStartNumber) Then
        cancelled = True
        Exit Function
    End If
End Function

Private Function PromptForStartOrContinue(ByVal prompt As String, ByVal lo As ListObject, ByVal iktszCol As Long) As Long
    Dim inputText As String
    inputText = Trim$(InputBox(prompt, "Kezdő iktsz", ""))

    If inputText = "" Then
        PromptForStartOrContinue = MaxIktszValue(lo, iktszCol) + 1
        Exit Function
    End If

    If Not TryParseLongInput(inputText, PromptForStartOrContinue) Then
        PromptForStartOrContinue = -1
        Exit Function
    End If
End Function

Private Function TryParseLongInput(ByVal inputText As String, ByRef outParsedValue As Long) As Boolean
    TryParseLongInput = False

    If Not IsNumeric(inputText) Then
        MsgBox "A megadott érték nem szám.", vbExclamation
        Exit Function
    End If

    outParsedValue = CLng(inputText)
    TryParseLongInput = True
End Function

Private Function MaxIktszValue(ByVal lo As ListObject, ByVal iktszCol As Long) As Long
    Dim lr As ListRow, currentVal As Variant
    MaxIktszValue = 0

    For Each lr In lo.ListRows
        currentVal = lr.Range.Cells(1, iktszCol).Value
        If IsNumeric(currentVal) Then
            If CLng(currentVal) > MaxIktszValue Then MaxIktszValue = CLng(currentVal)
        End If
    Next lr
End Function

Private Function FindListObjectInWorkbook(ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase$(Trim$(lo.Name)) = LCase$(Trim$(tableName)) Then
                Set FindListObjectInWorkbook = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function FindColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If LCase$(Trim$(lc.Name)) = LCase$(Trim$(columnName)) Then
            FindColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
End Function
