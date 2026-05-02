Option Explicit

' ==============================================
' modIktsz.bas
' Központi iktatószám (iktsz) kitöltési rutinok
'
' 3 használati mód wrapperrel (Ribbon/Menu):
'   1) Intézményi értesítés (lista): csoportos iktsz isk_nev szerint
'   2) Határozatok (lista): egyedi/szekvenciális iktsz soronként (nem ír felül)
'   3) Szóbeli kiértesítés (diakadat): feltételes + max utáni folytatás (nem ír felül)
'
' Megjegyzés:
' - A wrapper eljárásnevek szándékosan egyszerűek és Ribbon-kompatibilisek.
' - A belső függvények minimalizálják a duplikált tábla/oszlop keresést és a kezdőszám logikát.
' ==============================================

' -------- Ribbon / Menü belépési pontok --------

Public Sub Iktsz_KitoltesIskola(Optional control As IRibbonControl)
    ' lista táblában: azonos iskola -> azonos iktsz
    FillIktsz_GroupByKey _
        tableName:="lista", _
        keyColName:="isk_nev", _
        iktszColName:="iktsz", _
        promptTitle:="Intézményi iktatószám", _
        promptText:="Add meg a kezdő iktatószámot (intézményi értesítéshez):", _
        defaultStart:=1
End Sub

Public Sub Iktsz_KitoltesHatarozat(Optional control As IRibbonControl)
    ' lista táblában: egyedi, növekvő iktsz soronként, csak az üresekre
    FillIktsz_Sequential _
        tableName:="lista", _
        iktszColName:="iktsz", _
        eligibilityMode:=Eligibility_AllRows, _
        promptTitle:="Határozat iktatószám", _
        promptText:="Add meg a kezdő iktatószámot (üresen hagyva: meglévő max+1):"
End Sub

Public Sub Iktsz_KitoltesSzobeli(Optional control As IRibbonControl)
    ' diakadat táblában: csak a küldhető sorokra, üresekre, max+1 folytatással
    FillIktsz_Sequential _
        tableName:="diakadat", _
        iktszColName:="iktsz", _
        eligibilityMode:=Eligibility_SzobeliKiertesites, _
        promptTitle:="Szóbeli kiértesítés iktatószám", _
        promptText:="Kezdő iktatószám (üresen hagyva: meglévő max+1):"
End Sub

' Régi név megtartása (ha a customUI vagy más kód hívja)
Public Sub KitoltIktsz_TablaAutomatikusan(Optional control As IRibbonControl)
    ' Eredetileg oktazon alapú csoportosítás volt; ha azt akarod vissza:
    ' FillIktsz_GroupByKey "lista", "oktazon", "iktsz", "Iktatószám", "Kezdő iktsz:", 1
    ' Jelenleg a menü miatt inkább a határozat jellegű (egyedi) kitöltés a biztonságosabb default:
    Iktsz_KitoltesHatarozat control
End Sub

' -------- Közös magok --------

Private Enum EligibilityMode
    Eligibility_AllRows = 0
    Eligibility_SzobeliKiertesites = 1
End Enum

Private Sub FillIktsz_GroupByKey(ByVal tableName As String, _
                                ByVal keyColName As String, _
                                ByVal iktszColName As String, _
                                ByVal promptTitle As String, _
                                ByVal promptText As String, _
                                ByVal defaultStart As Long)
    Dim lo As ListObject
    Set lo = FindListObjectInWorkbook(tableName)
    If lo Is Nothing Then
        MsgBox "Nem található '" & tableName & "' nevű tábla.", vbCritical
        Exit Sub
    End If

    Dim keyIdx As Long, iktszIdx As Long
    keyIdx = ColIndexByName(lo, keyColName)
    iktszIdx = ColIndexByName(lo, iktszColName)

    If keyIdx = 0 Or iktszIdx = 0 Then
        MsgBox "Hiányzó oszlop: '" & keyColName & "' vagy '" & iktszColName & "' a '" & tableName & "' táblában.", vbCritical
        Exit Sub
    End If

    Dim startNum As Variant
    startNum = PromptStartNumber(promptTitle, promptText, defaultStart, allowBlank:=False, fallbackToMaxPlusOne:=False, lo:=lo, iktszIdx:=iktszIdx)
    If IsEmpty(startNum) Then Exit Sub

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim n As Long
    n = CLng(startNum)

    Dim r As ListRow
    For Each r In lo.ListRows
        Dim k As String
        k = Trim$(CStr(r.Range(1, keyIdx).Value))

        If k <> "" Then
            If Not dict.Exists(k) Then
                dict.Add k, n
                n = n + 1
            End If
            r.Range(1, iktszIdx).Value = dict(k)
        Else
            r.Range(1, iktszIdx).Value = ""
        End If
    Next r

    MsgBox "Kész: iktatószám kitöltve (csoportos: " & keyColName & ").", vbInformation
End Sub

Private Sub FillIktsz_Sequential(ByVal tableName As String, _
                                ByVal iktszColName As String, _
                                ByVal eligibilityMode As EligibilityMode, _
                                ByVal promptTitle As String, _
                                ByVal promptText As String)
    Dim lo As ListObject
    Set lo = FindListObjectInWorkbook(tableName)
    If lo Is Nothing Then
        MsgBox "Nem található '" & tableName & "' nevű tábla.", vbCritical
        Exit Sub
    End If

    Dim iktszIdx As Long
    iktszIdx = ColIndexByName(lo, iktszColName)
    If iktszIdx = 0 Then
        MsgBox "Hiányzó oszlop: '" & iktszColName & "' a '" & tableName & "' táblában.", vbCritical
        Exit Sub
    End If

    Dim startNum As Variant
    startNum = PromptStartNumber(promptTitle, promptText, 1, allowBlank:=True, fallbackToMaxPlusOne:=True, lo:=lo, iktszIdx:=iktszIdx)
    If IsEmpty(startNum) Then Exit Sub

    Dim n As Long
    n = CLng(startNum)

    Dim filled As Long
    filled = 0

    Dim r As ListRow
    For Each r In lo.ListRows
        If ShouldAssignSequential(lo, r, eligibilityMode) Then
            Dim curVal As String
            curVal = Trim$(CStr(r.Range(1, iktszIdx).Value))
            If curVal = "" Then
                r.Range(1, iktszIdx).Value = n
                n = n + 1
                filled = filled + 1
            End If
        End If
    Next r

    MsgBox "Kész: " & filled & " sor kapott iktatószámot.", vbInformation
End Sub

Private Function ShouldAssignSequential(ByVal lo As ListObject, ByVal r As ListRow, ByVal mode As EligibilityMode) As Boolean
    Select Case mode
        Case Eligibility_AllRows
            ShouldAssignSequential = True

        Case Eligibility_SzobeliKiertesites
            ' Feltételek (a jelenlegi kód logikájához igazítva):
            ' - bizottsag nem üres
            ' - datum_nap nem üres (vagy idopont_nap)
            ' - mail nem üres
            ' - idopont_kiadva <> "x" (ha van ilyen oszlop)
            Dim idxBiz As Long, idxDt As Long, idxMail As Long, idxKiadva As Long
            idxBiz = ColIndexByName(lo, "bizottsag")
            idxDt = ColIndexByName(lo, "datum_nap")
            If idxDt = 0 Then idxDt = ColIndexByName(lo, "idopont_nap")
            idxMail = ColIndexByName(lo, "mail")
            idxKiadva = ColIndexByName(lo, "idopont_kiadva")

            If idxBiz = 0 Or idxDt = 0 Or idxMail = 0 Then
                ' Ha hiányoznak kulcs oszlopok, inkább ne osszunk automatikusan.
                ShouldAssignSequential = False
                Exit Function
            End If

            Dim vBiz As String, vDt As String, vMail As String, vKiadva As String
            vBiz = Trim$(CStr(r.Range(1, idxBiz).Value))
            vDt = Trim$(CStr(r.Range(1, idxDt).Value))
            vMail = Trim$(CStr(r.Range(1, idxMail).Value))

            If idxKiadva <> 0 Then
                vKiadva = LCase$(Trim$(CStr(r.Range(1, idxKiadva).Value)))
            Else
                vKiadva = ""
            End If

            ShouldAssignSequential = (vBiz <> "") And (vDt <> "") And (vMail <> "") And (vKiadva <> "x")

        Case Else
            ShouldAssignSequential = False
    End Select
End Function

' -------- Segédfüggvények --------

Private Function FindListObjectInWorkbook(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase$(lo.Name) = LCase$(tableName) Then
                Set FindListObjectInWorkbook = lo
                Exit Function
            End If
        Next lo
    Next ws

    Set FindListObjectInWorkbook = Nothing
End Function

Private Function ColIndexByName(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If LCase$(Trim$(lc.Name)) = LCase$(Trim$(colName)) Then
            ColIndexByName = lc.Index
            Exit Function
        End If
    Next lc
    ColIndexByName = 0
End Function

Private Function PromptStartNumber(ByVal title As String, _
                                  ByVal prompt As String, _
                                  ByVal defaultValue As Long, _
                                  ByVal allowBlank As Boolean, _
                                  ByVal fallbackToMaxPlusOne As Boolean, _
                                  ByVal lo As ListObject, _
                                  ByVal iktszIdx As Long) As Variant
    Dim defText As String
    defText = CStr(defaultValue)

    Dim s As String
    s = InputBox(prompt, title, defText)
    s = Trim$(s)

    If s = "" Then
        If allowBlank And fallbackToMaxPlusOne Then
            PromptStartNumber = MaxNumericInColumn(lo, iktszIdx) + 1
            Exit Function
        End If
        PromptStartNumber = Empty
        Exit Function
    End If

    If Not IsNumeric(s) Then
        MsgBox "A megadott érték nem szám.", vbExclamation
        PromptStartNumber = Empty
        Exit Function
    End If

    PromptStartNumber = CLng(s)
End Function

Private Function MaxNumericInColumn(ByVal lo As ListObject, ByVal colIdx As Long) As Long
    On Error GoTo EH

    Dim mx As Long
    mx = 0

    Dim r As ListRow
    For Each r In lo.ListRows
        Dim v As Variant
        v = r.Range(1, colIdx).Value
        If IsNumeric(v) Then
            If CLng(v) > mx Then mx = CLng(v)
        Else
            Dim s As String
            s = Trim$(CStr(v))
            If s <> "" And IsNumeric(s) Then
                If CLng(s) > mx Then mx = CLng(s)
            End If
        End If
    Next r

    MaxNumericInColumn = mx
    Exit Function

EH:
    MaxNumericInColumn = 0
End Function
