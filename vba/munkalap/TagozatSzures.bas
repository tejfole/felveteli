Attribute VB_Name = "TagozatSzures"
Option Explicit

' ============================================================
' TAGOZAT SZÛRÉS – EGYBEN (az elgondolás szerint)
'
' Forrás táblák / lapok:
' - diakadat lap / diakadat tábla
' - rangsor lap / rangsor tábla  -> CSAK az "elut" oszlop számít ("x" = dolgozunk vele)
' - Szabályok lap / tbl_szabalyok tábla -> ezt a meglévõ rangsor-makród használja
'
' Kimenet:
' - tagozat lap / tagozatokszures tábla
'   fejlécek: nev | oktazon | osszpont | szamitott_rang
'
' Vezérlés:
' - A tagozat lap B1 cellája tartalmazza a tagozat kódot (1000/2000/3000/4000)
' - Szûrés feltétele: diakadat["j_" & B1] = "x"
'
' Rangsor:
' - A meglévõ (jól mûködõ) makró számolja: RangsorTolt_Klasszikus_SorrendDontos
' - Ez feltölti a diakadat[rangsor] oszlopot
' - A kimeneti szamitott_rang = diakadat[rangsor]
'
' Megjegyzés:
' - Ez a modul NEM tartalmaz OnTime/autorefresh kódot.
' - A B1-re reagáló automatikus futtatás a tagozat munkalap (sheet) moduljába kerül.
' ============================================================

' Ezt hívd (sheet eventbõl vagy gombból): rangsor újra + output újratölt
Public Sub Tagozat_Refresh(Optional ByVal silent As Boolean = True)
    On Error GoTo EH

    ' 1) Rangsor frissítés (diakadat[rangsor]) – a te bevált modulodból
    RangsorTolt_Klasszikus_SorrendDontos

    ' 2) Tagozat output újratöltés
    TagozatSzures_Elutasitottakbol

    Exit Sub

EH:
    If Not silent Then
        MsgBox "Tagozat_Refresh hiba: " & Err.Number & " - " & Err.Description, vbCritical
    End If
End Sub

' Újratölti a tagozat/tagozatokszures táblát:
' - csak rangsor[elut]="x"
' - és diakadat["j_" & B1]="x"
Public Sub TagozatSzures_Elutasitottakbol()
    On Error GoTo EH

    Dim wb As Workbook: Set wb = ThisWorkbook

    Dim loD As ListObject, loR As ListObject, loOut As ListObject
    Set loD = wb.Worksheets("diakadat").ListObjects("diakadat")
    Set loR = wb.Worksheets("rangsor").ListObjects("rangsor")
    Set loOut = wb.Worksheets("tagozat").ListObjects("tagozatokszures")

    If loD Is Nothing Or loR Is Nothing Or loOut Is Nothing Then
        Err.Raise vbObjectError + 2000, "TagozatSzures_Elutasitottakbol", _
                  "Hiányzik valamelyik tábla: diakadat / rangsor / tagozatokszures."
    End If

    ' ---- B1 -> tagozat kód ----
    Dim tagKod As String, szuresColName As String
    tagKod = NormalizeTagKod(wb.Worksheets("tagozat").Range("B1").value)
    If tagKod = "" Then Exit Sub
    szuresColName = "j_" & tagKod

    ' ---- diakadat oszlopok ----
    Dim cNev As Long, cOkt As Long, cPont As Long, cSzures As Long, cRang As Long
    cNev = LoCol(loD, "f_nev")
    If cNev = 0 Then cNev = LoCol(loD, "i_nev")
    cOkt = LoCol(loD, "oktazon")
    cPont = LoCol(loD, "p_mindossz")
    cSzures = LoCol(loD, szuresColName)
    cRang = LoCol(loD, "rangsor")

    If cNev = 0 Or cOkt = 0 Or cPont = 0 Or cRang = 0 Then
        Err.Raise vbObjectError + 2001, "TagozatSzures_Elutasitottakbol", _
                  "Hiányzó oszlop a diakadat táblában (f_nev/i_nev, oktazon, p_mindossz, rangsor)."
    End If
    If cSzures = 0 Then
        Err.Raise vbObjectError + 2002, "TagozatSzures_Elutasitottakbol", _
                  "Hiányzó szûrõ oszlop a diakadat táblában: " & szuresColName
    End If

    ' ---- rangsor oszlopok (CSAK elut kell) ----
    Dim rOkt As Long, rElut As Long
    rOkt = LoCol(loR, "oktazon")
    rElut = LoCol(loR, "elut")
    If rOkt = 0 Or rElut = 0 Then
        Err.Raise vbObjectError + 2003, "TagozatSzures_Elutasitottakbol", _
                  "Hiányzó oszlop a rangsor táblában (oktazon, elut)."
    End If

    ' ---- output oszlopok ----
    Dim oNev As Long, oOkt As Long, oPont As Long, oRang As Long
    oNev = LoCol(loOut, "nev")
    oOkt = LoCol(loOut, "oktazon")
    oPont = LoCol(loOut, "osszpont")
    oRang = LoCol(loOut, "szamitott_rang")
    If oNev = 0 Or oOkt = 0 Or oPont = 0 Or oRang = 0 Then
        Err.Raise vbObjectError + 2004, "TagozatSzures_Elutasitottakbol", _
                  "Hiányzó fejléc a tagozatokszures táblában (nev, oktazon, osszpont, szamitott_rang)."
    End If

    ' ---- Elutasítottak set (oktazon alapján) ----
    Dim elut As Object: Set elut = CreateObject("Scripting.Dictionary")
    elut.CompareMode = 1

    If loR.ListRows.Count > 0 Then
        Dim arrR As Variant: arrR = loR.DataBodyRange.value
        Dim i As Long, ok As String
        For i = 1 To UBound(arrR, 1)
            ok = Trim$(CStr(arrR(i, rOkt)))
            If ok <> "" Then
                If IsX(arrR(i, rElut)) Then elut(ok) = True
            End If
        Next i
    End If

    ' ---- Output ürítés ----
    Do While loOut.ListRows.Count > 0
        loOut.ListRows(1).Delete
    Loop

    ' ---- diakadat -> output ----
    If loD.ListRows.Count > 0 Then
        Dim arrD As Variant: arrD = loD.DataBodyRange.value
        Dim outRow As ListRow
        Dim i2 As Long, ok2 As String

        For i2 = 1 To UBound(arrD, 1)
            ok2 = Trim$(CStr(arrD(i2, cOkt)))
            If ok2 <> "" Then
                If elut.Exists(ok2) Then
                    If IsX(arrD(i2, cSzures)) Then
                        Set outRow = loOut.ListRows.add
                        outRow.Range(1, oNev).value = arrD(i2, cNev)
                        outRow.Range(1, oOkt).value = ok2
                        outRow.Range(1, oPont).value = arrD(i2, cPont)
                        outRow.Range(1, oRang).value = arrD(i2, cRang)
                    End If
                End If
            End If
        Next i2
    End If

    ' ---- Rendezés: osszpont DESC, szamitott_rang ASC ----
    On Error Resume Next
    loOut.Sort.SortFields.clear
    loOut.Sort.SortFields.add key:=loOut.ListColumns("osszpont").DataBodyRange, _
        SortOn:=xlSortOnValues, Order:=xlDescending
    loOut.Sort.SortFields.add key:=loOut.ListColumns("szamitott_rang").DataBodyRange, _
        SortOn:=xlSortOnValues, Order:=xlAscending
    With loOut.Sort
        .Header = xlYes
        .Apply
    End With
    On Error GoTo 0

    Exit Sub

EH:
    MsgBox "TagozatSzures_Elutasitottakbol hiba: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ----------------------------
' Helpers
' ----------------------------
Private Function LoCol(ByVal lo As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    LoCol = lo.ListColumns(colName).Index
    On Error GoTo 0
End Function

Private Function IsX(ByVal v As Variant) As Boolean
    IsX = (LCase$(Trim$(CStr(v))) = "x")
End Function

Private Function NormalizeTagKod(ByVal v As Variant) As String
    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    s = Replace(s, ChrW(160), "") ' NBSP
    s = Replace(s, " ", "")
    NormalizeTagKod = s
End Function

