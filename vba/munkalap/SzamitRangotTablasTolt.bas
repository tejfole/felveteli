Attribute VB_Name = "SzamitRangotTablasTolt"
Option Explicit

' ===============================================================
' KLASSZIKUS FELVÉTELI RANGSOR SZABÁLYTÁBLA SZERINT, SORRENDDEL
' ===============================================================
'
' Feltételek (pl. testvér, lakcím, hátrányos, pluszpont stb.)
' a "Szabályok" lapon lévõ Excel-táblából ("tbl_szabalyok") – 
' a táblában szereplõ SORREND a döntõ! Súlyt, típust figyelembe vesz,
' de a megszokott felvételi sorrend szerinti szabályokra épít.
'
' A felhasználó CSAK a szabálytáblát szerkeszti Excelben!
'
' A "diakadat" nevû ListObject (Excel-táblában) legyen a jelentkezõk táblája.
' A "p_mindossz" oszlop tartalmazza az összpontszámot.
' A "rangsor" oszlopba kerül a ranghely.
'
' ===============================================================

Sub RangsorTolt_Klasszikus_SorrendDontos()
    Dim wsAdat As Worksheet, wsSzabaly As Worksheet
    Dim diakTbl As ListObject, szabalyTbl As ListObject
    Dim pontRange As Range
    Dim i As Long, n As Long

    ' --- Módosítsd, ha más a lap/táblanév nálad ---
    Set wsAdat = Worksheets("diakadat")
    Set wsSzabaly = Worksheets("Szabályok")
    Set diakTbl = wsAdat.ListObjects("diakadat")
    Set szabalyTbl = wsSzabaly.ListObjects("tbl_szabalyok")
    Set pontRange = diakTbl.ListColumns("p_mindossz").DataBodyRange

    n = diakTbl.DataBodyRange.rows.Count
    Dim rangIdx As Long: rangIdx = diakTbl.ListColumns("rangsor").Index

    For i = 1 To n
        Dim pontSzam As Double: pontSzam = pontRange.Cells(i, 1).value
        Dim rang As Long: rang = 1
        Dim azonosPontszamuak As Long: azonosPontszamuak = 0

        ' --- Elsõ lépés: rangsor pontszám alapján ---
        Dim k As Long
        For k = 1 To n
            If IsNumeric(pontRange.Cells(k, 1).value) Then
                If pontRange.Cells(k, 1).value > pontSzam Then rang = rang + 1
                If pontRange.Cells(k, 1).value = pontSzam Then azonosPontszamuak = azonosPontszamuak + 1
            End If
        Next k

        ' --- Ha nincs döntetlen, fix a helyezés ---
        If azonosPontszamuak = 1 Then
            diakTbl.DataBodyRange.Cells(i, rangIdx).value = rang
            GoTo NextVersenyzo
        End If

        ' --- Döntetlenbontás: szabálytábla sorrend szerinti "elõnyösség" ---
        ' Összegyûjtjük az összes döntetlen indexét:
        Dim tieIndexes() As Long, tCount As Long
        ReDim tieIndexes(1 To azonosPontszamuak)
        tCount = 0
        For k = 1 To n
            If pontRange.Cells(k, 1).value = pontSzam Then
                tCount = tCount + 1
                tieIndexes(tCount) = k
            End If
        Next k

        ' Hány döntetlen versenyzõ "megelõzi" szabály szerinti bontásban ezt a jelentkezõt?
        Dim tiePoz As Long
        For k = 1 To azonosPontszamuak
            If tieIndexes(k) = i Then
                ' Saját magát ne hasonlítsa!
            Else
                Dim cmp As Integer
                cmp = Rangsor_Eloresorol(i, tieIndexes(k), diakTbl, szabalyTbl)
                If cmp = 1 Then
                    rang = rang + 1
                End If
                ' cmp = -1: marad, cmp = 0: teljes döntetlen, marad
            End If
        Next k

        diakTbl.DataBodyRange.Cells(i, rangIdx).value = rang

NextVersenyzo:
    Next i

    MsgBox "Rangsor klasszikus szabály-sorrend szerint frissítve!", vbInformation
End Sub

' =============== SEGÉDFÜGGVÉNY ===============
'
' Adott két jelentkezõ (sorA, sorB) közt végigmegy a szabálytáblán és az elsõ eltérésnél dönt:
'  -1: A megelõzi B-t
'   1: B megelõzi A-t
'   0: teljes döntetlen minden szabály alapján is
'
Function Rangsor_Eloresorol(sorA As Long, sorB As Long, diakTbl As ListObject, szabalyTbl As ListObject) As Integer
    Dim szabalySor As ListRow, oszlopnev As String, tipus As String, suly As Double, aktiv As String
    Dim colIdx As Long, aVal As Variant, bVal As Variant
    Dim j As Long

    For Each szabalySor In szabalyTbl.ListRows
        oszlopnev = szabalySor.Range(1, szabalyTbl.ListColumns("Oszlop_Név").Index).value
        tipus = szabalySor.Range(1, szabalyTbl.ListColumns("Típus").Index).value
        suly = val(szabalySor.Range(1, szabalyTbl.ListColumns("Súly").Index).value)
        aktiv = szabalySor.Range(1, szabalyTbl.ListColumns("Aktív").Index).value

        If UCase(aktiv) = "X" And oszlopnev <> "" Then
            colIdx = 0
            For j = 1 To diakTbl.ListColumns.Count
                If LCase(diakTbl.ListColumns(j).Name) = LCase(oszlopnev) Then
                    colIdx = diakTbl.ListColumns(j).Index
                    Exit For
                End If
            Next j

            If colIdx > 0 Then
                aVal = diakTbl.DataBodyRange.Cells(sorA, colIdx).value
                bVal = diakTbl.DataBodyRange.Cells(sorB, colIdx).value
                If LCase(tipus) = "prioritas" Then
                    If LCase(Trim(CStr(aVal))) = "x" And LCase(Trim(CStr(bVal))) <> "x" Then
                        Rangsor_Eloresorol = -1: Exit Function
                    ElseIf LCase(Trim(CStr(aVal))) <> "x" And LCase(Trim(CStr(bVal))) = "x" Then
                        Rangsor_Eloresorol = 1: Exit Function
                    End If
                ElseIf LCase(tipus) = "pluszpont" Then
                    If val(aVal) > val(bVal) Then
                        Rangsor_Eloresorol = -1: Exit Function
                    ElseIf val(aVal) < val(bVal) Then
                        Rangsor_Eloresorol = 1: Exit Function
                    End If
                End If
                ' Ha még mindig nincs eltérés: megy a következõ szabályhoz.
            End If
        End If
    Next szabalySor
    Rangsor_Eloresorol = 0 ' teljes döntetlen
End Function
