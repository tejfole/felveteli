Attribute VB_Name = "HatarozatValasztas"
Function Határozat(nev As String) As String
    Dim tblRangsor As ListObject
    Dim i As Long
    Dim nevOszlop As Long, felveszOszlop As Long
    Dim utolsoSor As Long

    ' Tábla elérése
    On Error Resume Next
    Set tblRangsor = ThisWorkbook.Sheets("rangsor").ListObjects("rangsor")
    On Error GoTo 0

    If tblRangsor Is Nothing Then
        Határozat = "Hiba: A 'rangsor' nevû tábla nem található!"
        Exit Function
    End If

    ' Oszlop indexek
    nevOszlop = tblRangsor.ListColumns("nev").Index
    felveszOszlop = tblRangsor.ListColumns("felvesz").Index

    ' Sor bejárása
    For i = 1 To tblRangsor.ListRows.Count
        If LCase(tblRangsor.DataBodyRange(i, nevOszlop).value) = LCase(nev) Then
            If LCase(tblRangsor.DataBodyRange(i, felveszOszlop).value) = "x" Then
                Határozat = "felveszem"
            Else
                Határozat = "nem nyert felvételt"
            End If
            Exit Function
        End If
    Next i

    ' Ha nem található
    Határozat = ""

End Function

