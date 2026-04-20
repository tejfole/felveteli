Attribute VB_Name = "Osszefuz"
Sub NyelvekBeirasaOsszefuzve()
Attribute NyelvekBeirasaOsszefuzve.VB_ProcData.VB_Invoke_Func = "P\n14"

    Dim tbl As ListObject
    Dim aktSor As ListRow
    Dim ertek As Variant
    Dim nyelv1Col As Long, ny_2Col As Long, nyelvOsszesCol As Long
    
    ' ?? Betöltjük a "Rangsor" nevû táblát
    Set tbl = ThisWorkbook.Sheets("lista").ListObjects("lista") ' ‹ módosítsd, ha más a lap neve

    ' ?? Oszlopszámok mentése (jobban átlátható és gyorsabb)
    nyelv1Col = tbl.ListColumns("ny_1").Index
    nyelv2Col = tbl.ListColumns("ny_2").Index
    nyelvOsszesCol = tbl.ListColumns("ny_osszefuz").Index ' ‹ ez legyen az összefûzött nyelvek oszlopa
    
    ' ?? Végigmegyünk minden soron
    For Each aktSor In tbl.ListRows
        ertek = aktSor.Range(1, tbl.ListColumns("tagozat").Index).value
        
        Select Case ertek
            Case 1000
                aktSor.Range(1, nyelv1Col).value = "angol"
                aktSor.Range(1, nyelv2Col).value = "spanyol"
            Case 2000
                aktSor.Range(1, nyelv1Col).value = "angol"
                aktSor.Range(1, nyelv2Col).value = "olasz"
            Case 3000
                aktSor.Range(1, nyelv1Col).value = "német"
                aktSor.Range(1, nyelv2Col).value = "angol"
            Case 4000
                aktSor.Range(1, nyelv1Col).value = "francia"
                aktSor.Range(1, nyelv2Col).value = "angol"
            Case 5000
                aktSor.Range(1, nyelv1Col).value = "angol"
                aktSor.Range(1, nyelv2Col).value = "német"
            Case Else
                aktSor.Range(1, nyelv1Col).value = ""
                aktSor.Range(1, nyelv2Col).value = ""
        End Select
        
        ' ?? Összefûzés
        aktSor.Range(1, nyelvOsszesCol).value = Trim( _
            aktSor.Range(1, nyelv1Col).value & " - " & aktSor.Range(1, nyelv2Col).value)
    Next aktSor

End Sub

