Attribute VB_Name = "pdfMentes"
Sub ListaSorokPDFMentesSablonrol()

    Dim wsLista As Worksheet
    Dim wsSablon As Worksheet
    Dim listaT As ListObject
    Dim i As Long
    Dim nev As String
    Dim szoveg As String
    Dim megszolit As String
    Dim fileName As String
    Dim filePath As String
    Dim folderPicker As FileDialog
    
    ' Lapok beállítása
    Set wsLista = ThisWorkbook.Sheets("lista")
    Set wsSablon = ThisWorkbook.Sheets("sablon")
    Set listaT = wsLista.ListObjects("lista")
    
    ' Mappaválasztó ablak
    Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With folderPicker
        .title = "?? Válaszd ki a PDF-ek mentési mappáját"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "?? Mûvelet megszakítva. Nem lett kiválasztva mappa.", vbExclamation
            Exit Sub
        End If
        filePath = .SelectedItems(1)
    End With
    
    ' Végigmegyünk a lista táblán
    For i = 1 To listaT.ListRows.Count
        nev = listaT.ListColumns("nev").DataBodyRange.Cells(i, 1).value
        szoveg = listaT.ListColumns("szoveg").DataBodyRange.Cells(i, 1).value
        megszolit = listaT.ListColumns("megszolit").DataBodyRange.Cells(i, 1).value
        
        ' Csak ha van szöveg
        If szoveg <> "" Then
            ' Feltöltjük a sablonlap mezõit
            With wsSablon
                .Range("B2").value = megszolit & " " & nev
                .Range("B4").value = szoveg
            End With
            
            ' PDF mentése a sablonlapról
            fileName = nev & ".pdf"
            wsSablon.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                fileName:=filePath & "\" & fileName, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
            
            ' Törlés a sablonlapról (opcionális, de szép tiszta marad)
            wsSablon.Range("B2:B4").ClearContents
        End If
    Next i
    
    MsgBox "? PDF-ek elkészültek a kiválasztott mappába!", vbInformation

End Sub


