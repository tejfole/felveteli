Attribute VB_Name = "EmailKeszites"
Sub EmailKeszitesOutlook()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim listaT As ListObject
    Dim i As Long
    Dim nev As String
    Dim emailcim As String
    Dim megszolit As String
    Dim szoveg As String
    Dim filePath As String
    Dim attachFile As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set listaT = ThisWorkbook.Sheets("lista").ListObjects("lista")
    
    filePath = ThisWorkbook.path & "\PDF_Ertesitok\"
    
    For i = 1 To listaT.ListRows.Count
        nev = listaT.ListColumns("nev").DataBodyRange.Cells(i, 1).value
        szoveg = listaT.ListColumns("szoveg").DataBodyRange.Cells(i, 1).value
        megszolit = listaT.ListColumns("megszolit").DataBodyRange.Cells(i, 1).value
        emailcim = listaT.ListColumns("email").DataBodyRange.Cells(i, 1).value ' kell legyen egy "email" oszlop
        
        If szoveg <> "" And emailcim <> "" Then
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = emailcim
                .Subject = "Felvételi Értesítés - " & nev
                .body = megszolit & " " & nev & "," & vbNewLine & vbNewLine & _
                        szoveg & vbNewLine & vbNewLine & _
                        "Üdvözlettel," & vbNewLine & "Felvételi Osztály"
                
                ' Csatoljuk a PDF-et, ha van
                attachFile = filePath & nev & ".pdf"
                If Dir(attachFile) <> "" Then
                    .Attachments.add attachFile
                End If
                
                .Display ' vagy .Send ha azonnal küldeni akarod
            End With
        End If
    Next i
    
    MsgBox "? Az összes e-mail elõnézetre megnyílt!", vbInformation

End Sub


