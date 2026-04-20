Attribute VB_Name = "PonthatarTorles"
Sub TorolPonthatarA14(Optional control As IRibbonControl)

    Dim wsAdatok As Worksheet

    On Error Resume Next
    Set wsAdatok = ThisWorkbook.Sheets("adatok")
    On Error GoTo 0

    If wsAdatok Is Nothing Then
        MsgBox "? A 'adatok' nevû munkalap nem található!", vbCritical
        Exit Sub
    End If

    wsAdatok.Range("A14").ClearContents

    MsgBox "? Ponthatár törölve az adatok!A14 cellából.", vbInformation

End Sub

