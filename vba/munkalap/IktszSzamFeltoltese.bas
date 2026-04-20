Attribute VB_Name = "IktszSzamFeltoltese"
Sub KitoltIktsz_TablaAutomatikusan(Optional control As IRibbonControl)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim oktazonCol As ListColumn, iktszCol As ListColumn
    Dim r As ListRow
    Dim dict As Object
    Dim kezdoSzam As Long
    Dim megtalalva As Boolean
    
    megtalalva = False
    ' Minden munkalapot végignéz
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = "lista" Then
                megtalalva = True
                Exit For
            End If
        Next tbl
        If megtalalva Then Exit For
    Next ws
    
    If Not megtalalva Then
        MsgBox "Nem található 'lista' nevû tábla egyik munkalapon sem!", vbCritical
        Exit Sub
    End If
    
    ' Oszlopok keresése a táblában
    For Each col In tbl.ListColumns
        Select Case LCase(col.Name)
            Case "oktazon": Set oktazonCol = col
            Case "iktsz": Set iktszCol = col
        End Select
    Next col
    
    If oktazonCol Is Nothing Or iktszCol Is Nothing Then
        MsgBox "Hiányzik az 'oktazon' vagy 'iktsz' oszlop a táblában!", vbCritical
        Exit Sub
    End If
    
    ' Kezdõ szám bekérése
    kezdoSzam = CLng(InputBox("Add meg a kezdõ iktsz számot:", "Kezdõ szám", 1))
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Táblán belüli sorok feldolgozása
    For Each r In tbl.ListRows
        Dim oktazonVal As String
        oktazonVal = Trim(r.Range(1, oktazonCol.Index).value)
        
        If oktazonVal <> "" Then
            If Not dict.Exists(oktazonVal) Then
                dict.add oktazonVal, kezdoSzam
                kezdoSzam = kezdoSzam + 1
            End If
            r.Range(1, iktszCol.Index).value = dict(oktazonVal)
        Else
            r.Range(1, iktszCol.Index).value = ""
        End If
    Next r
    
    MsgBox "Az iktsz oszlop sikeresen feltöltve a 'lista' táblában!", vbInformation

End Sub

