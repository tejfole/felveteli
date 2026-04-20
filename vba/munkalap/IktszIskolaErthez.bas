Attribute VB_Name = "IktszIskolaErthez"
Sub KitoltIktsz_TablaAutomatikusan()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim isknevCol As ListColumn, iktszCol As ListColumn
    Dim r As ListRow
    Dim dict As Object
    Dim kezdoSzam As Long
    Dim megtalalva As Boolean
    Dim col As ListColumn
    
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
        MsgBox "? Nem található 'lista' nevû tábla egyik munkalapon sem!", vbCritical
        Exit Sub
    End If
    
    ' Oszlopok keresése a táblában
    For Each col In tbl.ListColumns
        Select Case LCase(col.Name)
            Case "isk_nev": Set isknevCol = col
            Case "iktsz": Set iktszCol = col
        End Select
    Next col
    
    If isknevCol Is Nothing Or iktszCol Is Nothing Then
        MsgBox "? Hiányzik az 'isk_nev' vagy 'iktsz' oszlop a táblában!", vbCritical
        Exit Sub
    End If
    
    ' Kezdõ szám bekérése
    kezdoSzam = CLng(InputBox("Add meg a kezdõ iktatószámot:", "Kezdõ iktsz", 1))
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Táblán belüli sorok feldolgozása
    For Each r In tbl.ListRows
        Dim isknev As String
        isknev = Trim(r.Range(1, isknevCol.Index).value)
        
        If isknev <> "" Then
            If Not dict.Exists(isknev) Then
                dict.add isknev, kezdoSzam
                kezdoSzam = kezdoSzam + 1
            End If
            r.Range(1, iktszCol.Index).value = dict(isknev)
        Else
            r.Range(1, iktszCol.Index).value = ""
        End If
    Next r
    
    MsgBox "? Az iktsz oszlop sikeresen feltöltve az 'isk_nev' alapján!", vbInformation

End Sub

