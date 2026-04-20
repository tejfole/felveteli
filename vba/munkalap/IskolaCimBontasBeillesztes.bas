Attribute VB_Name = "IskolaCimBontasBeillesztes"
Sub BontsdEsMasoldAdatokLista()
    Dim wsAdatok As Worksheet, wsLista As Worksheet
    Dim tblIsk As ListObject, tblList As ListObject
    Dim i As Long
    Dim irsz As String, varos As String, maradek As String
    Dim cimTeljes As String
    Dim colCimOssze As Range
    Dim colIskIrzs As Range, colIskVaros As Range, colIskCim As Range
    Dim nSor As Long
    
    Set wsAdatok = ThisWorkbook.Worksheets("adatok")
    Set wsLista = ThisWorkbook.Worksheets("lista")
    
    Set tblIsk = wsAdatok.ListObjects("iskola")
    Set tblList = wsLista.ListObjects("lista")
    
    On Error Resume Next
    Set colCimOssze = tblIsk.ListColumns("cim_ossze").DataBodyRange
    Set colIskIrzs = tblList.ListColumns("isk_irsz").DataBodyRange
    Set colIskVaros = tblList.ListColumns("isk_varos").DataBodyRange
    Set colIskCim = tblList.ListColumns("isk_utca").DataBodyRange
    On Error GoTo 0
    
    If colCimOssze Is Nothing Then
        MsgBox "'cim_ossze' oszlop nem található az 'iskolak' táblában az 'adatok' munkalapon!", vbCritical
        Exit Sub
    End If
    If colIskIrzs Is Nothing Or colIskVaros Is Nothing Or colIskCim Is Nothing Then
        MsgBox "Nem található valamelyik oszlop ('isk_irsz', 'isk_varos', 'isk_utca') a 'listak' táblában a 'lista' munkalapon!", vbCritical
        Exit Sub
    End If
    
    nSor = tblIsk.ListRows.Count
    If tblList.ListRows.Count < nSor Then
        MsgBox "A 'listak' tábla kevesebb sort tartalmaz, mint az 'iskolak' tábla!", vbExclamation
        Exit Sub
    End If
    
    For i = 1 To nSor
        cimTeljes = colCimOssze.Cells(i).value
        Call CimSzetszed(cimTeljes, irsz, varos, maradek)
        
        colIskIrzs.Cells(i).value = irsz
        colIskVaros.Cells(i).value = varos
        colIskCim.Cells(i).value = maradek
    Next i
    
    MsgBox "Szétbontás kész!"
End Sub

Sub CimSzetszed(cim As String, ByRef irsz As String, ByRef varos As String, ByRef maradek As String)
    Dim parts() As String
    Dim idx As Long
    
    irsz = ""
    varos = ""
    maradek = ""
    
    If Len(Trim(cim)) = 0 Then Exit Sub
    
    parts = Split(Trim(cim), " ")
    
    If UBound(parts) < 1 Then Exit Sub
    
    irsz = parts(0)
    varos = parts(1)
    
    idx = InStr(cim, varos)
    If idx > 0 Then
        maradek = Trim(mid(cim, idx + Len(varos)))
    End If
End Sub


