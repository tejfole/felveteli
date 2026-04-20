Attribute VB_Name = "modIdopontDebug"
Option Explicit

Public Sub IDOPONT_EnableEvents()
    Application.EnableEvents = True
    MsgBox "EnableEvents = TRUE", vbInformation
End Sub

Public Sub IDOPONT_Test_Assign_ActiveRow()
    On Error GoTo EH

    Application.EnableEvents = True

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("diakadat")
    Dim lo As ListObject: Set lo = ws.ListObjects("diakadat")

    If lo.DataBodyRange Is Nothing Then
        MsgBox "A diakadat tábla üres.", vbExclamation
        Exit Sub
    End If

    ' Aktív cella sorából számoljuk a táblán belüli rowIdx-t
    Dim rowIdx As Long
    rowIdx = ActiveCell.Row - lo.DataBodyRange.Row + 1

    If rowIdx < 1 Or rowIdx > lo.ListRows.Count Then
        MsgBox "Kattints a diakadat tábla egyik adat sorára (a táblán belül) és futtasd újra.", vbExclamation
        Exit Sub
    End If

    Dim iBiz As Long: iBiz = lo.ListColumns("bizottsag").Index
    Dim iDt As Long: iDt = lo.ListColumns("datum_nap").Index

    Dim biz As Long
    biz = CLng(val(lo.DataBodyRange.Cells(rowIdx, iBiz).value))

    If biz < 1 Or biz > 10 Then
        MsgBox "Ezen a soron nincs 1–10 közötti bizottság szám a 'bizottsag' oszlopban.", vbExclamation
        Exit Sub
    End If

    ' Ha már van datum_nap, ne írjuk felül
    If Trim$(CStr(lo.DataBodyRange.Cells(rowIdx, iDt).value)) <> "" Then
        MsgBox "Ezen a soron már van datum_nap. (Nem írjuk felül.)", vbInformation
        Exit Sub
    End If

    ' Itt derül ki, hogy a kiosztó sub elérhetõ-e és mûködik-e
    AssignDatumNap_FromIdopontTabla lo, rowIdx, biz, 4

    MsgBox "Kiosztás lefutott (ha volt szabad idõpont).", vbInformation
    Exit Sub

EH:
    MsgBox "Hiba a kézi tesztben: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub
