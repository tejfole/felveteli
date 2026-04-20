Attribute VB_Name = "modIdopont"
Public Sub TEST_IdopontRendszer()
    On Error GoTo EH

    Dim loT As ListObject
    Set loT = GetIdopontTabla_V2()

    If loT Is Nothing Then
        MsgBox "NINCS idõpont tábla (GetIdopontTabla_V2 Nothing).", vbCritical
        Exit Sub
    End If

    MsgBox "OK: megvan az idõpont tábla." & vbCrLf & _
           "Lap: " & loT.parent.Name & vbCrLf & _
           "Tábla: " & loT.Name & vbCrLf & _
           "Sorok: " & loT.ListRows.Count, vbInformation

    ' teszt listázás: írjunk ki 1-2 aktív idõpontot
    If loT.ListRows.Count = 0 Then
        MsgBox "Az idõpont tábla üres. Vegyél fel idõpontot!", vbExclamation
        Exit Sub
    End If

    Dim iDt As Long: iDt = loT.ListColumns("datum_nap").Index
    Dim iAk As Long: iAk = loT.ListColumns("aktiv").Index

    Dim arr As Variant: arr = loT.DataBodyRange.value
    Dim r As Long, msg As String: msg = "Elsõ 10 sor:" & vbCrLf

    For r = 1 To Application.Min(10, UBound(arr, 1))
        msg = msg & r & ") " & CStr(arr(r, iDt)) & " | aktiv=" & CStr(arr(r, iAk)) & vbCrLf
    Next r

    MsgBox msg, vbInformation
    Exit Sub

EH:
    MsgBox "TESZT hiba: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

