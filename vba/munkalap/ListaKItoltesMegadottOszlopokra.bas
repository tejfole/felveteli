Attribute VB_Name = "ListaKItoltesMegadottOszlopokra"
Option Explicit

Private Function ColIndex(ByVal lo As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    ColIndex = lo.ListColumns(colName).Index
    On Error GoTo 0
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        Txt = ""
    Else
        Txt = Trim$(CStr(v & ""))
    End If
End Function

Public Sub ListaKitoltesCsakMegadottOszlopokba()
    On Error GoTo EH

    Dim wsL As Worksheet, wsD As Worksheet, wsR As Worksheet
    Dim loL As ListObject, loD As ListObject, loR As ListObject

    Set wsL = ThisWorkbook.Worksheets("lista")
    Set wsD = ThisWorkbook.Worksheets("diakadat")
    Set wsR = ThisWorkbook.Worksheets("rangsor")

    Set loL = wsL.ListObjects("lista")
    Set loD = wsD.ListObjects("diakadat")
    Set loR = wsR.ListObjects("rangsor")

    If loL.DataBodyRange Is Nothing Then Exit Sub
    If loD.DataBodyRange Is Nothing Then Exit Sub
    If loR.DataBodyRange Is Nothing Then Exit Sub

    Dim lOkt As Long
    Dim lNev As Long, lEmail As Long, lSzH As Long, lSzulI As Long
    Dim lOssz As Long, lTag As Long, lIskNev As Long, lIskMail As Long, lJelszo As Long

    lOkt = ColIndex(loL, "oktazon")
    lNev = ColIndex(loL, "a_nev")
    lEmail = ColIndex(loL, "email")
    lSzH = ColIndex(loL, "sz_h")
    lSzulI = ColIndex(loL, "szul_i")
    lOssz = ColIndex(loL, "osszpont")
    lTag = ColIndex(loL, "tagozat")
    lIskNev = ColIndex(loL, "isk_nev")
    lIskMail = ColIndex(loL, "isk_mail")
    lJelszo = ColIndex(loL, "jelszo")

    Dim dOkt As Long, dNev As Long, dEmail As Long, dSzH As Long, dSzulI As Long
    Dim dOssz As Long, dIskNev As Long, dIskMail As Long, dJelszo As Long

    dOkt = ColIndex(loD, "oktazon")
    dNev = ColIndex(loD, "f_a_nev")
    dEmail = ColIndex(loD, "mail")
    dSzH = ColIndex(loD, "f_szul_hely")
    dSzulI = ColIndex(loD, "f_szul_ido")
    dOssz = ColIndex(loD, "p_mindossz")
    dIskNev = ColIndex(loD, "isknev")
    dIskMail = ColIndex(loD, "i_mail")
    dJelszo = ColIndex(loD, "jelszo")

    Dim rOkt As Long, rTag As Long
    rOkt = ColIndex(loR, "oktazon")
    rTag = ColIndex(loR, "tagozat")

    If lOkt = 0 Then Exit Sub
    If dOkt = 0 Then Exit Sub
    If rOkt = 0 Then Exit Sub

    Dim arrD As Variant, arrR As Variant
    arrD = loD.DataBodyRange.value
    arrR = loR.DataBodyRange.value

    Dim dictD As Object, dictR As Object
    Set dictD = CreateObject("Scripting.Dictionary")
    Set dictR = CreateObject("Scripting.Dictionary")

    Dim i As Long, key As String

    For i = 1 To UBound(arrD, 1)
        key = Txt(arrD(i, dOkt))
        If Len(key) > 0 Then
            If Not dictD.Exists(key) Then dictD.add key, i
        End If
    Next i

    For i = 1 To UBound(arrR, 1)
        key = Txt(arrR(i, rOkt))
        If Len(key) > 0 Then
            If Not dictR.Exists(key) Then dictR.add key, i
        End If
    Next i

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim r As Long, rowD As Long, rowR As Long, okt As String

    For r = 1 To loL.ListRows.Count
        okt = Txt(loL.DataBodyRange.Cells(r, lOkt).value)
        If Len(okt) = 0 Then GoTo NextR

        If dictD.Exists(okt) Then
            rowD = CLng(dictD(okt))

            If lNev > 0 Then loL.DataBodyRange.Cells(r, lNev).value = arrD(rowD, dNev)
            If lEmail > 0 Then loL.DataBodyRange.Cells(r, lEmail).value = arrD(rowD, dEmail)
            If lSzH > 0 Then loL.DataBodyRange.Cells(r, lSzH).value = arrD(rowD, dSzH)
            If lSzulI > 0 Then loL.DataBodyRange.Cells(r, lSzulI).value = arrD(rowD, dSzulI)
            If lOssz > 0 Then loL.DataBodyRange.Cells(r, lOssz).value = arrD(rowD, dOssz)
            If lIskNev > 0 Then loL.DataBodyRange.Cells(r, lIskNev).value = arrD(rowD, dIskNev)
            If lIskMail > 0 Then loL.DataBodyRange.Cells(r, lIskMail).value = arrD(rowD, dIskMail)
            If lJelszo > 0 Then loL.DataBodyRange.Cells(r, lJelszo).value = arrD(rowD, dJelszo)
        End If

        If dictR.Exists(okt) Then
            rowR = CLng(dictR(okt))
            If lTag > 0 Then loL.DataBodyRange.Cells(r, lTag).value = arrR(rowR, rTag)
        End If

NextR:
    Next r

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

EH:
    MsgBox "Hiba: " & Err.Number & " - " & Err.Description, vbCritical
    Resume SafeExit
End Sub
