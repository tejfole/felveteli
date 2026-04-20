Attribute VB_Name = "SzuressNevListaTablaba"
Option Explicit

Public Function SzuressNev(Optional ByVal Valasztas As String = "", _
                           Optional ByVal SorIndex As Long = 0, _
                           Optional ByVal KeresettErtek As Variant = "x") As Variant
    On Error GoTo Hibakezeles

    Dim nevek As Variant
    nevek = BuildNevekLista(Valasztas, KeresettErtek)

    If IsEmpty(nevek) Then
        SzuressNev = ""
        Exit Function
    End If

    If IsArray(nevek) Then
        Dim lb As Long, ub As Long
        lb = LBound(nevek): ub = UBound(nevek)

        If ub < lb Then
            SzuressNev = ""
        ElseIf SorIndex <= 0 Then
            SzuressNev = Application.Transpose(nevek)
        ElseIf SorIndex >= lb And SorIndex <= ub Then
            SzuressNev = nevek(SorIndex)
        Else
            SzuressNev = ""
        End If
    Else
        SzuressNev = CStr(nevek)
    End If
    Exit Function

Hibakezeles:
    SzuressNev = ""
End Function

Private Function BuildNevekLista(ByVal Valasztas As String, ByVal KeresettErtek As Variant) As Variant
    On Error GoTo Hibakezeles

    Dim ws As Worksheet, tbl As ListObject, d As Variant
    Set ws = ThisWorkbook.Worksheets("rangsor")
    Set tbl = ws.ListObjects("rangsor")

    If tbl.DataBodyRange Is Nothing Then
        BuildNevekLista = Empty
        Exit Function
    End If
    d = tbl.DataBodyRange.value

    If Len(Valasztas) = 0 Then
        On Error Resume Next
        Valasztas = NormText(Application.Caller.Worksheet.Range("B1").value)
        On Error GoTo Hibakezeles
    Else
        Valasztas = NormText(Valasztas)
    End If

    Dim cNev As Long, cIras As Long, cElut As Long, cVissza As Long, cFelvesz As Long, cMastValaszt As Long
    cNev = GetRequiredColIndex(tbl, "nev")
    cIras = GetRequiredColIndex(tbl, "irasbeliossz")
    cElut = GetRequiredColIndex(tbl, "elut")
    cVissza = GetRequiredColIndex(tbl, "visszalepett")
    cFelvesz = GetRequiredColIndex(tbl, "felvesz")
    cMastValaszt = GetRequiredColIndex(tbl, "mastvalaszt")

    Dim cJ(1 To 4) As Long
    cJ(1) = GetRequiredColIndex(tbl, "j_1000")
    cJ(2) = GetRequiredColIndex(tbl, "j_2000")
    cJ(3) = GetRequiredColIndex(tbl, "j_3000")
    cJ(4) = GetRequiredColIndex(tbl, "j_4000")

    Dim out() As String, n As Long, i As Long, j As Long
    ReDim out(1 To UBound(d, 1) * 6)

    Dim keresX As Boolean
    keresX = (NormText(KeresettErtek) = "x")

    For i = 1 To UBound(d, 1)
        Dim nev As String
        nev = Trim$(CStr(d(i, cNev)))
        If Len(nev) = 0 Then GoTo NextI

        ' választásfüggõ kizárások
        Dim skipRow As Boolean
        skipRow = False

        Select Case Valasztas
            Case "elut", "elutkevespont", "kevespont", "mastvalaszt"
                If IsX(d(i, cVissza)) Then skipRow = True
                If IsX(d(i, cFelvesz)) Then skipRow = True

            Case "felvesz"
                If IsX(d(i, cVissza)) Then skipRow = True

            Case "visszalep", "visszalepett"
                ' nincs kizárás

            Case Else
                ' ismeretlen választás
                skipRow = True
        End Select

        If skipRow Then GoTo NextI

        Select Case Valasztas
            Case "elut", "elutkevespont"
                ' kevéspont (<55)
                If IsNumeric(d(i, cIras)) And CDbl(d(i, cIras)) < 55 Then
                    n = n + 1
                    out(n) = nev
                End If

                ' elut + j_* tagozatok
                If IsX(d(i, cElut)) Then
                    For j = 1 To 4
                        If IsX(d(i, cJ(j))) Then
                            n = n + 1
                            out(n) = nev
                            If Valasztas = "elutkevespont" Then Exit For
                        End If
                    Next j
                End If

            Case "kevespont"
                If IsNumeric(d(i, cIras)) And CDbl(d(i, cIras)) < 55 Then
                    n = n + 1
                    out(n) = nev
                End If

            Case "felvesz"
                If keresX Then
                    If IsX(d(i, cFelvesz)) Then
                        n = n + 1
                        out(n) = nev
                    End If
                Else
                    If NormText(d(i, cFelvesz)) = NormText(KeresettErtek) Then
                        n = n + 1
                        out(n) = nev
                    End If
                End If

            Case "mastvalaszt"
                If keresX Then
                    If IsX(d(i, cMastValaszt)) Then
                        n = n + 1
                        out(n) = nev
                    End If
                Else
                    If NormText(d(i, cMastValaszt)) = NormText(KeresettErtek) Then
                        n = n + 1
                        out(n) = nev
                    End If
                End If

            Case "visszalep", "visszalepett"
                If keresX Then
                    If IsX(d(i, cVissza)) Then
                        n = n + 1
                        out(n) = nev
                    End If
                Else
                    If NormText(d(i, cVissza)) = NormText(KeresettErtek) Then
                        n = n + 1
                        out(n) = nev
                    End If
                End If
        End Select

NextI:
    Next i

    If n = 0 Then
        BuildNevekLista = Empty
    Else
        ReDim Preserve out(1 To n)
        BuildNevekLista = out
    End If
    Exit Function

Hibakezeles:
    BuildNevekLista = Empty
End Function

Private Function GetRequiredColIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    On Error GoTo NemTalalhato
    GetRequiredColIndex = tbl.ListColumns(colName).Index
    Exit Function
NemTalalhato:
    Err.Raise vbObjectError + 513, "SzuressNev", "Hiányzó oszlop: " & colName
End Function

Private Function NormText(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace$(s, ChrW$(160), " ")
    s = Replace$(s, vbTab, " ")
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, ChrW$(8203), "")
    s = Replace$(s, ChrW$(65279), "")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    NormText = LCase$(s)
End Function

Private Function IsX(ByVal v As Variant) As Boolean
    IsX = (NormText(v) = "x")
End Function

