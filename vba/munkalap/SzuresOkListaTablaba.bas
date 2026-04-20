Attribute VB_Name = "SzuresOkListaTablaba"
Option Explicit

Public Function SzuressOk(Optional ByVal Valasztas As String = "", _
                          Optional ByVal SorIndex As Long = 0, _
                          Optional ByVal KeresettErtek As Variant = "x") As String
    On Error GoTo Hibakezeles

    Dim nevek As Variant, okok As Variant
    BuildSzuresLista Valasztas, KeresettErtek, nevek, okok

    If IsEmpty(okok) Then
        SzuressOk = ""
        Exit Function
    End If

    If SorIndex <= 0 Then
        SzuressOk = ""
    ElseIf SorIndex >= LBound(okok) And SorIndex <= UBound(okok) Then
        SzuressOk = CStr(okok(SorIndex))
    Else
        SzuressOk = ""
    End If

    Exit Function

Hibakezeles:
    SzuressOk = ""
End Function

Private Sub BuildSzuresLista(ByVal Valasztas As String, ByVal KeresettErtek As Variant, _
                             ByRef nevekOut As Variant, ByRef okokOut As Variant)
    On Error GoTo Hibakezeles

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim d As Variant

    Set ws = ThisWorkbook.Worksheets("rangsor")
    Set tbl = ws.ListObjects("rangsor")

    If tbl.DataBodyRange Is Nothing Then Exit Sub
    d = tbl.DataBodyRange.value

    If Len(Valasztas) = 0 Then
        On Error Resume Next
        Valasztas = NormText(Application.Caller.Worksheet.Range("B1").value)
        On Error GoTo Hibakezeles
    Else
        Valasztas = NormText(Valasztas)
    End If

    Dim cNev As Long, cIras As Long, cElut As Long, cVissza As Long
    cNev = GetRequiredColIndex(tbl, "nev")
    cIras = GetRequiredColIndex(tbl, "irasbeliossz")
    cElut = GetRequiredColIndex(tbl, "elut")
    cVissza = GetRequiredColIndex(tbl, "visszalepett")

    Dim cJ(1 To 4) As Long
    cJ(1) = GetRequiredColIndex(tbl, "j_1000")
    cJ(2) = GetRequiredColIndex(tbl, "j_2000")
    cJ(3) = GetRequiredColIndex(tbl, "j_3000")
    cJ(4) = GetRequiredColIndex(tbl, "j_4000")

    Dim nevek() As String, okok() As String
    Dim n As Long, i As Long, j As Long

    ReDim nevek(1 To UBound(d, 1) * 6)
    ReDim okok(1 To UBound(d, 1) * 6)

    For i = 1 To UBound(d, 1)
        If IsX(d(i, cVissza)) Then GoTo NextI

        Dim nev As String
        nev = Trim$(CStr(d(i, cNev)))
        If Len(nev) = 0 Then GoTo NextI

        Select Case Valasztas
            Case "elut", "elutkevespont"
                ' kevéspont (<55)
                If IsNumeric(d(i, cIras)) And CDbl(d(i, cIras)) < 55 Then
                    n = n + 1
                    nevek(n) = nev
                    okok(n) = "kevéspont"
                End If

                ' elutasítva + tagozatok
                If IsX(d(i, cElut)) Then
                    For j = 1 To 4
                        If IsX(d(i, cJ(j))) Then
                            n = n + 1
                            nevek(n) = nev
                            Select Case j
                                Case 1: okok(n) = "1000"
                                Case 2: okok(n) = "2000"
                                Case 3: okok(n) = "3000"
                                Case 4: okok(n) = "4000"
                            End Select

                            If Valasztas = "elutkevespont" Then Exit For
                        End If
                    Next j
                End If

            Case "kevespont"
                If IsNumeric(d(i, cIras)) And CDbl(d(i, cIras)) < 55 Then
                    n = n + 1
                    nevek(n) = nev
                    okok(n) = "kevéspont"
                End If

            Case Else
                ' más választásokhoz itt most nem adunk okot
        End Select

NextI:
    Next i

    If n = 0 Then Exit Sub

    ReDim Preserve nevek(1 To n)
    ReDim Preserve okok(1 To n)

    nevekOut = nevek
    okokOut = okok
    Exit Sub

Hibakezeles:
    ' hibánál üres marad
End Sub

Private Function GetRequiredColIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    On Error GoTo NemTalalhato
    GetRequiredColIndex = tbl.ListColumns(colName).Index
    Exit Function

NemTalalhato:
    Err.Raise vbObjectError + 513, "SzuressOk", _
              "Hiányzó oszlop a(z) '" & tbl.Name & "' táblában: '" & colName & "'"
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

