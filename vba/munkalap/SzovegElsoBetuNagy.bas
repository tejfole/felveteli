Attribute VB_Name = "SzovegElsoBetuNagy"
Function ElsoBetuNagy(szoveg As String) As String
    If Len(szoveg) > 0 Then
        ElsoBetuNagy = UCase(Left(szoveg, 1)) & mid(szoveg, 2)
    Else
        ElsoBetuNagy = ""
    End If
End Function

