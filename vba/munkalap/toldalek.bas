Attribute VB_Name = "toldalek"
Function Toldalek2(szam As Integer) As String
    Select Case szam
    
        Case 1, 2, 4, 7, 9, 10
            Toldalek2 = szam & "-es"
        Case 8, 3
            Toldalek2 = szam & "-as"
        Case 5
            Toldalek2 = szam & "-ös"
        Case 6
            Toldalek2 = szam & "-os"
        Case Else
            Toldalek2 = szam & "-" ' Alapértelmezett (nem kellene elõfordulnia)
    End Select
End Function

