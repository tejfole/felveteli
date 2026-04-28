Attribute VB_Name = "ReplaceInDocumentRange"
' Public, hibatűrő helyettesítő: cserél a megadott Range-ben, visszaadja, hogy sikerült-e
Public Function ReplaceInDocumentRange(rng As Object, findText As String, replaceText As String) As Boolean
    On Error GoTo ErrHandler
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .Execute Replace:=wdReplaceAll
    End With
    ReplaceInDocumentRange = True
    Exit Function
ErrHandler:
    ReplaceInDocumentRange = False
    Err.clear
End Function
