VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Munka2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("B1")) Is Nothing Then
       Call NyelvekBeirasaOsszefuzve   ' Saját makró neve
    End If
End Sub


