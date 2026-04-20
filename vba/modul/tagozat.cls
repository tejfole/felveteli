VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Munka3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo EH
    If Intersect(Target, Me.Range("B1")) Is Nothing Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Tagozat_Refresh True

SafeExit:
    Application.EnableEvents = prevEvents
    Exit Sub

EH:
    MsgBox "Worksheet_Change hiba: " & Err.Number & " - " & Err.Description, vbCritical
    Resume SafeExit
End Sub
