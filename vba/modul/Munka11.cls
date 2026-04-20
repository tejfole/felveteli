VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Munka11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const DIRTY_COL As Long = 26 ' Z

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    If Not Application.EnableEvents Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Target.Row < 2 Then Exit Sub
    If Target.Column = DIRTY_COL Then Exit Sub

    Application.EnableEvents = False

    Dim area As Range, c As Range
    For Each area In Target.Areas
        For Each c In area.Cells
            If c.Row >= 2 Then
                Me.Cells(c.Row, DIRTY_COL).value = 1
            End If
        Next c
    Next area

SafeExit:
    Application.EnableEvents = True
    On Error Resume Next
    ScheduleBiziMatrixCommit 1   ' <-- EZ a modBiziMatrixAuto-ban van
End Sub
