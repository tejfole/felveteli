VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Re-entrancy védelem: egy Save hívás alatt ne fusson kétszer a backup
Private savingBackup As Boolean

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    On Error GoTo Vege

    ' Csak akkor reagáljunk, ha az események engedélyezve vannak
    If Not Application.EnableEvents Then GoTo Vege

    ' Ha a lista munkalap B1 cellája változik
    If sh.Name = "lista" Then
        If Not Application.Intersect(Target, sh.Range("B1")) Is Nothing Then
            Dim hadEvents As Boolean
            hadEvents = Application.EnableEvents
            Application.EnableEvents = False

            SzovegRendelesStrukturaltTablakkal_WriteIfChanged

            If hadEvents Then Application.EnableEvents = True
        End If
    End If

Vege:
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo Vege

    If savingBackup Then GoTo Vege
    savingBackup = True

    ' Ne csináljunk backupot read-only megnyitásnál
    If ThisWorkbook.ReadOnly Then GoTo Vege

    ' Ha nincs elmentési helye (új fájl), nincs hova menteni másolatot
    If ThisWorkbook.path = "" Then GoTo Vege

    ' Csak ha ténylegesen módosult
    If Not ThisWorkbook.Saved Then
        SaveVersionedCopy_Logged
    End If

Vege:
    savingBackup = False
End Sub
