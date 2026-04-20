Attribute VB_Name = "UjEVInditas"
Option Explicit

Public Sub UjEvInditasa(Optional control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook: Set wb = ThisWorkbook

    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "Új év indítása:" & vbCrLf & _
        " - törli a nem megtartandó munkalapokat" & vbCrLf & _
        " - kiüríti a diakadat és rangsor táblákat" & vbCrLf & _
        " - 1 üres sort bennhagy a táblákban" & vbCrLf & vbCrLf & _
        "Biztosan folytatod?", _
        vbYesNo + vbExclamation, "Új év indítása")
    If resp <> vbYes Then Exit Sub

    Dim keep As Object
    Set keep = CreateObject("Scripting.Dictionary")
    keep.CompareMode = 1 ' textcompare

    ' --- megtartandó lapok ---
    keep("adatok") = True
    keep("diakadat") = True
    keep("rangsor") = True
    keep("lista") = True
    keep("tagozat") = True
    keep("TanteremLista") = True

    ' IDE tudsz még hozzáírni:
    ' keep("valami") = True
    ' keep("masiklap") = True

    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    ' Visszafelé törlünk, ez biztonságosabb
    Dim i As Long
    For i = wb.Worksheets.Count To 1 Step -1
        Dim ws As Worksheet
        Set ws = wb.Worksheets(i)

        If Not keep.Exists(ws.Name) Then
            ws.Delete
        End If
    Next i

    Application.DisplayAlerts = prevAlerts

    ' Táblák ürítése, 1 üres sor bennhagyásával
    ResetTableToSingleEmptyRow wb, "diakadat"
    ResetTableToSingleEmptyRow wb, "rangsor"

    MsgBox "Kész. A felesleges lapok törölve, a táblák kiürítve." & vbCrLf & _
           "Most mentsd el új néven, pl.: Felveteli_" & Year(Date) & ".xlsm", vbInformation
    Exit Sub

EH:
    Application.DisplayAlerts = prevAlerts
    MsgBox "Hiba: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Private Sub ResetTableToSingleEmptyRow(ByVal wb As Workbook, ByVal tableName As String)
    Dim lo As ListObject
    Set lo = FindTableByName(wb, tableName)
    If lo Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Ha nincs sor, adjunk hozzá egyet
    If lo.ListRows.Count = 0 Then
        lo.ListRows.add
    End If

    ' Ha több sor van, töröljük az elsõ után következõket
    Do While lo.ListRows.Count > 1
        lo.ListRows(lo.ListRows.Count).Delete
    Loop

    ' Az egyetlen megmaradt sor tartalmának ürítése
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.rows(1).ClearContents
    End If

    Application.ScreenUpdating = True
End Sub

Private Function FindTableByName(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If LCase$(lo.Name) = LCase$(tableName) Then
                Set FindTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

