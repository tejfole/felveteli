Attribute VB_Name = "PontSzamitasSzinezve"
Option Explicit

Public Sub SzamoljPontokatTombosen(Optional control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim tbl As ListObject
    Set tbl = FindTable("diakadat")
    If tbl Is Nothing Then
        Err.Raise vbObjectError + 100, , "A 'diakadat' tábla nem található!"
    End If
    
    If tbl.ListRows.Count = 0 Then GoTo Cleanup
    
    ' Oszlopindexek
    Dim colMagyar As Long, colMatek As Long, colBizonyitvany As Long
    Dim colIrasbeli As Long, colSzorzo As Long, colBizi As Long
    Dim colSzobeli As Long, colMindOssz As Long
    
    colMagyar = GetColIndex(tbl, "p_magyar")
    colMatek = GetColIndex(tbl, "p_matek")
    colBizonyitvany = GetColIndex(tbl, "p_bizonyitvany")
    
    colIrasbeli = GetColIndex(tbl, "irasbeliossz")
    colSzorzo = GetColIndex(tbl, "irasbeliossz+szorzo")
    colBizi = GetColIndex(tbl, "biziirasbeliossz")
    
    ' ÚJ: a szóbeli már kész összpont (import tölti)
    colSzobeli = GetColIndex(tbl, "szobeli")
    colMindOssz = GetColIndex(tbl, "p_mindossz")
    
    Dim dataArr As Variant
    dataArr = tbl.DataBodyRange.value
    
    Dim i As Long
    For i = 1 To UBound(dataArr, 1)
        Dim magyar As Double, matek As Double, bizonyitvany As Double
        Dim irasbeli As Double, szorzo As Double, bizi As Double
        Dim szobeli As Double, mindossz As Double
        
        magyar = SafeVal(dataArr(i, colMagyar))
        matek = SafeVal(dataArr(i, colMatek))
        bizonyitvany = SafeVal(dataArr(i, colBizonyitvany))
        
        ' Importból jön, mi csak beolvassuk + kerekítjük
        szobeli = Round(SafeVal(dataArr(i, colSzobeli)), 2)
        
        irasbeli = magyar + matek
        szorzo = Round(matek * 1.25 + magyar * 1.25, 2)
        bizi = Round(szorzo + bizonyitvany, 2)
        mindossz = Round(bizi + szobeli, 2)
        
        dataArr(i, colIrasbeli) = irasbeli
        dataArr(i, colSzorzo) = szorzo
        dataArr(i, colBizi) = bizi
        dataArr(i, colSzobeli) = szobeli
        dataArr(i, colMindOssz) = mindossz
    Next i
    
    tbl.DataBodyRange.value = dataArr
    
    ' Színezések
    ApplyColumnColor tbl, "irasbeliossz", RGB(180, 220, 255)
    ApplyColumnColor tbl, "irasbeliossz+szorzo", RGB(180, 220, 255)
    ApplyColumnColor tbl, "biziirasbeliossz", RGB(180, 220, 255)
    ApplyColumnColor tbl, "szobeli", RGB(180, 220, 255)
    ApplyColumnColor tbl, "p_mindossz", RGB(255, 204, 153)

Cleanup:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    Exit Sub

ErrorHandler:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    
    MsgBox "Hiba történt: " & Err.Description & vbCrLf & _
           "Sor: " & i & vbCrLf & _
           "Hibakód: " & Err.Number, vbCritical
End Sub

Private Function GetColIndex(tbl As ListObject, colName As String) As Long
    On Error GoTo EH
    GetColIndex = tbl.ListColumns(colName).Index
    Exit Function
EH:
    Err.Raise vbObjectError + 101, , "Hiányzó oszlop a táblában: " & colName
End Function

Private Function FindTable(tableName As String) As ListObject
    Dim ws As Worksheet, t As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each t In ws.ListObjects
            If t.Name = tableName Then
                Set FindTable = t
                Exit Function
            End If
        Next t
    Next ws
    Set FindTable = Nothing
End Function

Private Function SafeVal(ByVal value As Variant) As Double
    If IsEmpty(value) Or IsError(value) Then
        SafeVal = 0
    ElseIf IsNumeric(value) Then
        SafeVal = CDbl(value)
    Else
        SafeVal = val(CStr(value))
    End If
End Function

Private Sub ApplyColumnColor(tbl As ListObject, columnName As String, color As Long)
    On Error Resume Next
    tbl.ListColumns(columnName).DataBodyRange.Interior.color = color
    On Error GoTo 0
End Sub

