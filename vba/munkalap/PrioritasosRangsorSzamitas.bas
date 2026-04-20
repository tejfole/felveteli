Attribute VB_Name = "PrioritasosRangsorSzamitas"
Sub SzamoljEgyediPrioritasosRangsor(Optional control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Dim tbl As ListObject
    Dim dataArr As Variant
    Dim i As Long, n As Long
    Dim colPont As Long, colRangsor As Long
    Dim colHatranyos As Long, colIranyitoszam As Long, colTestver As Long
    Dim startTime As Double
    
    startTime = Timer
    
    ' Képernyõfrissítés kikapcsolása (KRITIKUS 400 diáknál!)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Tábla keresése
    Set tbl = FindTable("diakadat")
    If tbl Is Nothing Then
        MsgBox "?? A 'diakadat' tábla nem található!", vbCritical
        GoTo Cleanup
    End If
    
    If tbl.ListRows.Count = 0 Then
        MsgBox "?? A táblázat üres!", vbExclamation
        GoTo Cleanup
    End If

    ' Oszlop indexek (hibakezeléssel)
    On Error Resume Next
    colPont = tbl.ListColumns("p_mindossz").Index
    colHatranyos = tbl.ListColumns("f_hatranyos").Index
    colIranyitoszam = tbl.ListColumns("I_ker_irsz").Index
    colTestver = tbl.ListColumns("f_testver").Index
    colRangsor = tbl.ListColumns("rangsor").Index
    
    If Err.Number <> 0 Then
        MsgBox "? Hiányzó oszlop(ok) a táblázatban!" & vbCrLf & Err.Description, vbCritical
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler

    ' Adatok beolvasása
    dataArr = tbl.DataBodyRange.value
    n = UBound(dataArr, 1)

    ' Ideiglenes rendezési lista: sorIndex, pont, prioritás
    Dim rangLista() As Variant
    ReDim rangLista(1 To n, 1 To 3)

    ' Prioritások kiszámítása
    For i = 1 To n
        rangLista(i, 1) = i ' Eredeti sorindex
        rangLista(i, 2) = SafeVal(dataArr(i, colPont)) ' Pontszám
        
        ' Prioritás súlyozása
        Dim prior As Long
        prior = 0
        If IsChecked(dataArr(i, colHatranyos)) Then prior = prior + 4
        If IsChecked(dataArr(i, colIranyitoszam)) Then prior = prior + 2
        If IsChecked(dataArr(i, colTestver)) Then prior = prior + 1
        
        rangLista(i, 3) = prior
    Next i

    ' QUICKSORT RENDEZÉS (10-30x gyorsabb 400 diáknál!)
    Call QuickSortRangLista(rangLista, 1, n)

    ' Rangsor visszaírása
    For i = 1 To n
        Dim eredetiSorIndex As Long
        eredetiSorIndex = rangLista(i, 1)
        dataArr(eredetiSorIndex, colRangsor) = i
    Next i

    ' Visszaírás táblába
    tbl.DataBodyRange.value = dataArr

    ' Színezés
    Call SzinezzTopEsKevesPontokatRangsorban
    
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    If Err.Number = 0 Then
        Dim elapsed As Double
        elapsed = Round(Timer - startTime, 2)
        MsgBox "? Egyedi prioritásos rangsor kiszámítva!" & vbCrLf & _
               "Feldolgozott diákok: " & n & vbCrLf & _
               "Futási idõ: " & elapsed & " mp", vbInformation
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "? Hiba történt: " & Err.Description & vbCrLf & _
           "Hibakód: " & Err.Number, vbCritical
End Sub

' ========== SEGÉDFÜGGVÉNYEK ==========

Private Function FindTable(tableName As String) As ListObject
    ' Tábla keresése
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = tableName Then
                Set FindTable = tbl
                Exit Function
            End If
        Next tbl
    Next ws
    
    Set FindTable = Nothing
End Function

Private Function SafeVal(ByVal value As Variant) As Double
    ' Biztonságos számérték olvasás
    If IsEmpty(value) Then
        SafeVal = 0
    ElseIf IsError(value) Then
        SafeVal = 0
    ElseIf IsNumeric(value) Then
        SafeVal = CDbl(value)
    Else
        SafeVal = val(CStr(value))
    End If
End Function

Private Function IsChecked(ByVal value As Variant) As Boolean
    ' Ellenõrzi, hogy a cella "x"-et tartalmaz-e
    If IsEmpty(value) Then
        IsChecked = False
    ElseIf IsError(value) Then
        IsChecked = False
    Else
        Dim strVal As String
        strVal = LCase(Trim(CStr(value)))
        IsChecked = (strVal = "x" Or strVal = "igen" Or strVal = "true")
    End If
End Function

' ========== QUICKSORT RENDEZÉS (400 diákhoz optimalizált!) ==========

Private Sub QuickSortRangLista(ByRef rangLista() As Variant, ByVal low As Long, ByVal high As Long)
    ' QuickSort algoritmus - O(n log n) komplexitás
    ' 400 diáknál ~5000 összehasonlítás vs buborék ~160,000!
    
    If low < high Then
        Dim pivot As Long
        pivot = Partition(rangLista, low, high)
        
        ' Rekurzív rendezés
        QuickSortRangLista rangLista, low, pivot - 1
        QuickSortRangLista rangLista, pivot + 1, high
    End If
End Sub

Private Function Partition(ByRef rangLista() As Variant, ByVal low As Long, ByVal high As Long) As Long
    ' Partíciók létrehozása a pivot elem körül
    Dim pivotPont As Double
    Dim pivotPrior As Long
    Dim i As Long, j As Long
    
    ' Pivot elem kiválasztása (utolsó elem)
    pivotPont = rangLista(high, 2)
    pivotPrior = rangLista(high, 3)
    i = low - 1
    
    ' Elemek átrendezése a pivot körül
    For j = low To high - 1
        ' RENDEZÉSI LOGIKA:
        ' 1. Nagyobb pontszám elõrébb
        ' 2. Azonos pont esetén nagyobb prioritás elõrébb
        If rangLista(j, 2) > pivotPont Or _
           (rangLista(j, 2) = pivotPont And rangLista(j, 3) > pivotPrior) Then
            i = i + 1
            SwapRows rangLista, i, j
        End If
    Next j
    
    ' Pivot elem helyére rakása
    SwapRows rangLista, i + 1, high
    Partition = i + 1
End Function

Private Sub SwapRows(ByRef rangLista() As Variant, ByVal i As Long, ByVal j As Long)
    ' Két sor cseréje a tömbben
    Dim temp1 As Variant, temp2 As Variant, temp3 As Variant
    
    temp1 = rangLista(i, 1)
    temp2 = rangLista(i, 2)
    temp3 = rangLista(i, 3)
    
    rangLista(i, 1) = rangLista(j, 1)
    rangLista(i, 2) = rangLista(j, 2)
    rangLista(i, 3) = rangLista(j, 3)
    
    rangLista(j, 1) = temp1
    rangLista(j, 2) = temp2
    rangLista(j, 3) = temp3
End Sub

