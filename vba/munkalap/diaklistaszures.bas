Attribute VB_Name = "diaklistaszures"
Option Explicit

' =========================
' BEÁLLÍTÁSOK
' =========================
Public Const SRC_RANGSOR_SHEET As String = "rangsor"
Public Const SRC_RANGSOR_TABLE As String = "rangsor"

Public Const SRC_DATA_SHEET As String = "diakadat"
Public Const SRC_DATA_TABLE As String = "diakadat"

Public Const KEY_COL As String = "oktazon"
Public Const NAME_COL As String = "f_nev"

Public Const ACCEPT_COL As String = "felvesz"   ' "x" = felvett
Public Const REJECT_COL As String = "elut"      ' "x" = elutasított
Public Const FLAG_VALUE As String = "x"

' Felvettek lista cél
Public Const DST_SHEET As String = "diak_lista"
Public Const DST_TABLE As String = "diak_lista_tbl"
Public Const DST_TOPLEFT As String = "C1"

' Esélyesek (elutasítottak) cél
Public Const CAN_SHEET As String = "eselyesek"
Public Const CAN_TABLE As String = "eselyesek_tbl"
Public Const CAN_TOPLEFT As String = "C1"
Public Const MAX_HIANY As Long = 5  ' ennyi pont hiányig "esélyes"

' =========================
' KIMENET OSZLOPOK
' =========================
Private Function OutCols_Felvettek() As Variant
    OutCols_Felvettek = Array( _
        "sorszam", _
        "r_irasbeliossz", _
        "r_biziirasbeliossz", _
        "r_p_mindossz", _
        "oktazon", _
        "f_nev", _
        "irasbeliossz", _
        "biziirasbeliossz", _
        "szobeli", _
        "p_mindossz", _
        "p_bizonyitvany" _
    )
End Function

Private Function OutCols_Eselyesek() As Variant
    OutCols_Eselyesek = Array( _
        "sorszam", _
        "oktazon", _
        "f_nev", _
        "irasbeliossz", _
        "biziirasbeliossz", _
        "szobeli", _
        "kuszob_iras", _
        "kuszob_bizi", _
        "kuszob_szobeli", _
        "hiany_iras", _
        "hiany_bizi", _
        "hiany_szobeli", _
        "hiany_min", _
        "bekerulhet" _
    )
End Function

' ======================================================================
' 1) FRISSÍTÉS – Felvettek lista (oktazon alapján pontok a diakadatból)
' ======================================================================
Public Sub Frissites_Felvettek_DiakLista_OktazonAlapjan()

    Dim wsR As Worksheet, wsD As Worksheet, wsDst As Worksheet
    Dim loR As ListObject, loD As ListObject, loDst As ListObject
    Dim colS As Variant

    Dim vKeyR As Variant, vAcc As Variant
    Dim vKeyD As Variant, vNev As Variant, vIras As Variant, vBizi As Variant
    Dim vSzob As Variant, vPmind As Variant, vPbiz As Variant

    Dim dictAcc As Object         ' oktazon -> True
    Dim dictRowD As Object        ' oktazon -> row index in diakadat arrays

    Dim keepKeys() As String, nKeep As Long
    Dim i As Long, rD As Long

    Dim arrOut() As Variant
    Dim tmpIras() As Variant, tmpBizi() As Variant, tmpPmind() As Variant
    Dim keepIdx() As Long
    Dim rankIras As Object, rankBizi As Object, rankPmind As Object

    On Error GoTo ErrHandler

    Set wsR = GetSheetOrNothing(SRC_RANGSOR_SHEET)
    If wsR Is Nothing Then Err.Raise vbObjectError + 11, , "Hiányzó munkalap: " & SRC_RANGSOR_SHEET
    Set loR = GetTableOrNothing(wsR, SRC_RANGSOR_TABLE)
    If loR Is Nothing Then Err.Raise vbObjectError + 12, , "Hiányzó tábla: " & SRC_RANGSOR_TABLE & " (" & SRC_RANGSOR_SHEET & ")"

    Set wsD = GetSheetOrNothing(SRC_DATA_SHEET)
    If wsD Is Nothing Then Err.Raise vbObjectError + 13, , "Hiányzó munkalap: " & SRC_DATA_SHEET
    Set loD = GetTableOrNothing(wsD, SRC_DATA_TABLE)
    If loD Is Nothing Then Err.Raise vbObjectError + 14, , "Hiányzó tábla: " & SRC_DATA_TABLE & " (" & SRC_DATA_SHEET & ")"

    RequireColSafe loR, KEY_COL, SRC_RANGSOR_TABLE
    RequireColSafe loR, ACCEPT_COL, SRC_RANGSOR_TABLE

    RequireColSafe loD, KEY_COL, SRC_DATA_TABLE
    RequireColSafe loD, NAME_COL, SRC_DATA_TABLE
    RequireColSafe loD, "irasbeliossz", SRC_DATA_TABLE
    RequireColSafe loD, "biziirasbeliossz", SRC_DATA_TABLE
    RequireColSafe loD, "szobeli", SRC_DATA_TABLE
    RequireColSafe loD, "p_mindossz", SRC_DATA_TABLE
    RequireColSafe loD, "p_bizonyitvany", SRC_DATA_TABLE

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    colS = OutCols_Felvettek()
    Set wsDst = GetOrCreateSheet(DST_SHEET)

    DeleteTableIfExists wsDst, DST_TABLE
    wsDst.Range("C1:Z50000").ClearContents

    If loR.DataBodyRange Is Nothing Or loD.DataBodyRange Is Nothing Then
        Set loDst = CreateTable(wsDst, DST_TABLE, DST_TOPLEFT, colS)
        GoTo SafeExit
    End If

    ' rangsor (felvettek)
    vKeyR = loR.ListColumns(KEY_COL).DataBodyRange.value
    vAcc = loR.ListColumns(ACCEPT_COL).DataBodyRange.value

    Set dictAcc = CreateObject("Scripting.Dictionary")
    dictAcc.CompareMode = vbTextCompare

    For i = 1 To UBound(vKeyR, 1)
        If LCase$(Trim$(CStr(vAcc(i, 1)))) = FLAG_VALUE Then
            If Len(Trim$(CStr(vKeyR(i, 1)))) > 0 Then dictAcc(CStr(vKeyR(i, 1))) = True
        End If
    Next i

    ' diakadat (pontok) + index oktazon szerint
    vKeyD = loD.ListColumns(KEY_COL).DataBodyRange.value
    vNev = loD.ListColumns(NAME_COL).DataBodyRange.value
    vIras = loD.ListColumns("irasbeliossz").DataBodyRange.value
    vBizi = loD.ListColumns("biziirasbeliossz").DataBodyRange.value
    vSzob = loD.ListColumns("szobeli").DataBodyRange.value
    vPmind = loD.ListColumns("p_mindossz").DataBodyRange.value
    vPbiz = loD.ListColumns("p_bizonyitvany").DataBodyRange.value

    Set dictRowD = CreateObject("Scripting.Dictionary")
    dictRowD.CompareMode = vbTextCompare

    For i = 1 To UBound(vKeyD, 1)
        Dim kD As String
        kD = Trim$(CStr(vKeyD(i, 1)))
        If Len(kD) > 0 Then
            If Not dictRowD.Exists(kD) Then dictRowD.add kD, i
        End If
    Next i

    ' metszet
    ReDim keepKeys(1 To dictAcc.Count)
    nKeep = 0
    Dim k As Variant
    For Each k In dictAcc.keys
        If dictRowD.Exists(CStr(k)) Then
            nKeep = nKeep + 1
            keepKeys(nKeep) = CStr(k)
        End If
    Next k

    Set loDst = CreateTable(wsDst, DST_TABLE, DST_TOPLEFT, colS)
    If nKeep = 0 Then GoTo SafeExit

    ' rangszámok a felvettek között (írás/bizi/p_mindossz)
    ReDim tmpIras(1 To nKeep, 1 To 1)
    ReDim tmpBizi(1 To nKeep, 1 To 1)
    ReDim tmpPmind(1 To nKeep, 1 To 1)
    ReDim keepIdx(1 To nKeep)

    For i = 1 To nKeep
        keepIdx(i) = i
        rD = CLng(dictRowD(keepKeys(i)))
        tmpIras(i, 1) = vIras(rD, 1)
        tmpBizi(i, 1) = vBizi(rD, 1)
        tmpPmind(i, 1) = vPmind(rD, 1)
    Next i

    Set rankIras = BuildRankDictFromValues(tmpIras, keepIdx, nKeep)
    Set rankBizi = BuildRankDictFromValues(tmpBizi, keepIdx, nKeep)
    Set rankPmind = BuildRankDictFromValues(tmpPmind, keepIdx, nKeep)

    ReDim arrOut(1 To nKeep, 1 To UBound(colS) + 1)

    For i = 1 To nKeep
        rD = CLng(dictRowD(keepKeys(i)))

        arrOut(i, 1) = i
        arrOut(i, 2) = GetRank(rankIras, vIras(rD, 1))
        arrOut(i, 3) = GetRank(rankBizi, vBizi(rD, 1))
        arrOut(i, 4) = GetRank(rankPmind, vPmind(rD, 1))

        arrOut(i, 5) = keepKeys(i)
        arrOut(i, 6) = vNev(rD, 1)
        arrOut(i, 7) = vIras(rD, 1)
        arrOut(i, 8) = vBizi(rD, 1)
        arrOut(i, 9) = vSzob(rD, 1)
        arrOut(i, 10) = vPmind(rD, 1)
        arrOut(i, 11) = vPbiz(rD, 1)
    Next i

    WriteArrayToTable loDst, arrOut
    loDst.Range.Columns.AutoFit

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Hiba a FRISSÍTÉS során: " & Err.Description, vbExclamation
End Sub

' ======================================================================
' 2) NÉZZÜK – elutasítottak közül kik vannak közel (írás + bizi + szóbeli)
' ======================================================================
Public Sub Nezzuk_Elutasitottak_IrasbeliEsely()

    Dim wsR As Worksheet, wsD As Worksheet, wsC As Worksheet
    Dim loR As ListObject, loD As ListObject, loC As ListObject
    Dim colS As Variant

    Dim dictRowD As Object, dictAcc As Object
    Dim vKeyR As Variant, vAcc As Variant, vRej As Variant
    Dim vKeyD As Variant, vNev As Variant, vIras As Variant, vBizi As Variant, vSzob As Variant

    Dim i As Long, rD As Long
    Dim accIras() As Double, accBizi() As Double, accSzob() As Double, nAcc As Long
    Dim cutI As Double, cutB As Double, cutS As Double

    Dim outKeys() As String, nOut As Long
    Dim arrOut() As Variant

    On Error GoTo ErrHandler

    Set wsR = GetSheetOrNothing(SRC_RANGSOR_SHEET)
    If wsR Is Nothing Then Err.Raise vbObjectError + 21, , "Hiányzó munkalap: " & SRC_RANGSOR_SHEET
    Set loR = GetTableOrNothing(wsR, SRC_RANGSOR_TABLE)
    If loR Is Nothing Then Err.Raise vbObjectError + 22, , "Hiányzó tábla: " & SRC_RANGSOR_TABLE & " (" & SRC_RANGSOR_SHEET & ")"

    Set wsD = GetSheetOrNothing(SRC_DATA_SHEET)
    If wsD Is Nothing Then Err.Raise vbObjectError + 23, , "Hiányzó munkalap: " & SRC_DATA_SHEET
    Set loD = GetTableOrNothing(wsD, SRC_DATA_TABLE)
    If loD Is Nothing Then Err.Raise vbObjectError + 24, , "Hiányzó tábla: " & SRC_DATA_TABLE & " (" & SRC_DATA_SHEET & ")"

    RequireColSafe loR, KEY_COL, SRC_RANGSOR_TABLE
    RequireColSafe loR, ACCEPT_COL, SRC_RANGSOR_TABLE
    RequireColSafe loR, REJECT_COL, SRC_RANGSOR_TABLE

    RequireColSafe loD, KEY_COL, SRC_DATA_TABLE
    RequireColSafe loD, NAME_COL, SRC_DATA_TABLE
    RequireColSafe loD, "irasbeliossz", SRC_DATA_TABLE
    RequireColSafe loD, "biziirasbeliossz", SRC_DATA_TABLE
    RequireColSafe loD, "szobeli", SRC_DATA_TABLE

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If loD.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 25, , "A '" & SRC_DATA_TABLE & "' tábla üres."
    If loR.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 26, , "A '" & SRC_RANGSOR_TABLE & "' tábla üres."

    ' diakadat arrays
    vKeyD = loD.ListColumns(KEY_COL).DataBodyRange.value
    vNev = loD.ListColumns(NAME_COL).DataBodyRange.value
    vIras = loD.ListColumns("irasbeliossz").DataBodyRange.value
    vBizi = loD.ListColumns("biziirasbeliossz").DataBodyRange.value
    vSzob = loD.ListColumns("szobeli").DataBodyRange.value

    Set dictRowD = CreateObject("Scripting.Dictionary")
    dictRowD.CompareMode = vbTextCompare

    For i = 1 To UBound(vKeyD, 1)
        Dim kD As String
        kD = Trim$(CStr(vKeyD(i, 1)))
        If Len(kD) > 0 Then
            If Not dictRowD.Exists(kD) Then dictRowD.add kD, i
        End If
    Next i

    ' rangsor arrays
    vKeyR = loR.ListColumns(KEY_COL).DataBodyRange.value
    vAcc = loR.ListColumns(ACCEPT_COL).DataBodyRange.value
    vRej = loR.ListColumns(REJECT_COL).DataBodyRange.value

    Set dictAcc = CreateObject("Scripting.Dictionary")
    dictAcc.CompareMode = vbTextCompare

    ReDim accIras(1 To UBound(vKeyR, 1))
    ReDim accBizi(1 To UBound(vKeyR, 1))
    ReDim accSzob(1 To UBound(vKeyR, 1))
    nAcc = 0

    ' küszöbök: felvettek 10. legalacsonyabb (külön-külön oszlop)
    For i = 1 To UBound(vKeyR, 1)
        Dim kR As String
        kR = Trim$(CStr(vKeyR(i, 1)))
        If Len(kR) = 0 Then GoTo NextAcc

        If LCase$(Trim$(CStr(vAcc(i, 1)))) = FLAG_VALUE Then
            If dictRowD.Exists(kR) Then
                rD = CLng(dictRowD(kR))
                If IsNumeric(vIras(rD, 1)) And IsNumeric(vBizi(rD, 1)) And IsNumeric(vSzob(rD, 1)) Then
                    nAcc = nAcc + 1
                    accIras(nAcc) = CDbl(vIras(rD, 1))
                    accBizi(nAcc) = CDbl(vBizi(rD, 1))
                    accSzob(nAcc) = CDbl(vSzob(rD, 1))
                    If Not dictAcc.Exists(kR) Then dictAcc.add kR, True
                End If
            End If
        End If
NextAcc:
    Next i

    If nAcc = 0 Then
        MsgBox "Nincs felvett / vagy hiányzik pont (írás+bizi+szóbeli) a felvettekhez a diakadat táblában.", vbInformation
        GoTo SafeExit
    End If

    ReDim Preserve accIras(1 To nAcc)
    ReDim Preserve accBizi(1 To nAcc)
    ReDim Preserve accSzob(1 To nAcc)

    QuickSortDbl accIras, 1, nAcc
    QuickSortDbl accBizi, 1, nAcc
    QuickSortDbl accSzob, 1, nAcc

    cutI = IIf(nAcc >= 10, accIras(10), accIras(1))
    cutB = IIf(nAcc >= 10, accBizi(10), accBizi(1))
    cutS = IIf(nAcc >= 10, accSzob(10), accSzob(1))

    ' elutasítottak, akik közel vannak bármelyik küszöbhöz
    ReDim outKeys(1 To UBound(vKeyR, 1))
    nOut = 0

    For i = 1 To UBound(vKeyR, 1)
        Dim keyX As String
        keyX = Trim$(CStr(vKeyR(i, 1)))
        If Len(keyX) = 0 Then GoTo NextRej

        If LCase$(Trim$(CStr(vRej(i, 1)))) = FLAG_VALUE Then
            If Not dictAcc.Exists(keyX) Then
                If dictRowD.Exists(keyX) Then
                    rD = CLng(dictRowD(keyX))
                    If IsNumeric(vIras(rD, 1)) And IsNumeric(vBizi(rD, 1)) And IsNumeric(vSzob(rD, 1)) Then
                        Dim hi As Double, hB As Double, hS As Double, hMin As Double
                        hi = cutI - CDbl(vIras(rD, 1))
                        hB = cutB - CDbl(vBizi(rD, 1))
                        hS = cutS - CDbl(vSzob(rD, 1))
                        hMin = WorksheetFunction.Min(hi, hB, hS)
                        If hMin <= MAX_HIANY Then
                            nOut = nOut + 1
                            outKeys(nOut) = keyX
                        End If
                    End If
                End If
            End If
        End If
NextRej:
    Next i

    ' kiírás
    colS = OutCols_Eselyesek()
    Set wsC = GetOrCreateSheet(CAN_SHEET)

    DeleteTableIfExists wsC, CAN_TABLE
    wsC.Range("C1:Z50000").ClearContents
    Set loC = CreateTable(wsC, CAN_TABLE, CAN_TOPLEFT, colS)

    If nOut = 0 Then
        MsgBox "Nincs olyan elutasított, aki " & MAX_HIANY & " ponton belül lenne (írás/bizi/szóbeli küszöbhöz).", vbInformation
        GoTo SafeExit
    End If

    ReDim arrOut(1 To nOut, 1 To UBound(colS) + 1)

    For i = 1 To nOut
        rD = CLng(dictRowD(outKeys(i)))

        Dim hI2 As Double, hB2 As Double, hS2 As Double, hMin2 As Double
        hI2 = cutI - CDbl(vIras(rD, 1))
        hB2 = cutB - CDbl(vBizi(rD, 1))
        hS2 = cutS - CDbl(vSzob(rD, 1))
        hMin2 = WorksheetFunction.Min(hI2, hB2, hS2)

        arrOut(i, 1) = i
        arrOut(i, 2) = outKeys(i)
        arrOut(i, 3) = vNev(rD, 1)
        arrOut(i, 4) = vIras(rD, 1)
        arrOut(i, 5) = vBizi(rD, 1)
        arrOut(i, 6) = vSzob(rD, 1)

        arrOut(i, 7) = cutI
        arrOut(i, 8) = cutB
        arrOut(i, 9) = cutS

        arrOut(i, 10) = hI2
        arrOut(i, 11) = hB2
        arrOut(i, 12) = hS2
        arrOut(i, 13) = hMin2
        arrOut(i, 14) = "IGEN"
    Next i

    WriteArrayToTable loC, arrOut
    loC.Range.Columns.AutoFit

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Hiba a NÉZZÜK során: " & Err.Description, vbExclamation
End Sub

' ======================================================================
' 3) EXPORT – Felvettek
' ======================================================================
Public Sub Export_Felvettek_UjMunkafuzetbe()
    ExportTableToNewWorkbook DST_SHEET, DST_TABLE, "felvettek", "felvettek_export.xlsx"
End Sub

' ======================================================================
' 4) EXPORT – Esélyek
' ======================================================================
Public Sub Export_Eselyek_UjMunkafuzetbe()
    ExportTableToNewWorkbook CAN_SHEET, CAN_TABLE, "eselyesek", "eselyesek_export.xlsx"
End Sub

' ======================================================================
' KOMPATIBILITÁSI MAKRÓNEVEK (ha a gombok régi nevekre mutatnak)
' ======================================================================
Public Sub Frissites_Felvettek_DiakLista()
    Frissites_Felvettek_DiakLista_OktazonAlapjan
End Sub

Public Sub Nezzuk()
    Nezzuk_Elutasitottak_IrasbeliEsely
End Sub

Public Sub Export_felvettek()
    Export_Felvettek_UjMunkafuzetbe
End Sub

Public Sub Export_eselyek()
    Export_Eselyek_UjMunkafuzetbe
End Sub

' ======================================================================
' EXPORT SEGÉD
' ======================================================================
Private Sub ExportTableToNewWorkbook(sheetName As String, tableName As String, outSheetName As String, defaultFileName As String)
    Dim ws As Worksheet, lo As ListObject
    Dim wbNew As Workbook, wsNew As Worksheet
    Dim savePath As Variant

    On Error GoTo ErrHandler

    Set ws = GetSheetOrNothing(sheetName)
    If ws Is Nothing Then
        MsgBox "Nem találom a munkalapot: '" & sheetName & "'.", vbExclamation
        Exit Sub
    End If

    Set lo = GetTableOrNothing(ws, tableName)
    If lo Is Nothing Then
        MsgBox "Nem találom a táblát: '" & tableName & "'.", vbExclamation
        Exit Sub
    End If

    If lo.DataBodyRange Is Nothing Then
        MsgBox "A tábla üres, nincs exportálható adat.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Set wbNew = Workbooks.add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    wsNew.Name = outSheetName

    lo.HeaderRowRange.Copy
    wsNew.Range("A1").PasteSpecial xlPasteValues

    lo.DataBodyRange.Copy
    wsNew.Range("A2").PasteSpecial xlPasteValues

    Application.CutCopyMode = False
    wsNew.UsedRange.Columns.AutoFit

    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, FileFilter:="Excel munkafüzet (*.xlsx), *.xlsx")
    If savePath <> False Then
        Application.DisplayAlerts = False
        wbNew.SaveAs fileName:=CStr(savePath), FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        MsgBox "Export kész:" & vbCrLf & CStr(savePath), vbInformation
    Else
        MsgBox "Mentés megszakítva. Az export munkafüzet megnyitva maradt.", vbInformation
    End If

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Hiba export közben: " & Err.Description, vbExclamation
End Sub

' ======================================================================
' SEGÉDEK: lap/tábla/ellenõrzés
' ======================================================================
Private Function GetSheetOrNothing(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetOrNothing = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetTableOrNothing(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set GetTableOrNothing = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function GetOrCreateSheet(n As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(n)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.add
        GetOrCreateSheet.Name = n
    End If
End Function

Private Sub RequireColSafe(lo As ListObject, colName As String, tableName As String)
    If Not HasListColumn(lo, colName) Then
        Err.Raise vbObjectError + 501, , "Hiányzó oszlop: '" & colName & "' a(z) '" & tableName & "' táblában."
    End If
End Sub

Private Function HasListColumn(lo As ListObject, colName As String) As Boolean
    On Error GoTo NoCol
    Dim lc As ListColumn
    Set lc = lo.ListColumns(colName)
    HasListColumn = True
    Exit Function
NoCol:
    HasListColumn = False
End Function

Private Sub DeleteTableIfExists(ws As Worksheet, tableName As String)
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If LCase$(lo.Name) = LCase$(tableName) Then
            lo.Unlist
            Exit For
        End If
    Next lo
End Sub

Private Function CreateTable(ws As Worksheet, tableName As String, topLeftAddr As String, colS As Variant) As ListObject
    Dim i As Long, tl As Range, lo As ListObject
    Set tl = ws.Range(topLeftAddr)

    For i = LBound(colS) To UBound(colS)
        tl.Offset(0, i).value = colS(i)
    Next i

    Set lo = ws.ListObjects.add(xlSrcRange, tl.Resize(2, UBound(colS) + 1), , xlYes)
    lo.Name = tableName
    lo.TableStyle = "TableStyleMedium2"
    Set CreateTable = lo
End Function

Private Sub WriteArrayToTable(lo As ListObject, arr As Variant)
    lo.Resize lo.Range.Resize(UBound(arr, 1) + 1, UBound(arr, 2))
    lo.DataBodyRange.value = arr
End Sub

' ======================================================================
' RANG – holtverseny: azonos érték azonos rang; csökkenõ (jobb a nagyobb)
' ======================================================================
Private Function BuildRankDictFromValues(vCol As Variant, keepIdx() As Long, nKeep As Long) As Object
    Dim dictCount As Object, dictRank As Object
    Dim i As Long, r As Long, key As String, val As Double

    Set dictCount = CreateObject("Scripting.Dictionary")
    dictCount.CompareMode = vbBinaryCompare

    For i = 1 To nKeep
        r = keepIdx(i)
        If IsNumeric(vCol(r, 1)) Then
            val = CDbl(vCol(r, 1))
            key = CStr(val)
            If dictCount.Exists(key) Then
                dictCount(key) = CLng(dictCount(key)) + 1
            Else
                dictCount.add key, 1
            End If
        End If
    Next i

    Set dictRank = CreateObject("Scripting.Dictionary")
    dictRank.CompareMode = vbBinaryCompare
    If dictCount.Count = 0 Then
        Set BuildRankDictFromValues = dictRank
        Exit Function
    End If

    Dim scores() As Double, k As Variant, n As Long
    ReDim scores(1 To dictCount.Count)
    n = 0
    For Each k In dictCount.keys
        n = n + 1
        scores(n) = CDbl(k)
    Next k

    QuickSortDblDesc scores, 1, UBound(scores)

    Dim runningGreater As Long
    runningGreater = 0
    For i = 1 To UBound(scores)
        key = CStr(scores(i))
        dictRank(key) = runningGreater + 1
        runningGreater = runningGreater + CLng(dictCount(key))
    Next i

    Set BuildRankDictFromValues = dictRank
End Function

Private Function GetRank(dictRank As Object, v As Variant) As Variant
    If IsNumeric(v) Then
        Dim key As String
        key = CStr(CDbl(v))
        If dictRank.Exists(key) Then
            GetRank = dictRank(key)
        Else
            GetRank = vbNullString
        End If
    Else
        GetRank = vbNullString
    End If
End Function

Private Sub QuickSortDbl(ByRef a() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Double, tmp As Double
    i = lo: j = hi
    pivot = a((lo + hi) \ 2)

    Do While i <= j
        Do While a(i) < pivot: i = i + 1: Loop
        Do While a(j) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortDbl a, lo, j
    If i < hi Then QuickSortDbl a, i, hi
End Sub

Private Sub QuickSortDblDesc(ByRef a() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Double, tmp As Double
    i = lo: j = hi
    pivot = a((lo + hi) \ 2)

    Do While i <= j
        Do While a(i) > pivot: i = i + 1: Loop
        Do While a(j) < pivot: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortDblDesc a, lo, j
    If i < hi Then QuickSortDblDesc a, i, hi
End Sub


