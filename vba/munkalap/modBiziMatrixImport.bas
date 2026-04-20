Attribute VB_Name = "modBiziMatrixImport"
Option Explicit

' =========================
' Beállítások
' =========================
Private Const MATRIX_SHEET As String = "bizonyitvany_matrix"
Private Const DIRTY_COL As Long = 26              ' Z
Private Const SRC_HDR_SUBJECT As Long = 1         ' Tantárgy (összevont is lehet)
Private Const SRC_HDR_YEAR As Long = 2            ' 1-4 évf.
Private Const SRC_FIRST_DATA_ROW As Long = 3

' Mátrix oszlopok
Private Const MAT_COL_OKTAZON As Long = 1         ' A
Private Const MAT_COL_NEV As Long = 2             ' B
Private Const MAT_FIRST_SUBJ_COL As Long = 3      ' C...

' automatikus commit a mátrix szerkesztése után (debounce)
Public Const MATRIX_AUTO_COMMIT As Boolean = True
Public NextBiziCommit As Date
Public BiziCommitScheduled As Boolean

' =========================================================
' RIBBON / PUBLIC belépési pontok (a te Ribbon kódodhoz)
' =========================================================

' 1) Mátrix betöltés (forrásból)
Public Sub BiziMatrix_Build(Optional control As IRibbonControl)
    Import_Bizonyitvany_Matrix_Teljes
End Sub

' 2) Csak módosult (dirty) sorok alapján frissítés diakadat[p_bizonyitvany]-ba
Public Sub BiziMatrix_UpdateTarget_ChangedOnly(Optional control As IRibbonControl)
    BiziMatrix_UpdateTarget_ChangedOnly_Impl False
End Sub

' belsõ: silent módban is tudjon futni
Private Sub BiziMatrix_UpdateTarget_ChangedOnly_Impl(ByVal silent As Boolean)
    On Error GoTo EH

    Dim wbD As Workbook: Set wbD = ThisWorkbook

    Dim wsD As Worksheet: Set wsD = wbD.Worksheets("diakadat")
    Dim loD As ListObject: Set loD = wsD.ListObjects("diakadat")

    If loD.ListRows.Count = 0 Then
        If Not silent Then MsgBox "A diakadat tábla üres.", vbExclamation
        Exit Sub
    End If

    Dim colOktD As Long, colPBiziD As Long
    colOktD = GetLoCol(loD, "oktazon")
    colPBiziD = GetLoCol(loD, "p_bizonyitvany")
    If colOktD = 0 Or colPBiziD = 0 Then
        If Not silent Then MsgBox "Hiányzó oszlop a diakadat táblában: oktazon / p_bizonyitvany", vbCritical
        Exit Sub
    End If

    Dim wsM As Worksheet
    Set wsM = GetSheetIfExists(wbD, MATRIX_SHEET)
    If wsM Is Nothing Then
        If Not silent Then MsgBox "Nincs bizonyitvany_matrix lap. Elõbb töltsd be a mátrixot.", vbExclamation
        Exit Sub
    End If

    Dim lastRowM As Long
    lastRowM = wsM.Cells(wsM.rows.Count, MAT_COL_OKTAZON).End(xlUp).Row
    If lastRowM < 2 Then
        If Not silent Then MsgBox "A bizonyitvany_matrix üres.", vbExclamation
        Exit Sub
    End If

    Dim lastColM As Long
    lastColM = wsM.Cells(1, wsM.Columns.Count).End(xlToLeft).Column
    If lastColM < MAT_FIRST_SUBJ_COL Then
        If Not silent Then MsgBox "Nincs tantárgy oszlop a mátrixban.", vbExclamation
        Exit Sub
    End If
    If lastColM < DIRTY_COL Then lastColM = DIRTY_COL

    ' diakadat tömb + index
    Dim arrD As Variant: arrD = loD.DataBodyRange.value
    Dim idxD As Object: Set idxD = CreateObject("Scripting.Dictionary")
    idxD.CompareMode = 1

    Dim r As Long, ok As String
    For r = 1 To UBound(arrD, 1)
        ok = Trim$(CStr(arrD(r, colOktD)))
        If ok <> "" Then
            If Not idxD.Exists(ok) Then idxD.add ok, r
        End If
    Next r

    ' frissítés (csak dirty=1)
    Dim updated As Long, skipped As Long, missing As Long
    updated = 0: skipped = 0: missing = 0

    Dim rowM As Long
    For rowM = 2 To lastRowM
        If CLng(val(wsM.Cells(rowM, DIRTY_COL).value)) <> 1 Then GoTo NextM

        ok = Trim$(CStr(wsM.Cells(rowM, MAT_COL_OKTAZON).value))
        If ok = "" Then GoTo ClearDirty

        If Not idxD.Exists(ok) Then
            missing = missing + 1
            GoTo ClearDirty
        End If

        Dim sumP As Double: sumP = 0#
        Dim c As Long
        For c = MAT_FIRST_SUBJ_COL To lastColM
            If c <> DIRTY_COL Then
                sumP = sumP + GradeToNumberDbl(wsM.Cells(rowM, c).value)
            End If
        Next c

        sumP = Round2(sumP)

        Dim dRow As Long: dRow = CLng(idxD(ok))
        Dim cur As Variant: cur = arrD(dRow, colPBiziD)

        If Round2(NzDbl(cur)) <> sumP Then
            arrD(dRow, colPBiziD) = sumP
            updated = updated + 1
        Else
            skipped = skipped + 1
        End If

ClearDirty:
        wsM.Cells(rowM, DIRTY_COL).value = 0

NextM:
    Next rowM

    loD.DataBodyRange.value = arrD

    If Not silent Then
        MsgBox "Bizonyítvány frissítés kész." & vbCrLf & _
               "Változott (írtuk): " & updated & vbCrLf & _
               "Nem változott: " & skipped & vbCrLf & _
               "Célban nem talált oktazon: " & missing, vbInformation
    End If

    Exit Sub

EH:
    If Not silent Then MsgBox "Bizi frissítési hiba: " & Err.Description, vbCritical
End Sub

' 3) TELJES import: forrásból mátrix + azonnali p_bizonyitvany kitöltés + A–Z rendezés + pont újraszámolás
Public Sub Import_Bizonyitvany_Matrix_Teljes(Optional control As IRibbonControl)
    On Error GoTo EH

    Dim srcPath As String
    srcPath = PickExcelFile("Válaszd ki a BIZONYÍTVÁNYOS (forrás) Excel fájlt")
    If srcPath = "" Then Exit Sub

    Dim srcSheetName As String
    srcSheetName = InputBox("Forrás munkalap neve:", "Forrás munkalap", "Export")
    If Trim$(srcSheetName) = "" Then Exit Sub

    Dim keyHeader As String
    keyHeader = InputBox("Oktazon oszlop FEJLÉCE (a 2. fejléc sorban):", "Kulcs oszlop", "Oktatási azonosító")
    If Trim$(keyHeader) = "" Then Exit Sub

    Dim wbD As Workbook: Set wbD = ThisWorkbook
    Dim wsD As Worksheet: Set wsD = wbD.Worksheets("diakadat")
    Dim loD As ListObject: Set loD = wsD.ListObjects("diakadat")
    If loD.ListRows.Count = 0 Then
        MsgBox "A diakadat tábla üres.", vbExclamation
        Exit Sub
    End If

    Dim colOktD As Long, colNevD As Long
    colOktD = GetLoCol(loD, "oktazon")
    colNevD = GetLoCol(loD, "f_nev")
    If colNevD = 0 Then colNevD = GetLoCol(loD, "i_nev")

    If colOktD = 0 Then
        MsgBox "A diakadat táblában nincs 'oktazon' oszlop.", vbCritical
        Exit Sub
    End If

    ' Forrás megnyitás
    Application.ScreenUpdating = False

    Dim wbS As Workbook, wsS As Worksheet
    Set wbS = Workbooks.Open(srcPath, ReadOnly:=True)

    On Error Resume Next
    Set wsS = wbS.Worksheets(srcSheetName)
    On Error GoTo EH
    If wsS Is Nothing Then
        MsgBox "Nem találom a forrás munkalapot: " & srcSheetName, vbExclamation
        GoTo CleanExit
    End If

    ' Kulcs oszlop a forrásban (2. fejléc sor)
    Dim colKeyS As Long
    colKeyS = FindSubHeaderCol(wsS, SRC_HDR_YEAR, keyHeader)
    If colKeyS = 0 Then
        MsgBox "Nem találom a kulcs oszlopot a 2. fejléc sorban: " & keyHeader, vbCritical
        GoTo CleanExit
    End If

    ' Tantárgy oszlopok: csak 4. évf.
    Dim subjMap As Object
    Set subjMap = BuildSubjectToColMap_Year4(wsS, SRC_HDR_SUBJECT, SRC_HDR_YEAR)

    If subjMap.Count = 0 Then
        MsgBox "Nem találok 4. évf. oszlopokat. (Lehet: '4. évf.', 'IV. évf.', '4 évfolyam' stb.)", vbCritical
        GoTo CleanExit
    End If

    ' Forrás index: oktazon -> sor (elsõt vesszük)
    Dim srcIdx As Object: Set srcIdx = CreateObject("Scripting.Dictionary")
    srcIdx.CompareMode = 1

    Dim lastRowS As Long
    lastRowS = wsS.Cells(wsS.rows.Count, colKeyS).End(xlUp).Row

    Dim rr As Long, ok As String
    For rr = SRC_FIRST_DATA_ROW To lastRowS
        ok = Trim$(CStr(wsS.Cells(rr, colKeyS).value))
        If ok <> "" Then
            If Not srcIdx.Exists(ok) Then srcIdx.add ok, rr
        End If
    Next rr

    ' Mátrix lap elõkészítés
    Dim wsM As Worksheet
    Set wsM = EnsureSheet(wbD, MATRIX_SHEET)
    wsM.Cells.clear

    ' Tantárgyak "szép" listája ABC sorrendben (ékezetet megtartjuk!)
    Dim subjects() As String
    subjects = DictKeysToSortedArray_KeepOriginal(subjMap)

    ' Fejléc
    wsM.Cells(1, MAT_COL_OKTAZON).value = "oktazon"
    wsM.Cells(1, MAT_COL_NEV).value = "f_nev"

    Dim i As Long
    For i = LBound(subjects) To UBound(subjects)
        wsM.Cells(1, MAT_FIRST_SUBJ_COL + i).value = subjects(i)
    Next i

    wsM.Cells(1, DIRTY_COL).value = "dirty"

    ' diakadat tömb
    Dim arrD As Variant: arrD = loD.DataBodyRange.value

    ' Mátrix feltöltés a diakadat alapján (tehát biztosan ugyanaz a kör)
    Dim outRow As Long: outRow = 2
    Dim missInSource As Long: missInSource = 0

    Dim dRow As Long
    For dRow = 1 To UBound(arrD, 1)
        ok = Trim$(CStr(arrD(dRow, colOktD)))
        If ok = "" Then GoTo NextStudent

        wsM.Cells(outRow, MAT_COL_OKTAZON).value = ok
        If colNevD > 0 Then
            wsM.Cells(outRow, MAT_COL_NEV).value = CStr(arrD(dRow, colNevD))
        Else
            wsM.Cells(outRow, MAT_COL_NEV).value = ""
        End If

        If srcIdx.Exists(ok) Then
            Dim srcRow As Long: srcRow = CLng(srcIdx(ok))
            For i = LBound(subjects) To UBound(subjects)
                Dim colS As Long
                colS = subjMap(subjects(i))

                Dim v As Variant
                v = wsS.Cells(srcRow, colS).value

                ' 2 tizedes a mátrixban is
                If IsNumeric(v) And Not IsEmpty(v) Then
                    wsM.Cells(outRow, MAT_FIRST_SUBJ_COL + i).value = WorksheetFunction.Round(CDbl(v), 2)
                Else
                    wsM.Cells(outRow, MAT_FIRST_SUBJ_COL + i).value = v
                End If
            Next i
        Else
            missInSource = missInSource + 1
        End If

        ' Build után azonnal szeretnéd p_bizonyitvany-t tölteni -> jelöljük dirty-re
        wsM.Cells(outRow, DIRTY_COL).value = 1

        outRow = outRow + 1

NextStudent:
    Next dRow

    ' Formázás + rendezés A–Z (név)
    wsM.rows(1).Font.Bold = True
    wsM.Columns.AutoFit
    wsM.Columns(DIRTY_COL).Hidden = True

    ' 2 tizedes formátum a tantárgy mezõkre (adat sorok)
    Dim lastColFmt As Long
    lastColFmt = wsM.Cells(1, wsM.Columns.Count).End(xlToLeft).Column
    If lastColFmt >= MAT_FIRST_SUBJ_COL And outRow > 2 Then
        wsM.Range(wsM.Cells(2, MAT_FIRST_SUBJ_COL), wsM.Cells(outRow - 1, lastColFmt)).NumberFormat = "0.00"
    End If

    BiziMatrix_SortRowsByName wsM, outRow - 1

    ' Azonnali betöltés diakadat[p_bizonyitvany]-ba (dirty=1 sorok alapján)
    BiziMatrix_UpdateTarget_ChangedOnly_Impl True

    ' Pontok újraszámolása (ha nálad megvan)
    On Error Resume Next
    RecalcPontok_Automatikus
    On Error GoTo EH

    MsgBox "Bizonyítvány mátrix import kész." & vbCrLf & _
           "Mátrix sorok: " & (outRow - 2) & vbCrLf & _
           "Forrásban nem talált oktazon: " & missInSource, vbInformation

CleanExit:
    On Error Resume Next
    wbS.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Exit Sub

EH:
    MsgBox "Bizonyítvány import hiba: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' =========================================================
' Mátrix rendezés név szerint (A–Z)
' =========================================================
Private Sub BiziMatrix_SortRowsByName(ws As Worksheet, ByVal lastDataRow As Long)
    On Error GoTo EH
    If lastDataRow < 2 Then Exit Sub

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < DIRTY_COL Then lastCol = DIRTY_COL

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastDataRow, lastCol))

    rng.Sort Key1:=ws.Range("B2"), Order1:=xlAscending, Header:=xlYes, _
             MatchCase:=False, Orientation:=xlTopToBottom
    Exit Sub
EH:
End Sub

' =========================================================
' 4. évf oszlopok felderítése (2 fejléc sor, 1. sor tantárgy merge-elt is lehet)
' subjMap: original tantárgy név -> oszlopindex
' =========================================================
Private Function BuildSubjectToColMap_Year4(ws As Worksheet, headerRowSubject As Long, headerRowYear As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    Dim lastCol As Long
    lastCol = ws.Cells(headerRowYear, ws.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        If IsYear4(ws.Cells(headerRowYear, c).value) Then
            Dim subj As String
            subj = Trim$(SubjectHeaderText(ws, headerRowSubject, c))
            subj = Replace(subj, ChrW(160), " ")
            subj = Trim$(subj)

            If subj <> "" Then
                If Not d.Exists(subj) Then d.add subj, c
            End If
        End If
    Next c

    Set BuildSubjectToColMap_Year4 = d
End Function

Private Function IsYear4(ByVal v As Variant) As Boolean
    Dim s As String: s = NKey2(CStr(v))
    If (InStr(s, "4") > 0 Or InStr(s, "iv") > 0) And (InStr(s, "evf") > 0 Or InStr(s, "evfolyam") > 0) Then
        IsYear4 = True
    End If
End Function

Private Function SubjectHeaderText(ws As Worksheet, headerRowSubject As Long, col As Long) As String
    Dim cell As Range: Set cell = ws.Cells(headerRowSubject, col)
    If cell.MergeCells Then
        SubjectHeaderText = CStr(cell.MergeArea.Cells(1, 1).value)
    Else
        SubjectHeaderText = CStr(cell.value)
    End If
End Function

' 2. fejléc sorban keres (normalizált egyezéssel)
Private Function FindSubHeaderCol(ws As Worksheet, headerRowSub As Long, headerText As String) As Long
    Dim lastCol As Long: lastCol = ws.Cells(headerRowSub, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        If NKey2(CStr(ws.Cells(headerRowSub, c).value)) = NKey2(headerText) Then
            FindSubHeaderCol = c
            Exit Function
        End If
    Next c
    FindSubHeaderCol = 0
End Function

' =========================================================
' Jegy -> szám (szöveg is lehet)  [DOUBLE, 2 tizedesre]
' =========================================================
Private Function GradeToNumberDbl(ByVal v As Variant) As Double
    On Error GoTo EH

    If IsError(v) Or IsEmpty(v) Then Exit Function

    If IsNumeric(v) Then
        GradeToNumberDbl = Round2(CDbl(v))
        Exit Function
    End If

    Dim s As String
    s = NKey2(CStr(v))
    If s = "" Then Exit Function

    If InStr(s, "jeles") > 0 Or InStr(s, "kivalo") > 0 Then GradeToNumberDbl = 5#: Exit Function
    If InStr(s, "jo") > 0 Then GradeToNumberDbl = 4#: Exit Function
    If InStr(s, "kozepes") > 0 Then GradeToNumberDbl = 3#: Exit Function
    If InStr(s, "elegseges") > 0 Then GradeToNumberDbl = 2#: Exit Function
    If InStr(s, "elegtelen") > 0 Then GradeToNumberDbl = 1#: Exit Function

    ' Ha valami vegyes szöveg, megpróbáljuk a Val()-t, majd kerekítünk
    GradeToNumberDbl = Round2(CDbl(val(s)))
    Exit Function

EH:
End Function

Private Function NzDbl(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then NzDbl = CDbl(v)
End Function

Private Function Round2(ByVal x As Double) As Double
    ' Excel/VBA bankers rounding helyett: WorksheetFunction.Round általában a várt viselkedés
    Round2 = WorksheetFunction.Round(x, 2)
End Function

' =========================================================
' ListObject segédek / sheet segédek
' =========================================================
Private Function GetLoCol(lo As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    GetLoCol = lo.ListColumns(colName).Index
    On Error GoTo 0
End Function

Private Function EnsureSheet(wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = wb.Worksheets.add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

Private Function GetSheetIfExists(wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheetIfExists = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

' =========================================================
' Tantárgy lista ABC – az ORIGINAL (ékezetes) neveket rendezzük
' =========================================================
Private Function DictKeysToSortedArray_KeepOriginal(d As Object) As String()
    Dim keys As Variant: keys = d.keys
    Dim n As Long: n = d.Count
    Dim arr() As String
    ReDim arr(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1
        arr(i) = CStr(keys(i))
    Next i

    If n > 1 Then QuickSortStr arr, LBound(arr), UBound(arr)
    DictKeysToSortedArray_KeepOriginal = arr
End Function

Private Sub QuickSortStr(ByRef a() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, tmp As String
    i = first: j = last
    pivot = a((first + last) \ 2)

    Do While i <= j
        Do While NKey2(a(i)) < NKey2(pivot): i = i + 1: Loop
        Do While NKey2(a(j)) > NKey2(pivot): j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop

    If first < j Then QuickSortStr a, first, j
    If i < last Then QuickSortStr a, i, last
End Sub

' =========================================================
' Normalizáló (ékezet- és NBSP-tûrõ kereséshez)
' =========================================================
Private Function NKey2(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, ".", " ")
    t = Replace(t, vbTab, " ")
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop

    t = Replace(t, "á", "a")
    t = Replace(t, "é", "e")
    t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o")
    t = Replace(t, "ö", "o")
    t = Replace(t, "õ", "o")
    t = Replace(t, "ú", "u")
    t = Replace(t, "ü", "u")
    t = Replace(t, "û", "u")

    NKey2 = t
End Function

' =========================================================
' File picker
' =========================================================
Private Function PickExcelFile(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = title
        .Filters.clear
        .Filters.add "Excel fájlok", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            PickExcelFile = ""
        Else
            PickExcelFile = .SelectedItems(1)
        End If
    End With
End Function
