Attribute VB_Name = "TeremNevsorUIInstallOrRefresh"
Option Explicit

' ============================================================
' Teljes, önálló VBA modul — Névsor generálás + Ribbon callback +
' PDF mentés hálózatra + export log
'
' Ribbon:
' - customUI.xml onAction="Ribbon_TeremNevsor_Refresh"
' - customUI.xml onAction="Ribbon_TeremNevsor_Generate"
'
' Használat:
' 1) Illeszd be ezt a modult (Insert -> Module).
' 2) Futtasd egyszer: TeremNevsor_UI_InstallOrRefresh (vagy Ribbon Frissítés)
' 3) Válaszd: B1=Bizottság, B2=Nap, B3=Tanterem
' 4) Ribbon Generálás -> lista + nyomtatás/PDF kérdés
' ============================================================

' Lap/tábla nevek — igazítsd, ha szükséges
Private Const SHEET_DATA As String = "diakadat"
Private Const TABLE_DATA As String = "diakadat"

Private Const SHEET_ROSTER As String = "Névsor"
Private Const SHEET_ROOMLIST As String = "TanteremLista"

Private Const SHEET_SLOTS As String = "idopontok"
Private Const TABLE_SLOTS As String = "tbl_idopontok"

' UI cellák (függõleges elrendezés)
Private Const CELL_COMMITTEE As String = "B1"  ' Bizottság dropdown
Private Const CELL_DAY As String = "B2"        ' Nap dropdown
Private Const CELL_ROOM As String = "B3"       ' Tanterem dropdown

' Layout sorok
Private Const HEADER_TITLE_ROW As Long = 5     ' A5:C5 - címek (nyomtatás kezdete)
Private Const HEADER_VALUE_ROW As Long = 6     ' A6:C6 - értékek
Private Const LIST_HEADER_ROW As Long = 8      ' A8:C8 - lista fejléc
Private Const LIST_START_ROW As Long = 9       ' A9 - adatok kezdete

' Standard nyomtatási betûtípus
Private Const STD_FONT_NAME As String = "Calibri"
Private Const STD_FONT_SIZE As Long = 11

' Hálózati alapértelmezett mentési útvonal és log fájl neve
Private Const DEFAULT_PDF_FOLDER As String = "\\NS2\Felvételi\Data\Nevsor"
Private Const EXPORT_LOG_NAME As String = "export_log.csv"

' ============================================================
' Ribbon callbackek (EZEK KELLENEK A RIBBONHOZ)
' ============================================================
Public Sub Ribbon_TeremNevsor_Refresh(control As IRibbonControl)
    TeremNevsor_UI_InstallOrRefresh
End Sub

Public Sub Ribbon_TeremNevsor_Generate(control As IRibbonControl)
    TeremNevsor_Generalas
End Sub

' (Opcionális) ha a Ribbon XML mégis más névre hivatkozik:
Public Sub Ribbon_TeremNevsor_UI_InstallOrRefresh(control As IRibbonControl)
    TeremNevsor_UI_InstallOrRefresh
End Sub

Public Sub Ribbon_TeremNevsor_Generalas(control As IRibbonControl)
    TeremNevsor_Generalas
End Sub

' ============================================================
' UI telepítés / frissítés (futtasd egyszer)
' ============================================================
Public Sub TeremNevsor_UI_InstallOrRefresh()
    On Error GoTo ErrHandler

    Dim ws As Worksheet: Set ws = GetOrCreateSheet(SHEET_ROSTER, False)

    ' Tisztítjuk a felsõ területet (A1:C4)
    ws.Range("A1:C4").clear

    ' Label-ek (függõlegesen)
    ws.Range("A1").value = "Bizottság"
    ws.Range("A2").value = "Nap (yyyy. mm. dd)"
    ws.Range("A3").value = "Tanterem"
    ws.Range("A1:A3").Font.Bold = True

    ' Clear dropdown cells
    ws.Range(CELL_COMMITTEE & ":" & CELL_ROOM).ClearContents

    ' Apply dropdowns into B1,B2,B3
    ApplyCommitteeDropdown_FromDiakadat ws.Range(CELL_COMMITTEE)
    ApplyDayDropdown_FromIdopontokElseDiakadat ws.Range(CELL_DAY)
    ApplyRoomDropdown ws.Range(CELL_ROOM)

    ' Nyomtatási fejléc címei (A5:C5) és értékek sora (A6:C6)
    ws.Range("A" & HEADER_TITLE_ROW & ":C" & HEADER_TITLE_ROW).value = Array("Bizottság", "Nap", "Tanterem")
    ws.Range("A" & HEADER_VALUE_ROW & ":C" & HEADER_VALUE_ROW).ClearContents
    ws.rows(HEADER_TITLE_ROW).Font.Bold = True

    ' Lista fejléc (Sorszám | Név | Megjelent)
    ws.Range("A" & LIST_HEADER_ROW & ":C" & LIST_HEADER_ROW).value = Array("Sorszám", "Név", "Megjelent")
    ws.rows(LIST_HEADER_ROW).Font.Bold = True

    ws.Columns("A:C").AutoFit

    MsgBox "UI készen. Válaszd ki: B1 = Bizottság, B2 = Nap, B3 = Tanterem. Majd kattints Generál.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "UI telepítési hiba: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ============================================================
' Generálás (1 gomb)
' ============================================================
Public Sub TeremNevsor_Generalas()
    On Error GoTo ErrHandler

    Dim ws As Worksheet: Set ws = GetOrCreateSheet(SHEET_ROSTER, False)

    Dim committeeSel As String: committeeSel = NormalizeText(ws.Range(CELL_COMMITTEE).value)
    Dim daySel As String: daySel = NormalizeDayForUI(ws.Range(CELL_DAY).value)
    Dim roomSel As String: roomSel = Trim(CStr(ws.Range(CELL_ROOM).value & ""))

    If committeeSel = "" Then MsgBox "Válassz bizottságot (B1).", vbExclamation: Exit Sub
    If daySel = "" Then MsgBox "Válassz napot (B2).", vbExclamation: Exit Sub
    If roomSel = "" Then MsgBox "Válassz tantermet (B3).", vbExclamation: Exit Sub

    ' Töröljük az elõzõ lista régiót (A9:C...)
    ws.Range(ws.Cells(LIST_START_ROW, 1), ws.Cells(ws.rows.Count, 3)).ClearContents

    ' Fejléc értékek beírása A5:C6 (nyomtatás kezdete: 5. sor)
    ws.Range("A" & HEADER_TITLE_ROW & ":C" & HEADER_TITLE_ROW).value = Array("Bizottság", "Nap", "Tanterem")
    ws.Range("A" & HEADER_VALUE_ROW).value = committeeSel
    ws.Range("B" & HEADER_VALUE_ROW).value = daySel
    ws.Range("C" & HEADER_VALUE_ROW).value = roomSel
    With ws.Range("A" & HEADER_TITLE_ROW & ":C" & HEADER_VALUE_ROW)
        .Font.Name = STD_FONT_NAME
        .Font.Size = STD_FONT_SIZE
        .HorizontalAlignment = xlLeft
    End With

    ' Adatforrás beolvasása
    Dim wsData As Worksheet: Set wsData = ThisWorkbook.Worksheets(SHEET_DATA)
    Dim loData As ListObject: Set loData = wsData.ListObjects(TABLE_DATA)
    Dim dictCols As Object: Set dictCols = GetColumnIndexMap(loData)
    RequireCols dictCols, Array("f_nev", "bizottsag", "datum_nap")

    Dim outRow As Long: outRow = LIST_START_ROW
    Dim idx As Long: idx = 0

    Dim lr As ListRow
    For Each lr In loData.ListRows
        Dim nm As String: nm = Trim(CStr(lr.Range.Cells(1, dictCols("f_nev")).value & ""))
        If nm = "" Then GoTo NextRow

        Dim biz As String: biz = NormalizeText(lr.Range.Cells(1, dictCols("bizottsag")).value)
        Dim day2 As String: day2 = NormalizeDayForUI(lr.Range.Cells(1, dictCols("datum_nap")).value)
        If biz = "" Or day2 = "" Then GoTo NextRow

        If biz = committeeSel And day2 = daySel Then
            idx = idx + 1
            ws.Cells(outRow, 1).value = idx
            ws.Cells(outRow, 2).value = nm
            ws.Cells(outRow, 3).value = "" ' Megjelent
            outRow = outRow + 1
        End If
NextRow:
    Next lr

    ws.Columns("A:C").AutoFit

    ' Nyomtatási terület: 5. sortól a lista végéig
    Dim printLastRow As Long
    If outRow > LIST_START_ROW Then
        printLastRow = outRow - 1
    Else
        printLastRow = LIST_HEADER_ROW
    End If
    ws.PageSetup.PrintArea = ws.Range("A" & HEADER_TITLE_ROW & ":C" & printLastRow).Address

    With ws.Range("A" & HEADER_TITLE_ROW & ":C" & printLastRow).Font
        .Name = STD_FONT_NAME
        .Size = STD_FONT_SIZE
    End With

    If MsgBox("Nyomtassam most a névsort?", vbYesNo + vbQuestion, "Nyomtatás") = vbYes Then
        ws.PrintOut
    End If

    If MsgBox("Készítsek PDF-et a névsorból és mentsem ide: " & DEFAULT_PDF_FOLDER & " ?", vbYesNo + vbQuestion, "PDF export") = vbYes Then
        ExportRosterToPdf ws, printLastRow, committeeSel, daySel, roomSel
    End If

    Exit Sub

ErrHandler:
    MsgBox "Generálás hiba: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ============================================================
' PDF export + log
' ============================================================
Private Sub ExportRosterToPdf(ws As Worksheet, printLastRow As Long, committee As String, dayKey As String, room As String)
    On Error GoTo ErrHandler

    Dim basePath As String: basePath = DEFAULT_PDF_FOLDER

    If Not FolderExists(basePath) Then
        On Error Resume Next
        MkDirPath basePath
        On Error GoTo ErrHandler
    End If

    If Not FolderExists(basePath) Then
        If ThisWorkbook.path <> "" Then
            basePath = ThisWorkbook.path
        Else
            basePath = Environ$("USERPROFILE") & "\Desktop"
        End If
        MsgBox "A hálózati mappa nem elérhetõ. PDF ide kerül: " & basePath, vbExclamation, "Mentési hely"
        If Not FolderExists(basePath) Then MkDirPath basePath
    End If

    Dim fileName As String
    fileName = "Nevsor_" & SafeFileName(committee) & "_" & SafeFileName(dayKey) & "_" & SafeFileName(room) & ".pdf"

    Dim fullPath As String
    fullPath = basePath
    If Right(fullPath, 1) <> "\" Then fullPath = fullPath & "\"
    fullPath = fullPath & fileName

    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fullPath, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True

    AppendExportLog fullPath, committee, dayKey, room

    MsgBox "PDF elmentve: " & fullPath, vbInformation, "PDF mentés"
    Exit Sub

ErrHandler:
    MsgBox "PDF export hiba: " & Err.Number & " - " & Err.Description, vbExclamation, "Hiba"
End Sub

Private Sub AppendExportLog(fullPath As String, committee As String, dayKey As String, room As String)
    On Error GoTo ErrHandler

    Dim logFolder As String: logFolder = DEFAULT_PDF_FOLDER
    If Not FolderExists(logFolder) Then
        If ThisWorkbook.path <> "" Then
            logFolder = ThisWorkbook.path
        Else
            logFolder = Environ$("USERPROFILE") & "\Desktop"
        End If
    End If

    Dim logFull As String
    If Right(logFolder, 1) <> "\" Then logFull = logFolder & "\" & EXPORT_LOG_NAME Else logFull = logFolder & EXPORT_LOG_NAME

    Dim isNew As Boolean: isNew = (Dir(logFull) = "")

    Dim ff As Integer: ff = FreeFile
    Open logFull For Append As #ff
    If isNew Then Print #ff, "timestamp,user,fullpath,committee,day,room"

    Print #ff, """" & Format(Now, "yyyy-mm-dd HH:nn:ss") & """,""" & GetCurrentUserName() & """,""" & _
              Replace(fullPath, """", "'") & """,""" & Replace(committee, """", "'") & """,""" & Replace(dayKey, """", "'") & """,""" & Replace(room, """", "'") & """"
    Close #ff
    Exit Sub

ErrHandler:
    On Error Resume Next
    Close #ff
End Sub

' ============================================================
' Dropdown segédek
' ============================================================
Private Sub ApplyRoomDropdown(Target As Range)
    On Error GoTo ErrHandler
    Dim wsList As Worksheet
    On Error Resume Next
    Set wsList = ThisWorkbook.Worksheets(SHEET_ROOMLIST)
    On Error GoTo ErrHandler
    If wsList Is Nothing Then Exit Sub

    Dim last As Long: last = wsList.Cells(wsList.rows.Count, "A").End(xlUp).Row
    If last < 2 Then Exit Sub

    Dim listRange As Range: Set listRange = wsList.Range("A2:A" & last)
    Dim formula As String: formula = "='" & SHEET_ROOMLIST & "'!" & listRange.Address(True, True, xlA1)

    Target.Validation.Delete
    Target.Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=formula
    Target.Validation.IgnoreBlank = True
    Target.Validation.InCellDropdown = True
    Exit Sub
ErrHandler:
    Resume Next
End Sub

Private Sub ApplyCommitteeDropdown_FromDiakadat(Target As Range)
    Dim wsData As Worksheet: Set wsData = ThisWorkbook.Worksheets(SHEET_DATA)
    Dim loData As ListObject: Set loData = wsData.ListObjects(TABLE_DATA)
    Dim dictCols As Object: Set dictCols = GetColumnIndexMap(loData)
    RequireCols dictCols, Array("bizottsag")

    Dim arr As Variant
    arr = GetUniqueValuesFromColumn(loData, dictCols("bizottsag"))
    ApplyDropdownFromArray Target, arr
End Sub

Private Sub ApplyDayDropdown_FromIdopontokElseDiakadat(Target As Range)
    Dim arr As Variant
    arr = GetActiveDaysFromIdopontok()
    If IsEmpty(arr) Then arr = GetDaysFromDiakadat()
    ApplyDropdownFromArray Target, arr
End Sub

Private Sub ApplyDropdownFromArray(Target As Range, arr As Variant)
    On Error GoTo ErrHandler
    If IsEmpty(arr) Then Exit Sub

    Dim i As Long, s As String
    For i = LBound(arr) To UBound(arr)
        If s <> "" Then s = s & ","
        s = s & Replace(CStr(arr(i)), ",", " ")
    Next i

    Target.Validation.Delete
    Target.Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=s
    Target.Validation.IgnoreBlank = True
    Target.Validation.InCellDropdown = True
    Exit Sub
ErrHandler:
    Resume Next
End Sub

' ============================================================
' Idõpontok: tbl_idopontok(datum_nap, aktiv)
' ============================================================
Private Function GetActiveDaysFromIdopontok() As Variant
    On Error GoTo ErrHandler

    Dim wsSlots As Worksheet
    On Error Resume Next
    Set wsSlots = ThisWorkbook.Worksheets(SHEET_SLOTS)
    On Error GoTo ErrHandler
    If wsSlots Is Nothing Then GetActiveDaysFromIdopontok = Empty: Exit Function

    Dim loSlots As ListObject: Set loSlots = wsSlots.ListObjects(TABLE_SLOTS)
    Dim dictCols As Object: Set dictCols = GetColumnIndexMap(loSlots)
    If Not dictCols.Exists("datum_nap") Or Not dictCols.Exists("aktiv") Then
        GetActiveDaysFromIdopontok = Empty
        Exit Function
    End If

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lr As ListRow, v As Variant, dayKey As String
    For Each lr In loSlots.ListRows
        v = lr.Range.Cells(1, dictCols("aktiv")).value
        If IsActiveValue(v) Then
            dayKey = NormalizeDayForUI(lr.Range.Cells(1, dictCols("datum_nap")).value)
            If dayKey <> "" Then If Not dict.Exists(dayKey) Then dict.add dayKey, 1
        End If
    Next lr

    If dict.Count = 0 Then
        GetActiveDaysFromIdopontok = Empty
    Else
        GetActiveDaysFromIdopontok = DictionaryKeysToSortedArray(dict)
    End If
    Exit Function

ErrHandler:
    GetActiveDaysFromIdopontok = Empty
End Function

Private Function IsActiveValue(v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        IsActiveValue = CBool(v)
        Exit Function
    End If
    Dim s As String: s = LCase(Trim(CStr(v & "")))
    IsActiveValue = (s = "1" Or s = "true" Or s = "igen" Or s = "x" Or s = "yes")
End Function

Private Function GetDaysFromDiakadat() As Variant
    Dim wsData As Worksheet: Set wsData = ThisWorkbook.Worksheets(SHEET_DATA)
    Dim loData As ListObject: Set loData = wsData.ListObjects(TABLE_DATA)
    Dim dictCols As Object: Set dictCols = GetColumnIndexMap(loData)
    RequireCols dictCols, Array("datum_nap")
    GetDaysFromDiakadat = GetUniqueDaysFromColumnForUI(loData, dictCols("datum_nap"))
End Function

' ============================================================
' Normalizálás (yyyy. mm. dd)
' ============================================================
Private Function NormalizeText(v As Variant) As String
    NormalizeText = Trim(CStr(v & ""))
End Function

Private Function NormalizeDayForUI(v As Variant) As String
    If IsDate(v) Then
        NormalizeDayForUI = Format(CDate(v), "yyyy\. mm\. dd")
    Else
        Dim s As String: s = Trim(CStr(v & ""))
        If Len(s) >= 10 Then
            On Error Resume Next
            Dim tryDt As Date
            tryDt = CDate(Left$(s, 10))
            If Err.Number = 0 Then
                NormalizeDayForUI = Format(tryDt, "yyyy\. mm\. dd")
            Else
                NormalizeDayForUI = Left$(s, 10)
            End If
            On Error GoTo 0
        Else
            NormalizeDayForUI = s
        End If
    End If
End Function

Private Function GetUniqueDaysFromColumnForUI(lo As ListObject, colIndex As Long) As Variant
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lr As ListRow, dayKey As String
    For Each lr In lo.ListRows
        dayKey = NormalizeDayForUI(lr.Range.Cells(1, colIndex).value)
        If dayKey <> "" Then If Not dict.Exists(dayKey) Then dict.add dayKey, 1
    Next lr
    If dict.Count = 0 Then GetUniqueDaysFromColumnForUI = Empty Else GetUniqueDaysFromColumnForUI = DictionaryKeysToSortedArray(dict)
End Function

' ============================================================
' Unique helpers
' ============================================================
Private Function GetUniqueValuesFromColumn(lo As ListObject, colIndex As Long) As Variant
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lr As ListRow, v As String
    For Each lr In lo.ListRows
        v = Trim(CStr(lr.Range.Cells(1, colIndex).value & ""))
        If v <> "" Then If Not dict.Exists(v) Then dict.add v, 1
    Next lr
    If dict.Count = 0 Then GetUniqueValuesFromColumn = Empty Else GetUniqueValuesFromColumn = DictionaryKeysToSortedArray(dict)
End Function

Private Function DictionaryKeysToSortedArray(dict As Object) As Variant
    Dim arr() As String, i As Long
    ReDim arr(0 To dict.Count - 1)
    i = 0
    Dim k As Variant
    For Each k In dict.keys
        arr(i) = CStr(k): i = i + 1
    Next k
    QuickSortString arr, LBound(arr), UBound(arr)
    DictionaryKeysToSortedArray = arr
End Function

' ============================================================
' Utilities
' ============================================================
Public Function GetOrCreateSheet(sheetName As String, Optional clear As Boolean = False) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        On Error Resume Next
        ws.Name = sheetName
        On Error GoTo 0
    End If

    If clear Then ws.Cells.clear
    Set GetOrCreateSheet = ws
End Function

Private Function GetColumnIndexMap(lo As ListObject) As Object
    Dim dictCols As Object: Set dictCols = CreateObject("Scripting.Dictionary")
    Dim c As Long, hdr As String
    For c = 1 To lo.HeaderRowRange.Columns.Count
        hdr = Trim(LCase(CStr(lo.HeaderRowRange.Cells(1, c).value & "")))
        If hdr <> "" Then dictCols(hdr) = c
    Next c
    Set GetColumnIndexMap = dictCols
End Function

Private Sub RequireCols(dictCols As Object, required As Variant)
    Dim i As Long, k As String
    For i = LBound(required) To UBound(required)
        k = LCase(CStr(required(i)))
        If Not dictCols.Exists(k) Then Err.Raise vbObjectError + 2500, , "Hiányzó oszlop: " & k
    Next i
End Sub

Private Sub QuickSortString(arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As String, tmp As String
    i = first: j = last: pivot = arr((first + last) \ 2)
    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortString arr, first, j
    If i < last Then QuickSortString arr, i, last
End Sub

Private Function FolderExists(path As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(path, vbDirectory) <> "")
End Function

Private Sub MkDirPath(path As String)
    On Error GoTo ErrHandler
    Dim parts() As String
    Dim i As Long, cur As String

    If Len(path) >= 2 And Left(path, 2) = "\\" Then
        parts = Split(mid(path, 3), "\")
        If UBound(parts) >= 1 Then
            cur = "\\" & parts(0) & "\" & parts(1)
            For i = 2 To UBound(parts)
                cur = cur & "\" & parts(i)
                If Dir(cur, vbDirectory) = "" Then MkDir cur
            Next i
        End If
    Else
        parts = Split(path, "\")
        cur = parts(0)
        For i = 1 To UBound(parts)
            cur = cur & "\" & parts(i)
            If Dir(cur, vbDirectory) = "" Then MkDir cur
        Next i
    End If
    Exit Sub
ErrHandler:
    Err.clear
End Sub

Private Function GetCurrentUserName() As String
    On Error Resume Next
    GetCurrentUserName = Environ$("USERNAME")
    If GetCurrentUserName = "" Then GetCurrentUserName = Application.userName
End Function

Private Function SafeFileName(s As String) As String
    Dim bad As Variant: bad = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next i
    SafeFileName = Trim(s)
End Function

