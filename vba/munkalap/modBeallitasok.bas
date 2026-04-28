Attribute VB_Name = "modBeallitasok"
Option Explicit

Private Const SETTINGS_SHEET_NAME As String = "beallitasok"
Private Const SETTINGS_HEADER_ROW As Long = 1
Private Const SETTINGS_FIRST_DATA_ROW As Long = 2
Private Const COL_KEY As Long = 1
Private Const COL_DESCRIPTION As Long = 2
Private Const COL_VALUE As Long = 3
Private Const COL_DEFAULT As Long = 4

Private Const KEY_BACKUP_ROOT As String = "backup_root"
Private Const KEY_NEVSOR_PDF_FOLDER As String = "nevsor_pdf_folder"
Private Const KEY_SZOBELI_TEMPLATE As String = "szobeli_meghivo_template"
Private Const KEY_PONTOZO_TEMPLATE As String = "pontozolap_template"
Private Const KEY_OSSZESITO_TEMPLATE As String = "osszesitolap_template"
Private Const KEY_PONTOZO_OUTPUT_ROOT As String = "pontozo_output_root"

Private Const DEFAULT_BACKUP_ROOT As String = "\\NS2\Felvételi\Backup\"
Private Const DEFAULT_NEVSOR_PDF_FOLDER As String = "\\NS2\Felvételi\Data\Nevsor"
Private Const DEFAULT_SZOBELI_TEMPLATE As String = "\\NS2\Felvételi\outlooksablon\szobeli-behivo.oft"
Private Const DEFAULT_PONTOZO_TEMPLATE As String = "\\NS2\Felvételi\Data\PontozolapTemplate.docx"
Private Const DEFAULT_OSSZESITO_TEMPLATE As String = "\\NS2\Felvételi\Data\OsszesitolapTemplate.docx"
Private Const DEFAULT_PONTOZO_OUTPUT_ROOT As String = "\\NS2\Felvételi\Data\Pontozo\"

Public Sub Ribbon_Beallitasok(control As IRibbonControl)
    Beallitasok_Menu control
End Sub

Public Sub Beallitasok_Megnyitasa(Optional control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = EnsureSettingsSheet(True)
    ws.Activate
End Sub

Public Sub Beallitasok_Menu(Optional control As IRibbonControl)
    Dim choice As String

    Do
        choice = InputBox(BuildSettingsMenuText(), "Beállítások", "1")
        choice = Trim$(choice)
        If choice = "" Then Exit Sub

        Select Case choice
            Case "1"
                Beallitasok_Megnyitasa
            Case "2"
                PickFolderForSetting KEY_BACKUP_ROOT, "Válaszd ki a backup mappát"
            Case "3"
                PickFolderForSetting KEY_NEVSOR_PDF_FOLDER, "Válaszd ki a névsor PDF mappát"
            Case "4"
                PickFileForSetting KEY_SZOBELI_TEMPLATE, "Válaszd ki a szóbeli meghívó Outlook sablont", "Outlook sablonok", "*.oft"
            Case "5"
                PickFileForSetting KEY_PONTOZO_TEMPLATE, "Válaszd ki a pontozólap Word sablont", "Word dokumentumok", "*.doc;*.docx;*.docm"
            Case "6"
                PickFileForSetting KEY_OSSZESITO_TEMPLATE, "Válaszd ki az összesítőlap Word sablont", "Word dokumentumok", "*.doc;*.docx;*.docm"
            Case "7"
                PickFolderForSetting KEY_PONTOZO_OUTPUT_ROOT, "Válaszd ki a pontozó kimeneti mappát"
            Case Else
                MsgBox "Érvénytelen választás.", vbExclamation
        End Select
    Loop
End Sub

Public Function GetConfiguredBackupRoot() As String
    GetConfiguredBackupRoot = EnsureTrailingSlashLocal(GetSettingValue(KEY_BACKUP_ROOT, DEFAULT_BACKUP_ROOT))
End Function

Public Function GetConfiguredNevsorPdfFolder() As String
    GetConfiguredNevsorPdfFolder = GetSettingValue(KEY_NEVSOR_PDF_FOLDER, DEFAULT_NEVSOR_PDF_FOLDER)
End Function

Public Function GetConfiguredSzobeliTemplatePath() As String
    GetConfiguredSzobeliTemplatePath = GetSettingValue(KEY_SZOBELI_TEMPLATE, DEFAULT_SZOBELI_TEMPLATE)
End Function

Public Function GetConfiguredPontozolapTemplatePath() As String
    GetConfiguredPontozolapTemplatePath = GetSettingValue(KEY_PONTOZO_TEMPLATE, DEFAULT_PONTOZO_TEMPLATE)
End Function

Public Function GetConfiguredOsszesitoTemplatePath() As String
    GetConfiguredOsszesitoTemplatePath = GetSettingValue(KEY_OSSZESITO_TEMPLATE, DEFAULT_OSSZESITO_TEMPLATE)
End Function

Public Function GetConfiguredPontozoOutputRoot() As String
    GetConfiguredPontozoOutputRoot = EnsureTrailingSlashLocal(GetSettingValue(KEY_PONTOZO_OUTPUT_ROOT, DEFAULT_PONTOZO_OUTPUT_ROOT))
End Function

Private Function BuildSettingsMenuText() As String
    BuildSettingsMenuText = _
        "Válassz beállítást:" & vbCrLf & _
        "  1. Beállítások lap megnyitása" & vbCrLf & _
        "  2. Backup mappa" & vbCrLf & _
        "  3. Névsor PDF mappa" & vbCrLf & _
        "  4. Szóbeli meghívó Outlook sablon" & vbCrLf & _
        "  5. Pontozólap Word sablon" & vbCrLf & _
        "  6. Összesítőlap Word sablon" & vbCrLf & _
        "  7. Pontozó kimeneti mappa" & vbCrLf & _
        vbCrLf & _
        "A beállítások a " & SETTINGS_SHEET_NAME & " lapon is szerkeszthetők." & vbCrLf & _
        "Kilépés: Mégse / üres"
End Function

Private Function EnsureSettingsSheet(Optional ByVal activateSheet As Boolean = False) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SETTINGS_SHEET_NAME
    End If

    PrepareSettingsSheet ws

    If activateSheet Then
        ws.Visible = xlSheetVisible
        ws.Activate
    End If

    Set EnsureSettingsSheet = ws
End Function

Private Sub PrepareSettingsSheet(ws As Worksheet)
    ws.Visible = xlSheetVisible

    ws.Cells(SETTINGS_HEADER_ROW, COL_KEY).Value = "kulcs"
    ws.Cells(SETTINGS_HEADER_ROW, COL_DESCRIPTION).Value = "leírás"
    ws.Cells(SETTINGS_HEADER_ROW, COL_VALUE).Value = "érték"
    ws.Cells(SETTINGS_HEADER_ROW, COL_DEFAULT).Value = "alapértelmezett"
    ws.Rows(SETTINGS_HEADER_ROW).Font.Bold = True

    EnsureSettingRow ws, KEY_BACKUP_ROOT, "Automatikus verziómentés gyökér mappa", DEFAULT_BACKUP_ROOT
    EnsureSettingRow ws, KEY_NEVSOR_PDF_FOLDER, "Névsor PDF export mappa", DEFAULT_NEVSOR_PDF_FOLDER
    EnsureSettingRow ws, KEY_SZOBELI_TEMPLATE, "Szóbeli meghívó Outlook sablon", DEFAULT_SZOBELI_TEMPLATE
    EnsureSettingRow ws, KEY_PONTOZO_TEMPLATE, "Pontozólap Word sablon", DEFAULT_PONTOZO_TEMPLATE
    EnsureSettingRow ws, KEY_OSSZESITO_TEMPLATE, "Összesítőlap Word sablon", DEFAULT_OSSZESITO_TEMPLATE
    EnsureSettingRow ws, KEY_PONTOZO_OUTPUT_ROOT, "Pontozó dokumentumok gyökér mappája", DEFAULT_PONTOZO_OUTPUT_ROOT

    ws.Range("F1").Value = "A C oszlopba írd a saját értéket. Üresen hagyva az alapértelmezett lesz használva."
    ws.Columns("A:F").AutoFit
End Sub

Private Sub EnsureSettingRow(ws As Worksheet, ByVal settingKey As String, ByVal description As String, ByVal defaultValue As String)
    Dim rowIndex As Long
    rowIndex = FindSettingRow(ws, settingKey)

    If rowIndex = 0 Then
        rowIndex = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row + 1
        If rowIndex < SETTINGS_FIRST_DATA_ROW Then rowIndex = SETTINGS_FIRST_DATA_ROW
        ws.Cells(rowIndex, COL_KEY).Value = settingKey
    End If

    ws.Cells(rowIndex, COL_DESCRIPTION).Value = description
    ws.Cells(rowIndex, COL_DEFAULT).Value = defaultValue
End Sub

Private Function FindSettingRow(ws As Worksheet, ByVal settingKey As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_KEY).End(xlUp).Row
    If lastRow < SETTINGS_FIRST_DATA_ROW Then Exit Function

    Dim r As Long
    For r = SETTINGS_FIRST_DATA_ROW To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, COL_KEY).Value & "")), settingKey, vbTextCompare) = 0 Then
            FindSettingRow = r
            Exit Function
        End If
    Next r
End Function

Private Function GetSettingValue(ByVal settingKey As String, ByVal fallbackValue As String) As String
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim currentValue As String

    Set ws = EnsureSettingsSheet(False)
    rowIndex = FindSettingRow(ws, settingKey)
    If rowIndex = 0 Then
        GetSettingValue = fallbackValue
        Exit Function
    End If

    currentValue = Trim$(CStr(ws.Cells(rowIndex, COL_VALUE).Value & ""))
    If currentValue = "" Then currentValue = Trim$(CStr(ws.Cells(rowIndex, COL_DEFAULT).Value & ""))
    If currentValue = "" Then currentValue = fallbackValue
    GetSettingValue = currentValue
End Function

Private Sub SaveSettingValue(ByVal settingKey As String, ByVal settingValue As String)
    Dim ws As Worksheet
    Dim rowIndex As Long

    Set ws = EnsureSettingsSheet(False)
    rowIndex = FindSettingRow(ws, settingKey)
    If rowIndex = 0 Then Exit Sub

    ws.Cells(rowIndex, COL_VALUE).Value = Trim$(settingValue)
End Sub

Private Sub PickFolderForSetting(ByVal settingKey As String, ByVal title As String)
    Dim fd As FileDialog
    Dim currentValue As String

    currentValue = GetSettingValue(settingKey, "")

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = title
        If currentValue <> "" Then .InitialFileName = EnsureTrailingSlashLocal(currentValue)
        If .Show <> -1 Then Exit Sub
        SaveSettingValue settingKey, .SelectedItems(1)
    End With

    MsgBox "Beállítás elmentve.", vbInformation
End Sub

Private Sub PickFileForSetting(ByVal settingKey As String, ByVal title As String, ByVal filterName As String, ByVal filterPattern As String)
    Dim fd As FileDialog
    Dim currentValue As String

    currentValue = GetSettingValue(settingKey, "")

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = title
        .Filters.Clear
        .Filters.Add filterName, filterPattern
        .AllowMultiSelect = False
        If currentValue <> "" Then .InitialFileName = currentValue
        If .Show <> -1 Then Exit Sub
        SaveSettingValue settingKey, .SelectedItems(1)
    End With

    MsgBox "Beállítás elmentve.", vbInformation
End Sub

Private Function EnsureTrailingSlashLocal(ByVal pathValue As String) As String
    pathValue = Trim$(pathValue)
    If pathValue = "" Then Exit Function

    If Right$(pathValue, 1) = "\" Or Right$(pathValue, 1) = "/" Then
        EnsureTrailingSlashLocal = pathValue
    Else
        EnsureTrailingSlashLocal = pathValue & "\"
    End If
End Function
