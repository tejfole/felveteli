Attribute VB_Name = "OktazonMailCsvbe"
Sub ExportCSVFromActiveSheetTable_UniqueOktazon(Optional control As IRibbonControl)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim folderPath As String
    Dim csvFile As String
    Dim logFile As String
    Dim outputFile As Integer
    Dim logOutput As Integer
    Dim i As Long
    Dim oktazonCol As Long
    Dim emailCol As Long
    Dim rowValues As String
    Dim emailValue As String
    Dim oktazonValue As String
    Dim dict As Object

    ' --- Szótár az egyediséghez ---
    Set dict = CreateObject("Scripting.Dictionary")

    ' --- Aktív munkalap
    Set ws = ActiveSheet

    ' --- Tábla beállítása
    If ws.ListObjects.Count = 0 Then
        MsgBox "? Nincs tábla az aktív munkalapon!", vbCritical
        Exit Sub
    End If

    Set tbl = ws.ListObjects(1) ' Az elsõ tábla

    ' --- Oktazon és Email oszlopok keresése ---
    For i = 1 To tbl.ListColumns.Count
        If LCase(tbl.ListColumns(i).Name) = "oktazon" Then oktazonCol = i
        If LCase(tbl.ListColumns(i).Name) = "email" Then emailCol = i
    Next i

    If oktazonCol = 0 Or emailCol = 0 Then
        MsgBox "Hiba: Nem található 'oktazon' vagy 'email' oszlop a táblában!", vbCritical
        Exit Sub
    End If

    ' --- Mappa kiválasztása ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Válaszd ki a CSV mentési mappát"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "Nincs mappa kiválasztva!", vbCritical
            Exit Sub
        End If
    End With

    ' --- Fájlok elérési útjai ---
    csvFile = folderPath & "\cimzettek.csv"
    logFile = folderPath & "\hibas_cimek_log.txt"

    outputFile = FreeFile
    Open csvFile For Output As #outputFile

    logOutput = FreeFile
    Open logFile For Output As #logOutput

    ' --- Fejléc a CSV-ben ---
    Print #outputFile, "fajlnev;email"

    ' --- Adatok feldolgozása ---
    For i = 1 To tbl.ListRows.Count
        oktazonValue = Trim(tbl.DataBodyRange(i, oktazonCol).value)
        emailValue = Trim(tbl.DataBodyRange(i, emailCol).value)

        ' Ha nincs oktazon (üres sor), kihagyjuk
        If oktazonValue = "" Then GoTo SkipNext

        ' Ha már volt ilyen oktazon, kihagyjuk
        If dict.Exists(oktazonValue) Then GoTo SkipNext

        ' Ha az email helyes
        If IsValidEmail(emailValue) Then
            rowValues = oktazonValue & ";" & emailValue
            Print #outputFile, rowValues
            dict.add oktazonValue, True
        Else
            ' Hibás email logolása
            Print #logOutput, "Hibás e-mail: " & emailValue & " (sor: " & i + 1 & ")"
        End If

SkipNext:
    Next i

    Close #outputFile
    Close #logOutput

    MsgBox "? CSV és hibás e-mail log sikeresen elkészült!", vbInformation

End Sub

' --- Email cím validáció ---
Function IsValidEmail(ByVal email As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    re.Pattern = "^[\w\.\-]+@([\w\-]+\.)+[\w\-]{2,4}$"
    re.IgnoreCase = True
    re.Global = False

    IsValidEmail = re.Test(email)
End Function

