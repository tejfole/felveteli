Attribute VB_Name = "VerziokAutoMentese"
Option Explicit

Public Const BACKUP_ROOT As String = "\\NS2\Felvételi\Backup\"
Public Const BACKUP_SUBFOLDER_PREFIX As String = "FELVETELI_"

Public Sub SaveVersionedCopy_Logged(Optional control As IRibbonControl)
    ' Almappa év alapján: elsõként a munkafüzet nevébõl, fallback: aktuális év
    Dim y As Long
    y = GetYearFromWorkbookNameOrZero(ThisWorkbook.Name)
    If y = 0 Then y = Year(Date)

    Dim subFolder As String
    subFolder = BACKUP_SUBFOLDER_PREFIX & CStr(y)

    SaveVersionedCopy_Core_Logged ThisWorkbook, BACKUP_ROOT, subFolder, True
End Sub

Public Sub SaveVersionedCopy_Core_Logged( _
    ByVal wb As Workbook, _
    ByVal backupRoot As String, _
    ByVal subFolder As String, _
    Optional ByVal statusBar As Boolean = True _
)
    Dim logPath As String
    logPath = Environ$("TEMP") & "\SaveVersionBackup.log"
    LogWrite logPath, "=== START SaveVersionedCopy_Core_Logged ==="
    On Error GoTo EH

    If wb Is Nothing Then
        LogWrite logPath, "Workbook objektum üres - kilépés."
        Exit Sub
    End If
    LogWrite logPath, "Workbook Name: " & wb.Name
    LogWrite logPath, "Workbook Path: " & wb.path

    If wb.path = "" Then
        LogWrite logPath, "Workbook nincs elmentve (wb.Path = ''). Kilépés."
        Exit Sub
    End If

    subFolder = Trim$(subFolder)
    If subFolder = "" Then subFolder = BACKUP_SUBFOLDER_PREFIX & CStr(Year(Date))

    Dim folderPath As String
    folderPath = EnsureTrailingSlash(backupRoot) & EnsureTrailingSlash(subFolder)
    LogWrite logPath, "Computed folderPath: " & folderPath

    If Dir$(folderPath, vbDirectory) = "" Then
        LogWrite logPath, "Mappa nem létezik. Próbálok létrehozni: " & folderPath
        If Not MkDirRecursive_Logged(folderPath, logPath) Then
            LogWrite logPath, "MkDirRecursive visszautasította a létrehozást. Kilépés."
            GoTo ExitSub
        Else
            LogWrite logPath, "Mappa létrehozva (vagy már létezett): " & folderPath
        End If
    Else
        LogWrite logPath, "Mappa már létezik."
    End If

    Dim prefix As String
    prefix = Trim$(subFolder)
    If prefix = "" Then prefix = "BACKUP"

    Dim userName As String
    userName = Environ$("USERNAME")
    If userName = "" Then userName = "user"

    Dim ts As String
    ts = Format$(Now, "yyyymmdd_hhnnss")

    Dim baseName As String
    baseName = GetBaseName(wb.Name)

    Dim ext As String
    ext = GetExtensionOrDefault(wb.Name, ".xlsm")

    Dim desired As String
    desired = folderPath & prefix & "_" & baseName & "_" & ts & "_" & userName & ext
    LogWrite logPath, "Desired path: " & desired

    Dim fullPath As String
    fullPath = UniquePath(desired)
    LogWrite logPath, "Unique path: " & fullPath

    If Not CanWriteTest(fullPath, logPath) Then
        LogWrite logPath, "Nem sikerült írni a célra (CanWriteTest false). Kilépés."
        GoTo ExitSub
    End If

    LogWrite logPath, "Call wb.SaveCopyAs: " & fullPath
    wb.SaveCopyAs fullPath
    LogWrite logPath, "SaveCopyAs sikeres: " & fullPath

    If statusBar Then
        StatusBarTemp "Verzió mentve: " & fullPath, 3
    End If

ExitSub:
    LogWrite logPath, "=== END SaveVersionedCopy_Core_Logged ==="
    Exit Sub

EH:
    LogWrite logPath, "HIBA: " & Err.Number & " - " & Err.Description
    If statusBar Then
        StatusBarTemp "Verzió mentés hiba: " & Err.Description, 5
    End If
    Resume ExitSub
End Sub

' ---------------- helperek ----------------
Private Sub LogWrite(ByVal logFile As String, ByVal msg As String)
    On Error Resume Next
    Dim f As Integer: f = FreeFile
    Open logFile For Append As #f
    Print #f, Format(Now, "yyyy-MM-dd HH:mm:ss") & " - " & msg
    Close #f
End Sub

Private Function GetYearFromWorkbookNameOrZero(ByVal fileName As String) As Long
    ' Kikeresi az elsõ elõforduló 4 számjegyû évet (2000-2099) a fájlnévbõl.
    ' Példa: "Felveteli_2026.xlsm" -> 2026
    Dim base As String
    base = GetBaseName(fileName)

    Dim i As Long
    For i = 1 To Len(base) - 3
        Dim token As String
        token = mid$(base, i, 4)

        If IsNumeric(token) Then
            Dim y As Long
            y = CLng(token)
            If y >= 2000 And y <= 2099 Then
                GetYearFromWorkbookNameOrZero = y
                Exit Function
            End If
        End If
    Next i

    GetYearFromWorkbookNameOrZero = 0
End Function

Private Function MkDirRecursive_Logged(ByVal folderPath As String, ByVal logPath As String) As Boolean
    On Error GoTo ErrHandler
    Dim p As String: p = EnsureTrailingSlash(folderPath)
    If Dir$(p, vbDirectory) <> "" Then
        MkDirRecursive_Logged = True
        Exit Function
    End If

    Dim noSlash As String: noSlash = Left$(p, Len(p) - 1)
    Dim pos As Long: pos = InStrRev(noSlash, "\")
    If pos <= 0 Then
        LogWrite logPath, "MkDirRecursive: nem található '\' a path-ban: " & noSlash
        MkDirRecursive_Logged = False
        Exit Function
    End If
    Dim parent As String: parent = Left$(noSlash, pos - 1)
    If parent <> "" Then
        If Dir$(parent & "\", vbDirectory) = "" Then
            LogWrite logPath, "MkDirRecursive: elõször a szülõ mappa létrehozása: " & parent
            If Not MkDirRecursive_Logged(parent & "\", logPath) Then
                MkDirRecursive_Logged = False
                Exit Function
            End If
        End If
    End If

    MkDir noSlash
    LogWrite logPath, "MkDirRecursive: létrehozva: " & noSlash
    MkDirRecursive_Logged = True
    Exit Function

ErrHandler:
    LogWrite logPath, "MkDirRecursive HIBA: " & Err.Number & " - " & Err.Description
    MkDirRecursive_Logged = False
End Function

Private Function CanWriteTest(ByVal targetPath As String, ByVal logPath As String) As Boolean
    On Error GoTo ErrHandler
    Dim testPath As String
    Dim p As Long: p = InStrRev(targetPath, "\")
    If p = 0 Then GoTo ErrHandler
    testPath = Left$(targetPath, p - 1) & "\.__vbatest_" & Format(Now, "yyyymmdd_hhnnss") & ".tmp"
    Dim f As Integer: f = FreeFile
    Open testPath For Binary Access Write As #f
    Put #f, , "x"
    Close #f
    Kill testPath
    LogWrite logPath, "CanWriteTest OK in folder: " & Left$(targetPath, p - 1)
    CanWriteTest = True
    Exit Function
ErrHandler:
    LogWrite logPath, "CanWriteTest HIBA: " & Err.Number & " - " & Err.Description
    CanWriteTest = False
End Function

Private Sub StatusBarTemp(ByVal msg As String, ByVal seconds As Long)
    On Error Resume Next
    Dim prev As Variant
    prev = Application.statusBar
    Application.statusBar = msg
    Dim t As Date
    t = Now + TimeSerial(0, 0, seconds)
    Do While Now < t
        DoEvents
    Loop
    Application.statusBar = prev
End Sub

Private Function EnsureTrailingSlash(ByVal p As String) As String
    p = Trim$(p)
    If p = "" Then
        EnsureTrailingSlash = ""
    ElseIf Right$(p, 1) = "\" Then
        EnsureTrailingSlash = p
    Else
        EnsureTrailingSlash = p & "\"
    End If
End Function

Private Function GetBaseName(ByVal fileName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetBaseName = Left$(fileName, dotPos - 1)
    Else
        GetBaseName = fileName
    End If
End Function

Private Function GetExtensionOrDefault(ByVal fileName As String, ByVal defaultExt As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetExtensionOrDefault = mid$(fileName, dotPos)
    Else
        GetExtensionOrDefault = defaultExt
    End If
End Function

Private Function UniquePath(ByVal fullPath As String) As String
    If Dir$(fullPath) = "" Then
        UniquePath = fullPath
        Exit Function
    End If
    Dim p As Long
    p = InStrRev(fullPath, ".")
    Dim base As String, ext As String
    If p > 0 Then
        base = Left$(fullPath, p - 1)
        ext = mid$(fullPath, p)
    Else
        base = fullPath
        ext = ""
    End If
    Dim i As Long
    For i = 2 To 999
        Dim candidate As String
        candidate = base & "_" & Format$(i, "00") & ext
        If Dir$(candidate) = "" Then
            UniquePath = candidate
            Exit Function
        End If
    Next i
    UniquePath = base & "_" & Format$(CLng(Timer), "000000") & ext
End Function

