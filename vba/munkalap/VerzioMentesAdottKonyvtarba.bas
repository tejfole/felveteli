Attribute VB_Name = "VerzioMentesAdottKonyvtarba"
Option Explicit

' ============================================================
' Manuális verziómentés a "\Felvételi\Backup_Manual\" alá
'
' Cél mappa:
'   <Felvételi_root>\Backup_Manual\<BaseName>\
'
' Fájlnév idõbélyeggel (mindig egyedi):
'   <BaseName>_YYYYMMDD_HHMMSS.<ext>
'
' UNC-kompatibilis mappalétrehozás: FileSystemObject (FSO)
' ============================================================

Public Sub SaveVersionedCopy(Optional control As IRibbonControl)
    On Error GoTo EH

    Dim wb As Workbook: Set wb = ThisWorkbook

    If Len(wb.path) = 0 Then
        MsgBox "A fájlt elõször el kell menteni, hogy legyen elérési út!", vbCritical
        Exit Sub
    End If

    Dim felveteliRoot As String
    felveteliRoot = FindRootUpToFolder(wb.fullName, "Felvételi")
    If Len(felveteliRoot) = 0 Then
        MsgBox "A verziómentés csak a '\Felvételi\' könyvtár alól mûködik." & vbCrLf & _
               "Jelenlegi hely: " & wb.fullName, vbCritical
        Exit Sub
    End If

    Dim baseName As String, ext As String
    baseName = GetBaseName(wb.Name)
    ext = GetExtension(wb.Name)

    Dim targetFolder As String
    targetFolder = felveteliRoot & "Backup_Manual" & Application.PathSeparator & baseName

    EnsureFolder targetFolder ' FSO-s, UNC-safe

    Dim stamp As String
    stamp = Format(Now, "yyyymmdd_hhnnss")

    Dim fullPath As String
    fullPath = targetFolder & Application.PathSeparator & baseName & "_" & stamp & "." & ext

    wb.SaveCopyAs fullPath
    MsgBox "Manuális verzió mentve:" & vbCrLf & fullPath, vbInformation
    Exit Sub

EH:
    MsgBox "Manuális verziómentés hiba: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ----------------------------
' Root finder: visszaadja az útvonalat a megadott mappáig (végén \-sel)
' Példa:
'   input:  \\NS2\Felvételi\Data\valami\file.xlsm
'   folder: Felvételi
'   output: \\NS2\Felvételi\
' ----------------------------
Private Function FindRootUpToFolder(ByVal fullName As String, ByVal folderName As String) As String
    Dim p As String
    p = fullName

    Dim lastSep As Long
    lastSep = InStrRev(p, Application.PathSeparator)
    If lastSep = 0 Then Exit Function
    p = Left$(p, lastSep) ' ...\

    Dim token As String
    token = Application.PathSeparator & folderName & Application.PathSeparator

    Dim pos As Long
    pos = InStr(1, p, token, vbTextCompare)
    If pos = 0 Then Exit Function

    FindRootUpToFolder = Left$(p, pos + Len(token) - 1)
End Function

Private Sub EnsureFolder(ByVal fullPath As String)
    ' UNC és helyi útvonalon is biztosan létrehozza a teljes mappautat
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(fullPath) Then Exit Sub

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(fullPath)

    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
        EnsureFolder parentPath
    End If

    If Not fso.FolderExists(fullPath) Then
        fso.CreateFolder fullPath
    End If
End Sub

Private Function GetBaseName(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 1 Then
        GetBaseName = Left$(fileName, p - 1)
    Else
        GetBaseName = fileName
    End If
    GetBaseName = SanitizeFileName(GetBaseName)
End Function

Private Function GetExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 And p < Len(fileName) Then
        GetExtension = mid$(fileName, p + 1)
    Else
        GetExtension = "xlsm"
    End If
End Function

Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "_")
    Next i
    SanitizeFileName = s
End Function

