Attribute VB_Name = "modPBizonyitvanyRibbon"
Option Explicit

' Ribbon gombhoz: ÚJRATÖLTÉS (mátrixból) + p_bizonyitvany megjelenítés tizedesei
Public Sub PBizonyitvany_UjratoltesEsTizedesBeallitas(Optional control As IRibbonControl)
    On Error GoTo EH

    Dim decStr As String
    decStr = InputBox( _
        "Hány tizedesjeggyel jelenjen meg a diakadat[p_bizonyitvany]?" & vbCrLf & _
        "Írj 0–6 közötti számot. (pl. 2)", _
        "p_bizonyitvany – újratöltés + formázás", _
        "2")

    If Len(Trim$(decStr)) = 0 Then Exit Sub

    Dim decimals As Long
    decimals = CLng(val(decStr))
    If decimals < 0 Then decimals = 0
    If decimals > 6 Then decimals = 6

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 1) Mátrixból újratöltés: legyen teljes kör (ne csak dirty)!
    BiziMatrix_MarkAllDirty
    BiziMatrix_UpdateTarget_ChangedOnly

    ' 2) Megjelenítés tizedeseinek beállítása
    PBizonyitvany_ApplyFormat decimals

    MsgBox "Kész." & vbCrLf & _
           "p_bizonyitvany újratöltve és " & decimals & " tizedesre formázva.", vbInformation

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

EH:
    MsgBox "Hiba: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' Csak format beállítás – belsõ segéd
Private Sub PBizonyitvany_ApplyFormat(ByVal decimals As Long)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet: Set ws = wb.Worksheets("diakadat")
    Dim lo As ListObject: Set lo = ws.ListObjects("diakadat")

    Dim colIdx As Long
    On Error Resume Next
    colIdx = lo.ListColumns("p_bizonyitvany").Index
    On Error GoTo 0
    If colIdx = 0 Then Err.Raise vbObjectError + 901, "PBizonyitvany_ApplyFormat", "Nincs p_bizonyitvany oszlop a diakadat táblában."

    If lo.ListRows.Count = 0 Then Exit Sub

    Dim fmt As String
    fmt = "0"
    If decimals > 0 Then fmt = fmt & "." & String$(decimals, "0")

    lo.ListColumns(colIdx).DataBodyRange.NumberFormat = fmt
End Sub

' Minden mátrix sort dirty=1-re állít (hogy teljes újratöltés legyen)
Private Sub BiziMatrix_MarkAllDirty()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsM As Worksheet

    On Error Resume Next
    Set wsM = wb.Worksheets("bizonyitvany_matrix")
    On Error GoTo 0
    If wsM Is Nothing Then Err.Raise vbObjectError + 902, "BiziMatrix_MarkAllDirty", "Nincs bizonyitvany_matrix lap."

    Dim lastRowM As Long
    lastRowM = wsM.Cells(wsM.rows.Count, 1).End(xlUp).Row
    If lastRowM < 2 Then Exit Sub

    wsM.Range(wsM.Cells(2, 26), wsM.Cells(lastRowM, 26)).value = 1
End Sub

