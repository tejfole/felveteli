VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Munka5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit

    Dim hadEvents As Boolean
    hadEvents = Application.EnableEvents
    If Not hadEvents Then Exit Sub

    If Target Is Nothing Then Exit Sub
    If Target.Cells.CountLarge > 1 Then Exit Sub

    Dim lo As ListObject
    Set lo = Me.ListObjects("diakadat")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    ' -------------------------
    ' 1) AUTOMATIKUS PONTSZÁM ÚJRASZÁMOLÁS
    ' -------------------------
    Dim watchCols As Variant
    watchCols = Array("p_magyar", "p_matek", "p_bizonyitvany", "p_szovegalkotas", "p_kirako", "p_bemutatkozas")

    Dim i As Long
    For i = LBound(watchCols) To UBound(watchCols)
        Dim rngW As Range
        On Error Resume Next
        Set rngW = lo.ListColumns(CStr(watchCols(i))).DataBodyRange
        On Error GoTo SafeExit

        If Not rngW Is Nothing Then
            If Not Intersect(Target, rngW) Is Nothing Then
                Application.EnableEvents = False
                RecalcPontok_Automatikus
                GoTo SafeExit
            End If
        End If

        Set rngW = Nothing
    Next i

    ' -------------------------
    ' 2) IDÕPONT KIOSZTÁS (bizottsag -> datum_nap)
    ' -------------------------
    Dim rngBiz As Range
    On Error Resume Next
    Set rngBiz = lo.ListColumns("bizottsag").DataBodyRange
    On Error GoTo SafeExit
    If rngBiz Is Nothing Then GoTo SafeExit

    If Intersect(Target, rngBiz) Is Nothing Then GoTo SafeExit

    Dim v As String
    v = Trim$(CStr(Target.value))
    If v = "" Then GoTo SafeExit
    If Not IsNumeric(v) Then GoTo SafeExit

    Dim biz As Long
    biz = CLng(v)
    If biz < 1 Or biz > 10 Then
        MsgBox "A bizottság csak 1 és 10 közötti szám lehet.", vbExclamation
        GoTo SafeExit
    End If

    Dim rowIdx As Long
    rowIdx = Target.Row - lo.DataBodyRange.Row + 1
    If rowIdx < 1 Or rowIdx > lo.ListRows.Count Then GoTo SafeExit

    ' datum_nap oszlop (fallback: idopont_nap)
    Dim colDt As Long: colDt = 0
    On Error Resume Next
    colDt = lo.ListColumns("datum_nap").Index
    If colDt = 0 Then colDt = lo.ListColumns("idopont_nap").Index
    On Error GoTo SafeExit

    If colDt = 0 Then
        MsgBox "Hiányzó oszlop: datum_nap (vagy idopont_nap) a diakadat táblában.", vbExclamation
        GoTo SafeExit
    End If

    ' ha már van idõpont, ne írjuk felül
    If Trim$(CStr(lo.DataBodyRange.Cells(rowIdx, colDt).value)) <> "" Then GoTo SafeExit

    Application.EnableEvents = False
    AssignDatumNap_FromIdopontTabla lo, rowIdx, biz, 4

SafeExit:
    Application.EnableEvents = hadEvents
End Sub
