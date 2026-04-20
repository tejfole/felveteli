Attribute VB_Name = "modBiziMatrixAuto"
Option Explicit

Public NextBiziCommit As Date
Public BiziCommitScheduled As Boolean

' Mátrix lap hívja: 1 másodpercen belül "összevonva" commitol
Public Sub ScheduleBiziMatrixCommit(Optional ByVal secondsDelay As Double = 1)
    On Error Resume Next

    ' ha már be van ütemezve, töröljük és újraütemezzük (debounce)
    If BiziCommitScheduled Then
        Application.OnTime EarliestTime:=NextBiziCommit, Procedure:="BiziMatrix_AutoCommit", Schedule:=False
    End If

    NextBiziCommit = Now + TimeSerial(0, 0, secondsDelay)
    BiziCommitScheduled = True
    Application.OnTime EarliestTime:=NextBiziCommit, Procedure:="BiziMatrix_AutoCommit", Schedule:=True
End Sub

' Tényleges commit: csak akkor érdemes, ha vannak dirty sorok
Public Sub BiziMatrix_AutoCommit()
    Dim prevEvents As Boolean, prevCalc As XlCalculation, prevScreen As Boolean
    On Error GoTo SafeExit

    BiziCommitScheduled = False

    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1) Bizonyítvány pont frissítés a mátrixból a diakadatba (csak dirty sorok)
    BiziMatrix_UpdateTarget_ChangedOnly

    ' 2) Teljes pontok újraszámolása
    RecalcPontok_Automatikus

SafeExit:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Sub
