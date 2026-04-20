Attribute VB_Name = "BiziPontokRecalcAutomatikus"
Public Sub RecalcPontok_Automatikus()
    Dim prevEvents As Boolean
    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean
    
    On Error GoTo Vege
    
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call SzamoljPontokatTombosen
    
Vege:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Sub
