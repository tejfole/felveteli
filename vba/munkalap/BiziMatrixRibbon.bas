Attribute VB_Name = "BiziMatrixRibbon"
Option Explicit

Public Sub Ribbon_BizonyitvanyMatrix(control As IRibbonControl)
    BiziMatrix_Build
End Sub

Public Sub Ribbon_BizonyitvanyFrissites(control As IRibbonControl)
    ' 1. Bizonyítvány pontok frissítése
    BiziMatrix_UpdateTarget_ChangedOnly
    
    ' 2. AUTOMATIKUS újraszámolás (p_bizonyitvany változott)
    RecalcPontok_Automatikus
End Sub

Public Sub Ribbon_BizonyitvanyTeljes(control As IRibbonControl)
    ' 1. Mátrix betöltés
    BiziMatrix_Build
    
    ' 2. p_bizonyitvany frissítés
    BiziMatrix_UpdateTarget_ChangedOnly
    
    ' 3. AUTOMATIKUS teljes pont újraszámolás
    RecalcPontok_Automatikus
End Sub
