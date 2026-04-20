Attribute VB_Name = "modTablaGetters"
Option Explicit

Public Function GetIdopontTabla_V2() As ListObject
    On Error GoTo Fail
    Set GetIdopontTabla_V2 = ThisWorkbook.Worksheets("idopontok").ListObjects("tbl_idopontok")
    Exit Function
Fail:
    Set GetIdopontTabla_V2 = Nothing
End Function
