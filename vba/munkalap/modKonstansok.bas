Attribute VB_Name = "modKonstansok"
Option Explicit

Public Const KEVES_IRASBELI_KUSZOB As Long = 55

Public Function HonapNevekHU() As Variant
    HonapNevekHU = Array("", "január", "február", "március", "április", "május", "június", _
        "július", "augusztus", "szeptember", "október", "november", "december")
End Function

Public Function NapNevekHU() As Variant
    NapNevekHU = Array("vasárnap", "hétfõ", "kedd", "szerda", "csütörtök", "péntek", "szombat")
End Function

