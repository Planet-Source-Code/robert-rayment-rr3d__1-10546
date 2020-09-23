Attribute VB_Name = "Module1"
'LineTo and MoveToEx are much faster the VB equivalents
Declare Function LineTo Lib "gdi32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Declare Function MoveToEx Lib "gdi32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, _
lpPoint As POINTAPI) As Long

'Use:
'Dim pp As POINTAPI
'res& = LineTo(Object.hdc, x, y)
'res& = MoveToEx(Object.hdc, x, y, pp)


