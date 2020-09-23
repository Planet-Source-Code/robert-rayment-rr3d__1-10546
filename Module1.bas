Attribute VB_Name = "Module1"
'RR3D

Option Base 1              'Arrays starting at subscript 1

DefInt A-Q                 'Integers
DefSng R-Z                 'rst uvw xyz Real

'To shift cursor
Public Declare Sub SetCursorPos Lib "user32" _
(ByVal IX As Long, ByVal IY As Long)

Public Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long

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

Public Sub ReplacePI(inval$)
p1 = 1
Do
  pip = InStr(p1, inval$, "pi")
  If pip = 0 Then Exit Do
  ReplaceStr inval$, pip, pip + 1, "3.1415927"
  p1 = pip + 1
Loop
End Sub
Public Sub ReplaceLN(inval$)
p1 = 1
Do
  pip = InStr(p1, inval$, "ln")
  If pip = 0 Then Exit Do
  ReplaceStr inval$, pip, pip + 1, "log"
  p1 = pip + 1
Loop
End Sub

Public Sub ReplaceLOG(inval$)
p1 = 1
Do
  pip = InStr(p1, inval$, "log")
  If pip = 0 Then Exit Do
  'pip>log
  pob = InStr(pip, inval$, "(")
  FindMatchingClosingBracket inval$, pob, pout
  'pip>log(xxx)<pout
  stringinbr$ = Mid$(inval$, pip + 4, pout - (pip + 4))
  
  rep$ = "log(" + stringinbr$ + ")/log(10)"
  
  ReplaceStr inval$, pip, pout, rep$
  
  p1 = pip + Len(rep$)
  If p1 > Len(inval$) Then Exit Do
Loop
End Sub

Public Sub ReplaceASIN(inval$)
p1 = 1
Do
  pip = InStr(p1, inval$, "asin")
  If pip = 0 Then Exit Do
  'pip>log
  pob = InStr(pip, inval$, "(")
  FindMatchingClosingBracket inval$, pob, pout
  'pip>asin(xxx)<pout
  stringinbr$ = Mid$(inval$, pip + 4, (pout + 1) - (pip + 4))
  rep$ = "atn(" + stringinbr$ + "/sqr(1-" + stringinbr$ + "^2))"
  
  ReplaceStr inval$, pip, pout, rep$
  
  p1 = pip + Len(rep$)
  If p1 > Len(inval$) Then Exit Do
Loop
End Sub
Public Sub ReplaceACOS(inval$)
p1 = 1
Do
  pip = InStr(p1, inval$, "acos")
  If pip = 0 Then Exit Do
  'pip>log
  pob = InStr(pip, inval$, "(")
  FindMatchingClosingBracket inval$, pob, pout
  'pip>asin(xxx)<pout
  stringinbr$ = Mid$(inval$, pip + 4, (pout + 1) - (pip + 4))
  rep$ = "(pi/2)-atn(" + stringinbr$ + "/sqr(1-" + stringinbr$ + "^2))"
  
  ReplaceStr inval$, pip, pout, rep$
  
  p1 = pip + Len(rep$)
  If p1 > Len(inval$) Then Exit Do
Loop
ReplacePI inval$
End Sub
Public Sub ReplaceSINH(inval$)
p1 = 1
Do
  pip = InStr(p1, inval$, "sinh")
  If pip = 0 Then Exit Do
  'pip>log
  pob = InStr(pip, inval$, "(")
  FindMatchingClosingBracket inval$, pob, pout
  'pip>asin(xxx)<pout
  stringinbr$ = Mid$(inval$, pip + 4, (pout + 1) - (pip + 4))
  string2$ = "(-" + stringinbr$ + ")"
  rep$ = "(exp" + stringinbr$ + "-exp" + string2$ + ")/2"
  
  ReplaceStr inval$, pip, pout, rep$
  
  p1 = pip + Len(rep$)
  If p1 > Len(inval$) Then Exit Do
Loop

End Sub

Public Sub ReplaceCOSH(inval$)
p1 = 1
Do
  pip = InStr(p1, inval$, "cosh")
  If pip = 0 Then Exit Do
  'pip>log
  pob = InStr(pip, inval$, "(")
  FindMatchingClosingBracket inval$, pob, pout
  'pip>asin(xxx)<pout
  stringinbr$ = Mid$(inval$, pip + 4, (pout + 1) - (pip + 4))
  string2$ = "(-" + stringinbr$ + ")"
  rep$ = "(exp" + stringinbr$ + "+exp" + string2$ + ")/2"
  
  ReplaceStr inval$, pip, pout, rep$
  
  p1 = pip + Len(rep$)
  If p1 > Len(inval$) Then Exit Do
Loop

End Sub
Public Sub ReplaceTANH(inval$)
p1 = 1
Do
  pip = InStr(p1, inval$, "tanh")
  If pip = 0 Then Exit Do
  'pip>log
  pob = InStr(pip, inval$, "(")
  FindMatchingClosingBracket inval$, pob, pout
  'pip>asin(xxx)<pout
  stringinbr$ = Mid$(inval$, pip + 4, (pout + 1) - (pip + 4))
  rep$ = "(sinh" + stringinbr$ + ")/cosh" + stringinbr$
  ReplaceStr inval$, pip, pout, rep$
  
  p1 = pip + Len(rep$)
  If p1 > Len(inval$) Then Exit Do
Loop
ReplaceSINH inval$
ReplaceCOSH inval$
End Sub

Public Sub ReplaceStr(inval$, p1, p2, rep$)
'Replace sub-string p1->p2 in inval$ by rep$
ilenin = Len(inval$)
If p1 > p2 Or p1 > ilenin Or p2 > ilenin Then Exit Sub
If p1 = 1 Then
   inval$ = rep$ + Mid$(inval$, p2 + 1)
Else
   inval$ = Left$(inval$, p1 - 1) + rep$ + Mid$(inval$, p2 + 1)
End If
End Sub

Public Sub ReplaceXY(inval$, ByVal X, ByVal Y)
xstr$ = Trim$(Str$(X))
ystr$ = Trim$(Str$(Y))
p1 = 0
Do
 p1 = InStr(p1 + 1, inval$, "x")
 If p1 = 0 Then Exit Do
 'Check for exp
 If p1 = 1 Then
   ReplaceStr inval$, p1, p1, xstr$
 ElseIf p1 > 1 Then
   If Mid$(inval$, p1 - 1, 3) <> "exp" Then
      ReplaceStr inval$, p1, p1, xstr$
   End If
 End If
Loop

p1 = 0
Do
 p1 = InStr(p1 + 1, inval$, "y")
 If p1 = 0 Then Exit Do
 ReplaceStr inval$, p1, p1, ystr$
Loop

End Sub

Public Sub SqueezeSpaces(inval$)
'Squeeze out all spaces, trim & remove any leading +
inval$ = Trim$(inval$)
pp = InStr(1, inval$, "+")
If pp = 1 Then inval$ = Mid$(inval$, 2)
Do
  ps = InStr(1, inval$, " ")
  If ps = 0 Then Exit Do
  inval$ = Left(inval$, ps - 1) + Mid$(inval$, ps + 1)
Loop
End Sub

Public Function NumOccStr(inval$, c$)
'Find Number of occurences of character c$ in inval$
NumOccStr = 0
p1 = 1
Do
  p2 = InStr(p1, inval$, c$)
  If p2 <> 0 Then NumOccStr = NumOccStr + 1 Else Exit Function
  p1 = p2 + 1
Loop
End Function

Public Sub FindMatchingClosingBracket(inval$, pin, pout)
'pin is the position of an (
'pout is the position of the matching )
pob = InStr(pin + 1, inval$, "(")
If pob = 0 Then  '() no intermediate brackets
   pout = InStr(pin + 1, inval$, ")")
   Exit Sub
Else  '( @ pob before )
   nopbr = 0: nocbr = 0
   For k = pin To Len(inval$)
      c$ = Mid$(inval$, k, 1)
      If c$ = "(" Then nopbr = nopbr + 1
      If c$ = ")" Then nocbr = nocbr + 1
      If nopbr = nocbr Then
         pout = k
         Exit Sub
      End If
   Next k
End If
End Sub
