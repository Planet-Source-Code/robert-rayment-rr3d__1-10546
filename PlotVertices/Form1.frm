VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Plotting vertices from RR3D"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H80000007&
      Height          =   3450
      Left            =   285
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   226
      TabIndex        =   0
      Top             =   105
      Width           =   3450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim pp As POINTAPI

'Input Vertices -10 to 10 by -10 to 10
'ie 20 x20 = 400 points  sqr(400)=20

ReDim xs(0 To 0, 0 To 0), zs(0 To 0, 0 To 0)
Show
Open "Vertices.txt" For Input As #1
   Input #1, nj, ni
   ReDim xs(1 To ni, 1 To nj), zs(1 To ni, 1 To nj)
   For j = 1 To nj
   For i = 1 To ni
      Input #1, xs(i, j), zs(i, j)
   Next i
   Next j
Close

MousePointer = 11

'Draw X-lines
picDisplay.Cls 'Fast enough

picDisplay.ForeColor = QBColor(10)
For j = 1 To nj
   res& = MoveToEx(picDisplay.hdc, xs(1, j), zs(1, j), pp)
For i = 2 To ni
   res& = LineTo(picDisplay.hdc, xs(i, j), zs(i, j))
Next i
Next j

'Draw Y-lines
picDisplay.ForeColor = QBColor(11)
For i = 1 To ni
   res& = MoveToEx(picDisplay.hdc, xs(i, 1), zs(i, 1), pp)
For j = 2 To nj
   res& = LineTo(picDisplay.hdc, xs(i, j), zs(i, j))
Next j
Next i

MousePointer = 0
End Sub
