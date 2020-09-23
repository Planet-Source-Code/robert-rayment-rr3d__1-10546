VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   404
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11355
      Top             =   6885
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSaveVertices 
      Caption         =   "Save vertices"
      Height          =   255
      Left            =   660
      TabIndex        =   49
      Top             =   7860
      Width           =   1275
   End
   Begin VB.CommandButton cmdReadMe 
      Caption         =   "ReadMe.txt"
      Height          =   285
      Left            =   3645
      TabIndex        =   48
      Top             =   7875
      Width           =   1110
   End
   Begin VB.CheckBox chkYROT 
      BackColor       =   &H0000C000&
      Caption         =   "Y-rotation only"
      Height          =   255
      Left            =   600
      TabIndex        =   45
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox chkXROT 
      BackColor       =   &H00C0C000&
      Caption         =   "X-rotation only"
      Height          =   255
      Left            =   600
      TabIndex        =   44
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame frmPlotIntv 
      Caption         =   "Plot Interval"
      Height          =   615
      Left            =   720
      TabIndex        =   41
      Top             =   5100
      Width           =   1095
      Begin VB.TextBox txtPlotIntv 
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   225
         Width           =   615
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   255
         Left            =   840
         TabIndex        =   43
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   450
         _Version        =   327681
         Value           =   10
         BuddyControl    =   "txtPlotIntv"
         BuddyDispid     =   196614
         OrigLeft        =   780
         OrigTop         =   240
         OrigRight       =   975
         OrigBottom      =   495
         Max             =   28
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame frmXYXRes 
      Caption         =   "XYZ values"
      Height          =   1935
      Left            =   300
      TabIndex        =   34
      Top             =   5820
      Width           =   1875
      Begin VB.PictureBox picXYZ 
         Height          =   315
         Index           =   2
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   39
         Top             =   1380
         Width           =   1395
      End
      Begin VB.PictureBox picXYZ 
         Height          =   315
         Index           =   1
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   37
         Top             =   900
         Width           =   1395
      End
      Begin VB.PictureBox picXYZ 
         Height          =   315
         Index           =   0
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   35
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "Z"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "Y"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   900
         Width           =   195
      End
      Begin VB.Label Label7 
         Caption         =   "X"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   540
         Width           =   255
      End
   End
   Begin VB.Frame frmPERSPEC 
      Caption         =   "Perspec Dis"
      Height          =   615
      Left            =   720
      TabIndex        =   31
      Top             =   4380
      Width           =   1095
      Begin VB.TextBox txtPERSPEC 
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   240
         Width           =   615
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   840
         TabIndex        =   33
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   450
         _Version        =   327681
         Value           =   10
         BuddyControl    =   "txtPERSPEC"
         BuddyDispid     =   196619
         OrigLeft        =   780
         OrigTop         =   240
         OrigRight       =   975
         OrigBottom      =   495
         Increment       =   10
         Max             =   400
         Min             =   10
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Perspective"
      Height          =   195
      Left            =   2880
      TabIndex        =   30
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame frmZLIM 
      Caption         =   "Z +/- Limits"
      Height          =   615
      Left            =   720
      TabIndex        =   26
      Top             =   3660
      Width           =   1095
      Begin VB.TextBox txtZLIM 
         Height          =   255
         Left            =   180
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   240
         Width           =   555
      End
      Begin ComCtl2.UpDown UDZHILO 
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   240
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   450
         _Version        =   327681
         Value           =   10
         BuddyControl    =   "txtZLIM"
         BuddyDispid     =   196622
         OrigLeft        =   780
         OrigTop         =   240
         OrigRight       =   975
         OrigBottom      =   495
         Increment       =   10
         Max             =   400
         Min             =   10
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame frmXLO 
      Caption         =   "X-Low"
      Height          =   615
      Left            =   720
      TabIndex        =   20
      Top             =   2940
      Width           =   1095
      Begin ComCtl2.UpDown UDXLO 
         Height          =   315
         Left            =   720
         TabIndex        =   25
         Top             =   180
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   327681
         BuddyControl    =   "txtXLO"
         BuddyDispid     =   196624
         OrigLeft        =   780
         OrigTop         =   240
         OrigRight       =   975
         OrigBottom      =   495
         Max             =   12
         Min             =   -12
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtXLO 
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame frmXHI 
      Caption         =   "X-High"
      Height          =   615
      Left            =   720
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
      Begin ComCtl2.UpDown UDXHI 
         Height          =   315
         Left            =   720
         TabIndex        =   24
         Top             =   180
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   327681
         BuddyControl    =   "txtXHI"
         BuddyDispid     =   196626
         OrigLeft        =   780
         OrigTop         =   225
         OrigRight       =   975
         OrigBottom      =   495
         Max             =   12
         Min             =   -12
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtXHI 
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame frmYLO 
      Caption         =   "Y-Low"
      Height          =   615
      Left            =   720
      TabIndex        =   16
      Top             =   1500
      Width           =   1095
      Begin ComCtl2.UpDown UDYLO 
         Height          =   315
         Left            =   720
         TabIndex        =   23
         Top             =   180
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   327681
         BuddyControl    =   "txtYLO"
         BuddyDispid     =   196628
         OrigLeft        =   780
         OrigTop         =   195
         OrigRight       =   975
         OrigBottom      =   480
         Max             =   12
         Min             =   -12
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtYLO 
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame frmYHI 
      Caption         =   "Y-High"
      Height          =   615
      Left            =   720
      TabIndex        =   14
      Top             =   900
      Width           =   1095
      Begin ComCtl2.UpDown UDYHI 
         Height          =   315
         Left            =   720
         TabIndex        =   22
         Top             =   180
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   327681
         BuddyControl    =   "txtYHI"
         BuddyDispid     =   196630
         OrigLeft        =   765
         OrigTop         =   195
         OrigRight       =   960
         OrigBottom      =   510
         Max             =   12
         Min             =   -12
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtYHI 
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdCLEAR 
      Caption         =   "Clear item"
      Height          =   255
      Left            =   9180
      TabIndex        =   13
      Top             =   4680
      Width           =   915
   End
   Begin VB.CommandButton cmdLOAD 
      Caption         =   "Load"
      Height          =   255
      Left            =   8520
      TabIndex        =   12
      Top             =   4680
      Width           =   555
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "Save"
      Height          =   255
      Left            =   7860
      TabIndex        =   11
      Top             =   4680
      Width           =   555
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "Add"
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   4680
      Width           =   555
   End
   Begin MSScriptControlCtl.ScriptControl SCI 
      Left            =   11325
      Top             =   7620
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.CommandButton cmdEvaluate 
      Caption         =   "&Evaluate formula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4140
      TabIndex        =   4
      Top             =   4080
      Width           =   1695
   End
   Begin VB.ListBox lstFormulae 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   6300
      TabIndex        =   3
      Top             =   4965
      Width           =   4935
   End
   Begin VB.TextBox txtFormula 
      Height          =   315
      Left            =   5940
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4200
      Width           =   5535
   End
   Begin VB.PictureBox picFunctions 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   5940
      ScaleHeight     =   3615
      ScaleWidth      =   5415
      TabIndex        =   1
      Top             =   360
      Width           =   5475
   End
   Begin VB.PictureBox picDisplay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   3435
      Left            =   2460
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   360
      Width           =   3435
   End
   Begin VB.Label labTitle 
      AutoSize        =   -1  'True
      Caption         =   "LabTitle"
      Height          =   195
      Left            =   2520
      TabIndex        =   47
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label8 
      Caption         =   "Formulae.txt"
      Height          =   255
      Left            =   10320
      TabIndex        =   46
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "VALID ENTRIES"
      Height          =   315
      Left            =   7860
      TabIndex        =   29
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "LabInstructions"
      Height          =   3105
      Left            =   2580
      TabIndex        =   9
      Top             =   4680
      Width           =   3555
   End
   Begin VB.Label Label5 
      Caption         =   "Y, Z"
      Height          =   255
      Left            =   2100
      TabIndex        =   8
      Top             =   4200
      Width           =   315
   End
   Begin VB.Label Label4 
      Caption         =   "X"
      Height          =   195
      Left            =   3600
      TabIndex        =   7
      Top             =   3840
      Width           =   195
   End
   Begin VB.Label Label3 
      Caption         =   "Z, Y"
      Height          =   195
      Left            =   2100
      TabIndex        =   6
      Top             =   1200
      Width           =   375
   End
   Begin VB.Line Line3 
      X1              =   159
      X2              =   136
      Y1              =   256
      Y2              =   278
   End
   Begin VB.Line Line2 
      X1              =   152
      X2              =   152
      Y1              =   142
      Y2              =   100
   End
   Begin VB.Line Line1 
      X1              =   176
      X2              =   237
      Y1              =   264
      Y2              =   264
   End
   Begin VB.Label Label2 
      Caption         =   "Formulae"
      Height          =   255
      Left            =   6300
      TabIndex        =   5
      Top             =   4680
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'RR3D  by Robert Rayment  29/7/00  3/8/00

Option Base 1              'Arrays starting at subscript 1 by default

DefInt A-Q                 'Integers
DefSng R-Z                 'rst uvw xyz Real

Dim YH(), YL(), XH(), XL() 'Low & High x & y limits as read from file
Dim ZL()                   ''+/- z-limit as read from file
Dim PD()                   'Perspective distance as read from file

Dim svrx(), svry(), svrz() 'Original 3D function points
Dim rx(), ry(), rz()       'Rotated 3D function points
Dim xs(), ys(), zs()       'Screen points

Dim grxhi, gryhi         'High x & y grid limits

Dim scx, scy               'Convert grid vals to x,y vals
Dim xxlo, xxhi, yylo, yyhi 'Low & High x & y limits
Dim zlim                   '+/- z-limit

Dim xang, zang             'Rotation angle about x & z axes

Dim zpL, zpH               'Used to scale dowm plots
Dim xoff, yoff, zoff, scfx, scfy, scfz 'Function -> picDisplay Scale factors
Dim xmin, zmin             'Used with scale factors
Dim CheckPerspective       'Check1.Value

Dim xe, ze                 'Eye point (at y=0)
Dim ye                     'Perspective distance
Dim XMouse, YMouse         'Saved mouse coords at right-click MouseDown

Dim pp As POINTAPI         'Used in GetCursorPos no used in MoveToEx
Dim topoffset              'Height of form header, to adjust SetCursorPos

Dim culgreen&              'Plot X colors
Dim culcyan&               'Plot Y colors

Const pi = 3.1415927

Private Sub Check1_Click()
CheckPerspective = Check1.Value
End Sub

Private Sub cmdADD_Click()
A$ = txtFormula.Text
If A$ = "" Then Exit Sub
lstFormulae.AddItem A$
'Index = lstFormulae.ListCount
I = UBound(YH) + 1
ReDim Preserve YH(0 To I), YL(0 To I), XH(0 To I), XL(0 To I), ZL(0 To I), PD(0 To I)
YH(I - 1) = yyhi
YL(I - 1) = yylo
XH(I - 1) = xxhi
XL(I - 1) = xxlo
ZL(I - 1) = zlim
PD(I - 1) = ye
End Sub

Private Sub cmdSave_Click()
k = lstFormulae.ListCount
If k <= 0 Then Exit Sub
'Formula list
Open "Formulae.txt" For Output As #1
For I = 0 To k - 1
   A$ = lstFormulae.List(I)
   Print #1, A$; ","; YH(I); ","; YL(I); ","; XH(I); ","; XL(I);
   Print #1, ","; ZL(I); ","; PD(I)
Next I
Close
End Sub

Private Sub cmdCLEAR_Click()
Index = lstFormulae.ListIndex
If Index < 0 Then Exit Sub
lstFormulae.RemoveItem Index
'Move down YH YH XH XL PD
For I = Index To UBound(YL) - 1
   YH(I) = YH(I + 1)
   YL(I) = YL(I + 1)
   XH(I) = XH(I + 1)
   XL(I) = XL(I + 1)
   ZL(I) = ZL(I + 1)
   PD(I) = PD(I + 1)
Next I
I = UBound(YL) - 1
ReDim Preserve YH(0 To I), YL(0 To I), XH(0 To I), XL(0 To I), ZL(0 To I), PD(0 To I)

End Sub

Private Sub cmdLOAD_Click()
lstFormulae.Clear
ReDim YH(0 To 0), YL(0 To 0), XH(0 To 0), XL(0 To 0), ZL(0 To 0), PD(0 To 0)
'Formula list
Open "Formulae.txt" For Input As #1
I = 0
Do
   Input #1, A$, YH(I), YL(I), XH(I), XL(I), ZL(I), PD(I)
   I = I + 1
   ReDim Preserve YH(0 To I), YL(0 To I), XH(0 To I), XL(0 To I), ZL(0 To I), PD(0 To I)
   lstFormulae.AddItem A$
Loop Until EOF(1)
Close
End Sub

Private Sub cmdReadMe_Click()
MsgBox ("Make sure to close Notepad and do not change folder")
FileSpec$ = App.Path + "/ReadMe.txt"
n$ = "NOTEPAD.EXE " & FileSpec$
res& = Shell(n$, 1)

End Sub


Private Sub cmdEvaluate_Click()

On Error GoTo evalerror

yyhi = Val(txtYHI.Text)
yylo = Val(txtYLO.Text)
xxhi = Val(txtXHI.Text)
xxlo = Val(txtXLO.Text)
zlim = Val(txtZLIM.Text)
ye = Val(txtPERSPEC.Text)
If (yyhi = 0 And yylo = 0) Or (xxhi = 0 And xxlo = 0) Then
   MsgBox ("Zero limits can cause too many errors" + vbCrLf + "Change limits")
   Exit Sub
End If
If yylo > yyhi Then
   MsgBox ("Y-High must be higher than Y-Low!" + vbCrLf + "Change limits")
   Exit Sub
End If
If xxlo > xxhi Then
   MsgBox ("X-High must be higher than X-Low!" + vbCrLf + "Change limits")
   Exit Sub
End If

grxhi = Val(txtPlotIntv.Text)
gryhi = grxhi

'Original function points
ReDim svrx(1 To grxhi, 1 To gryhi)
ReDim svry(1 To gryhi, 1 To gryhi)
ReDim svrz(1 To grxhi, 1 To gryhi)
'Rotated function points
ReDim rx(1 To grxhi, 1 To gryhi)
ReDim ry(1 To gryhi, 1 To gryhi)
ReDim rz(1 To grxhi, 1 To gryhi)
'Display points
ReDim xs(1 To grxhi, 1 To gryhi)
ReDim ys(1 To grxhi, 1 To gryhi)
ReDim zs(1 To grxhi, 1 To gryhi)

Normal:
'EVALUATE STRING
A$ = txtFormula.Text
p = InStr(1, A$, ":")
If p = 0 Then p = Len(A$) + 1
txtFormula.Text = Left(A$, p - 1)

estring$ = LCase(txtFormula.Text)
SqueezeSpaces estring$
txtFormula.Text = estring$
nleftbrackets = NumOccStr(estring$, "(")
nrightbrackets = NumOccStr(estring$, ")")
If nleftbrackets <> nrightbrackets Then
   MsgBox ("Unmatched brackets")
   Exit Sub
End If
ReplaceLOG estring$
ReplaceASIN estring$
ReplaceACOS estring$
ReplaceSINH estring$
ReplaceCOSH estring$
ReplaceTANH estring$
ReplacePI estring$
ReplaceLN estring$

'Test
'MsgBox (estring$)
'Exit Sub

'Scale factors for converting grid points to x,y values
scx = (xxhi - xxlo) / (grxhi - 1)
scy = (yyhi - yylo) / (gryhi - 1)
For J = 1 To gryhi
For I = 1 To grxhi
   svrx(I, J) = scx * (I - 1) + xxlo
   svry(I, J) = scy * (J - 1) + yylo
Next I
Next J

'Fill svrz(i,j) with Func(x,y)
For J = 1 To gryhi
For I = 1 To grxhi
   X = svrx(I, J)
   Y = svry(I, J)
   cstring$ = estring$
   ReplaceXY cstring$, X, Y
   '---------------------------------------------
   svrz(I, J) = SCI.Eval(cstring$)  'MS Script Eval
   '---------------------------------------------
   If svrz(I, J) < -zlim Then svrz(I, J) = -zlim
   If svrz(I, J) > zlim Then svrz(I, J) = zlim
Next I
Next J

'Used with scaling factors
zpL = 0.2 * picDisplay.ScaleWidth
zpH = 0.8 * picDisplay.ScaleWidth

'Display initial points
'----------------------------------
picDisplay_MouseMove 1, 0, 0, 0
'----------------------------------

Exit Sub
'==================
evalerror:
res& = MsgBox("Error #" & CStr(Err.Number) & " " & Err.Description + vbCrLf + cstring$, 5)
'NB Err.Number 13 will probably be a misspelling
If res& = vbRetry Then
   'Check for Overflow or Division by zero
   If Err.Number = 6 Or Err.Number = 11 Then
      svrz(I, J) = zlim
   Else
      svrz(I, J) = 0
   End If
   Err.Clear
   Resume Next
Else
   On Error GoTo 0
   Exit Sub
End If
End Sub


Private Sub Form_Load()

Show
fwidth = Form1.ScaleWidth
fheight = Form1.ScaleHeight
Form1.DrawWidth = 3
Form1.Line (2, 2)-(fwidth - 4, fheight - 30), RGB(200, 200, 200), B
Form1.DrawWidth = 1
Caption = "RR3D  by  Robert Rayment"
'------------------------------------------------------------------
'INPUT TWO FILES

'--------  FUNCTION SYNTAX -----------------
With picFunctions
   .FontName = "MS Serif"
   .FontSize = 8
End With
On Error GoTo PlotterError
Open "Valid.txt" For Input As #1
Do
   Line Input #1, A$
   A$ = "  " + A$
   picFunctions.Print A$
Loop Until EOF(1)
Close

'--------  SAVED FORMULAE  -----------------
'Syntax:
'string formula[: description], yhi, ylo, xhi, xlo, zlim, perspective distance ye

FormulaList:
With lstFormulae
   .FontName = "MS Serif"
   .FontSize = 8
End With
lstFormulae.Clear
ReDim YH(0 To 0), YL(0 To 0), XH(0 To 0), XL(0 To 0), ZL(0 To 0), PD(0 To 0)
On Error GoTo FormulaeError
Open "Formulae.txt" For Input As #1
I = 0
Do
   Input #1, A$, YH(I), YL(I), XH(I), XL(I), ZL(I), PD(I)
   I = I + 1
   ReDim Preserve YH(0 To I), YL(0 To I), XH(0 To I), XL(0 To I), ZL(0 To I), PD(0 To I)
   lstFormulae.AddItem A$
Loop Until EOF(1)
Close
'------------------------------------------------------------------

Instructions:
On Error GoTo 0
FillInInstructions

'For adjusting SetCursorPos for 800x600 resolution
topoffset = 599 - ScaleHeight

'STARTING FORMULA
txtFormula.Text = "x^2-y^2"
labTitle.Caption = "Saddle"
'Initial x,y,z,zlim,ye limits
yyhi = 6
yylo = -6
xxhi = 6
xxlo = -6
zlim = 100
ye = 20 'perspective distance
txtYHI.Text = yyhi
txtYLO.Text = yylo
txtXHI.Text = xxhi
txtXLO.Text = xxlo
txtZLIM.Text = zlim
txtPERSPEC.Text = ye
txtPlotIntv.Text = 20

'DEFAULT GRIDS PLOT INTERVALS
grxhi = 10
gryhi = 10

'PLOT COLORS
culgreen& = QBColor(10)
culcyan& = QBColor(11)

'picDisplay.AutoRedraw = False    'Faster but slight flicker
picDisplay.AutoRedraw = True     'No flicker

'-----------------------
cmdEvaluate_Click    'Do starting plot
'-----------------------

Exit Sub
'=====================
PlotterError:
MsgBox ("Error cannot open Valid.txt" + vbCrLf + "Must be in same Folder as Application")
Close
Resume FormulaList
'=====================
FormulaeError:
MsgBox ("Cannot open or error in Formulae.txt" + vbCrLf + "Must be in same Folder as Application")
Resume Instructions
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub lstFormulae_Click()
'SELECT FORMULA FROM LIST
Index = lstFormulae.ListIndex
A$ = lstFormulae.List(Index)
k = InStr(1, A$, ":")
If k = 0 Then
   k = Len(A$) + 1
   txtFormula.Text = Left(A$, k - 1)
   labTitle.Caption = ""
Else
   txtFormula.Text = Left(A$, k - 1)
   labTitle.Caption = Mid$(A$, k + 1)
End If

yyhi = YH(Index)
yylo = YL(Index)
xxhi = XH(Index)
xxlo = XL(Index)
zlim = ZL(Index)
ye = PD(Index)
txtYHI.Text = yyhi
txtYLO.Text = yylo
txtXHI.Text = xxhi
txtXLO.Text = xxlo
txtZLIM.Text = zlim
txtPERSPEC.Text = ye
End Sub

Private Sub picDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
res& = GetCursorPos(pp) 'On whole window
mx = pp.X
my = pp.Y
codeval = KeyCode
Select Case codeval
Case 37: mx = mx - 1: SetCursorPos mx, my 'Left key
Case 38: my = my - 1: SetCursorPos mx, my 'Up
Case 39: mx = mx + 1: SetCursorPos mx, my 'Right
Case 40: my = my + 1: SetCursorPos mx, my 'Dn
Case 13
End Select

xp = mx - picDisplay.Left
yp = my - picDisplay.Top - topoffset
picDisplay_MouseMove 1, 0, xp, yp

End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then   'Save X,Y with right-click
   XMouse = X
   YMouse = Y
End If
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-----------------------------------

If Button = 0 Then       'Show values only
   FindClosestij X, Y, I, J
   For n = 0 To 2
      picXYZ(n).Cls
   Next n
   picXYZ(0).Print svrx(I, J)
   picXYZ(1).Print svry(I, J)
   picXYZ(2).Print svrz(I, J)
   Exit Sub
End If
   
'Hold X or Y for X,Y rotations only
If chkYROT = 1 Then Y = YMouse
If chkXROT = 1 Then X = XMouse

'-----------------------------------
CalculateScreenPoints X, Y
'-----------------------------------

'Dim pp As POINTAPI
'res& = LineTo(Form1.hdc, xx&, yy&)
'res& = MoveToEx(Form1.hdc, xx&, yy&, pp)

picDisplay.Cls 'Simplest & Fast enough

'Draw X-lines
picDisplay.ForeColor = culgreen&
For J = 1 To gryhi
   res& = MoveToEx(picDisplay.hdc, xs(1, J), zs(1, J), pp)
For I = 1 + 1 To grxhi
   res& = LineTo(picDisplay.hdc, xs(I, J), zs(I, J))
Next I
Next J

'Draw Y-lines
picDisplay.ForeColor = culcyan&
For I = 1 To grxhi
   res& = MoveToEx(picDisplay.hdc, xs(I, 1), zs(I, 1), pp)
For J = 1 + 1 To gryhi
   res& = LineTo(picDisplay.hdc, xs(I, J), zs(I, J))
Next J
Next I
End Sub

Private Sub CalculateScreenPoints(X, Y)
'X,Y cursor position from MouseDown

'Get angles based on cursor position
xcen = picDisplay.ScaleWidth / 2
ycen = picDisplay.ScaleHeight / 2
zang = (pi / 2) * ((X - xcen) / xcen)  'zang about z-axis
xang = -(pi / 2) * ((Y - ycen) / ycen) 'xang about x-axis

'Apply rotation to original data about z-axis
For J = 1 To gryhi
For I = 1 To grxhi
   rx(I, J) = svrx(I, J) * Cos(zang) + svry(I, J) * Sin(zang)
   ry(I, J) = svry(I, J) * Cos(zang) - svrx(I, J) * Sin(zang)
   rz(I, J) = svrz(I, J)
Next I
Next J
'Apply rotation about x-axis
For J = 1 To gryhi
For I = 1 To grxhi
   rx(I, J) = rx(I, J)
   ry(I, J) = ry(I, J) * Cos(xang) - rz(I, J) * Sin(xang)
   rz(I, J) = ry(I, J) * Sin(xang) + rz(I, J) * Cos(xang)
Next I
Next J

If CheckPerspective = 1 Then
   
   'Find the intercept at plane y=0 (ie the screen plane) of the
   'line connecting the eye point (xe,ye,ye) with each function
   'point (rx(),ry(),rz()) in turn.
   'The display intercept points will be modified & unscaled
   'rx(i,j) & rz(i,j)

   'EYE POINT  ye settable
   xe = 0: ze = 0

   For J = 1 To gryhi
   For I = 1 To grxhi
      zd = (ye - ry(I, J))
      If zd = 0 Then
         rx(I, J) = rx(I, J)
      Else
         rx(I, J) = -ye * (xe - rx(I, J)) / zd
      End If
      If zd = 0 Then
         rz(I, J) = rz(I, J)
      Else
         rz(I, J) = -ye * (ze - rz(I, J)) / zd
      End If
   Next I
   Next J

End If

'-------  FIND MAX & MINS -------------------

xmax = -10000: xmin = 10000
zmax = -10000: zmin = 10000
For I = 1 To grxhi
For J = 1 To gryhi
   If rx(I, J) > xmax Then xmax = rx(I, J)
   If rx(I, J) < xmin Then xmin = rx(I, J)
   If rz(I, J) > zmax Then zmax = rz(I, J)
   If rz(I, J) < zmin Then zmin = rz(I, J)
Next J
Next I

'-------  FIND SCALE FACTORS  ----------

'Horizontal Scaling
'xs=scfx*(x-xmin)+xoff
'zpL = 0.2 * picDisplay.ScaleWidth
'zpH = 0.8 * picDisplay.ScaleWidth

If (xmax - xmin) = 0 Then
   scfx = 1
Else
   scfx = (zpH - zpL) / (xmax - xmin)
End If
xoff = zpL

'Vertical scaling
'zs=scfz*(z-zmin)+zoff
'zpL = 0.2 * picDisplay.ScaleHeight
'zpH = 0.8 * picDisplay.ScaleHeight
If (zmax - zmin) = 0 Then
   scfz = 1
   zoff = 0.5 * picDisplay.ScaleHeight
Else
   scfz = -(zpH - zpL) / (zmax - zmin)
   zoff = zpH
End If

'-----  GET PLOTTING POINTS  ----------------------

'Get picDisplay points xs(), zs()
For J = 1 To gryhi
For I = 1 To grxhi
   xs(I, J) = scfx * (rx(I, J) - xmin) + xoff
   zs(I, J) = scfz * (rz(I, J) - zmin) + zoff
Next I
Next J

End Sub

Private Sub FindClosestij(X, Y, ip, jp)
'In:  x,y picDisplay coords
'Out: ip,jp of closest xs(i,j),zs(i,j)
zdiss = 10000
For J = 1 To gryhi
For I = 1 To grxhi
   zdis = Abs(X - xs(I, J)) + Abs(Y - zs(I, J))
   If zdis < zdiss Or zdis = 0 Then
      zdiss = zdis
      idis = I
      jdis = J
   End If
Next I
If zdiss = 0 Then Exit For
Next J

ip = idis
jp = jdis

End Sub

Private Sub FillInInstructions()
With Label6
   .FontName = "MS Serif"
   .FontSize = 8
End With


A$ = "INSTRUCTIONS" + vbCrLf
A$ = A$ + "Click on formula to transfer to text box, modify" + vbCrLf
A$ = A$ + "if wanted, then press Evaluate formula." + vbCrLf
'a$ = a$ + vbCrLf
A$ = A$ + "CONTROLS:" + vbCrLf
A$ = A$ + "Move Mouse over display to show X,Y,Z values" + vbCrLf
A$ = A$ + "Move MouseDown over display to rotate" + vbCrLf
A$ = A$ + "Right-Click-MouseDown to save X & Y" + vbCrLf
A$ = A$ + "   for X or Y axis rotation only" + vbCrLf
A$ = A$ + "Cursor Keys can also be used" + vbCrLf
'a$ = a$ + vbCrLf
A$ = A$ + "ERRORS:" + vbCrLf
A$ = A$ + "Ignore Run-time errors, key No" + vbCrLf
A$ = A$ + "Other errors, try Retry to go on, Cancel to go" + vbCrLf
A$ = A$ + "back to formulae." + vbCrLf
A$ = A$ + "'Spiking', rotate further or change plot limits." + vbCrLf

Label6.Caption = A$
End Sub

Private Sub cmdSaveVertices_Click()
Title$ = "Save Screen Vertices"
Choice$ = "Vertex text file files(*.txt)|*.txt"
InitDir$ = App.Path
OpenSaveDialog Title$, Choice$, SaveFile$, InitDir$
If SaveFile$ <> "" Then
   Open SaveFile$ For Output As #1
   asub1 = gryhi
   asub2 = grxhi
   Print #1, asub1; ","; asub2
   For J = 1 To gryhi
   For I = 1 To grxhi
      Print #1, xs(I, J); ","; zs(I, J)
   Next I
   Next J
   Close
End If
End Sub

Private Sub OpenSaveDialog(Title$, Choice$, SaveFile$, InitDir$)
CommonDialog1.DialogTitle = Title$
CommonDialog1.Flags = &H2  '&H2 checks if file exists
CommonDialog1.CancelError = True
On Error GoTo Acancel
CommonDialog1.Filter = Choice$
CommonDialog1.InitDir = InitDir$
CommonDialog1.ShowSave
SaveFile$ = CommonDialog1.FileName
Exit Sub
'============
Acancel:
Close
SaveFile$ = ""
Exit Sub
Resume
End Sub

