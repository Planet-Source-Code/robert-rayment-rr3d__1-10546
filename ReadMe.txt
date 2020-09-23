RR3D  by Robert Rayment    30/7/00

Z = Func(X,Y)

RR3D is a program for displaying three-dimensional Cartesian formulae as a rotatable 
and scaleable wire-frame with and without perspective.

Included is a list of formulae which can be added to or editted with your own functions.
These are contained in a file -  Formulae.txt  - which should be in the same folder as the
application.  The format is shown below.  

Apart from the normal VB maths, inverse and hypergbolic functions can also be
evaluated.  Valid entries are contained in a file -  Valid.txt  -, which should also be in the 
same folder as the application.  Valid entries are shown below.

The Microsoft Script Control (comes with VB6) Eval function simplifies calculations
enormously and its use is demonstrated here.  A complication arises with errors.  Script
will throw out a run-time error which should be ignored. In other words key No in the
run-time error message box.  If debugged it can crash the program.  RR3D will deal with 
most errors.

Common errors are:-

1.  Unmatched brackets
2.  Overflow and division by zero
3.  Misspelling of functions
4.  Illegal values for a particular function.

All of these can be dealt with by re-writing the formula and/or changing the X,Y high/low
limits.  The limit of Z values can be set and if exceeded will be set to this limit.  When this 
happens flat areas will show up in the display.

The number of plotting intervals and perspective depth can be set.  Perspective, 
rotation about Y and rotation about X can be switched on and off.  X,Y.Z values of 
the  function are shown by simply moving the mouse over the display.  With perspective
on, many functions will 'spike' when the perspective depth gets inside the function.
This is normal, simply rotate to a different angle or increase the depth.  If the redrawing
gets a bit too slow decrease the number of plot intervals.  Note that the mouse can be
moved outside the display to get angles with larger multiples of pi.  

The display vertices can be saved at any time for reproducing the image in another
program.  An example is shown in the PlotVertices folder.  This example could be
extended to save images as bitmaps or to print them.

Some easy to use API's are included:
MoveToEx and LineTo,  these are very much faster than the VB equivalents
SetCursorPos and GetCursorPos,  these are absolute positions of the mouse and 
	need adjusting for the location in Controls.  They are used  in conjunction 
	with the cursor keys.
	
Instructions included in the program are:-
----------------------------------------------------------
INSTRUCTIONS
Click on formula to transfer to text box modify
 if wanted, then press Evaluate formula.
CONTROLS:
Move Mouse over display to show X,Y,Z values
Move MouseDown over display to rotate
Right-Click-MouseDown to save X & Y
   for X or Y axis rotation only
Cursor Keys can also be used
ERRORS:
Ignore Run-time errors, key No
Other errors, try Retry to go on, Cancel to go
  back to formulae.
'Spiking', rotate further or change plot limits.
----------------------------------------------------------

INTERESTING FORMULAE

(x^2+y^2): Roll to a sphere, 6 ,-6 , 6 ,-6 , 100 , 20 
with perspective on this can be rotated into what looks
like a ball.  Rotateand click with the right-button will save the
last X,Y coords, check Y rotation only, then moving the 
mouse with the left-button down creates the illusion of a
rotating sphere.

sin(sqr((x/6)^2+(y/6)^2)): Center spike, 10 ,-10 , 10 ,-10 , 100 , 100 
tanh(x^2+y^2): Center spike 2, 2 ,-2 , 2 ,-2 , 100 , 50 
exp(-(x^2+y^2)): Edge spike,3,-3,5,0,100,100
when rotated, by moving the mouse up and down, appears
to show a spike rising out of a the center and edge of a sheet
respectively.

tanh(x+y): Flying sheet, 2 ,-2 , 2 ,-2 , 100 , 50 
Flapping sheet of paper which can roll up into a tube.

FORMULAE.TXT FORMAT

[optional]
On each line:-

String formula[: description], Y-high, Y-low, X-high, X-low, Z-limit, Perspective depth(ey)

eg

(x^2+y^2): 3D quadratic, 6 ,-6 , 6 ,-6 , 100 , 20 

Note the semi-colon after the formula and the placement of commas.

VALID SYNTAX FOR FORMULAE

SPACES	ignored			CASE     upper or lower
VARIABLES	x, y		BRACKETS ( )
NUMBERS eg 1, 1.1, -3.77		pi = 3.1415927
OPERATIONS ^  *  /  +  -		Operation priority    ^  *   /  +  -
-------------------------------------FUNCTIONS--------------------------------------------------
sin(x)     sine		cos(x)    cosine
tan(x)    tangent		atn(x)    arctan    +/- pi/2
ln(x)       natural log		exp(x)   exponential     x<= 88	
abs(x)   absolute		cint(x)    round to integer
sqr(x)    square root
------------------------------DERIVED FUNCTIONS----------------------------------------
log(x)    base 10 log		= ln(x) / ln(10)
asin(x)   arcsin		= atn(x / sqr(1 - x^2))    x<1
acos(x)  arccos		= pi / 2 - asin(x)              x<1
sinh(x)   hyperbolic sin	= (exp(x) -  exp( - x)) / 2
cosh(x)  hyperbolic cos	= (exp(x) + exp( - x)) / 2
tanh(x)  hyperbolic tan	= sinh(x) / cosh(x)
