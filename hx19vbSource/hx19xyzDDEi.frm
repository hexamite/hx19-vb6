VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Hx19 3D interface"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5400
   Icon            =   "hx19xyzDDEi.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "hx19"
   ScaleHeight     =   10230
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Set"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Text            =   "99999,10000 >"
      Top             =   9840
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "< 0,0 (x,z)"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Text            =   "99000,99000 >"
      Top             =   5280
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   4815
      TabIndex        =   9
      Top             =   5640
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Text            =   "< 0,0 (x,y)"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TX"
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Sync"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "DDE"
      ToolTipText     =   "Dynamic Data Exchange"
      Top             =   480
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Log"
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      ToolTipText     =   "log incoming data \dataFiles\"
      Top             =   480
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Left            =   2280
      Top             =   2520
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "Ready for data"
      ToolTipText     =   "tag, X,Y,Z, time(10mS),record,#receivers,(detecting receivers)"
      Top             =   480
      Width           =   4095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2280
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   4815
      TabIndex        =   8
      Top             =   960
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim linebuffer(200)
Dim rxIds(1024) As Integer
Dim pt(1024) As Integer
Dim lx(1024)
Dim l2x(1024)
Dim ly(1024)
Dim lz(1024)
Dim rx(1024, 21)
Dim xt(1024, 21)
Dim xin(300)
Dim s1x, s2x, s1y, s1z, ofx%, ofy%, ofz%
Dim xtSorted(21)
Dim rxSorted(21)
Dim diff(200)
Dim rxA(200)
Dim rxB(200)
Dim acdiff(200)
Dim temp$
Dim xx(2000)
Dim yy(2000)
Dim zz(2000)
Dim good, xp, yp, zp, nn, scount, count100ms, norx%
Dim xstep, ystep, zstep, minPoints, maxPoints, zmin, zmax, xmin, xmax, ymin, ymax, fx%, fy%, fz%

Dim xst(1024), yst(1024), zst(1024)
Dim xbt(1024), xf(1024)
Dim ybt(1024), yf(1024)
Dim zbt(1024), zf(1024)
Dim alpha, beta

Private Sub Check1_Click()
If Check1.Value = 1 Then
logfile = Format(Date, "mmm dd yy ") + Format(Time, "hh mm ss") + ".xyz"
On Error GoTo fixit
Open temp + "\dataFiles\" + logfile For Output As 1
Else
Close 1
End If
GoTo fdone
fixit:
MkDir (temp + "\dataFiles\")
DoEvents
Open temp + "\dataFiles\" + logfile For Output As 1
Resume Next
fdone:
End Sub
Private Sub Check3_Click()
If Check3.Value = 1 Then checkOut "M&$" Else checkOut "M&%"

End Sub

Private Sub Command1_Click()
checkOut Text3
End Sub


Private Sub Form_Activate()
Dim dummy$
Form1.Caption = CurDir("")
temp = CurDir("")
i = 0
'********** Get configuration file
cnfig = temp + "\hx19xyzDDEi.txt"
xxmap = temp + "\map.txt"
Form1.Caption = cnfig
Open cnfig For Input As 1

Input #1, cport, dummy
Input #1, alpha, dummy
Input #1, beta, dummy
Input #1, minPoints, dummy
Input #1, maxPoints, dummy
Input #1, ofx, dummy
Input #1, ofy, dummy
Input #1, fx, dummy
Input #1, fy, dummy
Input #1, ofz, dummy
Input #1, fz, dummy


Close

    Form1.MSComm1.CommPort = cport
    Form1.MSComm1.Settings = "256000,N,8,1"
    Form1.MSComm1.InputLen = 1
    Form1.MSComm1.PortOpen = True

'****** check parameters
If minPoints < 1 Then minPoints = 1
If maxPoints < 1 Then maxPoints = 1
If minPoints > 20 Then minPoints = 20
If maxPoints > 20 Then maxPoints = 20
'the following checks to see if the Z coordinates are within capacity of the system
If zmin < 0 Then zmin = 0
If zmin > 16000 Then zmin = 16000
If zmax < 0 Then zmax = 0
If zmax > 16000 Then zmax = 16000
'the following scales the display window
s1x = Picture1.ScaleWidth / (fx - ofx)
s2x = Picture2.ScaleWidth / (fx - ofx)
s1y = Picture1.ScaleHeight / (fy - ofy)
s1z = Picture2.ScaleHeight / (fz - ofz)

Text4 = "< " + Str(ofx) + "," + Str(ofy) + " (x,y)"
Text5 = Str(fx) + "," + Str(fy) + " (x,y) >"
Text7 = "< " + Str(ofx) + "," + Str(ofz) + " (x,z)"
Text8 = Str(fx) + "," + Str(fz) + " (x,z) >"

'mdir = temp + "\mapImages\"
Form1.Caption = xxmap

'dogrid
Open xxmap For Input As 1
  Picture1.DrawWidth = 3
  Picture2.DrawWidth = 3
Do
Input #1, id
If id > 0 And id < 1024 Then
  Input #1, mx, my, mz
  rxIds(norx) = id
  xx(id) = mx: yy(id) = my: zz(id) = mz
  xs1 = s1x * (mx - ofx): ys1 = s1y * (my - ofy): xs2 = s2x * (mx - ofx): zs1 = s1z * (mz - ofz)
  Picture1.Line (xs1, ys1)-(xs1 + 100, ys1 + 100), , B
  Picture2.Line (xs2, zs1)-(xs2 + 100, zs1 + 100), , B
  norx = norx + 1
End If
Loop Until EOF(1)
Close 1
  Picture1.DrawWidth = 1
  Picture2.DrawWidth = 1
  

'************ Start main program
Text1 = "Connect your serial port to network reader"


Timer1.Interval = 10

    Form1.MSComm1.InputLen = 1 'get serial characters one by one
    Form1.MSComm1.InBufferCount = 0
    Form1.MSComm1.OutBufferCount = 0
    
    Form1.MSComm1.Output = "$"
    Do: DoEvents: Loop Until Form1.MSComm1.OutBufferCount = 0
    
getready:
 Do
   Do: DoEvents: Loop While Form1.MSComm1.InBufferCount < 1
   cc = Form1.MSComm1.Input
 Loop Until cc = Chr(13)        '":"
 Text1 = "Activity: configure hx19 and/or click sync"
   Do: DoEvents: Loop While Form1.MSComm1.InBufferCount < 1
   cc = Form1.MSComm1.Input
If cc <> "X" Then GoTo getready
 

 Do
    Do: DoEvents: Loop While Form1.MSComm1.InBufferCount < 1    'this loop waits for a character from the hx19ms connected to the USB port.
    cc = Form1.MSComm1.Input
    cline = cline + cc

stsize = stsize + 1
If cc = "X" And marker = True Then

For i = 1 To Len(cline) - 1
cc = Mid(cline, i, 1)
If cc = "R" Then Do: i = i + 1: zxn = Mid(cline, i, 1): ctmp = ctmp + zxn: DoEvents: Loop Until zxn = " ": xin(1) = Val(ctmp): ctmp = "":
If cc = "U" Then Do: i = i + 1: zxn = Mid(cline, i, 1): ctmp = ctmp + zxn: DoEvents: Loop Until zxn = " ": xin(0) = Val(ctmp): ctmp = "":
If cc = "P" Then Do: i = i + 1: zxn = Mid(cline, i, 1): ctmp = ctmp + zxn: DoEvents: Loop Until zxn = " ": xin(0) = Val(ctmp): ctmp = "":
If cc = "C" Then Do: i = i + 1: zxn = Mid(cline, i, 1): ctmp = ctmp + zxn: DoEvents: Loop Until zxn = " ": xin(2) = Val(ctmp): ctmp = "":
If cc = "A" Then Do: i = i + 1: zxn = Mid(cline, i, 1): ctmp = ctmp + zxn: DoEvents: Loop Until zxn = " ": xin(2) = Val(ctmp): ctmp = "":

If (xin(0) > 0 And xin(0) < 1024 And cc = Chr(13)) And pt(xin(0)) < 21 Then
  rx(xin(0), pt(xin(0))) = xin(1): xt(xin(0), pt(xin(0))) = xin(2): pt(xin(0)) = pt(xin(0)) + 1
  'the above parses the ascii characters into numbers, and places the receiverID into the rx array, and the distance measured into the xt array
  'the pt array states how many receivers detected a given tag
  xin(0) = 0: xin(1) = 0: xin(2) = 0
End If
Next
cline = ""
xProcess
End If

If cc = Chr(13) Then marker = True Else: marker = False
'carriage return and marker is used to determine the end of a record and a beginning of a new one
 Loop
  
End Sub


Private Sub xProcess()
For i = 1 To 1023
If pt(i) >= minPoints Then positionTag (i) Else pt(i) = 0
Next
End Sub

Private Sub positionTag(tag)
Static prox%

freeze = count100ms
scount = scount + 1
If count100ms > 8640000 Then count100ms = 0: scount = 0

If pt(tag) > 20 Then pt(tag) = 20
If pt(tag) < 1 Then GoTo skipover

'the following places receiver ID and receiver distance for the tag being positioned into arrays ordered with the closest receiver first
For j = 0 To pt(tag) - 1
 xlow = 100000000000#
 For i = 0 To pt(tag) - 1
  If xt(tag, i) <= xlow And xt(tag, i) <= 15000 And xt(tag, i) > 0 And rx(tag, i) > 0 Then xlow = xt(tag, i): mn = i
 Next
 xtSorted(j) = xlow: rxSorted(j) = rx(tag, mn): rx(tag, mn) = 0:
Next
'the following is for display only and not important
receivers = ""
For i = 0 To pt(tag) - 1
receivers = receivers + Format(rxSorted(i), " 0")
Next


If maxPoints = 1 Or pt(tag) < 2 Then    'if not enough receivers exist to compute 3d or 2d then place the point close to the mapped location of the detecting receiver
 xp = xx(rxSorted(0)) + 100 * Rnd()
 yp = yy(rxSorted(0)) + 100 * Rnd()
 zp = zz(rxSorted(0))
Text1 = Format(tag, " 0") + Format(xp, " 0") + Format(yp, " 0") + Format(zp, " 0") + Format(freeze, " 0") + Format(scount, " 0 ") + Format(pt(tag), " 0 ") + "(" + receivers + ")"
If Check1.Value = 1 Then Print #1, tag, Round(xp, 0), Round(yp, 0), Round(zp, 0), freeze, scount, pt(tag); "("; receivers; ")"

 GoTo skipover
End If

If pt(tag) < minPoints Then GoTo skipover
If maxPoints < pt(tag) Then pt(tag) = maxPoints
'for 3d or 2d the following is executed
'following task
ii = 0
pts = mm
For j = 0 To pt(tag) - 2
 For i = j + 1 To pt(tag) - 1
  di = (xtSorted(i) - xtSorted(j))
   maxInterval = Sqr((xx(rxSorted(i)) - xx(rxSorted(j))) ^ 2 + (yy(rxSorted(i)) - yy(rxSorted(j))) ^ 2)
   If maxInterval > di And di >= 0 Then
   diff(ii) = di                            'this array holds the difference in the time of flight relative to a pair of receivers and is used to estimate the x and y coordinates
   rxA(ii) = rxSorted(i): rxB(ii) = rxSorted(j)
   ii = ii + 1
  End If
 Next
Next
good = ii   'the good variable states how many pairs qualified the sorting proceedure

ymax = -10000: xmax = -10000
ymin = 100000: xmin = 100000

'the following looks for the position of the receivers in the map file, and determines the range by which the computational iteration is to take place.
For i = 0 To good - 1
If xx(rxA(i)) > xmax Then xmax = xx(rxA(i))
If xx(rxA(i)) < xmin Then xmin = xx(rxA(i))
If xx(rxB(i)) > xmax Then xmax = xx(rxB(i))
If xx(rxB(i)) < xmin Then xmin = xx(rxB(i))
If yy(rxA(i)) > ymax Then ymax = yy(rxA(i))
If yy(rxA(i)) < ymin Then ymin = yy(rxA(i))
If yy(rxB(i)) > ymax Then ymax = yy(rxB(i))
If yy(rxB(i)) < ymin Then ymin = yy(rxB(i))
Next

z = 0
dIteration 0    'first itereation is started for Z=0
getZ (pt(tag))  'Z is estimated and feed to next iteration and so forth.
dIteration zp
getZ (pt(tag))
dIteration zp
getZ (pt(tag))

If xp < 0 Or xp > 100000 Then xp = xf(tag)  'if error is unacceptable we us x forcasted from the double exponential filter.
If yp < 0 Or yp > 100000 Then yp = yf(tag)
If zp < 0 Or zp > 15000 Then zp = zf(tag)

'the followin is a double exponential filter as presented by wikipedia see (http://en.wikipedia.org/wiki/Exponential_smoothing)
'there are probably better ways of filtering, but this is fairly good
If alpha <> 0 Then exponentialFilter tag
Text1 = Format(tag, " 0") + Format(xp, " 0") + Format(yp, " 0") + Format(zp, " 0") + Format(freeze, " 0") + Format(scount, " 0 ") + Format(pt(tag), " 0 ") + "(" + receivers + ")"
If Check1.Value = 1 Then Print #1, tag, Round(xp, 0), Round(yp, 0), Round(zp, 0), freeze, scount, pt(tag); "("; receivers; ")"
sixa% = tag
'the following just plots the position of the tag on the window set in the setup file
xyDotPlot sixa
xzDotPlot sixa
prox = (prox + 1) And 7: If prox = 0 Then refreshMap ': dogrid
skipover:
pt(tag) = 0
End Sub
Private Sub dIteration(zz1)

'this routine selects the steps at which the x, y and z are to be computationally tested against measured results
'at first a coarse minima is found, then the steps get finer and finer until a precise minima is found
'note that when a coarse minima is found, for next iteration there is a slight backstep and then the next steps are devided by 2
'so it gets finer and finer.

xstep = (xmax - xmin) / 4
xmin = xmin - 4 * xstep
xmax = xmax + 4 * xstep

ystep = (ymax - ymin) / 4
ymin = ymin - 4 * ystep
ymax = ymax + 4 * ystep
zzz = zz1
dscanXYZ xmin, xmax, ymin, ymax, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
  i = xstep: xstep = xstep / 2
  j = ystep: ystep = ystep / 2
'  l = zstep: zstep = zstep / 2
dscanXYZ xp - i, xp + i, yp - j, yp + j, zzz
donn:
End Sub

Private Sub dscanXYZ(xS, xE, yS, yE, zpa)

'the following is about finding at which xyz the distances measured from tag to receivers agree.
'the formula is distance^2=(receiverXlocation-x)^2+(receiverYlocation-y)^2+(receiverZlocation-z)^2
'this routine finds the x y and z where the difference between calculated distance and measured distance is minimum

minima = 1E+19
DoEvents
If xstep < 1 Then GoTo dsdoner
If ystep < 1 Then GoTo dsdoner
If (xE - xS) < 1 Then xxs = 1 Else xxs = xstep
If (yE - yS) < 1 Then yys = 1 Else yys = ystep
z = zpa
zp = 0
For x = xS To xE Step xxs          'ScanStep
  For y = yS To yE Step yys        'ScanStep
   ad = 0
    For c = 0 To good - 1
     xa = x - xx(rxA(c))
     ya = y - yy(rxA(c))
     za = z - zz(rxA(c))
     xb = x - xx(rxB(c))
     yb = y - yy(rxB(c))
     zb = z - zz(rxB(c))
     dd = Sqr(xa * xa + ya * ya + za * za) - Sqr(xb * xb + yb * yb + zb * zb)
     acdiff(c) = (dd - diff(c)) ^ 2 / good
     ad = ad + acdiff(c)    '((dd - difference(c)) ^ 2) / good
    Next
     If ad < minima Then minima = ad: xp = x: yp = y
 Next y
Next x
dsdoner:
End Sub

Private Sub getZ(nrx)
'find the Z distance based on known x and y
    zav = 0
    For i = 0 To nrx - 1
     xa = xp - xx(rxSorted(i))
     ya = yp - yy(rxSorted(i))
     za = zp - zz(rxSorted(i))
     kc = xtSorted(i) * xtSorted(i) - xa * xa - ya * ya
     If kc >= 0 Then za = Sqr(kc)
     zp = za + rxSorted(i)
     zav = zav + zp
    Next
    zp = zav / nrx
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
count100ms = count100ms + 1
End Sub

Private Sub checkOut(temst$)
Dim xx$, xsum%
'the following prepairs a string to be sent to the hx19 network, with the appropriate checksum
 xsum = 0
 For i = 1 To Len(temst)            'compute the checksum of the string
  xx = Mid(temst, i, 1)
  xsum = xsum + Asc(xx)             'accumulate ASCII codes
 Next
 temst = temst + "/" + Hex(xsum)    'append the checksum in hexadecimal format
 Form1.MSComm1.Output = temst + Chr(13)

End Sub
Private Sub xzDotPlot(xztag%)

Picture2.DrawWidth = 7
Picture2.ForeColor = QBColor(15)
Picture2.FillColor = QBColor(15)

ccx = s2x * (l2x(xztag) - ofx)
ccz = s1z * (lz(xztag) - ofz)
'Print ccx, ccz, Picture2.ScaleWidth, Picture2.ScaleHeight
If (ccx >= 0 And ccx < Picture2.ScaleWidth And ccz >= 0 And ccz < Picture2.ScaleHeight) Then Picture2.PSet (ccx, ccz)
'Picture2.PSet (ccx, ccz)

xu = xztag And 15
If xu = 15 Then xu = 0
DoEvents
Picture2.ForeColor = QBColor(xu)
Picture2.FillColor = QBColor(xu)
ccx = s2x * (xp - ofx)
ccz = s1z * (zp - ofz)
If (ccx >= 0 And ccx < Picture2.ScaleWidth And ccz >= 0 And ccz < Picture2.ScaleHeight) Then Picture2.PSet (ccx, ccz) Else GoTo xzend
'Picture2.PSet (ccx, ccz)
'Print ccx, ccz
l2x(xztag) = xp: lz(xztag) = zp
xzend:
Picture2.DrawWidth = 1


End Sub

Private Sub xyDotPlot(xytag%)

Picture1.DrawWidth = 7
Picture1.ForeColor = QBColor(15)
Picture1.FillColor = QBColor(15)
ccx = s1x * (lx(xytag) - ofx)
ccy = s1y * (ly(xytag) - ofy)

'Print ccx, ccz, Picture2.ScaleWidth, Picture2.ScaleHeight


If (ccx >= 0 And ccx < Picture1.ScaleWidth And ccy >= 0 And ccy < Picture1.ScaleHeight) Then Picture1.PSet (ccx, ccy)
xu = xytag And 15
If xu = 15 Then xu = 0
Picture1.ForeColor = QBColor(xu)
Picture1.FillColor = QBColor(xu)
ccx = s1x * (xp - ofx)
ccy = s1y * (yp - ofy)
If (ccx >= 0 And ccx < Picture1.ScaleWidth And ccy >= 0 And ccy < Picture1.ScaleHeight) Then Picture1.PSet (ccx, ccy) Else GoTo xyend
'Picture1.PSet (s1x * (xp + ofx), s1y * (yp + ofy))
lx(xytag) = xp: ly(xytag) = yp
xyend:
Picture1.DrawWidth = 1
End Sub

Private Sub dogrid()
Picture2.ForeColor = QBColor(0)
Picture2.FillColor = QBColor(0)
Text5.ZOrder 0

Picture2.DrawStyle = 2
ss = Picture2.ScaleWidth / 10
For q = 1 To 9
Picture2.Line (ss * q, 0)-(ss * q, Picture2.ScaleHeight)
Next
ss = Picture2.ScaleHeight / 10
For q = 1 To 9
Picture2.Line (0, ss * q)-(Picture2.ScaleWidth, ss * q)
Next
Picture2.DrawStyle = 0

  Picture1.ForeColor = QBColor(0)
  Picture1.FillColor = QBColor(0)
  Text4.ZOrder 0
  Picture1.DrawStyle = 2
  ss = Picture1.ScaleWidth / 100
  For q = 1 To 99
   Picture1.Line (ss * q, 0)-(ss * q, Picture1.ScaleHeight)
  Next
  ss = Picture1.ScaleHeight / 100
  For q = 1 To 99
   Picture1.Line (0, ss * q)-(Picture1.ScaleWidth, ss * q)
  Next
  Picture1.DrawStyle = 0

End Sub


Private Sub refreshMap()
  Picture1.DrawWidth = 3
  Picture2.DrawWidth = 3
For i = 0 To norx - 1
  mx = xx(rxIds(i)): my = yy(rxIds(i)): mz = zz(rxIds(i))
  xs1 = s1x * (mx - ofx): ys1 = s1y * (my - ofy): xs2 = s2x * (mx - ofx): zs1 = s1z * (mz - ofz)
  Picture1.Line (xs1, ys1)-(xs1 + 100, ys1 + 100), , B
  Picture2.Line (xs2, zs1)-(xs2 + 100, zs1 + 100), , B
Next
  Picture1.DrawWidth = 1
  Picture2.DrawWidth = 1

End Sub
Private Sub exponentialFilter(tag)
Dim xxl, yyl, zzl
xxl = xst(tag)
xst(tag) = alpha * xp + (1 - alpha) * (xst(tag) + xbt(tag))
xbt(tag) = beta * (xst(tag) - xxl) + (1 - beta) * xbt(tag)
xf(tag) = xst(tag) + xbt(tag)

yyl = yst(tag)
yst(tag) = alpha * yp + (1 - alpha) * (yst(tag) + ybt(tag))
ybt(tag) = beta * (yst(tag) - yyl) + (1 - beta) * ybt(tag)
yf(tag) = yst(tag) + ybt(tag)

zzl = zst(tag)
zst(tag) = alpha * zp + (1 - alpha) * (zst(tag) + zbt(tag))
zbt(tag) = beta * (zst(tag) - zzl) + (1 - beta) * zbt(tag)
zf(tag) = zst(tag) + zbt(tag)

xp = xst(tag): yp = yst(tag): zp = zst(tag)
End Sub


