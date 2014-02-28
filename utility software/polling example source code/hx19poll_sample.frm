VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Load To Flash"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "hx19poll_sample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Left            =   7920
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Left            =   7920
      Top             =   840
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   7575
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7800
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text6 
      Height          =   600
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim port%
Dim linebuffer(100)
Dim nn%
Dim sensorPollString(1000) As String
Dim sensors%, poll%, timeout1 As Boolean

Private Sub Form_Activate()
Dim hx19ms$, hx19rx$

Open "pollSetup.txt" For Input As 1
Input #1, hx19ms                    'get hx19ms setup string from file
Input #1, hx19rx                    'get hx19rx setup string from file
Input #1, port                      'get port number
Input #1, pollRate                  'get poll rate
Close 1
'****************************** set up the USB com port
Form1.MSComm1.CommPort = port
Form1.MSComm1.Settings = "256000,N,8,1"
Form1.MSComm1.InputLen = 0
Form1.MSComm1.PortOpen = True
    
Form1.MSComm1.InputLen = 0                                  'get serial characters one by one
Form1.MSComm1.InBufferCount = 0
Form1.MSComm1.OutBufferCount = 0
'---------------------------------------------------
sendString hx19rx                                           'send the content of the hx19ms string to the hx19ms
timeout1 = False
Timer2.Interval = 1000
Do: DoEvents: Loop Until timeout1 = True                    'wait 1 second

sendString hx19ms                                           'send the content of the hx19ms string to the hx19ms
timeout1 = False
Timer2.Interval = 1000
Do: DoEvents: Loop Until timeout1 = True                    'wait 1 second

cc = Form1.MSComm1.Input                                    'the hx19ms will respond to a correct string with acknowledge
Text6 = cc                                                  'display hx19ms acknowledgement

Open "receiverList.txt" For Input As 1
Do
Input #1, rx
sensorPollString(sensors) = rx
sensors = sensors + 1
Loop Until EOF(1)
Close

Form1.MSComm1.InputLen = 1                                  'get serial characters one by one
Form1.MSComm1.InBufferCount = 0                             'clear serial input buffers
Timer1.Interval = pollRate * 10                             'timer resolution set at 10 milliseconds

Open "hx19msLog.txt" For Output As 1
Do
    Do: DoEvents: Loop While Form1.MSComm1.InBufferCount < 1    'loop if there is no data in the serial buffers
    cc = Form1.MSComm1.Input                                    'get data characters one by one
    xin = xin + cc                                              'collect them into a buffer
  If cc = "#" Then
    Print #1, xin
    tScroll xin + vbCrLf
    xin = ""
  End If
Loop

End Sub

Private Sub Form_Terminate()
sendString "M1&%"
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
sendString "M1&%"
End
End Sub

Private Sub sendString(hx19$)

xsum = 0
Text6 = hx19
 xsum = 0
 temst = hx19
 For i = 1 To Len(temst)                                    'compute the checksum of the string
  xx = Mid(temst, i, 1)
  xsum = xsum + Asc(xx)                                     'accumulate ASCII codes
 Next
 temst = temst + "/" + Hex(xsum)                            'append the checksum in hexadecimal format
 Form1.MSComm1.Output = temst + Chr(13)
Do: DoEvents: Loop Until Form1.MSComm1.OutBufferCount = 0   'wait until the whole string has been sent

End Sub

Private Sub Timer1_Timer()
sendString (sensorPollString(poll))
poll = poll + 1
If poll = sensors Then poll = 0
End Sub

Private Sub tScroll(nline)       'scrolls 16 lines of text through textwindow
Dim jj%

linebuffer(nn) = nline
nn = (nn + 1) And 15
jj = nn + 1
Text3 = ""
Do
 Text3 = Text3 + linebuffer(jj)
 jj = (jj + 1) And 15
Loop Until jj = nn

End Sub

Private Sub Timer2_Timer()
timeout1 = True
End Sub
