VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "hx19access"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "hx19v2Access.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   14190
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "X"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Height          =   5535
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Broadcast hx19setup.txt"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Log"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Text            =   "Send String>"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   390
      Left            =   2280
      TabIndex        =   5
      Text            =   "String Size>"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sync Mode"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TX"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   975
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
      Height          =   6015
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   8175
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5040
      Top             =   2160
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
      Top             =   600
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim linebuffer(100) As String            'only used for the scrolling routine, not linked to hx19 operations
Dim nn%
Dim engage As Boolean
Dim haltDisplay As Boolean

Private Sub Command2_Click()
Dim conf$
Open "hx19setup.txt" For Input As 1
Do
Input #1, conf
If Len(conf) < 4 Then GoTo sdone
Text7 = Text7 + conf + vbCrLf
checkOut conf
   Do: DoEvents: cc = Form1.MSComm1.Input: Loop Until cc = "k" Or cc = "#"
   Do: DoEvents: cc = Form1.MSComm1.Input: Loop Until cc = chr13
Loop Until EOF(1)
sdone:
Text7 = Text7 + "DONE" + vbCrLf
Close
End Sub

Private Sub Form_Activate()
Dim cc As String

Text4.Visible = False
Command1.Visible = False
Check1.Visible = True

'The following code examines the com port selected to see if it is open and the device available
'if there is an error, it gives the user a chance to enter in the correct port number
'use Control Panel to determine which port is being used for the hx19 system
'this code section is not important for using the hx19
    On Error GoTo fixit
        Open "port.txt" For Input As 1
         Input #1, Port
        Close 1
        GoTo allOK
fixit:
        Check1.Visible = False
        Text4.Visible = True
        Command1.Visible = True

Text6 = "Type port number into the blue window above and click OK (find the correct port number under Device Manager)"
        
        Do: DoEvents: Loop Until engage = True
        Open "port.txt" For Output As 1
            Print #1, Val(Text4)
        Close 1
        Text4.Visible = False
        Command1.Visible = False
        Check1.Visible = True
        Port = Val(Text4)
allOK:
Text6 = ""
'END PORT NUMBER CHECK --------------------------------------------------------------------------

'The following code segment is the heart of the interface to the hx19ms

    Form1.MSComm1.CommPort = Port       'this port number must corrispond to the hx19ms usb port
    Form1.MSComm1.Settings = "256000,N,8,1" '256Kbaud, no parity, 8bit interchange and 1 stop bit
    Form1.MSComm1.PortOpen = True
    Form1.MSComm1.InputLen = 1                  'get serial characters one by one
    Form1.MSComm1.InBufferCount = 0             'make sure there is no data residue from last application in buffers
    Form1.MSComm1.OutBufferCount = 0            'make sure output buffer is empty
    
ticnt = 0
 Do
    Do:                                                 'loop until there's a character in input buffer
     DoEvents
    Loop While Form1.MSComm1.InBufferCount < 1
                    'using MSComm1 control component supplied with visual basic 6
    cc = Form1.MSComm1.Input   'input a character from hx19ms Mscomm1
      
     If cc = Chr(13) Then                           'if it is a data delimiter (carriage return) then display line
       If Check3.Value = 0 Then tScroll comline + vbCrLf   'tScroll scrolls display down one line
       If Check2.Value = 1 Then Print #1, comline       'save data on file 1 when log option is selected
       comline = ""                                 'clear this data line and prepare for then next
      Else
       comline = comline + cc                       'while there's no delimiter collect data into a line
       If Check3.Value = 0 Then Text3 = Text3 + cc  'accumulate data into the text to be scrolled
     End If
 Loop                   'continuous loop

End Sub

 
Private Sub Check1_Click()
    'Any hx19ms receiving the command $ will initiate syncronized strobing
    'the command % stops the sync sequence
    If Check1.Value = 1 Then checkOut "M&$" Else checkOut "M&%"
End Sub

Private Sub Command3_Click()
    checkOut Text6                          'here the content of Text6 is transmitted via usb com port to the hx19ms
End Sub

Private Sub checkOut(temst As String)
Dim xsum As Integer, xx As String
    'this routine sums up all Ascii characters entered, and creates an hx19 accepted checksum.
     xsum = 0
     For i = 1 To Len(temst)                'compute the checksum of the string
       xx = Mid(temst, i, 1)
       xsum = xsum + Asc(xx)                    'accumulate ASCII codes
    Next
    temst = temst + "/" + Hex(xsum)         'append the checksum in hexadecimal format
    Form1.MSComm1.Output = temst + Chr(13)  'sends a string to the hx19ms with the correct checksum attached
End Sub

Private Sub Text6_Change()
    'This routine keeps track of the total characters entered, should not exeed 116 characters.
    Text5 = Format(Len(Text6), "#")
End Sub

'Remaining routines are unimportant and secondary to the understanding of the hx19

Private Sub Check2_Click()
    'data coming from the hx19ms is text format and can be viewed using windows notepad
     If Check2.Value = 1 Then
        Open "hx19access.txt" For Output As 1             'save incoming hx19ms text data on file called hx19.log
      Else
       Close 1
     End If
End Sub

Private Sub Command1_Click()
     engage = True              'in case of a com error, user may input correction and then continue
End Sub

Private Sub tScroll(nline)                      'scrolls 16 lines of text through text window
Dim jj%
'routine scrolls down one line for display purposes only, it has otherwise nothing to do with hx19 system
    linebuffer(nn) = nline
    nn = (nn + 1) And 15
    jj = nn + 1
    Text3 = ""
    Do
     Text3 = Text3 + linebuffer(jj)
     jj = (jj + 1) And 15
    Loop Until jj = nn
End Sub

Private Sub Form_Terminate()
    Close                                       'make sure no files are left open when the program ends
    End                                             'if end isn't executed before termination, com port may be left open
End Sub

Private Sub Form_Unload(Cancel As Integer)
'to be absolutely sure
    Close                                           'make sure no files are left open when the program ends
    End                                             'if end isn't executed before termination, com port may be left open
End Sub


