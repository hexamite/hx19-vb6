VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Hexamite Hx19TX_comPin"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "hx19comPin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "Status"
      Top             =   2880
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ComPin TX"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1335
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
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   4815
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9960
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text6 
      Height          =   360
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Log"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim port%
Dim engage As Boolean


Private Sub Form_Activate()
Dim cc As String

Text4.Visible = False
Command1.Visible = False
Check2.Visible = True

checkComportAvailability

Text6 = ""
'The following code segment is the heart of the interface to the hx19ms

    Form1.MSComm1.CommPort = port                   'this port number must corrispond to the hx19ms usb port
    Form1.MSComm1.Settings = "256000,N,8,1"         '256Kbaud, no parity, 8bit interchange and 1 stop bit
    Form1.MSComm1.PortOpen = True
    Form1.MSComm1.InputLen = 1                      'get serial characters one by one
    Form1.MSComm1.InBufferCount = 0                 'make sure there is no data residue from last application in buffers
    Form1.MSComm1.OutBufferCount = 0                'make sure output buffer is empty
    
 Do
    Do                                             'loop until there's a character in input buffer
     DoEvents
    Loop While Form1.MSComm1.InBufferCount < 1
     cc = Form1.MSComm1.Input
     If cc = Chr(13) Then cc = vbCrLf
     Text3 = Text3 + cc
 Loop                                               'continuous loop

End Sub

Private Sub Command3_Click()
Text2 = "No Response": Text3 = ""
    Form1.MSComm1.Output = "?"                      'sends a string to the hx19ms with the correct checksum attached
    Do: DoEvents: p = Form1.MSComm1.Input: Loop Until p = "*" Or p = "<"
     Text3 = p                                 'accumulate data into the text to be scrolled
    If p = "<" Then
     Form1.MSComm1.Output = "?"                     'sends a string to the hx19ms with the correct checksum attached
     Do: DoEvents: p = Form1.MSComm1.Input: Loop Until p = "*"
     Form1.MSComm1.Output = Text6 + Chr(13)
     Text3 = Text3 + p
    Else
     If p = "*" Then Form1.MSComm1.Output = Text6 + Chr(13)
    End If
Text2 = "String Sent through ComPin"
End Sub


Private Sub Text3_DblClick()
Text3 = ""
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

Private Sub Form_Terminate()
    Close                                       'make sure no files are left open when the program ends
    End                                             'if end isn't executed before termination, com port may be left open
End Sub

Private Sub Form_Unload(Cancel As Integer)
'to be absolutely sure
    Close                                           'make sure no files are left open when the program ends
    End                                             'if end isn't executed before termination, com port may be left open
End Sub

Private Sub checkComportAvailability()
'The following code examines the com port selected to see if it is open and the device available
'if there is an error, it gives the user a chance to enter in the correct port number
'use Control Panel to determine which port is being used for the hx19 system
'this code section is not important for using the hx19
    On Error GoTo fixit
        Open "comPinPort.txt" For Input As 1
         Input #1, port
        Close 1
        Exit Sub
fixit:
        Check2.Visible = False
        Text4.Visible = True
        Command1.Visible = True

        Text6 = "Type port number into the blue window above and click OK (find the correct port number under Device Manager)"
        
        Do: DoEvents: Loop Until engage = True
        Open "comPinPort.txt" For Output As 1
            Print #1, Val(Text4)
        Close 1
        Text4.Visible = False
        Command1.Visible = False
        Check2.Visible = True
        port = Val(Text4)
'END PORT NUMBER CHECK --------------------------------------------------------------------------
End Sub

