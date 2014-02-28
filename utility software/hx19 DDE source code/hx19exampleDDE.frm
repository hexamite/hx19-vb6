VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hx17 Hexamite Ltd"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   FillStyle       =   0  'Solid
   Icon            =   "hx19exampleDDE.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   1  'Arrow
   ScaleHeight     =   4080
   ScaleWidth      =   4095
   Begin VB.CheckBox Check1 
      Caption         =   "Freeze"
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "hx19exampleDDE.frx":1CCA
      Top             =   360
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this program works with realtime data from the hx19 system
Dim linebuffer(100)
Dim nn As Integer
Private Sub text1_Change()
If Check1.Value = 0 Then tScroll Text1 + vbCrLf         'everytime the linktopic of hx19xyz changes text1, this program jumps to this subroutine
End Sub

Private Sub Form_Load()
Text1.LinkTopic = "HX19|hx19"    'hx19 is the title for the program hx19xyzDDE fetching data from the hx19 controller
Text1.LinkItem = "text1"         'DDE data from hx19xyzDDE is directed through text1 of this program
'the program hx19xyzDDE must be running on the computer for this link to become active
Text1.LinkMode = 1               'this sets the link process active
End Sub

Private Sub tScroll(nline)       'scrolls 16 lines of text
Dim jj%

linebuffer(nn) = nline
nn = (nn + 1) And 15
jj = nn + 1
Text2 = ""
Do
 Text2 = Text2 + linebuffer(jj)
 jj = (jj + 1) And 15
Loop Until jj = nn

End Sub


