VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   5295
   ClientLeft      =   6945
   ClientTop       =   2955
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Papyrus"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "traverse.frx":0000
   LinkTopic       =   "Form7"
   Picture         =   "traverse.frx":014A
   ScaleHeight     =   5295
   ScaleWidth      =   7455
   Begin VB.TextBox Text1 
      Height          =   570
      Left            =   120
      TabIndex        =   7
      Text            =   "Search"
      Top             =   4440
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      MousePointer    =   2  'Cross
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Reissue Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Return Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Issue Boooks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Book Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Profile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   135
      Left            =   7800
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Form2.Command1.Visible = True
Form2.cmd.Visible = False

Form2.txtin(4).Enabled = False
Form2.Command2.Value = True


End Sub

Private Sub Command2_Click()
stdbook.Show
stdbook.Command3.Value = True

End Sub

Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Command6_Click()
y = MsgBox("Are you sure you want to Log Out?", vbQuestion + vbYesNo + vbDefaultButton2, "Log Out!")
If y = 6 Then
Form1.Show
Unload Me
End If
End Sub

Private Sub Command7_Click()
MsgBox ("This is Our College Library Management Software. For More details about this software or our college go to this link http://www.nsec.ac.in Thank You!!")
End Sub

Private Sub Form_Load()
Form7.Caption = Form1.txtuname.Text
Label1.Caption = Form1.Label1.Caption
If Form1.txtuname.Text = "admin" Then
Text1.Visible = True
Unload Form1

End If
End Sub

