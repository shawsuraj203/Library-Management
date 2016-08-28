VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Library Management"
   ClientHeight    =   4800
   ClientLeft      =   5910
   ClientTop       =   3270
   ClientWidth     =   9525
   Icon            =   "1st.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   Picture         =   "1st.frx":014A
   ScaleHeight     =   4800
   ScaleWidth      =   9525
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\study\vb\lib.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "profile"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Create a &New Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "New User?"
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton cmdlog 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Log In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Log In"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox txtuname 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   4
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   7440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      DataField       =   "no of books"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblpass 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lbluname 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connect With Us"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3840
      TabIndex        =   1
      Top             =   1320
      Width           =   2385
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Library Management"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   6105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim result As Integer
Private Sub cmdexit_Click()
result = MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion + vbDefaultButton2 + vbSystemModal, "Exit?")
If result = 6 Then
Unload Me
ElseIf result = 7 Then

End If
End Sub

Private Sub cmdlog_Click()

f = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If Data1.Recordset.Fields("User") = txtuname.Text Then
If txtpass.Text = Data1.Recordset.Fields("password") Then
If txtuname = "admin" Then
Label1.Caption = ""
Else
Label1.Caption = Data1.Recordset.Fields("no of books")
End If
f = 1
Label2.Caption = f

Form7.Show
Unload Me

Else
MsgBox ("Incorrect Password!!")
End If
Exit Do
Else
Data1.Recordset.MoveNext
End If
Loop
If f = 0 Then
MsgBox ("Create a new account")
End If

End Sub

Private Sub cmdnew_Click()
Form2.Show
Form2.cmd.Value = True
Unload Me

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path + "/lib.mdb"
End Sub

