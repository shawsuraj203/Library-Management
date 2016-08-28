VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Book Details"
   ClientHeight    =   8280
   ClientLeft      =   5850
   ClientTop       =   2205
   ClientWidth     =   9990
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "book.frx":0000
   LinkTopic       =   "Form3"
   Picture         =   "book.frx":014A
   ScaleHeight     =   8280
   ScaleWidth      =   9990
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "D:\study\vb\lib.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "profile"
      Top             =   8640
      Width           =   1140
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      TabIndex        =   25
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox txtsearch 
      Height          =   735
      Left            =   1800
      TabIndex        =   24
      Top             =   6720
      Width           =   4095
   End
   Begin VB.CommandButton cmdlend 
      Caption         =   "Issue"
      Height          =   495
      Left            =   3840
      MousePointer    =   2  'Cross
      TabIndex        =   23
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\study\vb\lib.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "lend"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\study\vb\lib.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Book"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdprev 
      Caption         =   "Prev"
      Height          =   495
      Left            =   1800
      TabIndex        =   21
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next"
      Height          =   495
      Left            =   5880
      TabIndex        =   20
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete"
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   3960
      TabIndex        =   18
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      Height          =   495
      Left            =   1800
      TabIndex        =   17
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdbrws 
      Caption         =   "Browse...."
      Height          =   495
      Left            =   11520
      TabIndex        =   16
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox txtav 
      DataField       =   "Availability"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   4320
      Width           =   3855
   End
   Begin VB.TextBox txtpr 
      DataField       =   "Price"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      Top             =   4920
      Width           =   3855
   End
   Begin VB.TextBox txtrack 
      DataField       =   "Rack No"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Top             =   3720
      Width           =   3855
   End
   Begin VB.TextBox txtnb 
      DataField       =   "No of Books"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   12
      Top             =   3120
      Width           =   3855
   End
   Begin VB.TextBox txtedt 
      DataField       =   "Edition"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox txtpub 
      DataField       =   "Publisher"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox txtauthor 
      DataField       =   "Author"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox txtname 
      DataField       =   "Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label lblnbl 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9240
      TabIndex        =   27
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblnb 
      BackStyle       =   0  'Transparent
      Caption         =   "No of Books Taken:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   26
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   375
      Index           =   8
      Left            =   2760
      TabIndex        =   22
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   10800
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lbl 
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "Author:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1680
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "Availability :"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   5
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "Rack No."
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1680
      TabIndex        =   4
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "No of Books:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   3
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "Edition:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "Publisher"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lbl 
      Caption         =   "Book Name:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Menu Fille 
      Caption         =   "File"
      Begin VB.Menu Print 
         Caption         =   "Print"
      End
      Begin VB.Menu Reload 
         Caption         =   "Reload"
      End
      Begin VB.Menu sperate 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Cut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu Copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu Back 
      Caption         =   "Back"
   End
   Begin VB.Menu logout 
      Caption         =   "Log out"
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Back_Click()
Form7.Show
End Sub

Private Sub cmdadd_Click()
Data1.Recordset.AddNew
End Sub

Private Sub cmddel_Click()
Data1.Recordset.Delete
End Sub

Private Sub cmdedit_Click()
Data1.Recordset.Edit
Data1.Recordset.Update
End Sub

Private Sub cmdlend_Click()
If lblnbl.Caption = "10" Then
MsgBox ("You are exceeding to maximum allowed no of Books")
cmdlend.Visible = False
Else
Data2.Recordset.AddNew
Data2.Recordset.Fields("user") = lbl(8).Caption
Data2.Recordset.Fields("name") = txtname.Text
Data2.Recordset.Fields("author") = txtauthor.Text
Data2.Recordset.Fields("pub") = txtpub.Text
Data2.Recordset.Fields("edition") = txtedt.Text
Data2.Recordset.Fields("price") = txtpr.Text
lblnbl.Caption = Val(lblnbl.Caption) + 1
Data2.Recordset.Update
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
If Data3.Recordset.Fields("User") = lbl(8).Caption Then
Data3.Recordset.Edit
Data3.Recordset.Fields("no of books") = Val(Data3.Recordset.Fields("no of books")) + 1
Data3.Recordset.Update
Exit Do
Else
Data3.Recordset.MoveNext
End If
Loop
End If
End Sub

Private Sub cmdnext_Click()
If txtname.Text = "" Then
Data1.Recordset.MoveFirst
Else
Data1.Recordset.MoveNext
End If
End Sub

Private Sub cmdprev_Click()
If txtname.Text = "" Then
Data1.Recordset.MoveLast
Else
Data1.Recordset.MovePrevious
End If

End Sub

Private Sub cmdsearch_Click()
Dim f As Integer
f = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If Data1.Recordset.Fields("Name") = txtsearch.Text Then
txtname.Text = Data1.Recordset.Fields("Name")
txtauthor.Text = Data1.Recordset.Fields("Author")
txtpub.Text = Data1.Recordset.Fields("Publisher")
txtedt.Text = Data1.Recordset.Fields("Edition")
txtnb.Text = Data1.Recordset.Fields("No of Books")
txtrack.Text = Data1.Recordset.Fields("Rack No")
txtav.Text = Data1.Recordset.Fields("Availability")
txtpr.Text = Data1.Recordset.Fields("Price")
f = 1
Exit Do
Else
Data1.Recordset.MoveNext
End If
Loop
If f = 0 Then
x = MsgBox("Book not Found!!", vbExclamation + vbOKOnly, "Search Result")
End If
End Sub

Private Sub Command1_Click()
y = MsgBox("Are you sure you want to Log Out?", vbQuestion + vbYesNo + vbDefaultButton2, "Log Out!")
If y = 6 Then
Form1.Show
Unload Me
End If
End Sub

Private Sub Form_Load()
If Form7.Caption = "admin" Then
cmdadd.Visible = True
cmdedit.Visible = True
cmddel.Visible = True
cmdlend.Visible = False
lblnb.Visible = False
lblnbl.Visible = False
Else
cmdadd.Visible = False
cmdedit.Visible = False
cmddel.Visible = False
End If

lbl(8).Caption = Form7.Caption
lblnbl.Caption = Form7.Label1.Caption

For i = 0 To 7
lbl(i).BackStyle = 0
lbl(i).ForeColor = &HFF&

Next
Data1.DatabaseName = App.Path + "/lib.mdb"
Data2.DatabaseName = App.Path + "/lib.mdb"
Data3.DatabaseName = App.Path + "/lib.mdb"
Unload Form1


End Sub



Private Sub Help_Click()
MsgBox ("This is Our College Library Management Software. For More details about this software or our college go to this link http://www.nsec.ac.in Thank You!!")
End Sub

Private Sub logout_Click()
y = MsgBox("Are you sure you want to Log Out?", vbQuestion + vbYesNo + vbDefaultButton2, "Log Out!")
If y = 6 Then
Form1.Show
Unload Me
End If
End Sub

