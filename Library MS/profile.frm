VERSION 5.00
Begin VB.Form Form2 
   Caption         =   " v"
   ClientHeight    =   10200
   ClientLeft      =   6270
   ClientTop       =   675
   ClientWidth     =   9405
   Icon            =   "profile.frx":0000
   LinkTopic       =   "Form2"
   MousePointer    =   2  'Cross
   Picture         =   "profile.frx":014A
   ScaleHeight     =   10200
   ScaleWidth      =   9405
   Begin VB.CommandButton Command2 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   28
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   27
      Top             =   9360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\study\vb\lib.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   12000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "profile"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List1 
      DataField       =   "year"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "profile.frx":187A6
      Left            =   4920
      List            =   "profile.frx":187B6
      TabIndex        =   24
      Top             =   2520
      Width           =   3615
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   22
      Top             =   9360
      Width           =   1575
   End
   Begin VB.TextBox txtin 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   10
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   21
      Top             =   8160
      Width           =   3615
   End
   Begin VB.TextBox txtin 
      DataField       =   "password"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   9
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   20
      Top             =   7320
      Width           =   3615
   End
   Begin VB.TextBox txtin 
      DataField       =   "add"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   8
      Left            =   4920
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   5760
      Width           =   3615
   End
   Begin VB.TextBox txtin 
      DataField       =   "mob"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   4920
      MaxLength       =   10
      TabIndex        =   18
      Top             =   5160
      Width           =   3615
   End
   Begin VB.TextBox txtin 
      DataField       =   "email"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   4920
      TabIndex        =   17
      Top             =   4560
      Width           =   3615
   End
   Begin VB.TextBox txtin 
      DataField       =   "roll"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4920
      TabIndex        =   16
      Top             =   3840
      Width           =   3615
   End
   Begin VB.TextBox txtin 
      DataField       =   "user"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4920
      TabIndex        =   15
      Top             =   3120
      Width           =   3615
   End
   Begin VB.TextBox txtin 
      DataField       =   "stream"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4920
      TabIndex        =   14
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox txtin 
      DataField       =   "course"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4920
      TabIndex        =   13
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txtin 
      DataField       =   "name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4920
      TabIndex        =   12
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   135
      Left            =   11640
      TabIndex        =   26
      Top             =   6120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Your enrollment no will be your username."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1080
      TabIndex        =   23
      Top             =   9000
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ReEnter Password:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   1320
      TabIndex        =   11
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   9
      Left            =   1320
      TabIndex        =   10
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Address:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   1320
      TabIndex        =   9
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mob No:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1320
      TabIndex        =   8
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Id:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   7
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "University Roll No:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   1320
      TabIndex        =   6
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enrollment No:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   5
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stream:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Course:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "Basic Details"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   3600
      TabIndex        =   0
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim f As Integer

Private Sub cmd_Click()
Data1.Recordset.AddNew
Data1.Recordset.Fields("no of books") = "0"
End Sub

Private Sub cmdclear_Click()
For i = 0 To 10
txtin(i).Text = ""
If i = 2 Then
i = i + 1
End If

Next i


End Sub

Private Sub cmdsave_Click()
z = 0
w = 0
j = 0

For i = 0 To 10
If txtin(i).Text = "" Then
z = 1
Else
If txtin(9).Text = txtin(10).Text Then
Form1.Show

y = 1
Label2.Caption = y

Else
w = 1

End If
End If
If i = 2 Then
i = i + 1
End If
Next i
If z = 1 Then
MsgBox ("Complete the form first before moving")
End If
If w = 1 Then
MsgBox ("passwords didn't matched!!")
End If




End Sub



Private Sub Command1_Click()
Data1.Recordset.Edit
Data1.Recordset.Update
End Sub

Private Sub Command2_Click()
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If Data1.Recordset.Fields("user") = Form7.Caption Then
txtin(0).Text = Data1.Recordset.Fields("name")
txtin(1).Text = Data1.Recordset.Fields("course")
txtin(2).Text = Data1.Recordset.Fields("stream")

txtin(4).Text = Data1.Recordset.Fields("user")
txtin(5).Text = Data1.Recordset.Fields("roll")
txtin(6).Text = Data1.Recordset.Fields("email")
txtin(7).Text = Data1.Recordset.Fields("mob")
txtin(8).Text = Data1.Recordset.Fields("add")
txtin(9).Text = Data1.Recordset.Fields("password")

Exit Do
Else
Data1.Recordset.MoveNext
End If
Loop
End Sub

Private Sub Form_Load()
y = 0
MsgBox ("To Create new Account first click the add button Or to view or edit your profile first click the show button.")
Unload Form1
Data1.DatabaseName = App.Path + "/lib.mdb"

End Sub
