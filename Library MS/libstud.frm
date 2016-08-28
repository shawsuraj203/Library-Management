VERSION 5.00
Begin VB.Form stdbook 
   Caption         =   "Books Details"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11415
   FillColor       =   &H000000FF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "libstud.frx":0000
   LinkTopic       =   "Form5"
   MousePointer    =   2  'Cross
   Picture         =   "libstud.frx":014A
   ScaleHeight     =   8400
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton block 
      Caption         =   "Block"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   9480
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SHOW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   18
      Top             =   7680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "D:\study\vb\lib.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "lend"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\study\vb\lib.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "profile"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   17
      Top             =   7680
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PREV"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   16
      Top             =   7680
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      DataField       =   "price"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   735
      Index           =   7
      Left            =   6000
      TabIndex        =   10
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      DataField       =   "edition"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   735
      Index           =   3
      Left            =   6000
      TabIndex        =   9
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      DataField       =   "pub"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   735
      Index           =   2
      Left            =   6000
      TabIndex        =   8
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      DataField       =   "author"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   6000
      TabIndex        =   7
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      DataField       =   "name"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   6000
      TabIndex        =   6
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      DataField       =   "no of books"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7560
      TabIndex        =   5
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      DataField       =   "user"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      DataField       =   "name"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "iSSUED boOKS DETAILS"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   2040
      TabIndex        =   19
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   5
      Left            =   3360
      TabIndex        =   15
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Publisher"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Index           =   4
      Left            =   3600
      TabIndex        =   14
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edition"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Index           =   3
      Left            =   3480
      TabIndex        =   13
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Index           =   1
      Left            =   3600
      TabIndex        =   12
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Index           =   0
      Left            =   3600
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of books taken"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   7800
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enrollment no."
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student's Name:"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "stdbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub

Private Sub block_Click()
Data1.Recordset.Edit
Data1.Recordset.Update

End Sub

Private Sub Command1_Click()
Data2.Recordset.MovePrevious
Do While Not Data2.Recordset.EOF
If Data2.Recordset.Fields("user") = Text5.Text Then
Text4(0).Text = Data2.Recordset.Fields("name")
Text4(1).Text = Data2.Recordset.Fields("author")
Text4(2).Text = Data2.Recordset.Fields("pub")
Text4(3).Text = Data2.Recordset.Fields("edition")
Text4(7).Text = Data2.Recordset.Fields("price")
Exit Do
Else
Data2.Recordset.MovePrevious
End If
Loop


End Sub

Private Sub Command2_Click()
Data2.Recordset.MoveNext
Do While Not Data2.Recordset.EOF
If Data2.Recordset.Fields("user") = Text5.Text Then
Text4(0).Text = Data2.Recordset.Fields("name")
Text4(1).Text = Data2.Recordset.Fields("author")
Text4(2).Text = Data2.Recordset.Fields("pub")
Text4(3).Text = Data2.Recordset.Fields("edition")
Text4(7).Text = Data2.Recordset.Fields("price")
Exit Do
Else
Data2.Recordset.MoveNext
End If
Loop
If Text4(0).Text = "" Then
MsgBox ("No more books")

Command1.Value = True
End If
End Sub

Private Sub Command3_Click()
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If Data1.Recordset.Fields("user") = Text5.Text Then
Text1.Text = Data1.Recordset.Fields("name")
Text2.Text = Data1.Recordset.Fields("user")
Text3.Text = Data1.Recordset.Fields("no of books")
Exit Do
Else
Data1.Recordset.MoveNext
End If
Loop
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
If Data2.Recordset.Fields("user") = Text5.Text Then
Text4(0).Text = Data2.Recordset.Fields("name")
Text4(1).Text = Data2.Recordset.Fields("author")
Text4(2).Text = Data2.Recordset.Fields("pub")
Text4(3).Text = Data2.Recordset.Fields("edition")
Text4(7).Text = Data2.Recordset.Fields("price")
Exit Do
Else
Data2.Recordset.MoveNext
End If
Loop



End Sub

Private Sub Form_Load()
If Form7.Caption = "admin" Then
Text5.Text = Form7.Text1.Text
Text3.Enabled = True
block.Visible = True

Else
Text5.Text = Form7.Caption
End If
Data1.DatabaseName = App.Path + "/lib.mdb"
Data2.DatabaseName = App.Path + "/lib.mdb"

End Sub

