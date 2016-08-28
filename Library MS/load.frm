VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form6 
   Caption         =   "Library Management System"
   ClientHeight    =   8115
   ClientLeft      =   4020
   ClientTop       =   1905
   ClientWidth     =   13455
   Icon            =   "load.frx":0000
   LinkTopic       =   "Form6"
   MousePointer    =   2  'Cross
   ScaleHeight     =   8115
   ScaleWidth      =   13455
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7800
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1560
      Top             =   8280
   End
   Begin VB.PictureBox Picture1 
      Height          =   7815
      Left            =   0
      Picture         =   "load.frx":014A
      ScaleHeight     =   7755
      ScaleWidth      =   13275
      TabIndex        =   0
      Top             =   120
      Width           =   13335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim strPath As String


Private Sub Timer1_Timer()
Timer1.Interval = Rnd * 300 + 10
ProgressBar1.Value = ProgressBar1.Value + 2
If ProgressBar1.Value = "100" Then
Form1.Show
Unload Me
End If
End Sub



   
