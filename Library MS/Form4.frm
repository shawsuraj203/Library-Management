VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6600
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10050
   LinkTopic       =   "Form4"
   ScaleHeight     =   6600
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   109838337
      CurrentDate     =   42121
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   109838337
      CurrentDate     =   42121
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1695
      Left            =   2280
      TabIndex        =   1
      Top             =   4680
      Width           =   4455
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2895
      Left            =   15720
      TabIndex        =   0
      Top             =   1200
      Width           =   255
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4080
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2760
      Width           =   5175
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "Window"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

