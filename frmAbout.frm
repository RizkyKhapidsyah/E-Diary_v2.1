VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "xtremeprohacker@yahoo.com.sg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Hacker's Technology Corp. Softwares"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Lim Meng Huey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "This Software Is Proudly Presented By : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
End Sub
