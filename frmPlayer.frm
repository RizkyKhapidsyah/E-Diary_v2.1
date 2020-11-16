VERSION 5.00
Begin VB.Form frmPlayer 
   Caption         =   "MP3 Player "
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   6510
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrPosition 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1380
      Top             =   5370
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   450
      Left            =   4965
      TabIndex        =   7
      Top             =   5370
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Player"
      Height          =   5040
      Left            =   2385
      TabIndex        =   6
      Top             =   120
      Width           =   4005
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1980
         TabIndex        =   9
         Top             =   2355
         Width           =   1515
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   435
         TabIndex        =   8
         Top             =   2370
         Width           =   1500
      End
      Begin VB.Label lblSecond 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1965
         TabIndex        =   15
         Top             =   3465
         Width           =   975
      End
      Begin VB.Shape Shape7 
         Height          =   720
         Left            =   1965
         Top             =   3240
         Width           =   990
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Sec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   2235
         TabIndex        =   14
         Top             =   2955
         Width           =   660
      End
      Begin VB.Label lblMinute 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   915
         TabIndex        =   13
         Top             =   3465
         Width           =   975
      End
      Begin VB.Shape Shape5 
         Height          =   720
         Left            =   915
         Top             =   3240
         Width           =   990
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   375
         TabIndex        =   11
         Top             =   1800
         Width           =   3165
      End
      Begin VB.Label lblSongName 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1125
         Left            =   465
         TabIndex        =   10
         Top             =   465
         Width           =   2910
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   1230
         TabIndex        =   12
         Top             =   2955
         Width           =   660
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   915
         Top             =   2925
         Width           =   1020
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   1965
         Top             =   2925
         Width           =   1020
      End
   End
   Begin VB.FileListBox myFile 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   150
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   2835
      Width           =   2145
   End
   Begin VB.DirListBox myDirectory 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   165
      TabIndex        =   1
      Top             =   1200
      Width           =   2115
   End
   Begin VB.DriveListBox myDrive 
      Height          =   315
      Left            =   165
      TabIndex        =   0
      Top             =   450
      Width           =   2130
   End
   Begin VB.PictureBox myMP3 
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   16
      Top             =   5325
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Media"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   285
      TabIndex        =   5
      Top             =   165
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   180
      Top             =   120
      Width           =   2130
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Folder(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   255
      TabIndex        =   4
      Top             =   900
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   165
      Top             =   870
      Width           =   2130
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   255
      TabIndex        =   3
      Top             =   2550
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   285
      Left            =   165
      Top             =   2520
      Width           =   2130
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SongName As String
Dim SecPosition As Integer
Dim Less As Integer
Dim Second As Integer
Dim Minute As Integer

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPlay_Click()

myMP3.OpenFile SongName
myMP3.Play
tmrPosition.Enabled = True

' Counter Start
Second = 0
Minute = 0

End Sub

Private Sub cmdStop_Click()

tmrPosition.Enabled = False
lblSecond.Caption = 0
myMP3.Stop

End Sub

Private Sub myDirectory_Change()

myFile.Path = myDirectory.Path

End Sub

Private Sub myDrive_Change()

myDirectory.Path = myDrive.Drive

End Sub

Private Sub myFile_Click()

myMP3.Stop
tmrPosition.Enabled = False

SongName = myFile.Path & "\" & myFile.FileName
lblSongName.Caption = myFile.FileName
lblInfo.Caption = myMP3.GetInfo

End Sub

Private Sub tmrPosition_Timer()


Second = Second + 1
lblSecond.Caption = Second

If Second = 59 Then
    Minute = Minute + 1
    lblMinute.Caption = Minute
    Second = 0
End If

End Sub
