VERSION 5.00
Begin VB.Form frmDelete 
   Caption         =   "Delete Entry"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   Icon            =   "frmDelete.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   5415
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Text            =   "Select Date"
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Select Date :"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Time :"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblDate 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblTime 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Weather :"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Date :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblWeather 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblEntry 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   4695
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
    Reply = MsgBox("Do You Realy Want To Delete This Entry?", 4, "Delete?")
    If Reply = vbYes Then
        Set Rs = New ADODB.Recordset
        Rs.Open "Delete * From tEntry Where UserName = '" & UserName & "' And Date = '" & dDate & "'", Cn, 1, 4
        Unload Me
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    dDate = cboDate.Text
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * From tEntry Where UserName = '" & UserName & "' And Date = '" & dDate & "'", Cn, 1, 4
    If Rs.RecordCount > 0 Then
        lblDate.Enabled = True
        lblTime.Enabled = True
        lblWeather.Enabled = True
        lblEntry.Enabled = True
        lblDate.Caption = dDate
        lblTime.Caption = Rs("Time")
        lblWeather.Caption = Rs("Weather")
        lblEntry.Caption = Rs("Entry")
        cmdDelete.Enabled = True
    Else
        MsgBox "No Entry Found."
        cboDate.Text = "Select Date"
    End If
    Rs.Close
    Set Rs = Nothing
End Sub

Private Sub Form_Load()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * From tEntry Where UserName ='" & UserName & "'", Cn, 1, 4

    Do Until Rs.EOF
        cboDate.AddItem Rs("Date")
        Rs.MoveNext
    Loop
    
    Rs.Close
    Set Rs = Nothing

End Sub
