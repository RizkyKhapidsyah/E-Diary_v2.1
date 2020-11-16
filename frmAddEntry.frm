VERSION 5.00
Begin VB.Form frmAddEntry 
   Caption         =   "Add Entry"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   Icon            =   "frmAddEntry.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   5310
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtEntry 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   5055
   End
   Begin VB.ComboBox cboWeather 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Text            =   "Select Weather"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Weather :"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblTime 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Time :"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Date :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * From tEntry", Cn, 1, 4
    Rs.AddNew
    Rs("UserName") = UserName
    Rs("Date") = lblDate.Caption
    Rs("Time") = lblTime.Caption
    Rs("Weather") = cboWeather.Text
    Rs("Entry") = txtEntry.Text
    Rs.Update
    Rs.UpdateBatch
    Rs.MoveNext
    Rs.Close
    Set Rs = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    lblDate.Caption = Date
    lblTime.Caption = Time
    cboWeather.AddItem "Cloudy"
    cboWeather.AddItem "Rainy"
    cboWeather.AddItem "Sunny"
    cboWeather.AddItem "Windy"
    Connection
End Sub
