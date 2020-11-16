VERSION 5.00
Begin VB.Form frmFind 
   Caption         =   "Find Entry"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   5235
   Begin VB.TextBox txtEntry 
      Enabled         =   0   'False
      Height          =   1815
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2400
      Width           =   4935
   End
   Begin VB.ComboBox cboWeather 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   10
      Text            =   "Select Weather"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New Search"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Text            =   "Select Date"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Date :"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Weather :"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblTime 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblDate 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Time :"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Select Date :"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    lblDate.Enabled = False
    lblTime.Enabled = False
    cboWeather.Enabled = False
    txtEntry.Enabled = False
    cboDate.Text = "Select Date"
    cmdSearch.Caption = "Search"
End Sub

Private Sub cmdSearch_Click()
    If cmdSearch.Caption = "Search" Then
        dDate = cboDate.Text
        Set Rs = New ADODB.Recordset
        Rs.Open "Select * From tEntry Where UserName = '" & UserName & "' And Date = '" & dDate & "'", Cn, 1, 4
        If Rs.RecordCount > 0 Then
            lblDate.Enabled = True
            lblTime.Enabled = True
            cboWeather.Enabled = True
            txtEntry.Enabled = True
            lblDate.Caption = dDate
            lblTime.Caption = Rs("Time")
            cboWeather.Text = Rs("Weather")
            txtEntry.Text = Rs("Entry")
            cmdSearch.Caption = "Edit"
        Else
            MsgBox "No Entry Found."
            cboDate.Text = "Select Date"
        End If
    Else
        Rs("Weather") = cboWeather.Text
        Rs("Entry") = txtEntry.Text
        Rs.Update
        Rs.UpdateBatch
        Rs.MoveNext
        cmdNew.Value = True
    End If
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
        
    cboWeather.AddItem "Cloudy"
    cboWeather.AddItem "Rainy"
    cboWeather.AddItem "Sunny"
    cboWeather.AddItem "Windy"
End Sub

