VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1035
   ScaleWidth      =   7320
   Begin VB.ComboBox cboUser 
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Text            =   "Select User Name"
      Top             =   120
      Width           =   3735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   390
      Left            =   840
      Top             =   2160
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   688
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Password :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Login :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDone_Click()
On Error GoTo Error1
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * From tUser Where UserName = '" & cboUser.Text & "'", Cn, 1, 4
    
    If txtPassword.Text = Rs("Password") Then
        UserName = cboUser.Text
        frmMain.mnuDiary.Enabled = True
        frmMain.mnuUserDetail = True
        Unload frmLogin
        frmLogin.Hide
        frmMain.Show
        frmMain.Caption = "Electronic Diary v2.1 - By Lim Meng Huey         User: " & UserName
    Else
        MsgBox "Sorry, your password is incorrect.", 48, "Login Failed."
        cboUser.Text = ""
        txtPassword.Text = ""
        cboUser.SetFocus
    End If
    
    Exit Sub

Error1:
    MsgBox Error, 48, "Error"
    cboUser.Text = ""
    txtPassword.Text = ""
    cboUser.SetFocus
    Exit Sub
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Connection
    
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * From tUser", Cn, 1, 4
      
    Do Until Rs.EOF
        cboUser.AddItem Rs("UserName")
        Rs.MoveNext
    Loop
      
    Rs.Close
    Set Rs = Nothing
End Sub
