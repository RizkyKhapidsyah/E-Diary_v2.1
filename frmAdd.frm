VERSION 5.00
Begin VB.Form frmAdd 
   Caption         =   "Add User"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1665
   ScaleWidth      =   7350
   Begin VB.TextBox txtCPassword 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtPassword 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtLogin 
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirm Password :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Password :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Login :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    On Error GoTo FixError
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * From tUser", Cn, 1, 4
        
    If txtPassword.Text = txtCPassword.Text Then
        If Rs("UserName") <> txtLogin.Text Then
            Rs.AddNew
            Rs("UserName") = txtLogin.Text
            Rs("Password") = txtPassword.Text
            Rs.Update
            Rs.UpdateBatch
            Rs.MoveNext
            Unload frmAdd
            frmAdd.Hide
        Else
            MsgBox "The login name you typed have been used. Please select another.", vbExclamation, "UserName Used"
        End If
    Else
        MsgBox "Your Passwords Are Not Same.", 48, "Passwords Incorrect."
        txtPassword.Text = ""
        txtCPassword.Text = ""
        txtPassword.SetFocus
    End If
    Exit Sub
FixError:
    MsgBox Error, 48, "Error."
    Exit Sub
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Connection
End Sub
