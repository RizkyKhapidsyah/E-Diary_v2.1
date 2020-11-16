VERSION 5.00
Begin VB.Form frmPassword 
   Caption         =   "Change Password"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3960
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1560
   ScaleWidth      =   3960
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtCNPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtNPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirm New Password :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Password :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    Connection
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * From tUser Where UserName = '" & UserName & "'", Cn, 1, 4
    If txtNPass.Text = txtCNPass.Text Then
        Rs("Password") = txtCNPass.Text
        Rs.Update
        Rs.UpdateBatch
        Rs.MoveNext
        Unload Me
    Else
        MsgBox "Either You Didn't Enter Your New Password Or Your Passwords Are Not Same.", 0, "Error."
        txtNPass.Text = ""
        txtCNPass.Text = ""
        txtNPass.SetFocus
    End If
    Rs.Close
    Set Rs = Nothing
End Sub
