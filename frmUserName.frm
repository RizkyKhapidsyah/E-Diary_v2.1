VERSION 5.00
Begin VB.Form frmUserName 
   Caption         =   "Change UserName"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmUserName.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1290
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New UserName :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()
    Connection
    
    If txtUserName.Text <> "" Then
        reply = MsgBox("Do you want to change your username? This will delete all your entries you entered before.", vbExclamation + 4, "Change UserName?")
        If reply = vbYes Then
            Set Rs = New ADODB.Recordset
            Rs.Open "Select * From tUser Where UserName = '" & UserName & "'", Cn, 1, 4
            
            OldUserName = UserName
            Rs("UserName") = txtUserName.Text
            UserName = txtUserName.Text
            Rs.Update
            Rs.UpdateBatch
            Rs.MoveNext
            Unload Me
            frmMain.Show
            frmMain.Caption = "Electronic Diary v2.1 - By Lim Meng Huey       User: " & UserName
        
            Rs.Close
            Set Rs = Nothing
            
            Set Rs = New ADODB.Recordset
            Rs.Open "Delete * From tEntry Where UserName = '" & OldUserName & "'", Cn, 1, 4
        Else
            txtUserName.SetFocus
        End If
    Else
        MsgBox ("Please Input Your New UserName."), 48, "Input New UserName."
        txtUserName.SetFocus
    End If
    
End Sub
