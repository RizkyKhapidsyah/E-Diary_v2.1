VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSecret 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Secret"
   ClientHeight    =   3780
   ClientLeft      =   1770
   ClientTop       =   3645
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSecret.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3780
   ScaleWidth      =   5895
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox txtPassword2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2760
      Width           =   5175
   End
   Begin VB.TextBox txtPassword1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2040
      Width           =   5175
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1680
      TabIndex        =   4
      Top             =   1020
      Width           =   1200
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4320
      TabIndex        =   3
      Top             =   1020
      Width           =   1200
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3000
      TabIndex        =   2
      Top             =   1020
      Width           =   1200
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "txtFile"
      Top             =   480
      Width           =   5175
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   360
      TabIndex        =   0
      Top             =   1020
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog cdlOne 
      Left            =   3165
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   1.17491e-38
   End
   Begin VB.Label lblPassword2 
      Caption         =   "Enter password again:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblPassword1 
      Caption         =   "Enter password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblFile 
      Caption         =   "File:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   1755
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
         HelpContextID   =   101
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search For Help On..."
         HelpContextID   =   102
      End
      Begin VB.Menu mnuHelpDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmSecret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    'Prompt user for filename
    cdlOne.DialogTitle = "Secret"
    cdlOne.Flags = cdlOFNHideReadOnly
    cdlOne.Filter = "All files (*.*)|*.*"
    cdlOne.CancelError = True
    On Error Resume Next
    cdlOne.ShowOpen
    'Grab filename
    If Err = 0 Then
        txtFile.Text = cdlOne.FileName
    End If
    On Error GoTo 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEncrypt_Click()
    'Make sure both passwords match exactly
    If txtPassword1.Text <> txtPassword2.Text Then
        MsgBox "The two passwords are not the same!", _
            vbExclamation, "Secret"
        Exit Sub
    End If
    'Encrypt file
    MousePointer = vbHourglass
    cmdEncrypt.Enabled = False
    cmdDecrypt.Enabled = False
    cmdView.Enabled = False
    cmdBrowse.Enabled = False
    Refresh
    Encrypt
    txtFile_Change
    MousePointer = vbDefault
End Sub

Private Sub cmdDecrypt_Click()
    MousePointer = vbHourglass
    cmdEncrypt.Enabled = False
    cmdDecrypt.Enabled = False
    cmdView.Enabled = False
    cmdBrowse.Enabled = False
    Refresh
    Decrypt
    txtFile_Change
    MousePointer = vbDefault
End Sub

Private Sub cmdView_Click()
    Dim strA As String
    Dim lngZndx As Long
    MousePointer = vbHourglass
    'Get file contents
    Open txtFile.Text For Binary As #1
    strA = Space$(LOF(1))
    Get #1, , strA
    Close #1
    Do
        lngZndx = InStr(strA, Chr$(0))
        If lngZndx = 0 Or lngZndx > 5000 Then Exit Do
        Mid$(strA, lngZndx, 1) = Chr$(1)
    Loop
    'Display file contents
    MousePointer = vbDefault
    frmView.rtfView.Text = strA
    frmView.Caption = "Secret - " & txtFile.Text
    frmView.Show
End Sub

Private Sub Form_Load()
    'Center this form
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2
    'Disable most command buttons
    cmdEncrypt.Enabled = False
    cmdDecrypt.Enabled = False
    cmdView.Enabled = False
    'Initialize filename field
    txtFile.Text = ""
End Sub

Private Sub mnuAbout_Click()
    'Set properties
frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuContents_Click()
    cdlOne.HelpFile = App.Path & "\..\..\Help\Mvbdw.hlp"
    cdlOne.HelpCommand = cdlHelpContents
    cdlOne.ShowHelp
End Sub

Private Sub mnuSearch_Click()
    cdlOne.HelpFile = App.Path & "\..\..\Help\Mvbdw.hlp"
    cdlOne.HelpCommand = cdlHelpPartialKey
    cdlOne.ShowHelp
End Sub

Private Sub txtFile_Change()
    Dim lngFileLen As Long
    Dim strHead As String
    'Check to see whether file exists
    On Error Resume Next
    lngFileLen = Len(Dir(txtFile.Text))
    'Disable buttons if filename isn't valid
    If Err <> 0 Or lngFileLen = 0 Or Len(txtFile.Text) = 0 Then
        cmdEncrypt.Enabled = False
        cmdDecrypt.Enabled = False
        cmdView.Enabled = False
        lblPassword1.Enabled = False
        txtPassword1.Enabled = False
        lblPassword2.Enabled = False
        txtPassword2.Enabled = False
        txtPassword2.Text = ""
        Exit Sub
    End If
    'Get first 8 bytes of selected file
    Open txtFile.Text For Binary As #1
    strHead = Space(8)
    Get #1, , strHead
    Close #1
    'Check to see whether file is already encrypted
    If strHead = "[Secret]" Then
        cmdEncrypt.Enabled = False
        cmdDecrypt.Enabled = True
        lblPassword2.Enabled = False
        txtPassword2.Enabled = False
        txtPassword2.Text = ""
    Else
        cmdEncrypt.Enabled = True
        cmdDecrypt.Enabled = False
        lblPassword2.Enabled = True
        txtPassword2.Enabled = True
    End If
    lblPassword1.Enabled = True
    txtPassword1.Enabled = True
    cmdBrowse.Enabled = True
    cmdView.Enabled = True
End Sub

Sub Encrypt()
    Dim strHead As String
    Dim strT As String
    Dim strA As String
    Dim cphX As New Cipher
    Dim lngN As Long
    Open txtFile.Text For Binary As #1
    'Load entire file into strA
    strA = Space$(LOF(1))
    Get #1, , strA
    Close #1
    'Prepare header string with salt characters
    strT = Hash(Date & Str(Timer))
    strHead = "[Secret]" & strT & Hash(strT & txtPassword1.Text)
    'Do the encryption
    cphX.KeyString = strHead
    cphX.Text = strA
    cphX.DoXor
    cphX.Stretch
    strA = cphX.Text
    'Write header
    Open txtFile.Text For Output As #1
    Print #1, strHead
    'Write encrypted data
    lngN = 1
    Do
        Print #1, Mid(strA, lngN, 70)
        lngN = lngN + 70
    Loop Until lngN > Len(strA)
    Close #1
End Sub

Sub Decrypt()
    Dim strHead As String
    Dim strA As String
    Dim strT As String
    Dim cphX As New Cipher
    Dim lngN As Long
    'Get header (first 18 bytes of encrypted file)
    Open txtFile.Text For Input As #1
    Line Input #1, strHead
    Close #1
    'Check for correct password
    strT = Mid(strHead, 9, 8)
    If InStr(strHead, Hash(strT & txtPassword1.Text)) <> 17 Then
        MsgBox "Sorry, this is not the correct password!", _
            vbExclamation, "Secret"
        Exit Sub
    End If
    'Get file contents
    Open txtFile.Text For Input As #1
    'Read past the header
    Line Input #1, strHead
    'Read and build the contents string
    Do Until EOF(1)
        Line Input #1, strT
        strA = strA & strT
    Loop
    Close #1
    'Decrypted file contents
    cphX.KeyString = strHead
    cphX.Text = strA
    cphX.Shrink
    cphX.DoXor
    strA = cphX.Text
    'Replace file with decrypted version
    Kill txtFile.Text
    Open txtFile.Text For Binary As #1
    Put #1, , strA
    Close #1
End Sub

Function Hash(strA As String) As String
    Dim cphHash As New Cipher
    cphHash.KeyString = strA & "123456"
    cphHash.Text = strA & "123456"
    cphHash.DoXor
    cphHash.Stretch
    cphHash.KeyString = cphHash.Text
    cphHash.Text = "123456"
    cphHash.DoXor
    cphHash.Stretch
    Hash = cphHash.Text
End Function

