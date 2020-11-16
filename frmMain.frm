VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Electronic Diary v2.1 - By Lim Meng Huey"
   ClientHeight    =   5265
   ClientLeft      =   510
   ClientTop       =   855
   ClientWidth     =   6750
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0442
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Diary\dbDairy.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Diary\dbDairy.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&User"
      Begin VB.Menu mnuLogin 
         Caption         =   "&Login"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuNewUser 
         Caption         =   "&New User"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnuDiary 
      Caption         =   "&Diary"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Entries"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find Entries"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Entries"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuUserDetail 
      Caption         =   "User &Details"
      Begin VB.Menu mnuChangeUserName 
         Caption         =   "Change &UserName"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change &Password"
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "&Other Features"
      Begin VB.Menu mnuSecret 
         Caption         =   "Encrypt/Decrypt &Files"
      End
      Begin VB.Menu mnuIcqPager 
         Caption         =   "&Icq Pager"
      End
      Begin VB.Menu mnuBrowser 
         Caption         =   "&Web Browser"
      End
      Begin VB.Menu mnuMp3 
         Caption         =   "&MP3 Player"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    mnuDiary.Enabled = False
    mnuUserDetail.Enabled = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    MsgBox "Thanks For Using This Electronic Diary. For More Free Excelent Softwares, Please Contact Lim Meng Huey(xtrmeprohacker@yahoo.com.sg).", , "Thank You."
End Sub

Private Sub mnuAbout_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuAdd_Click()
    Load frmAddEntry
    frmAddEntry.Show
End Sub

Private Sub mnuBrowser_Click()
    Load frmBrowser
    frmBrowser.Show
End Sub

Private Sub mnuChangePassword_Click()
    On Error GoTo FixError
    Connection
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * From tUser Where UserName = '" & UserName & "'", Cn, 1, 4
    OPass = InputBox("Input Old Password", "Old Password.")
    If OPass = Rs("Password") Then
        Load frmPassword
        frmPassword.Show
    Else
        MsgBox "Your Old Password Is Incorrect.", 0, "Old Password Incorrect."
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Sub
FixError:
    MsgBox Error, 48, "Error."
    Exit Sub
End Sub

Private Sub mnuChangeUserName_Click()
    Load frmUserName
    frmUserName.Show
End Sub

Private Sub mnuDelete_Click()
    Load frmDelete
    frmDelete.Show
End Sub

Private Sub mnuDiary_Click()
    Connection
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * From tEntry Where UserName = '" & UserName & "' And Date = '" & Date & "'", Cn, 1, 4
    If Rs.RecordCount > 0 Then
        mnuAdd.Enabled = False
    Else
        mnuAdd.Enabled = True
    End If
End Sub

Private Sub mnuExit_Click()
    CloseProgram
End Sub

Private Sub mnuFind_Click()
    Load frmFind
    frmFind.Show
End Sub

Private Sub mnuIcqPager_Click()
    Load frmIcqPager
    frmIcqPager.Show
End Sub

Private Sub mnuLogin_Click()
    Load frmLogin
    frmLogin.Show
End Sub

Private Sub mnuMp3_Click()
    Load frmPlayer
    frmPlayer.Show
End Sub

Private Sub mnuNewUser_Click()
    Load frmAdd
    frmAdd.Show
End Sub

Private Sub mnuSecret_Click()
    Load frmSecret
    frmSecret.Show
End Sub
