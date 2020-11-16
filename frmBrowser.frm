VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "Web Browser "
   ClientHeight    =   8310
   ClientLeft      =   285
   ClientTop       =   630
   ClientWidth     =   11685
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar Status 
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   10080
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command12 
      Height          =   300
      Left            =   14280
      Picture         =   "frmBrowser.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Launch Website"
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox URLS 
      Height          =   315
      Left            =   1080
      TabIndex        =   14
      Top             =   960
      Width           =   10485
   End
   Begin SHDocVwCtl.WebBrowser Browser 
      Height          =   6735
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   11685
      ExtentX         =   20611
      ExtentY         =   11880
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Edit"
      Height          =   680
      Left            =   8680
      Picture         =   "frmBrowser.frx":0754
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Edit the currently loaded page using Notepad."
      Top             =   100
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Print"
      Height          =   680
      Left            =   7840
      Picture         =   "frmBrowser.frx":0C0E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Prints the currently loaded page."
      Top             =   100
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Mail"
      Height          =   680
      Left            =   7000
      Picture         =   "frmBrowser.frx":1088
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Loads you default E-Mail Client."
      Top             =   100
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "History"
      Height          =   680
      Left            =   6160
      Picture         =   "frmBrowser.frx":157A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Displays your history folder."
      Top             =   100
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Favourites"
      Height          =   680
      Left            =   5200
      Picture         =   "frmBrowser.frx":1A30
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Displays you favourites folder."
      Top             =   100
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Search"
      Height          =   680
      Left            =   4360
      Picture         =   "frmBrowser.frx":1E6E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Loads you current default search engine."
      Top             =   100
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Home"
      Height          =   680
      Left            =   3520
      Picture         =   "frmBrowser.frx":2360
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Loads your current homepage in the browser."
      Top             =   100
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "Refresh"
      Height          =   680
      Left            =   2680
      Picture         =   "frmBrowser.frx":2852
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Refreshes the data on the currently loaded page."
      Top             =   100
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   680
      Left            =   1840
      Picture         =   "frmBrowser.frx":2C70
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Stops the current page from loading."
      Top             =   100
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Forward"
      Height          =   680
      Left            =   1000
      Picture         =   "frmBrowser.frx":3126
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Loads the next page in the browser."
      Top             =   100
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   680
      Left            =   180
      Picture         =   "frmBrowser.frx":33D8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Loads the previous page in the browser."
      Top             =   100
      Width           =   735
   End
   Begin ComCtl3.CoolBar Toolbar 
      Height          =   885
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   1561
      BandCount       =   2
      VariantHeight   =   0   'False
      _CBWidth        =   15255
      _CBHeight       =   885
      _Version        =   "6.0.8169"
      MinHeight1      =   825
      NewRow1         =   0   'False
      MinHeight2      =   360
      NewRow2         =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileGo 
         Caption         =   "&Go"
         Shortcut        =   ^G
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBack 
         Caption         =   "&Back"
         Shortcut        =   ^B
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileForward 
         Caption         =   "&Forward"
         Shortcut        =   ^F
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileStop 
         Caption         =   "&Stop"
         Shortcut        =   ^S
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   ^R
      End
      Begin VB.Menu line6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileHome 
         Caption         =   "&Home"
         Shortcut        =   ^H
      End
      Begin VB.Menu line7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSearch 
         Caption         =   "Se&arch"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewFav 
         Caption         =   "&Favourites"
         Shortcut        =   ^O
      End
      Begin VB.Menu line56 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewHis 
         Caption         =   "&History"
         Shortcut        =   ^I
      End
      Begin VB.Menu line67 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewMail 
         Caption         =   "&Mail"
         Shortcut        =   ^M
      End
      Begin VB.Menu line78 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu line89 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewEdit 
         Caption         =   "&Edit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Browser.Navigate URLS.Text
URLS.AddItem URLS.Text
End If

End Sub

Private Sub Browser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
Status.Max = ProgressMax
Status.Value = Progress

End Sub
Private Sub Command1_Click()
Browser.GoBack
End Sub

Private Sub Command10_Click()
Printer.Print Browser.Document
End Sub

Private Sub Command11_Click()
frmNotepad.Show
Me.WindowState = 1
End Sub

Private Sub Command12_Click()
Browser.Navigate URLS.Text
URLS.AddItem URLS.Text

End Sub

Private Sub Command2_Click()
Browser.GoForward

End Sub

Private Sub Command3_Click()
Browser.Stop

End Sub

Private Sub Command4_Click()
Browser.Refresh
End Sub

Private Sub Command5_Click()
Browser.GoHome
End Sub

Private Sub Command6_Click()
Browser.GoSearch
End Sub

Private Sub Command7_Click()
Call ShellExecute(hwnd, "Open", "C:\Windows\Favourites\", "", App.Path, 1)

End Sub

Private Sub Command8_Click()
Call ShellExecute(hwnd, "Open", "C:\Windows\History\", "", App.Path, 1)

End Sub

Private Sub Command9_Click()
MsgBox "Not available yet.", vbInformation, Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload frmAbout
End Sub
Private Sub mnuFileBack_Click()
Command1_Click
End Sub
Private Sub mnuFileExit_Click()
Unload Me

End Sub

Private Sub mnuFileForward_Click()
Command2_Click
End Sub

Private Sub mnuFileGo_Click()
Command12_Click
End Sub

Private Sub mnuFileHome_Click()
Command5_Click
End Sub

Private Sub mnuFileRefresh_Click()
Command4_Click
End Sub

Private Sub mnuFileSearch_Click()
Command6_Click
End Sub

Private Sub mnuFileStop_Click()
Command3_Click
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
Me.WindowState = 1
End Sub

Private Sub mnuViewEdit_Click()
Command11_Click

End Sub

Private Sub mnuViewFav_Click()
Command7_Click

End Sub

Private Sub mnuViewHis_Click()
Command8_Click

End Sub

Private Sub mnuViewMail_Click()
Command9_Click

End Sub

Private Sub mnuViewPrint_Click()
Command10_Click

End Sub
Private Sub URLS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command12_Click
End If
End Sub
