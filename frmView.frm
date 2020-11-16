VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmView 
   Caption         =   "Secret - View File"
   ClientHeight    =   3735
   ClientLeft      =   1395
   ClientTop       =   2010
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3735
   ScaleWidth      =   7230
   Begin RichTextLib.RichTextBox rtfView 
      Height          =   1575
      Left            =   660
      TabIndex        =   0
      Top             =   540
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   2778
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmView.frx":0000
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnBeenHereDoneThis As Boolean
Private Sub Form_Load()
    mblnBeenHereDoneThis = False
End Sub
Private Sub Form_Resize()
    'Center this form, but only the first time
    If mblnBeenHereDoneThis = False Then
        Me.Left = (Screen.Width - Me.Width) \ 2
        Me.Top = (Screen.Height - Me.Height) \ 2
        mblnBeenHereDoneThis = True
    End If
    'Size RichTextBox to fill form
    rtfView.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
                                                                            
