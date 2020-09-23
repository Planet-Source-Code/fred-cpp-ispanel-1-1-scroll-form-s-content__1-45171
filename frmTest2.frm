VERSION 5.00
Begin VB.Form frmTest2 
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   268
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   343
   StartUpPosition =   3  'Windows Default
   Begin prjISPanelTest.ISPanel ispContents 
      Align           =   1  'Align Top
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   5741
      BackColor       =   -2147483633
      BorderStyle     =   10
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   9255
         Left            =   360
         ScaleHeight     =   9255
         ScaleWidth      =   12135
         TabIndex        =   1
         Top             =   240
         Width           =   12135
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   9240
            Left            =   0
            Picture         =   "frmTest2.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   12120
         End
      End
   End
End
Attribute VB_Name = "frmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    ispContents.Attach Picture1
End Sub

Private Sub Form_Resize()
    DoEvents    'If you don't use doevents here, the form doesn't update the size,
                'and controls have wrong size or position.
    ispContents.Height = ScaleHeight - 16
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ispContents.Detach
End Sub
