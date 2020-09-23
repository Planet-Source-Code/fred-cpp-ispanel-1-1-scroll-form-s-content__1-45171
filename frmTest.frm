VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "ISPanel Demo"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   571
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDocument 
      Height          =   3735
      Left            =   6000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   33
      Text            =   "frmTest.frx":0000
      Top             =   240
      Width           =   1935
   End
   Begin prjISPanelTest.ISPanel ISPanel1 
      Align           =   3  'Align Left
      Height          =   4605
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   8123
      BackColor       =   -2147483633
      BorderWidth     =   0
      Begin VB.PictureBox pCommands 
         BorderStyle     =   0  'None
         Height          =   15480
         Left            =   120
         ScaleHeight     =   15480
         ScaleWidth      =   5295
         TabIndex        =   1
         Top             =   120
         Width           =   5295
         Begin VB.Frame Frame1 
            Caption         =   "Border Style"
            Height          =   3615
            Left            =   240
            TabIndex        =   12
            Top             =   1320
            Width           =   4815
            Begin VB.OptionButton optBorderStyle 
               Caption         =   "None"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   21
               Top             =   240
               Width           =   2295
            End
            Begin VB.OptionButton optBorderStyle 
               Caption         =   "Raised Outer"
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   20
               Top             =   600
               Width           =   2295
            End
            Begin VB.OptionButton optBorderStyle 
               Caption         =   "Sunken Outer"
               Height          =   255
               Index           =   2
               Left            =   360
               TabIndex        =   19
               Top             =   960
               Value           =   -1  'True
               Width           =   2295
            End
            Begin VB.OptionButton optBorderStyle 
               Caption         =   "Raised Inner"
               Height          =   255
               Index           =   3
               Left            =   360
               TabIndex        =   18
               Top             =   1320
               Width           =   2295
            End
            Begin VB.OptionButton optBorderStyle 
               Caption         =   "Sunken Inner"
               Height          =   255
               Index           =   4
               Left            =   360
               TabIndex        =   17
               Top             =   1680
               Width           =   2295
            End
            Begin VB.OptionButton optBorderStyle 
               Caption         =   "Edge Raised"
               Height          =   255
               Index           =   5
               Left            =   360
               TabIndex        =   16
               Top             =   2040
               Width           =   2295
            End
            Begin VB.OptionButton optBorderStyle 
               Caption         =   "Edge Sunken"
               Height          =   255
               Index           =   6
               Left            =   360
               TabIndex        =   15
               Top             =   2400
               Width           =   2295
            End
            Begin VB.OptionButton optBorderStyle 
               Caption         =   "Edge Etched"
               Height          =   255
               Index           =   7
               Left            =   360
               TabIndex        =   14
               Top             =   2760
               Width           =   2295
            End
            Begin VB.OptionButton optBorderStyle 
               Caption         =   "Edge Bump"
               Height          =   255
               Index           =   8
               Left            =   360
               TabIndex        =   13
               Top             =   3120
               Width           =   2295
            End
         End
         Begin VB.Data Data1 
            Caption         =   "DemoData"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   495
            Left            =   360
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   5880
            Width           =   3495
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   360
            TabIndex        =   11
            Top             =   6720
            Width           =   3495
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   7440
            Width           =   3495
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   8160
            Width           =   3495
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   8880
            Width           =   3495
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Top             =   9600
            Width           =   3495
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   360
            TabIndex        =   6
            Top             =   10320
            Width           =   3495
         End
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   360
            TabIndex        =   5
            Top             =   11040
            Width           =   3495
         End
         Begin VB.ListBox List1 
            Height          =   2205
            Left            =   360
            TabIndex        =   4
            Top             =   12000
            Width           =   3495
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Do Something"
            Height          =   495
            Left            =   360
            TabIndex        =   3
            Top             =   14400
            Width           =   1575
         End
         Begin VB.CommandButton cmdMore 
            Caption         =   "Show another Example"
            Height          =   495
            Left            =   2040
            TabIndex        =   2
            Top             =   14400
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "This is just a test."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   5175
         End
         Begin VB.Label Label2 
            Caption         =   $"frmTest.frx":001D
            Height          =   735
            Left            =   240
            TabIndex        =   31
            Top             =   600
            Width           =   4815
         End
         Begin VB.Label Label3 
            Caption         =   "This Is just"
            Height          =   255
            Left            =   360
            TabIndex        =   30
            Top             =   6480
            Width           =   3495
         End
         Begin VB.Label Label4 
            Caption         =   "A Demo"
            Height          =   255
            Left            =   360
            TabIndex        =   29
            Top             =   7200
            Width           =   3495
         End
         Begin VB.Label Label5 
            Caption         =   "These controls"
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   7920
            Width           =   3495
         End
         Begin VB.Label Label6 
            Caption         =   "haven't a real"
            Height          =   255
            Left            =   360
            TabIndex        =   27
            Top             =   8640
            Width           =   3495
         End
         Begin VB.Label Label7 
            Caption         =   "utility"
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   9360
            Width           =   3495
         End
         Begin VB.Label Label8 
            Caption         =   "Caption..."
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   10080
            Width           =   3495
         End
         Begin VB.Label Label9 
            Caption         =   "another Caption...."
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   10800
            Width           =   3495
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   4080
            Y1              =   5280
            Y2              =   5280
         End
         Begin VB.Label Label10 
            Caption         =   "More Controls..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   23
            Top             =   5400
            Width           =   3495
         End
         Begin VB.Line Line2 
            X1              =   240
            X2              =   4080
            Y1              =   11520
            Y2              =   11520
         End
         Begin VB.Label Label11 
            Caption         =   "More Controls..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   22
            Top             =   11640
            Width           =   3495
         End
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdMore_Click()
    frmTest2.Show
End Sub

Private Sub Form_Load()
    Dim ni As Integer
    For ni = 0 To 50
        List1.AddItem Rnd(50)
    Next ni
    ISPanel1.Attach pCommands
End Sub

Private Sub Form_Resize()
    'ISPannel doesn't need code to be realigned, when has a child control.
    On Error Resume Next
    txtDocument.Move ISPanel1.Width + 2, txtDocument.Top, ScaleWidth - ISPanel1.Width - 4, ScaleHeight - txtDocument.Top - 4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ISPanel1.Detach
End Sub

Private Sub optBorderStyle_Click(Index As Integer)
    'For this demo: Change the Border Style!
    Select Case Index
        Case 0
            ISPanel1.BorderStyle = ISNone
        Case 1
            ISPanel1.BorderStyle = ISRaisedOuter
        Case 2
            ISPanel1.BorderStyle = ISSunkenOuter
        Case 3
            ISPanel1.BorderStyle = ISRaisedInner
        Case 4
            ISPanel1.BorderStyle = ISSunkenInner
        Case 5
            ISPanel1.BorderStyle = ISEdge_Raised
        Case 6
            ISPanel1.BorderStyle = ISEdge_Sunken
        Case 7
            ISPanel1.BorderStyle = ISEdge_Etched
        Case 8
            ISPanel1.BorderStyle = ISEdge_Bump
    End Select
End Sub
