VERSION 5.00
Begin VB.UserControl ISPanel 
   Alignable       =   -1  'True
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2730
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   182
   ToolboxBitmap   =   "ISPanel.ctx":0000
   Begin VB.VScrollBar VScroll 
      Height          =   1575
      Left            =   2400
      Max             =   115
      SmallChange     =   100
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   0
      Max             =   80
      SmallChange     =   100
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2415
   End
   Begin VB.PictureBox pCorner 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2400
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   2
      Top             =   1440
      Width           =   315
   End
   Begin VB.PictureBox pView 
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   120
      ScaleHeight     =   117
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   3
      Top             =   120
      Width           =   2595
   End
   Begin VB.Image curMove 
      Height          =   480
      Left            =   2520
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "ISPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Name: ISPanel 1.1
'
'   Original Author: Fred.Cpp
'   Date: 02/18/2002
'   e-mail:  alfredo_cp@hotmail.com
'   e-mail2: fred_cpp@yahoo.com.mx
'
'   Modified by: Elias Barbosa
'   Date: 02/19/2002
'   e-mail: elias@eb8.com
'
'   Modified Again by: Fred.cpp
'   Date: 09/02/2003
'   e-mail:  alfredo_cp@hotmail.com
'   e-mail2: fred_cpp@yahoo.com.mx
'
'Description: This control is very useful for
' those who need more space on their forms.
' Run this example.
'
'How to use:
' 1. Insert a ISPanel Control into your Form.
' 2. Insert a Picture Box into the ISPanel Control.
' 2. Insert other controls (Such us Command Buttons,
'    Text Boxs...) into the Picture Box.
' 3. In the Form_Load Event call the Attach Function.
' 4. In the Form_QueryUnload event call the Detach
'    Function.
'
'Notes:
'   The Control captures the events of the Picture Box,
'   so, if you resize the Picture Box, the control
'   adjust the scrollbars. Also, if you resize the
'   ISPanel Control, it adjust its properties.
'
'   New Features:
'   1 - Can Move contents pressing and holding
'       the left mouse button.
'
'   2 - Can change the Border style



Option Explicit

'Constant Declarations
' 3D border styles

Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

' Border flags
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const BF_DIAGONAL = &H10

Private Const WM_SIZE = &H5

Private Type PointAPI
        X As Long
        Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum ISBorderStyle
    ISNone = 0
    ISRaisedOuter = 1
    ISSunkenOuter = 2
    ISRaisedInner = 4
    ISSunkenInner = 8
    ISEdge_Raised = EDGE_RAISED
    ISEdge_Sunken = EDGE_SUNKEN
    ISEdge_Etched = EDGE_ETCHED
    ISEdge_Bump = EDGE_BUMP
End Enum

Private gScaleX As Single
Private gScaleY As Single
Private lPrevParent As Long
Private m_Rt As RECT
Private WithEvents pChild As PictureBox
Attribute pChild.VB_VarHelpID = -1
Private WithEvents pcLabel As Label
Attribute pcLabel.VB_VarHelpID = -1
Private bMove As Boolean
Private ptX As Long, ptY As Long

'Default Property Values:
Const m_Def_Align = 0
Const m_def_BackColor = &H8000000C

'Property Variables:
Private m_Align As Integer                      'Align of the Container Control
Private m_Backcolor As OLE_COLOR                'BackColor
Private m_BorderWidth As Integer                ' Depending the border style
Private m_BorderStyle As ISBorderStyle

'Event Declarations:
Event Resize()
Event Scroll()



' API Declarations
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Sub HScroll_Scroll()
    UpdatePos
End Sub

Private Sub pChild_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        bMove = True
        ptX = X
        ptY = Y
        pChild.MousePointer = 5
    End If
End Sub

Private Sub pChild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
        Dim tmpX As Single
        Dim tmpY As Single
    If Button = vbLeftButton Then
        tmpX = pChild.Left + X - ptX
        tmpY = pChild.Top + Y - ptY
        
        If tmpX >= -1 Then
            tmpX = -1
        ElseIf tmpX <= -HScroll.Max * Screen.TwipsPerPixelX Then
            tmpX = -HScroll.Max * Screen.TwipsPerPixelX
        End If
        
        If tmpY >= -1 Then
            tmpY = -1
        ElseIf tmpY <= -VScroll.Max * Screen.TwipsPerPixelY Then
            tmpY = -VScroll.Max * Screen.TwipsPerPixelY
        End If
        Debug.Print "X:" & tmpX, HScroll.Max, pChild.Width
        Debug.Print "Y:" & tmpY
        pChild.Left = tmpX
        pChild.Top = tmpY
        HScroll.Value = -pChild.Left / Screen.TwipsPerPixelX
        VScroll.Value = -pChild.Top / Screen.TwipsPerPixelY
    End If
End Sub

Private Sub pChild_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        bMove = False
        pChild.MousePointer = 99
    End If
End Sub

Private Sub UserControl_Paint()
    Call DrawEdge(hdc, m_Rt, m_BorderStyle, BF_RECT)
End Sub

Private Sub UserControl_Resize()
    Dim loff As Integer
    Dim loffV As Integer
    Dim loffH As Integer
    Dim sV As Single
    Dim sH As Single
    
    On Error Resume Next
    
    'Vertical additional space...
    loffV = 2
    'Horizontal addidional space...
    loffH = 2
    
    Call VScroll.Move(ScaleWidth - VScroll.Width - loffV, 2, VScroll.Width, ScaleHeight - HScroll.Height - loffH)
    Call HScroll.Move(2, ScaleHeight - HScroll.Height - loffH, ScaleWidth - VScroll.Width - loffV, HScroll.Height)
    Call pCorner.Move(ScaleWidth - VScroll.Width - loffV, ScaleHeight - HScroll.Height - loffH, VScroll.Width, HScroll.Height)
    Call pView.Move(2, 2, ScaleWidth - IIf(VScroll.Visible, VScroll.Width, 0) - 4, ScaleHeight - IIf(HScroll.Visible, HScroll.Height, 0) - 4)
    
    m_Rt.Bottom = 0: m_Rt.Top = 0: m_Rt.Right = ScaleWidth: m_Rt.Bottom = ScaleHeight
    
    HScroll.Min = 1
    VScroll.Min = 1
    HScroll.LargeChange = UserControl.ScaleWidth
    VScroll.LargeChange = UserControl.ScaleHeight
    'Get Initial Data
    sH = (pChild.ScaleWidth - pView.ScaleWidth)
    sV = (pChild.ScaleHeight - pView.ScaleHeight)
    
    'Modify Vertical ScrollBar
    If sV = 0 Then
        VScroll.Max = 1
        VScroll.Left = ScaleWidth
        loffV = 1
        VScroll.Width = 0
        VScroll.Visible = False
    ElseIf sV < 0 Then
        VScroll.Max = 1
        VScroll.Left = ScaleWidth
        loffV = 1
        VScroll.Width = 0
        VScroll.Visible = False
    Else
        VScroll.Visible = True
        VScroll.Max = sV
        VScroll.Width = 17
        loffV = 2
    End If
    
    'Modify Horizontal Scrollbar
    If sH = 0 Then
        HScroll.Max = 1
        HScroll.Height = 0
        loffH = 1
        HScroll.Visible = False
    ElseIf sH < 0 Then
        HScroll.Max = 1
        HScroll.Visible = False
        HScroll.Height = 0
        loffH = 1
    Else
        HScroll.Max = sH
        HScroll.Visible = True
        HScroll.Height = 17
        loffH = 2
    End If
    
    Call VScroll.Move(ScaleWidth - VScroll.Width - loffV, 2, VScroll.Width, ScaleHeight - HScroll.Height - loffH)
    Call HScroll.Move(2, ScaleHeight - HScroll.Height - loffH, ScaleWidth - VScroll.Width - loffV, HScroll.Height)
    Call pCorner.Move(ScaleWidth - VScroll.Width - loffV, ScaleHeight - HScroll.Height - loffH, VScroll.Width, HScroll.Height)
    Call pView.Move(2, 2, ScaleWidth - IIf(VScroll.Visible, 0, VScroll.Width) - 4, ScaleHeight - IIf(HScroll.Visible, 0, HScroll.Height) - 4)
    If HScroll.Height < 5 Or VScroll.Width < 5 Then
        pCorner.Visible = False
    Else
        pCorner.Visible = True
    End If
    RaiseEvent Resize
    
End Sub

Private Sub pChild_Resize()
    Call UserControl_Resize
    
End Sub

'==================================================
'======= Following are the Subs that will    ======
'======= initialize and save the properties. ======
'==================================================

'Get property values from property bags...
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Backcolor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    pView.BackColor = m_Backcolor
    m_BorderWidth = PropBag.ReadProperty("BorderWidth", 2)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", 2)
End Sub

'Write the property values to the property bags...
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_Backcolor, m_def_BackColor)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, 2)
    Call PropBag.WriteProperty("BorderWidth", m_BorderWidth, 2)
End Sub

Private Sub UserControl_InitProperties()
    gScaleX = Screen.TwipsPerPixelX
    gScaleY = Screen.TwipsPerPixelY
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_Backcolor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_Backcolor = New_BackColor
    pView.BackColor = New_BackColor
    pCorner.BackColor = New_BackColor
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As ISBorderStyle
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As ISBorderStyle)
    m_BorderStyle = New_BorderStyle
    Cls
    UserControl_Paint
    PropertyChanged "BorderStyle"
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'==================================
'======= Following are some  ======
'======= complementary Subs. ======
'==================================

Private Sub VScroll_Change()
    UpdatePos
End Sub
   
Private Sub HScroll_Change()
    UpdatePos
End Sub

Private Sub VScroll_Scroll()
    UpdatePos
End Sub

Sub UpdatePos()
    'Called when Scrolls have Changed
    On Error Resume Next
    If bMove Then Exit Sub
    pChild.Move (-HScroll.Value) * Screen.TwipsPerPixelX, (-VScroll.Value) * Screen.TwipsPerPixelY
    pView.SetFocus
    RaiseEvent Scroll
   
End Sub

Public Sub Attach(newChild As PictureBox)
    Set pChild = newChild
    lPrevParent = SetParent(newChild.hwnd, pView.hwnd)
    pChild.Move 0, 0
    Set pChild.MouseIcon = curMove.Picture
    pChild.MousePointer = 99
    pChild.ScaleMode = 3
    UserControl_Resize
    UpdatePos
    
End Sub

Public Sub Detach()
    SetParent pChild.hwnd, lPrevParent
    Set pChild = Nothing
End Sub

