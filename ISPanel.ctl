VERSION 5.00
Begin VB.UserControl ISPanel 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5745
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   5745
   Begin VB.Timer tmrChangeFocus 
      Interval        =   10
      Left            =   4920
      Top             =   2040
   End
   Begin VB.VScrollBar VScroll 
      Height          =   4815
      Left            =   4560
      Max             =   115
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   120
      Max             =   80
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5280
      Width           =   3975
   End
   Begin VB.PictureBox pCorner 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4260
      ScaleHeight     =   495
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   5100
      Width           =   555
   End
   Begin VB.PictureBox pView 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   5235
      Left            =   0
      ScaleHeight     =   5235
      ScaleWidth      =   4635
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   4635
   End
   Begin VB.Image curMove 
      Height          =   480
      Left            =   5040
      Top             =   660
      Width           =   480
   End
End
Attribute VB_Name = "ISPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'how to use:
' 1.- Insert a ISPanel Control in your form
' 2.- Insert controls in a  picture Box
' 3.- In the Form Load Event call the Attatch Function
' 4.- In the Query Unload event call the detatch Function
' 5.- If using END in your application will cause VB IDE to shutdown
'     Must call Detach before End
' Notes:
'   the Control Captures the events of the Picturebox,
'        so if you resize the picturebox, the control adjust the scrollbars.
'   Also, if you resize the ISPanel control, it adjust his properties
'
'   Feedback is GREATLY appreciated... Votes Would be nice ;)
'Author: Fred_Cpp
'   mail:  alfredo_cp@notmail.com
'   mail2: fred_cpp@yahoo.com
'***********************************************************************
'3/30/2005 Added
' David J
'   Mousewheel capability
'   When a new control has focus that is out of viewing range scrolls
'   to within viewing range
'***********************************************************************

Option Explicit
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Private Enum State
    Normal
    hover
End Enum

Private gScaleX As Single
Private gScaleY As Single
Private InOut As Boolean
Private iState As State
'Default Property Values:
Const m_def_Enabled = True
Const m_def_BorderStyle = 0
Const m_Def_Align = 0
Const m_def_BackColor = &H8000000C
'Property Variables:
Private m_Enabled As Boolean
Private m_BorderStyle As Integer                'What BorderStyle to Use??
Private m_Align As Integer                      'Align of the Container Control
Private m_BackColor As OLE_COLOR                'BackColor

Private sZoom As Single                         'Zoom Value
Private psWidth As Single, psHeight As Single   'Paper Size
Private lPrevParent As Long
Private sPrevX As Single
Private sPrevY As Single
Private WithEvents pChild As PictureBox
Attribute pChild.VB_VarHelpID = -1

'Event Declarations:
Event Resize()
'Constant Declarations
Private Const WM_SIZE = &H5
' API Declarations

'**************************************************************************
'3/30/2005
'Added to move scroll bars when field with focus is out of viewing range
' David J
'**************************************************************************
Private strControlName As String
Private intControlIndex As Integer
Private varControl As Variant
'**************************************************************************
'End New Additions
'**************************************************************************

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long




' Commons Functions Support
Private Function InBox() As Boolean
    Dim mpos As POINTAPI
    Dim oRect As RECT
    GetCursorPos mpos
    GetWindowRect lPrevParent, oRect
    If mpos.X >= oRect.Left And mpos.X <= oRect.Right And _
        mpos.Y >= oRect.Top And mpos.Y <= oRect.Bottom Then
        InBox = True
    Else
        InBox = False
   End If
End Function

Private Sub HScroll_Scroll()
    UpdatePos
End Sub

Private Sub pChild_Resize()
    UserControl_Resize
End Sub

Private Sub DragObj(hwnd As Long)
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, 0&
End Sub

'**************************************************************************
'3/30/2005
'Added to move scroll bars when field with focus is out of viewing range
' David J
'**************************************************************************
Private Sub tmrChangeFocus_Timer()
    'Determine if in design view
    If UserControl.Ambient.UserMode = True Then
        Dim strThisControl, strControlContainer As String
        Dim intThisIndex, intResult, intNumTries As Integer
        Dim lngCtrlTop, lngCalcDiff, lngCtrlValue, lngLastHwnd As Long
        Dim myControl As Control
        Dim blnFoundOne, blnDone As Boolean
        On Error Resume Next
        'Get Current Control Name and Index
        strThisControl = Screen.ActiveControl.Name
        intThisIndex = Screen.ActiveControl.Index
        'Determine if control name or index has changed
        If strThisControl <> strControlName Or (intThisIndex <> intControlIndex And Not IsNull(intThisIndex)) Then
            strControlName = strThisControl
            'Determine if Index changed
            If intThisIndex <> intControlIndex And Not IsNull(intThisIndex) Then
                intControlIndex = intThisIndex
            End If
            lngCtrlValue = 0
            Set varControl = Screen.ActiveControl
            On Error Resume Next
            Dim objControlContainer As Object
            Set objControlContainer = varControl.Container
            Do Until objControlContainer Is Nothing
                lngLastHwnd = objControlContainer.hwnd
                lngCtrlValue = lngCtrlValue + objControlContainer.Top
                Err.Clear
                Set objControlContainer = objControlContainer.Container
                If Err.Number <> 0 Then Exit Do
            Loop
            'If the Active Control is in the IsPanel Control
            If lngLastHwnd = pChild.hwnd Then
                lngCtrlTop = lngCtrlValue + varControl.Top - 50
                lngCtrlValue = lngCtrlValue + varControl.Top + varControl.Height + 75
                'If the top of the Active Control is outside of the viewing
                '   area then change the Vertical Scroll bar to place it in view
                If lngCtrlTop < 0 Then
                    VScroll.Value = VScroll.Value + lngCtrlTop
                ElseIf lngCtrlValue > pChild.Height - VScroll.Max Then
                    lngCalcDiff = lngCtrlValue - (pChild.Height - VScroll.Max)
                    VScroll.Value = VScroll.Value + lngCalcDiff
                End If
            Else
                Exit Sub
            End If
        End If
    Else
        tmrChangeFocus.Enabled = False
    End If
    DoEvents
End Sub
'**************************************************************************
'3/30/2005
'Added for use with Mouse wheel
' David J
'**************************************************************************
Public Function ScrollUp() As Boolean
    If InBox = True Then
        If VScroll.Value >= VScroll.SmallChange Then
            VScroll.Value = VScroll.Value - VScroll.SmallChange
        ElseIf VScroll.Value = 1 Then
            If HScroll.Value >= HScroll.SmallChange Then
                HScroll.Value = HScroll.Value - HScroll.SmallChange
            Else
                HScroll.Value = 1
            End If
        Else
            VScroll.Value = 1
        End If
        ScrollUp = True
    Else
        ScrollUp = False
    End If
End Function
Public Function ScrollDown() As Boolean
    If InBox = True Then
        If VScroll.Value <= VScroll.Max - VScroll.SmallChange Then
            VScroll.Value = VScroll.Value + VScroll.SmallChange
        ElseIf VScroll.Value = VScroll.Max Then
            If HScroll.Value <= HScroll.Max - HScroll.SmallChange Then
                HScroll.Value = HScroll.Value + HScroll.SmallChange
            Else
                HScroll.Value = HScroll.Max
            End If
        Else
            VScroll.Value = VScroll.Max
        End If
        ScrollDown = True
    Else
        ScrollDown = False
    End If
End Function
'**************************************************************************
'End New Additions
'**************************************************************************

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    pView.BackColor = m_BackColor
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Dim loff As Integer
    loff = 45
    VScroll.Move UserControl.Width - VScroll.Width - loff, 0, VScroll.Width, UserControl.Height - HScroll.Height - loff
    HScroll.Move 0, UserControl.Height - HScroll.Height - loff, UserControl.Width - VScroll.Width - loff, HScroll.Height
    pCorner.Move UserControl.Width - VScroll.Width - loff, UserControl.Height - HScroll.Height - loff, VScroll.Width, HScroll.Height
    Dim sV As Single
    Dim sH As Single
    pView.Move 0, 0, Width - VScroll.Width, Height - HScroll.Height
    HScroll.Min = 1
    VScroll.Min = 1
    sH = pChild.Width - pView.Width
    sV = pChild.Height - pView.Height
    'Modify Vertical ScrollBar
    '*********************************************************
    '3/30/2005
    'Added to disable the scrollbar if it cant be scrolled
    ' David J
    '*********************************************************
    If sV <= 0 Then
        VScroll.Max = 1
        VScroll.Enabled = False
    Else
        VScroll.Max = sV
        VScroll.Enabled = True
    End If
    'Modify Horizontal Scrollbar
    If sH <= 0 Then
        HScroll.Max = 1
        HScroll.Enabled = False
    Else
        HScroll.Max = sH
        HScroll.Enabled = True
    End If
    
    HScroll.LargeChange = UserControl.Width
    HScroll.SmallChange = 135
    
    VScroll.LargeChange = UserControl.Height
    VScroll.SmallChange = 135
    RaiseEvent Resize
End Sub

Private Sub UserControl_Terminate()
    '***************************************
    '3/30/2005 Added
    ' David J
    '***************************************
    UnHook
    tmrChangeFocus.Enabled = False
    '***************************************
    'End New Additions
    '***************************************
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
End Sub

Private Sub UserControl_Initialize()
    'Initialization Code
    psWidth = 8000
    psHeight = 11500
    '***************************************
    '3/30/2005 Added
    ' David J
    '***************************************
    tmrChangeFocus.Enabled = True
    '***************************************
    'End New Additions
    '***************************************
End Sub

Private Sub UserControl_InitProperties()
    gScaleX = Screen.TwipsPerPixelX
    gScaleY = Screen.TwipsPerPixelY
    m_Enabled = m_def_Enabled
    m_BorderStyle = m_def_BorderStyle
End Sub

Private Sub UserControl_Paint()
    If iState = Normal Then
        DrawFlat
    ElseIf iState = hover Then
        DrawRaised
    End If
End Sub

Private Sub DrawFlat()
    Cls
End Sub

Private Sub DrawRaised()
    Line (0, 0)-(Width, 0), vb3DShadow
    Line (0, 0)-(0, Height), vb3DShadow
    Line (Width - 15, 0)-(Width - 15, Height - 15), vb3DHighlight
    Line (0, Height - 15)-(Width - 15, Height - 15), vb3DHighlight
    
    Line (15, 15)-(ScaleWidth - 30, 15), vb3DHighlight
    Line (15, 15)-(15, ScaleHeight - 30), vb3DHighlight
    Line (ScaleWidth - 30, 15)-(ScaleWidth - 30, ScaleHeight - 30), vb3DShadow
    Line (15, ScaleHeight - 30)-(ScaleWidth - 30, ScaleHeight - 30), vb3DShadow
End Sub


'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property


'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    UserControl_Paint
    PropertyChanged "BorderStyle"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    pView.BackColor = New_BackColor
    UserControl_Paint
    PropertyChanged "BackColor"
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'   Functionality Routines
Private Sub VScroll_Change()
    UpdatePos
End Sub
   
Private Sub HScroll_Change()
    UpdatePos
End Sub

Sub UpdatePos()
    'Called when Scrolls Heve Changed
    On Error Resume Next
    pChild.Move -HScroll.Value, -VScroll.Value
    pView.SetFocus
    varControl.SetFocus
End Sub

Private Sub pChild_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pChild.MousePointer = 99
    sPrevX = HScroll.Max - HScroll.Value - X + pView.Width
    sPrevY = VScroll.Max - VScroll.Value - Y + pView.Height
End Sub

Private Sub pChild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    
    Dim vx As Single
    Dim vy As Single
    vx = 1 + (X + sPrevX) / 2
    vy = 1 + (Y - sPrevY) / 2
    'Check X Value
    If vx < HScroll.Max And vx > HScroll.Min Then
        HScroll.Value = vx
    ElseIf vx > HScroll.Max Then
        HScroll.Value = HScroll.Max
    ElseIf vx < HScroll.Min Then
        HScroll.Value = HScroll.Min
    End If
    'Check Y Value
    If vy < VScroll.Max And vy > VScroll.Min Then
        VScroll.Value = vy
    ElseIf vy > VScroll.Max Then
        VScroll.Value = VScroll.Max
    ElseIf vy < VScroll.Min Then
        VScroll.Value = VScroll.Min
    End If

End Sub

Private Sub pChild_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    pChild.MousePointer = 0
End Sub
Public Sub MoveScrollBar(NewValue As Long)
    If VScroll.Value < VScroll.Max Then
        VScroll.Value = VScroll.Value + 1
    End If
    If VScroll.Value > 0 Then
        VScroll.Value = VScroll.Value - 1
    End If
End Sub
Public Sub Attatch(newChild As Object)
    If TypeOf newChild Is PictureBox Then
        Set pChild = newChild
        lPrevParent = SetParent(newChild.hwnd, pView.hwnd)
        pChild.Move 0, 0
        pChild.MouseIcon = curMove.Picture
        pChild.MousePointer = 0
        UserControl_Resize
        UpdatePos
        Hook Me.hwnd
    Else
        MsgBox "Object being attached must be a Picturebox", vbCritical, "Incorrect Object"
    End If
End Sub

Public Sub Detatch()
    On Error Resume Next
    UnHook
    SetParent pChild.hwnd, lPrevParent
    Set pChild = Nothing
End Sub

Private Sub VScroll_Scroll()
    UpdatePos
End Sub
