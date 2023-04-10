VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "claProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###########################################################################################################
'###########################################################################################################
'###
'###                                       ProgressBar API Control
'###        Author:         Brian Phillips
'###        Date:           16/03/2023  (v. 0.1)
'###        Description:    Implementation Win32 API ProgressBar control for VBA
'###
'###
'###########################################################################################################
'###########################################################################################################
Option Explicit

'###########################################################################################################
'###    Constant Declarations
'###########################################################################################################
'Style Constants
Private Const PBS_SMOOTH As Long = &H1
Private Const PBS_VERTICAL As Long = &H4
Private Const PBS_MARQUEE As Long = &H8
Private Const PBS_SMOOTHREVERSE As Long = &H10

Private Const WS_BORDER As Long = &H800000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_OVERLAPPED As Long = &H0
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_VISIBLE As Long = &H10000000

Private Const WS_EX_CLIENTEDGE As Long = &H200
Private Const WS_EX_LEFT As Long = &H0
Private Const WS_EX_LTRREADING As Long = &H0
Private Const WS_EX_RIGHTSCROLLBAR As Long = &H0

'Message Constants
Private Const CCM_FIRST As Long = &H2000
Private Const CCM_SETBKCOLOR As Long = (CCM_FIRST + 1)
Private Const WM_USER As Long = &H400

Private Const PBM_DELTAPOS As Long = (WM_USER + 3)
Private Const PBM_GETBARCOLOR As Long = (WM_USER + 15)
Private Const PBM_GETBKCOLOR As Long = (WM_USER + 14)
Private Const PBM_GETPOS As Long = (WM_USER + 8)
Private Const PBM_GETRANGE As Long = (WM_USER + 7)
Private Const PBM_GETSTATE As Long = (WM_USER + 17)
Private Const PBM_GETSTEP As Long = (WM_USER + 13)
Private Const PBM_SETBARCOLOR As Long = (WM_USER + 9)
Private Const PBM_SETBKCOLOR As Long = CCM_SETBKCOLOR
Private Const PBM_SETMARQUEE As Long = (WM_USER + 10)
Private Const PBM_SETPOS As Long = (WM_USER + 2)
Private Const PBM_SETRANGE32 As Long = (WM_USER + 6)
Private Const PBM_SETSTATE As Long = (WM_USER + 16)
Private Const PBM_SETSTEP As Long = (WM_USER + 4)
Private Const PBM_STEPIT As Long = (WM_USER + 5)

'State Constants
Private Const PBST_NORMAL As Long = &H1
Private Const PBST_ERROR As Long = &H2
Private Const PBST_PAUSED As Long = &H3

'Misc Constants
Private Const PROGRESS_CLASS As String = "msctls_progress32"

'Constants for Scale functions
Private Const vbPixels As Integer = 999
Private Const vbInches As Integer = 1440
Private Const vbTwips As Integer = 1

'###########################################################################################################
'###    Enum/Type Declarations
'###########################################################################################################
Enum eOrientationConstants
    ccOrientationHorizontal = 0
    ccOrientationVertical = 1
End Enum

Enum eScrollingConstants
    ccScrollingSmooth = 1
    ccScrollingStandard = 0
End Enum

Enum eStateConstants
    pbNormal = PBST_NORMAL
    pbError = PBST_ERROR
    pbPause = PBST_PAUSED
End Enum

'Ensure that the following two enumerations are either uncommented, or if using multiple of my controls, contained
'in a public module

Enum eBorderStyleConstants
    ccNone = 0
    ccFixedSingle = 1
End Enum

Enum eAppearanceConstants
     ccFlat = 0
    cc3D = 1
End Enum

'###########################################################################################################
'###    API Function Declarations
'###########################################################################################################
'For Control
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

'For Scale functions
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

'###########################################################################################################
'###    Private Variable Declarations
'###########################################################################################################
Private pAppearance As eAppearanceConstants
Private pBorderStyle As eBorderStyleConstants
Private pHeight As Integer
Private pHwnd As OLE_HANDLE
Private pHwndParent As OLE_HANDLE
Private pLeft As Integer
Private pMarquee As Boolean
Private pOrientation As eOrientationConstants
Private pScrolling As eScrollingConstants
Private pState As eStateConstants
Private pTarget As TextBox
Private pTop As Integer
Private pWidth As Integer
Private pVisible As Boolean

'###########################################################################################################
'###    Property Get/Let/Set routines
'###########################################################################################################
Public Property Get Appearance() As eAppearanceConstants: Appearance = pAppearance: End Property
Public Property Let Appearance(iAppearance As eAppearanceConstants)
If iAppearance <> pAppearance Then
    pAppearance = iAppearance
    Call Me.Refresh
End If
End Property

Public Property Get BackColor() As OLE_COLOR: BackColor = SendMessage(pHwnd, PBM_GETBKCOLOR, 0&, 0&): End Property
Public Property Let BackColor(iBackColor As OLE_COLOR): Call SendMessage(pHwnd, PBM_SETBKCOLOR, ByVal iBackColor, 0&): End Property

Public Property Get BarColor() As OLE_COLOR: BarColor = SendMessage(pHwnd, PBM_GETBARCOLOR, 0&, 0&): End Property
Public Property Let BarColor(iBarColor As OLE_COLOR): Call SendMessage(pHwnd, PBM_SETBARCOLOR, ByVal iBarColor, 0&): End Property

Public Property Get BorderStyle() As eBorderStyleConstants: BorderStyle = pBorderStyle: End Property
Public Property Let BorderStyle(iBorderStyle As eBorderStyleConstants): pBorderStyle = iBorderStyle: Call Me.Refresh: End Property

Public Property Get Height() As Integer: Height = pHeight: End Property
Public Property Let Height(iHeight As Integer): pHeight = CInt(ScaleY(CDbl(iHeight), vbTwips, vbPixels)): Call Me.Refresh: End Property

Public Property Get hWnd() As OLE_HANDLE: hWnd = pHwnd: End Property

Public Property Get Left() As Integer: Left = pLeft: End Property
Public Property Let Left(iLeft As Integer): pLeft = CInt(ScaleX(CDbl(iLeft), vbTwips, vbPixels)): Call Me.Refresh: End Property

Public Property Get Marquee() As Boolean: Marquee = pMarquee: End Property
Public Property Let Marquee(iMarquee As Boolean): pMarquee = iMarquee: Call Me.Refresh: End Property

Public Property Get Max() As Long: Max = SendMessage(pHwnd, PBM_GETRANGE, False, 0&): End Property
Public Property Let Max(iMax As Long): Call SendMessage(pHwnd, PBM_SETRANGE32, ByVal Min, ByVal iMax): End Property

Public Property Get Min() As Long: Min = SendMessage(pHwnd, PBM_GETRANGE, True, 0&): End Property
Public Property Let Min(iMin As Long): Call SendMessage(pHwnd, PBM_SETRANGE32, ByVal iMin, ByVal Max): End Property

Public Property Get Orientation() As eOrientationConstants: Orientation = pOrientation: End Property
Public Property Let Orientation(iOrientation As eOrientationConstants): pOrientation = iOrientation: Call Me.Refresh: End Property

Public Property Get Scrolling() As eScrollingConstants: Scrolling = pScrolling: End Property
Public Property Let Scrolling(iScrolling As eScrollingConstants): pScrolling = iScrolling: Call Me.Refresh: End Property

Public Property Get State() As eStateConstants: State = SendMessage(pHwnd, PBM_GETSTATE, 0&, 0&): End Property
Public Property Let State(iState As eStateConstants): Call SendMessage(pHwnd, PBM_SETSTATE, iState, 0&): End Property

Public Property Get StepInterval() As Long: StepInterval = SendMessage(pHwnd, PBM_GETSTEP, 0&, 0&): End Property
Public Property Let StepInterval(iStepInterval As Long): Call SendMessage(pHwnd, PBM_SETSTEP, ByVal iStepInterval, 0&): End Property

Public Property Get Target() As TextBox: Set Target = pTarget: End Property
Public Property Set Target(iTarget As TextBox)
Set pTarget = iTarget

'Set initial dimensions/coordinates of control based on target TextBox placed on form
pLeft = ScaleX(pTarget.Left, vbTwips, vbPixels)
pTop = ScaleY(pTarget.Top, vbTwips, vbPixels)
pWidth = ScaleX(pTarget.Width, vbTwips, vbPixels)
pHeight = ScaleY(pTarget.Height, vbTwips, vbPixels)

If TypeOf pTarget.Parent Is Form Then: pHwndParent = fGetClientHandle(pTarget.Parent)

iTarget.Visible = False
End Property

Public Property Get Top() As Integer: Top = pTop: End Property
Public Property Let Top(iTop As Integer): pTop = CInt(ScaleY(CDbl(iTop), vbTwips, vbPixels)): Call Me.Refresh: End Property

Public Property Get Value() As Long: Value = SendMessage(pHwnd, PBM_GETPOS, 0&, 0&): End Property
Public Property Let Value(iValue As Long): Call SendMessage(pHwnd, PBM_SETPOS, ByVal iValue, 0&): End Property

Public Property Get Visible() As Boolean: Visible = pVisible: End Property
Public Property Let Visible(iVisible As Boolean): pVisible = iVisible: Call Me.Refresh: End Property

Public Property Get Width() As Integer: Width = pWidth: End Property
Public Property Let Width(iWidth As Integer): pWidth = CInt(ScaleX(CDbl(iWidth), vbTwips, vbPixels)): Call Me.Refresh: End Property

'###########################################################################################################
'###                                Public Functions/Subroutines                                         ###
'###########################################################################################################

'***********************************************************************************************************
'***    Draws the control on the userform
'***********************************************************************************************************
Public Function Draw(Optional hWndTarget As OLE_HANDLE) As OLE_HANDLE
If Not hWndTarget = 0 Then: pHwndParent = hWndTarget
If pHwndParent = 0 Then: Exit Function

Call Me.Refresh

Draw = pHwnd
End Function

'***********************************************************************************************************
'***    Increases/decreases the value of progress bar by passed value
'***********************************************************************************************************
Public Function DeltaPos(Value As Long) As Long: DeltaPos = SendMessage(pHwnd, PBM_DELTAPOS, ByVal Value, 0&): End Function

'***********************************************************************************************************
'***    Retain values and redraw the control
'***********************************************************************************************************
Public Sub Refresh()
Dim lStyle As Long, lExStyle As Long
Dim tBackColor As OLE_COLOR
Dim tBarColor As OLE_COLOR
Dim tMax As Long
Dim tMin As Long
Dim tStepInterval As Long
Dim tValue As Long
Dim tState As eStateConstants

'Backup values before redraw
tBackColor = Me.BackColor
tBarColor = Me.BarColor
tMax = Me.Max
tMin = Me.Min
tStepInterval = Me.StepInterval
tValue = Me.Value
tState = Me.State

'Destroy previous control
Call DestroyWindow(pHwnd)

'Redraw the control
lStyle = lStyle Or WS_CHILD Or (-pVisible * WS_VISIBLE) Or WS_OVERLAPPED Or (pBorderStyle * WS_BORDER) Or (pOrientation * PBS_VERTICAL) Or _
         (pScrolling * PBS_SMOOTH) Or (pScrolling * PBS_SMOOTHREVERSE) Or (-pMarquee * PBS_MARQUEE)
                               
lExStyle = lExStyle Or (pAppearance * WS_EX_CLIENTEDGE) Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR

pHwnd = CreateWindowEx(ByVal lExStyle, PROGRESS_CLASS, vbNullString, ByVal lStyle, pLeft, pTop, pWidth, pHeight, pHwndParent, 0, 0, vbNullString)

Call SendMessage(pHwnd, PBM_SETMARQUEE, ByVal pMarquee, ByVal 0&)

'Restore former values
Me.BackColor = tBackColor
Me.BarColor = tBarColor
Me.Max = tMax
Me.Min = tMin
Me.StepInterval = tStepInterval
Me.Value = tValue
Me.State = tState

End Sub

'***********************************************************************************************************
'***    Increases/decreases the value of progress bar by the step interval
'***********************************************************************************************************
Public Function StepIt() As Long: StepIt = SendMessage(pHwnd, PBM_STEPIT, 0&, 0&): End Function

'###########################################################################################################
'###    Private Functions/Subroutines
'###########################################################################################################

'***********************************************************************************************************
'***    Class initialization
'***********************************************************************************************************
Private Sub Class_Initialize()
'Set properties to default values
pBorderStyle = ccNone
pAppearance = ccFlat
pMarquee = False
pOrientation = ccOrientationHorizontal
pScrolling = ccScrollingSmooth
pVisible = True

'Create a dummy control
pHwnd = CreateWindowEx(0&, PROGRESS_CLASS, vbNullString, 0&, 0, 0, 0, 0, pHwndParent, 0, 0, vbNullString)
End Sub

'***********************************************************************************************************
'***    Class termination
'***********************************************************************************************************
Private Sub Class_Terminate()
Call DestroyWindow(pHwnd)
End Sub

Private Function ScaleX(Units As Double, ScaleFrom As Double, ScaleTo As Double) As Double: ScaleX = asy("X", Units, ScaleFrom, ScaleTo): End Function
Private Function ScaleY(Units As Double, ScaleFrom As Double, ScaleTo As Double) As Double: ScaleY = asy("Y", Units, ScaleFrom, ScaleTo): End Function

Private Function asy(S As String, Q As Double, U0 As Double, U1 As Double) As Double
Dim vbapix As Single, k As Integer

vbapix = tpp(S)
k = IIf(S = "X", 1, 2)
asy = IIf(U0 = vbPixels, Q * vbapix, Q * U0)
asy = IIf(U1 = vbPixels, asy / vbapix, asy / U1)

End Function

Private Function tpp(S As String) As Single
Dim axe As Long, gdc As Long
axe = IIf(S = "X", 88, 90)
gdc = GetDC(0): tpp = vbInches / GetDeviceCaps(gdc, axe)
ReleaseDC 0, gdc

End Function
