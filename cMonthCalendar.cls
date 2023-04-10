VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMonthCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###########################################################################################################
'###########################################################################################################
'###
'###                                       Month Calendar API Control
'###        Author:         Brian Phillips
'###        Date:           07/04/2023  (v0.1)
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
Private Const MCS_DAYSTATE As Long = &H1
Private Const MCS_MULTISELECT As Long = &H2
Private Const MCS_WEEKNUMBERS As Long = &H4
Private Const MCS_NOTODAYCIRCLE As Long = &H8
Private Const MCS_NOTODAY As Long = &H10
Private Const MCS_NOTRAILINGDATES As Long = &H40
Private Const MCS_SHORTDAYSOFWEEK As Long = &H80
Private Const MCS_NOSELCHANGEONNAV As Long = &H100

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

'Private Const MCM_FIRST As Long = &H1000
'Private Const MCM_GETCOLOR As Long = (MCM_FIRST + 11)
'Private Const MCM_GETCURSEL As Long = (MCM_FIRST + 1)
'Private Const MCM_GETFIRSTDAYOFWEEK As Long = (MCM_FIRST + 16)
'Private Const MCM_GETMAXSELCOUNT As Long = (MCM_FIRST + 3)
'Private Const MCM_GETMAXTODAYWIDTH As Long = (MCM_FIRST + 21)
'Private Const MCM_GETMINREQRECT As Long = (MCM_FIRST + 9)
'Private Const MCM_GETMONTHDELTA As Long = (MCM_FIRST + 19)
'Private Const MCM_GETMONTHRANGE As Long = (MCM_FIRST + 7)
'Private Const MCM_GETRANGE As Long = (MCM_FIRST + 17)
'Private Const MCM_GETSELRANGE As Long = (MCM_FIRST + 5)
'Private Const MCM_GETUNICODEFORMAT As Long = CCM_GETUNICODEFORMAT
'Private Const MCM_HITTEST As Long = (MCM_FIRST + 14)
'Private Const MCM_SETCOLOR As Long = (MCM_FIRST + 10)
'Private Const MCM_SETCURSEL As Long = (MCM_FIRST + 2)
'Private Const MCM_SETDAYSTATE As Long = (MCM_FIRST + 8)
'Private Const MCM_SETFIRSTDAYOFWEEK As Long = (MCM_FIRST + 15)
'Private Const MCM_SETMAXSELCOUNT As Long = (MCM_FIRST + 4)
'Private Const MCM_SETMONTHDELTA As Long = (MCM_FIRST + 20)
'Private Const MCM_SETRANGE As Long = (MCM_FIRST + 18)
'Private Const MCM_SETSELRANGE As Long = (MCM_FIRST + 6)
'Private Const MCM_SETTODAY As Long = (MCM_FIRST + 12)
'Private Const MCM_SETUNICODEFORMAT As Long = CCM_SETUNICODEFORMAT

'Misc Constants
Private Const MONTHCAL_CLASS As String = "SysMonthCal32"

'Constants for Scale functions
Private Const vbPixels As Integer = 999
Private Const vbInches As Integer = 1440
Private Const vbTwips As Integer = 1

'###########################################################################################################
'###    Enum/Type Declarations
'###########################################################################################################
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Enum eMonthCalDayConstants
    Monday = 0
    Tuesday = 1
    Wednesday = 2
    Thursday = 3
    Friday = 4
    Saturday = 5
    Sunday = 6
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


Private pMaxDate As Date
Private pMinDate As Date
Private pShowToday As Boolean
Private pShowTodayCircle As Boolean
Private pShowWeekNumbers As Boolean

Private pToday As Date




Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
