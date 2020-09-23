Attribute VB_Name = "winuser"
Option Explicit


Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As DEVMODE, ByVal dwFlags As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (lpszDeviceName As Any, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetCaretBlinkTime Lib "user32" () As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDoubleClickTime Lib "user32" () As Long
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowUnicode Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Public Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long


    Public ICONMETRICS As ICONMETRICS
    Public Type ICONMETRICS
        cbSize As Long
        iHorzSpacing As Long
        iVertSpacing As Long
        iTitleWrap As Long
        lfFont As LOGFONT
    End Type
    
    Public STICKYKEYS As STICKYKEYS
    Public Type STICKYKEYS
        cbSize As Long
        dwFlags As Long
    End Type
    
    Public WINDOWPLACEMENT As WINDOWPLACEMENT
    Public Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
    End Type

    Public Const ARW_BOTTOMLEFT = &H0
    Public Const ARW_BOTTOMRIGHT = &H1
    Public Const ARW_TOPLEFT = &H2
    Public Const ARW_TOPRIGHT = &H3
    Public Const ARW_STARTMASK = &H3
    Public Const ARW_STARTRIGHT = &H1
    Public Const ARW_STARTTOP = &H2

    Public Const ARW_LEFT = &H0
    Public Const ARW_RIGHT = &H0
    Public Const ARW_UP = &H4
    Public Const ARW_DOWN = &H4
    Public Const ARW_HIDE = &H8

    Public Const CDS_UPDATEREGISTRY = &H1
    Public Const CDS_TEST = &H2
    Public Const CDS_FULLSCREEN = &H4
    Public Const CDS_GLOBAL = &H8
    Public Const CDS_SET_PRIMARY = &H10
    Public Const CDS_VIDEOPARAMETERS = &H20
    Public Const CDS_RESET = &H40000000
    Public Const CDS_NORESET = &H10000000

    Public Const DISP_CHANGE_SUCCESSFUL = 0
    Public Const DISP_CHANGE_RESTART = 1
    Public Const DISP_CHANGE_FAILED = -1
    Public Const DISP_CHANGE_BADMODE = -2
    Public Const DISP_CHANGE_NOTUPDATED = -3
    Public Const DISP_CHANGE_BADFLAGS = -4
    Public Const DISP_CHANGE_BADPARAM = -5

    Public Const GW_HWNDFIRST = 0
    Public Const GW_HWNDLAST = 1
    Public Const GW_HWNDNEXT = 2
    Public Const GW_HWNDPREV = 3
    Public Const GW_OWNER = 4
    Public Const GW_CHILD = 5
    Public Const GW_ENABLEDPOPUP = 6
    Public Const GW_MAX = 6

    Public Const GWL_WNDPROC = (-4)

    Public Const SKF_STICKYKEYSON = &H1
    Public Const SKF_AVAILABLE = &H2
    Public Const SKF_HOTKEYACTIVE = &H4
    Public Const SKF_CONFIRMHOTKEY = &H8
    Public Const SKF_HOTKEYSOUND = &H10
    Public Const SKF_INDICATOR = &H20
    Public Const SKF_AUDIBLEFEEDBACK = &H40
    Public Const SKF_TRISTATE = &H80
    Public Const SKF_TWOKEYSOFF = &H100
    Public Const SKF_LALTLATCHED = &H10000000
    Public Const SKF_LCTLLATCHED = &H4000000
    Public Const SKF_LSHIFTLATCHED = &H1000000
    Public Const SKF_RALTLATCHED = &H20000000
    Public Const SKF_RCTLLATCHED = &H8000000
    Public Const SKF_RSHIFTLATCHED = &H2000000
    Public Const SKF_LWINLATCHED = &H40000000
    Public Const SKF_RWINLATCHED = &H80000000
    Public Const SKF_LALTLOCKED = &H100000
    Public Const SKF_LCTLLOCKED = &H40000
    Public Const SKF_LSHIFTLOCKED = &H10000
    Public Const SKF_RALTLOCKED = &H200000
    Public Const SKF_RCTLLOCKED = &H80000
    Public Const SKF_RSHIFTLOCKED = &H20000
    Public Const SKF_LWINLOCKED = &H400000
    Public Const SKF_RWINLOCKED = &H800000
    
    Public Const SPI_GETBEEP = 1
    Public Const SPI_SETBEEP = 2
    Public Const SPI_GETMOUSE = 3
    Public Const SPI_SETMOUSE = 4
    Public Const SPI_GETBORDER = 5
    Public Const SPI_SETBORDER = 6
    Public Const SPI_GETKEYBOARDSPEED = 10
    Public Const SPI_SETKEYBOARDSPEED = 11
    Public Const SPI_LANGDRIVER = 12
    Public Const SPI_ICONHORIZONTALSPACING = 13
    Public Const SPI_GETSCREENSAVETIMEOUT = 14
    Public Const SPI_SETSCREENSAVETIMEOUT = 15
    Public Const SPI_GETSCREENSAVEACTIVE = 16
    Public Const SPI_SETSCREENSAVEACTIVE = 17
    Public Const SPI_GETGRIDGRANULARITY = 18
    Public Const SPI_SETGRIDGRANULARITY = 19
    Public Const SPI_SETDESKWALLPAPER = 20
    Public Const SPI_SETDESKPATTERN = 21
    Public Const SPI_GETKEYBOARDDELAY = 22
    Public Const SPI_SETKEYBOARDDELAY = 23
    Public Const SPI_ICONVERTICALSPACING = 24
    Public Const SPI_GETICONTITLEWRAP = 25
    Public Const SPI_SETICONTITLEWRAP = 26
    Public Const SPI_GETMENUDROPALIGNMENT = 27
    Public Const SPI_SETMENUDROPALIGNMENT = 28
    Public Const SPI_SETDOUBLECLKWIDTH = 29
    Public Const SPI_SETDOUBLECLKHEIGHT = 30
    Public Const SPI_GETICONTITLELOGFONT = 31
    Public Const SPI_SETDOUBLECLICKTIME = 32
    Public Const SPI_SETMOUSEBUTTONSWAP = 33
    Public Const SPI_SETICONTITLELOGFONT = 34
    Public Const SPI_GETFASTTASKSWITCH = 35
    Public Const SPI_SETFASTTASKSWITCH = 36
    Public Const SPI_SETDRAGFULLWINDOWS = 37
    Public Const SPI_GETDRAGFULLWINDOWS = 38
    Public Const SPI_GETNONCLIENTMETRICS = 41
    Public Const SPI_SETNONCLIENTMETRICS = 42
    Public Const SPI_GETMINIMIZEDMETRICS = 43
    Public Const SPI_SETMINIMIZEDMETRICS = 44
    Public Const SPI_GETICONMETRICS = 45
    Public Const SPI_SETICONMETRICS = 46
    Public Const SPI_SETWORKAREA = 47
    Public Const SPI_GETWORKAREA = 48
    Public Const SPI_SETPENWINDOWS = 49
    Public Const SPI_GETFILTERKEYS = 50
    Public Const SPI_SETFILTERKEYS = 51
    Public Const SPI_GETTOGGLEKEYS = 52
    Public Const SPI_SETTOGGLEKEYS = 53
    Public Const SPI_GETMOUSEKEYS = 54
    Public Const SPI_SETMOUSEKEYS = 55
    Public Const SPI_GETSHOWSOUNDS = 56
    Public Const SPI_SETSHOWSOUNDS = 57
    Public Const SPI_GETSTICKYKEYS = 58
    Public Const SPI_SETSTICKYKEYS = 59
    Public Const SPI_GETACCESSTIMEOUT = 60
    Public Const SPI_SETACCESSTIMEOUT = 61
    Public Const SPI_GETSERIALKEYS = 62
    Public Const SPI_SETSERIALKEYS = 63
    Public Const SPI_GETSOUNDSENTRY = 64
    Public Const SPI_SETSOUNDSENTRY = 65
    Public Const SPI_GETHIGHCONTRAST = 66
    Public Const SPI_SETHIGHCONTRAST = 67
    Public Const SPI_GETKEYBOARDPREF = 68
    Public Const SPI_SETKEYBOARDPREF = 69
    Public Const SPI_GETSCREENREADER = 70
    Public Const SPI_SETSCREENREADER = 71
    Public Const SPI_GETANIMATION = 72
    Public Const SPI_SETANIMATION = 73
    Public Const SPI_GETFONTSMOOTHING = 74
    Public Const SPI_SETFONTSMOOTHING = 75
    Public Const SPI_SETDRAGWIDTH = 76
    Public Const SPI_SETDRAGHEIGHT = 77
    Public Const SPI_SETHANDHELD = 78
    Public Const SPI_GETLOWPOWERTIMEOUT = 79
    Public Const SPI_GETPOWEROFFTIMEOUT = 80
    Public Const SPI_SETLOWPOWERTIMEOUT = 81
    Public Const SPI_SETPOWEROFFTIMEOUT = 82
    Public Const SPI_GETLOWPOWERACTIVE = 83
    Public Const SPI_GETPOWEROFFACTIVE = 84
    Public Const SPI_SETLOWPOWERACTIVE = 85
    Public Const SPI_SETPOWEROFFACTIVE = 86
    Public Const SPI_SETCURSORS = 87
    Public Const SPI_SETICONS = 88
    Public Const SPI_GETDEFAULTINPUTLANG = 89
    Public Const SPI_SETDEFAULTINPUTLANG = 90
    Public Const SPI_SETLANGTOGGLE = 91
    Public Const SPI_GETWINDOWSEXTENSION = 92
    Public Const SPI_SETMOUSETRAILS = 93
    Public Const SPI_GETMOUSETRAILS = 94
    Public Const SPI_GETSNAPTODEFBUTTON = 95
    Public Const SPI_SETSNAPTODEFBUTTON = 96
    Public Const SPI_SETSCREENSAVERRUNNING = 97
    Public Const SPI_SCREENSAVERRUNNING = SPI_SETSCREENSAVERRUNNING
    Public Const SPI_GETMOUSEHOVERWIDTH = 98
    Public Const SPI_SETMOUSEHOVERWIDTH = 99
    Public Const SPI_GETMOUSEHOVERHEIGHT = 100
    Public Const SPI_SETMOUSEHOVERHEIGHT = 101
    Public Const SPI_GETMOUSEHOVERTIME = 102
    Public Const SPI_SETMOUSEHOVERTIME = 103
    Public Const SPI_GETWHEELSCROLLLINES = 104
    Public Const SPI_SETWHEELSCROLLLINES = 105
    Public Const SPI_GETMENUSHOWDELAY = 106
    Public Const SPI_SETMENUSHOWDELAY = 107
    Public Const SPI_GETSHOWIMEUI = 110
    Public Const SPI_SETSHOWIMEUI = 111
    Public Const SPI_GETMOUSESPEED = 112
    Public Const SPI_SETMOUSESPEED = 113
    Public Const SPI_GETSCREENSAVERRUNNING = 114
    Public Const SPI_GETDESKWALLPAPER = 115
    
    Public Const SPI_GETACTIVEWINDOWTRACKING = &H1000
    Public Const SPI_SETACTIVEWINDOWTRACKING = &H1001
    Public Const SPI_GETMENUANIMATION = &H1002
    Public Const SPI_SETMENUANIMATION = &H1003
    Public Const SPI_GETCOMBOBOXANIMATION = &H1004
    Public Const SPI_SETCOMBOBOXANIMATION = &H1005
    Public Const SPI_GETLISTBOXSMOOTHSCROLLING = &H1006
    Public Const SPI_SETLISTBOXSMOOTHSCROLLING = &H1007
    Public Const SPI_GETGRADIENTCAPTIONS = &H1008
    Public Const SPI_SETGRADIENTCAPTIONS = &H1009
    Public Const SPI_GETKEYBOARDCUES = &H100A
    Public Const SPI_SETKEYBOARDCUES = &H100B
    Public Const SPI_GETMENUUNDERLINES = SPI_GETKEYBOARDCUES
    Public Const SPI_SETMENUUNDERLINES = SPI_SETKEYBOARDCUES
    Public Const SPI_GETACTIVEWNDTRKZORDER = &H100C
    Public Const SPI_SETACTIVEWNDTRKZORDER = &H100D
    Public Const SPI_GETHOTTRACKING = &H100E
    Public Const SPI_SETHOTTRACKING = &H100F
    Public Const SPI_GETMENUFADE = &H1012
    Public Const SPI_SETMENUFADE = &H1013
    Public Const SPI_GETSELECTIONFADE = &H1014
    Public Const SPI_SETSELECTIONFADE = &H1015
    Public Const SPI_GETTOOLTIPANIMATION = &H1016
    Public Const SPI_SETTOOLTIPANIMATION = &H1017
    Public Const SPI_GETTOOLTIPFADE = &H1018
    Public Const SPI_SETTOOLTIPFADE = &H1019
    Public Const SPI_GETCURSORSHADOW = &H101A
    Public Const SPI_SETCURSORSHADOW = &H101B
    Public Const SPI_GETUIEFFECTS = &H103E
    Public Const SPI_SETUIEFFECTS = &H103F
    Public Const SPI_GETFOREGROUNDLOCKTIMEOUT = &H2000
    Public Const SPI_SETFOREGROUNDLOCKTIMEOUT = &H2001
    Public Const SPI_GETACTIVEWNDTRKTIMEOUT = &H2002
    Public Const SPI_SETACTIVEWNDTRKTIMEOUT = &H2003
    Public Const SPI_GETFOREGROUNDFLASHCOUNT = &H2004
    Public Const SPI_SETFOREGROUNDFLASHCOUNT = &H2005
    Public Const SPI_GETCARETWIDTH = &H2006
    Public Const SPI_SETCARETWIDTH = &H2007

    Public Const SPIF_UPDATEINIFILE = &H1
    Public Const SPIF_SENDWININICHANGE = &H2
    Public Const SPIF_SENDCHANGE = SPIF_SENDWININICHANGE

    Public Const SM_CXSCREEN = 0
    Public Const SM_CYSCREEN = 1
    Public Const SM_CXVSCROLL = 2
    Public Const SM_CYHSCROLL = 3
    Public Const SM_CYCAPTION = 4
    Public Const SM_CXBORDER = 5
    Public Const SM_CYBORDER = 6
    Public Const SM_CXDLGFRAME = 7
    Public Const SM_CYDLGFRAME = 8
    Public Const SM_CYVTHUMB = 9
    Public Const SM_CXHTHUMB = 10
    Public Const SM_CXICON = 11
    Public Const SM_CYICON = 12
    Public Const SM_CXCURSOR = 13
    Public Const SM_CYCURSOR = 14
    Public Const SM_CYMENU = 15
    Public Const SM_CXFULLSCREEN = 16
    Public Const SM_CYFULLSCREEN = 17
    Public Const SM_CYKANJIWINDOW = 18
    Public Const SM_MOUSEPRESENT = 19
    Public Const SM_CYVSCROLL = 20
    Public Const SM_CXHSCROLL = 21
    Public Const SM_DEBUG = 22
    Public Const SM_SWAPBUTTON = 23
    Public Const SM_RESERVED1 = 24
    Public Const SM_RESERVED2 = 25
    Public Const SM_RESERVED3 = 26
    Public Const SM_RESERVED4 = 27
    Public Const SM_CXMIN = 28
    Public Const SM_CYMIN = 29
    Public Const SM_CXSIZE = 30
    Public Const SM_CYSIZE = 31
    Public Const SM_CXFRAME = 32
    Public Const SM_CYFRAME = 33
    Public Const SM_CXMINTRACK = 34
    Public Const SM_CYMINTRACK = 35
    Public Const SM_CXDOUBLECLK = 36
    Public Const SM_CYDOUBLECLK = 37
    Public Const SM_CXICONSPACING = 38
    Public Const SM_CYICONSPACING = 39
    Public Const SM_MENUDROPALIGNMENT = 40
    Public Const SM_PENWINDOWS = 41
    Public Const SM_DBCSENABLED = 42
    Public Const SM_CMOUSEBUTTONS = 43
    Public Const SM_CMETRICS = 44
    Public Const SM_CXSIZEFRAME = SM_CXFRAME
    Public Const SM_CYSIZEFRAME = SM_CYFRAME
    Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
    Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
    Public Const SM_SECURE = 44
    Public Const SM_CXEDGE = 45
    Public Const SM_CYEDGE = 46
    Public Const SM_CXMINSPACING = 47
    Public Const SM_CYMINSPACING = 48
    Public Const SM_CXSMICON = 49
    Public Const SM_CYSMICON = 50
    Public Const SM_CYSMCAPTION = 51
    Public Const SM_CXSMSIZE = 52
    Public Const SM_CYSMSIZE = 53
    Public Const SM_CXMENUSIZE = 54
    Public Const SM_CYMENUSIZE = 55
    Public Const SM_ARRANGE = 56
    Public Const SM_CXMINIMIZED = 57
    Public Const SM_CYMINIMIZED = 58
    Public Const SM_CXMAXTRACK = 59
    Public Const SM_CYMAXTRACK = 60
    Public Const SM_CXMAXIMIZED = 61
    Public Const SM_CYMAXIMIZED = 62
    Public Const SM_NETWORK = 63
    Public Const SM_CLEANBOOT = 67
    Public Const SM_CXDRAG = 68
    Public Const SM_CYDRAG = 69
    Public Const SM_SHOWSOUNDS = 70
    Public Const SM_CXMENUCHECK = 71   ' Use instead of GetMenuCheckMarkDimensions()!
    Public Const SM_CYMENUCHECK = 72
    Public Const SM_SLOWMACHINE = 73
    Public Const SM_MIDEASTENABLED = 74
    Public Const SM_MOUSEWHEELPRESENT = 75
    Public Const SM_XVIRTUALSCREEN = 76
    Public Const SM_YVIRTUALSCREEN = 77
    Public Const SM_CXVIRTUALSCREEN = 78
    Public Const SM_CYVIRTUALSCREEN = 79
    Public Const SM_CMONITORS = 80
    Public Const SM_SAMEDISPLAYFORMAT = 81
    Public Const SM_IMMENABLED = 82

    Public Const SW_HIDE = 0
    Public Const SW_SHOWNORMAL = 1
    Public Const SW_NORMAL = 1
    Public Const SW_SHOWMINIMIZED = 2
    Public Const SW_SHOWMAXIMIZED = 3
    Public Const SW_MAXIMIZE = 3
    Public Const SW_SHOWNOACTIVATE = 4
    Public Const SW_SHOW = 5
    Public Const SW_MINIMIZE = 6
    Public Const SW_SHOWMINNOACTIVE = 7
    Public Const SW_SHOWNA = 8
    Public Const SW_RESTORE = 9
    Public Const SW_SHOWDEFAULT = 10
    Public Const SW_FORCEMINIMIZE = 11
    Public Const SW_MAX = 11

    Public Const WH_MIN = (-1)
    Public Const WH_MSGFILTER = (-1)
    Public Const WH_JOURNALRECORD = 0
    Public Const WH_JOURNALPLAYBACK = 1
    Public Const WH_KEYBOARD = 2
    Public Const WH_GETMESSAGE = 3
    Public Const WH_CALLWNDPROC = 4
    Public Const WH_CBT = 5
    Public Const WH_SYSMSGFILTER = 6
    Public Const WH_MOUSE = 7
    Public Const WH_HARDWARE = 8
    Public Const WH_DEBUG = 9
    Public Const WH_SHELL = 10
    Public Const WH_FOREGROUNDIDLE = 11
    Public Const WH_CALLWNDPROCRET = 12
    'Public Const WH_KEYBOARD_LL = 13
    'Public Const WH_MOUSE_LL = 14
    
    Public Const WM_CLOSE = &H10
    Public Const WM_DESTROY = &H2
    Public Const WM_LBUTTONDBLCLK = &H203
    Public Const WM_LBUTTONDOWN = &H201
    Public Const WM_LBUTTONUP = &H202
    Public Const WM_MBUTTONDBLCLK = &H209
    Public Const WM_MBUTTONDOWN = &H207
    Public Const WM_MBUTTONUP = &H208
    Public Const WM_MDIDESTROY = &H221
    Public Const WM_NCDESTROY = &H82
    Public Const WM_RBUTTONDBLCLK = &H206
    Public Const WM_RBUTTONDOWN = &H204
    Public Const WM_RBUTTONUP = &H205
    Public Const WM_SETTEXT = &HC
    
    Public Const WM_MOUSEMOVE = &H200
    
    Public Const WPF_ASYNCWINDOWPLACEMENT = &H4
    Public Const WPF_RESTORETOMAXIMIZED = &H2
    Public Const WPF_SETMINPOSITION = &H1


Public Function Get_KeyboardType() As String
    'Sub type based on number, 1 to 7 types
    Select Case GetKeyboardType(0)
        Case 0: Get_KeyboardType = "Unknown / Not Specified"
        Case 1: Get_KeyboardType = "IBM PC/XT ( ) or compatible (83-key)"
        Case 2: Get_KeyboardType = "Olivetti ICO (102-key) keyboard"
        Case 3: Get_KeyboardType = "IBM PC/AT (84-key) or similar"
        Case 4: Get_KeyboardType = "IBM enhanced (101- or 102-key)"
        Case 5: Get_KeyboardType = "Nokia 1050 and similar"
        Case 6: Get_KeyboardType = "Nokia 9140 and similar"
        Case 7: Get_KeyboardType = "Japanese"
        Case Else: Get_KeyboardType = "Unknown / Not Specified"
    End Select
End Function

Public Function Get_KeyboardFuncKeys() As String
    'Gives functionkey # based on number, 1 to 7 types
    Select Case GetKeyboardType(2)
        Case 0:
            apiError = 0
            Failed "GetKeyboardType"
        Case 1: Get_KeyboardFuncKeys = "10"
        Case 2: Get_KeyboardFuncKeys = "12/18"
        Case 3: Get_KeyboardFuncKeys = "10"
        Case 4: Get_KeyboardFuncKeys = "12"
        Case 5: Get_KeyboardFuncKeys = "10"
        Case 6: Get_KeyboardFuncKeys = "24"
        Case 7: Get_KeyboardFuncKeys = "10"
        Case Else: Get_KeyboardFuncKeys = "Hardware dependent and specified by the OEM"
    End Select
End Function

Public Function Get_KeyboardLayout() As String
    Dim KeyboardLayout As String * 9 'Only 8 bytes long + 1 for identifier
    If GetKeyboardLayoutName(KeyboardLayout) = 0 Then
        Failed "GetKeyboardLayoutName"
    Else
        Get_KeyboardLayout = Fix_NullTermStr(KeyboardLayout)
    End If
End Function

