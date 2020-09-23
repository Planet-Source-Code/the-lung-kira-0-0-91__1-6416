Attribute VB_Name = "Module"
Option Explicit

    Public apiError As Long
    Public errMsg As Boolean
    Public ComputerName As String
    Public UserName As String
    Public WinVer As Long
    Public WinID As String
    
    Public Dirs As Dirs
    Public Type Dirs
        AdminTools As String
        AppData As String
        AppPath As String
        Cache As String
        CommonFiles As String
        Cookies As String
        Current As String
        Desktop As String
        Favorites As String
        Fonts As String
        History As String
        LocalAppData As String
        MediaPath As String
        MyPictures As String
        NetHood As String
        Personal As String
        PrintHood As String
        Programs As String
        Recent As String
        sendto As String
        StartMenu As String
        Startup As String
        System As String
        Temp As String
        Templates As String
        Windows As String
    End Type

    Public GUID As GUID
    Public Type GUID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(8) As Byte
    End Type
    
    Public MpgInfo As MpgInfo
    Public Type MpgInfo
        Sync As String
        Version As String
        Layer As Byte
        Error_Protection As Integer 'Bool
        Bitrate_Index As Integer
        Sampling_Freq As Long
        Padding As String
        Extension As String
        Mode As String
        Mode_Extn As String
        Copyright As Integer 'Bool
        Original As Integer 'Bool
        Emphasis As String
    End Type
    
    Public MpgTag As MpgTag
    Public Type MpgTag
        Tag As Boolean
        Title As String * 30
        Artist As String * 30
        Album As String * 30
        Year As String * 4
        Comments As String * 30
        Genre As Byte
    End Type

    'Holds info which is not currently displayed
    Public WinsockData As WinsockData
    Public Type WinsockData
        Description As String
        SystemStatus As String
    End Type
'--------------------------------------
    Public WindowListName() As String
    Public WindowListhWnd() As Long
    Public WindowListNum As Long
    Public MouseMovX As Double
    Public MouseMovY As Double
    Public MouseMovTmpX As Integer
    Public MouseMovTmpY As Integer
    Public MouseWarp As Double
    Public WindowKiller() As String
    Public WindowKillerNum As Long
    
    Public PingNumber As Integer
    Public PingTimeout As Integer
    Public PingTTL As Byte
    Public PerfMonInterval As Integer
        
    Public ScreenEdge As ScreenEdge
    Public Type ScreenEdge
        X As Integer
        Y As Integer
    End Type
    

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim tmpString As String * 512
    
    apiError = GetWindowText(hwnd, tmpString, 512)
        
    ReDim Preserve WindowListName(WindowListNum) 'Resize arrays
    ReDim Preserve WindowListhWnd(WindowListNum)
    WindowListName(WindowListNum) = Fix_NullTermStr(tmpString)
    WindowListhWnd(WindowListNum) = hwnd
        
    WindowListNum = WindowListNum + 1 'Increment
    
    EnumWindowsProc = 1
End Function

Public Function App_Shutdown()
    WinSockEnd 'Shutdown winsock
    
    'Clean up hook
    If UnhookWindowsHookEx(mh_Hook) = 0 Then Failed "UnhookWindowsHookEx"
    If FreeLibrary(mh_Instance) = 0 Then Failed "FreeLibrary"
    
    'Write data to file
    Open Dirs.AppPath & "\main.dat" For Output As #1
        'Need a # not true false
        Print #1, CByte(frmMain.mnuMouseMovOO.Checked)
        Print #1, CByte(frmMain.mnuMouseWarpOO.Checked)
        Print #1, CByte(frmMain.mnuWindowKillerOO.Checked)
        
        Print #1, MouseMovX
        Print #1, MouseMovY
        Print #1, MouseWarp
        Print #1, PingNumber
        Print #1, PingTimeout
        Print #1, PingTTL
        Print #1, frmMain.timerWindowKiller.Interval
        Print #1, PerfMonInterval
        Print #1, CByte(errMsg)
    Close #1
    
    If WindowKillerNum > 0 Then 'Dont write no data
        Dim tmpInt As Integer
        Dim tmpArray() As String
        Dim tmpCount As Integer
    
        ReDim Preserve tmpArray(1)
        tmpCount = 1
        tmpArray(1) = WindowKiller(1)
        
        For tmpInt = 1 To WindowKillerNum 'Cycle through WindowKiller array
            If Not tmpArray(tmpCount) = WindowKiller(tmpInt) Then 'If not duplicate
                tmpCount = tmpCount + 1 'Increment
                ReDim Preserve tmpArray(tmpCount) 'Resize array
                tmpArray(tmpCount) = WindowKiller(tmpInt) 'Add item
            End If
        Next tmpInt
    
        Open Dirs.AppPath & "\wk.dat" For Output As #1
            For tmpInt = 1 To tmpCount 'Cycle through tmpArray
                Print #1, tmpArray(tmpInt) 'Put in file
            Next tmpInt
        Close #1
    End If
    
    'Remove icon from system tray
    With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = frmMain.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = vbNull
        .hIcon = frmMain.Icon
        .szTip = Chr(0) 'Clear
    End With
    If Shell_NotifyIcon(NIM_DELETE, NOTIFYICONDATA) = 0 Then
        Failed "Shell_NotifyIcon"
    End If
    
    End
End Function

Public Function App_Startup()
    frmMain.Hide
    If App.PrevInstance = True Then End
    
    'Add icon to system tray
    With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = frmMain.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = 512
        .hIcon = frmMain.Icon
        .szTip = frmMain.Caption + Chr(0) 'Tooltip text
    End With
    If Shell_NotifyIcon(NIM_ADD, NOTIFYICONDATA) = 0 Then
        Failed "Shell_NotifyIcon"
    End If
    
    'Fill Variables
    ComputerName = Get_ComputerName
    UserName = Get_UserName
    
    'Startup Windows Info
    OsVersionInfo.dwOSVersionInfoSize = Len(OsVersionInfo) 'Size of the structure
    If GetVersionEx(OsVersionInfo) = 0 Then
        Failed "GetVersionEx"
    End If

    Select Case OsVersionInfo.dwPlatformId
        Case VER_PLATFORM_WIN32_NT: WinID = "WIN32_NT"
        Case VER_PLATFORM_WIN32_WINDOWS: WinID = "WIN32_WINDOWS"
        Case VER_PLATFORM_WIN32s: WinID = "WIN32s"
    End Select
    
    WinVer = Right("0" & OsVersionInfo.dwMajorVersion, 1) & _
                      Right("00" & OsVersionInfo.dwMinorVersion, 2) & _
                      Right("0000" & (OsVersionInfo.dwBuildNumber And &HFFFF&), 4)
    
    If cpuid_avail = 1 Then cpu_id 'Initializes function
    Directories
    WinSockStart 'Startup winsock 2.2
    InstallMouseHook
'--------------------------------------
    'Read all the settings from file
    If Dir(Dirs.AppPath & "\main.dat") <> "" Then 'Cant open a blank file
        Dim aryData() As String
        Get_Data Dirs.AppPath & "\main.dat", 12, aryData()
    
        frmMain.mnuMouseMovOO.Checked = aryData(1)
        frmMain.mnuMouseWarpOO.Checked = aryData(2)
        frmMain.mnuWindowKillerOO.Checked = aryData(3)
        MouseMovX = aryData(4)
        MouseMovY = aryData(5)
        MouseWarp = aryData(6)
    
        PingNumber = aryData(7)
        PingTimeout = aryData(8)
        PingTTL = aryData(9)
        frmMain.timerWindowKiller.Interval = aryData(10)
        PerfMonInterval = aryData(11)
        errMsg = aryData(12)
    Else
        'Set defaults
        frmPing.hsNumber.Value = 1
        frmPing.hsTimeout.Value = 5000
        frmPing.hsTTL.Value = 128
        frmMain.timerWindowKiller.Interval = 200
        frmPerfMon.hsInterval.Value = 1000
    End If
    
    If Dir(Dirs.AppPath & "\wk.dat") <> "" Then
        Get_Data Dirs.AppPath & "\wk.dat", 0, WindowKiller()
    End If
'--------------------------------------
    'Get it going
    If frmMain.mnuMouseMovOO.Checked = True Then
        GetCursorPos POINTAPI 'Dumps info to pointapi
    
        'Gives point of reference for starting
        MouseMovTmpX = POINTAPI.X
        MouseMovTmpY = POINTAPI.Y
    End If
    If frmMain.mnuMouseWarpOO.Checked = True Then
        'Converts the twips to pixels and sets the edges
        ScreenEdge.X = Screen.Width \ Screen.TwipsPerPixelX
        ScreenEdge.Y = Screen.Height \ Screen.TwipsPerPixelY
    End If
End Function

Public Function Directories()
    With Dirs
        .AdminTools = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Administrative Tools"))
        .AppData = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "AppData"))
        .AppPath = Fix_Dir(App.path)
        .Cache = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cache"))
        .CommonFiles = Fix_Dir(GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "CommonFilesDir"))
        .Cookies = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cookies"))
        .Current = Get_CurrentDirectory
        .Desktop = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop"))
        .Favorites = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Favorites"))
        .Fonts = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Fonts"))
        .History = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "History"))
        .LocalAppData = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Local AppData"))
        .MediaPath = Fix_Dir(GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "MediaPath"))
        .MyPictures = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "My Pictures"))
        .NetHood = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "NetHood"))
        .Personal = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal"))
        .PrintHood = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "PrintHood"))
        .Programs = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Programs"))
        .Recent = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Recent"))
        .sendto = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "SendTo"))
        .StartMenu = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Start Menu"))
        .Startup = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup"))
        .System = Get_SystemDirectory
        .Temp = Get_TempPath
        .Templates = Fix_Dir(GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Templates"))
        .Windows = Get_WindowsDirectory
    End With
End Function

