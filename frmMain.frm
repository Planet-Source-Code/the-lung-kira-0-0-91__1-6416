VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kira"
   ClientHeight    =   495
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   1020
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   1020
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHM 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer timerWindowKiller 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   480
      Top             =   0
   End
   Begin VB.Menu mnuMain 
      Caption         =   ""
      Begin VB.Menu mnuFileFormat 
         Caption         =   "File Formats"
         Begin VB.Menu mnuMpeg 
            Caption         =   "Mpeg"
         End
         Begin VB.Menu mnuNE 
            Caption         =   "New Executable"
            Begin VB.Menu mnuSFI_NE 
               Caption         =   "String File Info Editor NE"
            End
         End
         Begin VB.Menu mnuPE 
            Caption         =   "Portable Executable"
            Begin VB.Menu mnuSFI_PE 
               Caption         =   "String File Info Editor PE"
            End
         End
      End
      Begin VB.Menu mnuHardware 
         Caption         =   "Hardware"
         Begin VB.Menu mnuCmos 
            Caption         =   "Cmos Contents"
         End
         Begin VB.Menu mnuDrives 
            Caption         =   "Drives"
            Begin VB.Menu mnuDiskSpace 
               Caption         =   "Disk Space"
            End
            Begin VB.Menu mnuVolumeInfo 
               Caption         =   "Volume Info"
            End
         End
         Begin VB.Menu mnuMemoryStatus 
            Caption         =   "Memory Status"
         End
         Begin VB.Menu mnuPowerStatus 
            Caption         =   "Power Status"
         End
         Begin VB.Menu mnuProcessor 
            Caption         =   "Processor"
            Begin VB.Menu mnuCPUID 
               Caption         =   "CPUID"
            End
            Begin VB.Menu mnuProcessorInfo 
               Caption         =   "Processor Info"
            End
         End
      End
      Begin VB.Menu mnuInternetNetwork 
         Caption         =   "Internet / Network"
         Begin VB.Menu mnuGetIPHost 
            Caption         =   "Get IP/Host"
         End
         Begin VB.Menu mnuIP_Stats 
            Caption         =   "IP Stats"
         End
         Begin VB.Menu mnuICPM_Stats 
            Caption         =   "ICMP Stats"
         End
         Begin VB.Menu mnuNetworkInfo 
            Caption         =   "Network Info"
         End
         Begin VB.Menu mnuPing 
            Caption         =   "Ping"
         End
         Begin VB.Menu mnuTCP_Stats 
            Caption         =   "TCP Stats"
         End
         Begin VB.Menu mnuUDP_Stats 
            Caption         =   "UDP Stats"
         End
      End
      Begin VB.Menu mnuPeripherials 
         Caption         =   "Peripherials"
         Begin VB.Menu mnuKeyboard 
            Caption         =   "Keyboard"
            Begin VB.Menu mnuKeyboardSettings 
               Caption         =   "Keyboard Settings"
            End
            Begin VB.Menu mnuKeyboardInfo 
               Caption         =   "Keyboard Info"
            End
            Begin VB.Menu mnuStickyKeys 
               Caption         =   "Sticky Keys"
            End
         End
         Begin VB.Menu mnuDisplay 
            Caption         =   "Display"
            Begin VB.Menu mnuDisplaySettings 
               Caption         =   "Display Settings"
            End
            Begin VB.Menu mnuMonitorInfo 
               Caption         =   "Monitor Info"
            End
         End
         Begin VB.Menu mnuMouse 
            Caption         =   "Mouse"
            Begin VB.Menu mnuMouseInfo 
               Caption         =   "Mouse Info"
            End
            Begin VB.Menu mnuMouseMov 
               Caption         =   "Mouse Movements"
            End
            Begin VB.Menu mnuMouseSettings 
               Caption         =   "Mouse Settings"
            End
            Begin VB.Menu mnuMouseWarp 
               Caption         =   "Mouse Warp"
            End
         End
      End
      Begin VB.Menu mnuSoftware 
         Caption         =   "Software"
         Begin VB.Menu mnuIE5 
            Caption         =   "Internet Explorer 5"
            Begin VB.Menu mnuIECache 
               Caption         =   "Cache Hit/Miss"
            End
            Begin VB.Menu mnuIEHistory 
               Caption         =   "History Viewer"
            End
            Begin VB.Menu mnuIEOptions 
               Caption         =   "IE Options"
            End
         End
      End
      Begin VB.Menu mnuWin 
         Caption         =   "Windows"
         Begin VB.Menu mnuCachedPasswords 
            Caption         =   "Cached Passwords"
         End
         Begin VB.Menu mnuDirectories 
            Caption         =   "Directories"
         End
         Begin VB.Menu mnuErrors 
            Caption         =   "Errors"
         End
         Begin VB.Menu mnuFiles 
            Caption         =   "Files"
            Begin VB.Menu mnuFileAttributes 
               Caption         =   "File Attributes"
            End
            Begin VB.Menu mnuFileTime 
               Caption         =   "File Time"
            End
            Begin VB.Menu mnuSharedFiles 
               Caption         =   "Shared Files"
            End
         End
         Begin VB.Menu mnuIcons 
            Caption         =   "Icons"
            Begin VB.Menu mnuIconInfo 
               Caption         =   "Icon Info"
            End
            Begin VB.Menu mnuIconSettings 
               Caption         =   "Icon Settings"
            End
         End
         Begin VB.Menu mnuMenuSettings 
            Caption         =   "Menu Settings"
         End
         Begin VB.Menu mnuPerfMon 
            Caption         =   "Performance Monitor"
         End
         Begin VB.Menu mnuUpTime 
            Caption         =   "Up Time"
         End
         Begin VB.Menu mnuUser 
            Caption         =   "User"
            Begin VB.Menu mnuCompUserName 
               Caption         =   "Computer/User Name"
            End
            Begin VB.Menu mnuRegistered 
               Caption         =   "Registered To"
            End
         End
         Begin VB.Menu mnuWindowInfo 
            Caption         =   "Window Info"
         End
         Begin VB.Menu mnuWindowKiller 
            Caption         =   "Window Killer"
         End
         Begin VB.Menu mnuWindows 
            Caption         =   "Windows"
         End
         Begin VB.Menu mnuWinInfo 
            Caption         =   "Windows Info"
         End
      End
      Begin VB.Menu mnuBreak0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnOff 
         Caption         =   "On / Off"
         Begin VB.Menu mnuErrMsg 
            Caption         =   "Error Messages"
         End
         Begin VB.Menu mnuMouseMovOO 
            Caption         =   "Mouse Movements"
         End
         Begin VB.Menu mnuMouseWarpOO 
            Caption         =   "Mouse Warp"
         End
         Begin VB.Menu mnuWindowKillerOO 
            Caption         =   "Window Killer"
         End
      End
      Begin VB.Menu mnuBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Begin VB.Menu mnuAbout 
            Caption         =   "About"
         End
         Begin VB.Menu mnuAboutShell 
            Caption         =   "About Shell"
         End
      End
      Begin VB.Menu mnuBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenAll 
         Caption         =   "Open All"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close All"
      End
      Begin VB.Menu mnuBreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call App_Startup
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmpLong As Single
    tmpLong = X / Screen.TwipsPerPixelX
    
    Select Case tmpLong 'For system tray icon
        Case WM_LBUTTONUP
            apiError = SetForegroundWindow(Me.hwnd) 'Make sure its on top
            frmMain.PopupMenu mnuMain
        Case WM_RBUTTONUP
            apiError = SetForegroundWindow(Me.hwnd) 'Make sure its on top
            frmMain.PopupMenu mnuMain
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call App_Shutdown
End Sub

Private Sub mnuAbout_Click()
    frmMainAbout.Show
End Sub

Private Sub mnuAboutShell_Click()
    If ShellAbout(0&, "", "", Me.Icon) = 0 Then
        Failed "ShellAbout"
    End If
End Sub

Private Sub mnuCachedPasswords_Click()
    frmCachedPasswords.Show
End Sub

Private Sub mnuCloseAll_Click()
    'Close all manually
    Unload frmCachedPasswords
    Unload frmCmos
    Unload frmCompUser
    Unload frmCPUID
    Unload frmDirectories
    Unload frmDiskSpace
    Unload frmDiskVolume
    Unload frmDisplaySettings
    Unload frmErrors
    Unload frmFileAttributes
    Unload frmFileTime
    Unload frmGetIPHost
    Unload frmIconInfo
    Unload frmIconSettings
    Unload frmICMP_Stats
    Unload frmIECache
    Unload frmIEHistory
    Unload frmIEOptions
    Unload frmIP_Stats
    Unload frmKeyboardInfo
    Unload frmKeyboardSettings
    Unload frmMainAbout
    Unload frmMemoryStatus
    Unload frmMenuSettings
    Unload frmMonitorInfo
    Unload frmMouseInfo
    Unload frmMouseMovement
    Unload frmMouseSettings
    Unload frmMouseWarp
    Unload frmMpeg
    Unload frmNetworkInfo
    Unload frmPerfMon
    Unload frmPing
    Unload frmPowerStatus
    Unload frmProcessorInfo
    Unload frmRegistered
    Unload frmSFI_NE
    Unload frmSFI_PE
    Unload frmSharedFiles
    Unload frmStickyKeys
    Unload frmTCP_Stats
    Unload frmUDP_Stats
    Unload frmUpTime
    Unload frmWindowKiller
    Unload frmWindowInfo
    Unload frmWindows
    Unload frmWinInfo
End Sub

Private Sub mnuCmos_Click()
    frmCmos.Show
End Sub

Private Sub mnuCompUserName_Click()
    frmCompUser.Show
End Sub

Private Sub mnuCPUID_Click()
    frmCPUID.Show
End Sub

Private Sub mnuDirectories_Click()
    frmDirectories.Show
End Sub

Private Sub mnuDiskSpace_Click()
    frmDiskSpace.Show
End Sub

Private Sub mnuDisplaySettings_Click()
    frmDisplaySettings.Show
End Sub

Private Sub mnuErrMsg_Click()
    'If checked then uncheck, vice versa
    If mnuErrMsg.Checked = False Then 'Off to on
        mnuErrMsg.Checked = True
        errMsg = True
    Else 'On to off
        mnuErrMsg.Checked = False
        errMsg = False
    End If
End Sub

Private Sub mnuErrors_Click()
    frmErrors.Show
End Sub

Private Sub mnuExit_Click()
    Call App_Shutdown
End Sub

Private Sub mnuFileAttributes_Click()
    frmFileAttributes.Show
End Sub

Private Sub mnuFileTime_Click()
    frmFileTime.Show
End Sub

Private Sub mnuIconInfo_Click()
    frmIconInfo.Show
End Sub

Private Sub mnuIconSettings_Click()
    frmIconSettings.Show
End Sub

Private Sub mnuICPM_Stats_Click()
    frmICMP_Stats.Show
End Sub

Private Sub mnuIECache_Click()
    frmIECache.Show
End Sub

Private Sub mnuIEHistory_Click()
    frmIEHistory.Show
End Sub

Private Sub mnuIEOptions_Click()
    frmIEOptions.Show
End Sub

Private Sub mnuIP_Stats_Click()
    frmIP_Stats.Show
End Sub

Private Sub mnuKeyboardInfo_Click()
    frmKeyboardInfo.Show
End Sub

Private Sub mnuGetIPHost_Click()
    frmGetIPHost.Show
End Sub

Private Sub mnuKeyboardSettings_Click()
    frmKeyboardSettings.Show
End Sub

Private Sub mnuMemoryStatus_Click()
    frmMemoryStatus.Show
End Sub

Private Sub mnuMenuSettings_Click()
    frmMenuSettings.Show
End Sub

Private Sub mnuMonitorInfo_Click()
    frmMonitorInfo.Show
End Sub

Private Sub mnuMouseInfo_Click()
    frmMouseInfo.Show
End Sub

Private Sub mnuMouseMov_Click()
    frmMouseMovement.Show
End Sub

Private Sub mnuMouseMovOO_Click()
    'If checked then uncheck, vice versa
    If mnuMouseMovOO.Checked = False Then 'Off to on
        mnuMouseMovOO.Checked = True
                
        GetCursorPos POINTAPI 'Dumps info to pointapi

        'Gives point of reference for starting
        MouseMovTmpX = POINTAPI.X
        MouseMovTmpY = POINTAPI.Y
    Else 'On to off
        mnuMouseMovOO.Checked = False
    End If
End Sub

Private Sub mnuMouseSettings_Click()
    frmMouseSettings.Show
End Sub

Private Sub mnuMouseWarp_Click()
    frmMouseWarp.Show
End Sub

Private Sub mnuMouseWarpOO_Click()
    'If checked then uncheck, vice versa
    If mnuMouseWarpOO.Checked = False Then 'Off to on
        mnuMouseWarpOO.Checked = True
        
        'Converts the twips to pixels and sets the edges
        ScreenEdge.X = Screen.Width \ Screen.TwipsPerPixelX
        ScreenEdge.Y = Screen.Height \ Screen.TwipsPerPixelY
    Else 'On to off
        mnuMouseWarpOO.Checked = False
    End If
End Sub

Private Sub mnuMpeg_Click()
    frmMpeg.Show
End Sub

Private Sub mnuNetworkInfo_Click()
    frmNetworkInfo.Show
End Sub

Private Sub mnuOpenAll_Click()
    'Show all the forms - manually
    frmCachedPasswords.Show
    frmCmos.Show
    frmCompUser.Show
    frmCPUID.Show
    frmDirectories.Show
    frmDiskSpace.Show
    frmDiskVolume.Show
    frmDisplaySettings.Show
    frmErrors.Show
    frmFileAttributes.Show
    frmFileTime.Show
    frmGetIPHost.Show
    frmIconInfo.Show
    frmIconSettings.Show
    frmICMP_Stats.Show
    frmIECache.Show
    frmIEHistory.Show
    frmIEOptions.Show
    frmIP_Stats.Show
    frmKeyboardInfo.Show
    frmKeyboardSettings.Show
    frmMainAbout.Show
    frmMemoryStatus.Show
    frmMenuSettings.Show
    frmMonitorInfo.Show
    frmMouseInfo.Show
    frmMouseMovement.Show
    frmMouseSettings.Show
    frmMouseWarp.Show
    frmMpeg.Show
    frmNetworkInfo.Show
    frmPerfMon.Show
    frmPing.Show
    frmPowerStatus.Show
    frmProcessorInfo.Show
    frmRegistered.Show
    frmSFI_NE.Show
    frmSFI_PE.Show
    frmSharedFiles.Show
    frmStickyKeys.Show
    frmTCP_Stats.Show
    frmUDP_Stats.Show
    frmUpTime.Show
    frmWindowKiller.Show
    frmWindowInfo.Show
    frmWindows.Show
    frmWinInfo.Show
End Sub

Private Sub mnuPerfMon_Click()
    frmPerfMon.Show
End Sub

Private Sub mnuPing_Click()
    frmPing.Show
End Sub

Private Sub mnuPowerStatus_Click()
    frmPowerStatus.Show
End Sub

Private Sub mnuProcessorInfo_Click()
    frmProcessorInfo.Show
End Sub

Private Sub mnuRegistered_Click()
    frmRegistered.Show
End Sub

Private Sub mnuSFI_NE_Click()
    frmSFI_NE.Show
End Sub

Private Sub mnuSFI_PE_Click()
    frmSFI_PE.Show
End Sub

Private Sub mnuSharedFiles_Click()
    frmSharedFiles.Show
End Sub

Private Sub mnuStickyKeys_Click()
    frmStickyKeys.Show
End Sub

Private Sub mnuTCP_Stats_Click()
    frmTCP_Stats.Show
End Sub

Private Sub mnuUDP_Stats_Click()
    frmUDP_Stats.Show
End Sub

Private Sub mnuUpTime_Click()
    frmUpTime.Show
End Sub

Private Sub mnuVolumeInfo_Click()
    frmDiskVolume.Show
End Sub

Private Sub mnuWindowInfo_Click()
    frmWindowInfo.Show
End Sub

Private Sub mnuWindowKiller_Click()
    frmWindowKiller.Show
End Sub

Private Sub mnuWindowKillerOO_Click()
    'If checked then uncheck, vice versa
    If mnuWindowKillerOO.Checked = False Then 'Off to on
        mnuWindowKillerOO.Checked = True
        timerWindowKiller.Enabled = True
    Else 'On to off
        mnuWindowKillerOO.Checked = False
        timerWindowKiller.Enabled = False
    End If
End Sub

Private Sub mnuWindows_Click()
    frmWindows.Show
End Sub

Private Sub mnuWinInfo_Click()
    frmWinInfo.Show
End Sub

Private Sub picHM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GetCursorPos POINTAPI
    
    If frmMain.mnuMouseMovOO.Checked = True Then
        'Calculates X
        If MouseMovTmpX > POINTAPI.X Then 'If placement of cursor has changed
            MouseMovX = (MouseMovTmpX - POINTAPI.X) + MouseMovX 'Calculates difference between current pos and old pos
            MouseMovTmpX = POINTAPI.X 'Resets the point of reference
        End If
        If MouseMovTmpX < POINTAPI.X Then 'If placement of cursor has changed
            MouseMovX = (POINTAPI.X - MouseMovTmpX) + MouseMovX 'Calculates difference between current pos and old pos
            MouseMovTmpX = POINTAPI.X 'Resets the point of reference
        End If
        
        'Calculates Y
        If MouseMovTmpY > POINTAPI.Y Then 'If placement of cursor has changed
            MouseMovY = (MouseMovTmpY - POINTAPI.Y) + MouseMovY 'Calculates difference between current pos and old pos
            MouseMovTmpY = POINTAPI.Y 'Resets the point of reference
        End If
        If MouseMovTmpY < POINTAPI.Y Then 'If placement of cursor has changed
            MouseMovY = (POINTAPI.Y - MouseMovTmpY) + MouseMovY 'Calculates difference between current pos and old pos
            MouseMovTmpY = POINTAPI.Y 'Resets the point of reference
        End If
        
        frmMouseMovement.txtX.Text = MouseMovX
        frmMouseMovement.txtY.Text = MouseMovY
        frmMouseMovement.txtTotal.Text = MouseMovX + MouseMovY
    End If
    
    If frmMain.mnuMouseWarpOO.Checked = True Then
        If POINTAPI.X = ScreenEdge.X - 1 Then  'If at right edge reset to left
            SetCursorPos 1, POINTAPI.Y
            MouseWarp = MouseWarp + 1 'Increments total
        Else
            If POINTAPI.X = 0 Then 'If at left edge reset to right
                SetCursorPos ScreenEdge.X - 2, POINTAPI.Y
                MouseWarp = MouseWarp + 1 'Increments total
            End If
        End If
        
        If POINTAPI.Y = ScreenEdge.Y - 1 Then 'If at bottom edge reset to top
            SetCursorPos POINTAPI.X, 1
            MouseWarp = MouseWarp + 1 'Increments total
        Else
            If POINTAPI.Y = 0 Then 'If at top edge reset to bottom
                SetCursorPos POINTAPI.X, ScreenEdge.Y - 2
                MouseWarp = MouseWarp + 1 'Increments total
            End If
        End If
        
        frmMouseWarp.txtTotal.Text = MouseWarp
    End If
End Sub

Private Sub timerWindowKiller_Timer()
    'Errors are caused by changes in array while its going
    On Error Resume Next

    If WindowKillerNum > 0 Then 'Cant send no messages
        Dim tmpInt As Integer
        Dim tmpHandle
        
        For tmpInt = 1 To WindowKillerNum 'Cycle through array
            tmpHandle = FindWindow(vbNullString, WindowKiller(tmpInt))
        
            If tmpHandle > 0 Then
                'SendMessage tmpHandle, WM_CLOSE, 0, 0
                SendMessage tmpHandle, WM_DESTROY, 0, 0
                SendMessage tmpHandle, WM_NCDESTROY, 0, 0
                'SendMessage tmpHandle, WM_MDIDESTROY, 0, 0
            End If
        Next tmpInt
    Else
        timerWindowKiller.Enabled = False
        mnuWindowKillerOO.Checked = False
    End If
End Sub
