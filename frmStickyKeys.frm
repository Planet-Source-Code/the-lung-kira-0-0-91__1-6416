VERSION 5.00
Begin VB.Form frmStickyKeys 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sticky Keys"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmStickyKeys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRWINLOCKED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   49
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkLWINLOCKED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   41
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkRWINLATCHED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   33
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkLWINLATCHED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   25
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkRSHIFTLOCKED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   47
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkRCTLLOCKED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   45
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkRALTLOCKED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   43
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkLSHIFTLOCKED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   39
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkLCTLLOCKED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   37
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkLALTLOCKED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   35
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkRSHIFTLATCHED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   31
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkRCTLLATCHED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   29
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkRALTLATCHED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   27
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkLSHIFTLATCHED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   23
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkLCTLLATCHED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkLALTLATCHED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkTWOKEYSOFF 
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkTRISTATE 
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkSTICKYKEYSON 
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkINDICATOR 
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkHOTKEYSOUND 
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkHOTKEYACTIVE 
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkCONFIRMHOTKEY 
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkAVAILABLE 
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkAUDIBLEFEEDBACK 
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   6480
      TabIndex        =   50
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblRWINLOCKED 
      Caption         =   "Right Winkey Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   48
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblLWINLOCKED 
      Caption         =   "Left Winkey Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   40
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblRWINLATCHED 
      Caption         =   "Right WinKey Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   32
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblLWINLATCHED 
      Caption         =   "Left WinKey Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   24
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblRSHIFTLOCKED 
      Caption         =   "Right Shift Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   46
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblRCTLLOCKED 
      Caption         =   "Right Ctrl Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   44
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblRALTLOCKED 
      Caption         =   "Right Alt Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   42
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblLSHIFTLOCKED 
      Caption         =   "Left Shift Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   38
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLCTLLOCKED 
      Caption         =   "Left Ctrl Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   36
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLALTLOCKED 
      Caption         =   "Left Alt Locked"
      Height          =   255
      Left            =   5160
      TabIndex        =   34
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblRSHIFTLATCHED 
      Caption         =   "Right Shift Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   30
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblRCTLLATCHED 
      Caption         =   "Right Ctrl Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblRALTLATCHED 
      Caption         =   "Right Alt Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   26
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblLSHIFTLATCHED 
      Caption         =   "Left Shift Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLCTLLATCHED 
      Caption         =   "Left Ctrl Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLALTLATCHED 
      Caption         =   "Left Alt Latched"
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblTWOKEYSOFF 
      Caption         =   "Two Keys Off"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblTRISTATE 
      Caption         =   "Tristate"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblSTICKYKEYSON 
      Caption         =   "Sticky Keys On"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblINDICATOR 
      Caption         =   "Indicator"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblHOTKEYSOUND 
      Caption         =   "Hotkey Sound"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblHOTKEYACTIVE 
      Caption         =   "Hotkey Active"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblCONFIRMHOTKEY 
      Caption         =   "Confirm Hotkey"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblAVAILABLE 
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblAUDIBLEFEEDBACK 
      Caption         =   "Audible Feed Back"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmStickyKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Needed for flags
Dim AUDIBLEFEEDBACK As Long
Dim AVAILABLE As Long
Dim CONFIRMHOTKEY As Long
Dim HOTKEYACTIVE As Long
Dim HOTKEYSOUND As Long
Dim INDICATOR As Long
Dim STICKYKEYSON As Long
Dim TRISTATE  As Long
Dim TWOKEYSOFF As Long
Dim LWINLATCHED As Long
Dim RWINLATCHED  As Long

Private Sub cmdApply_Click()
    STICKYKEYS.dwFlags = &H0 'Clear
    
    'If chk = 1 then add flag to var else var is nothing
    If chkAUDIBLEFEEDBACK.Value = 1 Then
        AUDIBLEFEEDBACK = SKF_AUDIBLEFEEDBACK
    Else
        AUDIBLEFEEDBACK = &H0
    End If
    If chkAVAILABLE.Value = 1 Then
        AVAILABLE = SKF_AVAILABLE
    Else
        AVAILABLE = &H0
    End If
    If chkCONFIRMHOTKEY.Value = 1 Then
        CONFIRMHOTKEY = SKF_CONFIRMHOTKEY
    Else
        CONFIRMHOTKEY = &H0
    End If
    If chkHOTKEYACTIVE.Value = 1 Then
        HOTKEYACTIVE = SKF_HOTKEYACTIVE
    Else
        HOTKEYACTIVE = &H0
    End If
    If chkHOTKEYSOUND.Value = 1 Then
        HOTKEYSOUND = SKF_HOTKEYSOUND
    Else
        HOTKEYSOUND = &H0
    End If
    If chkINDICATOR.Value = 1 Then
        INDICATOR = SKF_INDICATOR
    Else
        INDICATOR = &H0
    End If
    If chkSTICKYKEYSON.Value = 1 Then
        STICKYKEYSON = SKF_STICKYKEYSON
    Else
        STICKYKEYSON = &H0
    End If
    If chkTRISTATE.Value = 1 Then
        TRISTATE = SKF_TRISTATE
    Else
        TRISTATE = &H0
    End If
    If chkTWOKEYSOFF.Value = 1 Then
        TWOKEYSOFF = SKF_TWOKEYSOFF
    Else
        TWOKEYSOFF = &H0
    End If
    
    'Set flags according to variables
    STICKYKEYS.dwFlags = AUDIBLEFEEDBACK Or AVAILABLE Or CONFIRMHOTKEY Or HOTKEYACTIVE Or HOTKEYSOUND Or INDICATOR Or STICKYKEYSON Or TRISTATE Or TWOKEYSOFF
    
    If SystemParametersInfo(SPI_SETSTICKYKEYS, 0, STICKYKEYS, 0) = 0 Then
        Failed "SystemParametersInfo"
    End If
End Sub

Private Sub Form_Load()
    STICKYKEYS.cbSize = Len(STICKYKEYS) 'Set length
    
    If SystemParametersInfo(SPI_GETSTICKYKEYS, 0, STICKYKEYS, 0) = 0 Then
        Failed "SystemParametersInfo"
        
        'Disable check boxes
        chkAUDIBLEFEEDBACK.Enabled = False
        chkAVAILABLE.Enabled = False
        chkCONFIRMHOTKEY.Enabled = False
        chkHOTKEYACTIVE.Enabled = False
        chkHOTKEYSOUND.Enabled = False
        chkINDICATOR.Enabled = False
        chkSTICKYKEYSON.Enabled = False
        chkTRISTATE.Enabled = False
        chkTWOKEYSOFF.Enabled = False
        
        cmdApply.Enabled = False
    Else 'Pull settings
        'From flags set check boxes on
        If STICKYKEYS.dwFlags And SKF_AUDIBLEFEEDBACK Then chkAUDIBLEFEEDBACK.Value = 1
        If STICKYKEYS.dwFlags And SKF_AVAILABLE Then chkAVAILABLE.Value = 1
        If STICKYKEYS.dwFlags And SKF_CONFIRMHOTKEY Then chkCONFIRMHOTKEY.Value = 1
        If STICKYKEYS.dwFlags And SKF_HOTKEYACTIVE Then chkHOTKEYACTIVE.Value = 1
        If STICKYKEYS.dwFlags And SKF_HOTKEYSOUND Then chkHOTKEYSOUND.Value = 1
        If STICKYKEYS.dwFlags And SKF_INDICATOR Then chkINDICATOR.Value = 1
        If STICKYKEYS.dwFlags And SKF_STICKYKEYSON Then chkSTICKYKEYSON.Value = 1
        If STICKYKEYS.dwFlags And SKF_TRISTATE Then chkTRISTATE.Value = 1
        If STICKYKEYS.dwFlags And SKF_TWOKEYSOFF Then chkTWOKEYSOFF.Value = 1
    End If
    
    If WinVer > 4010000 Then 'If 98/2k
        'Latched
        If STICKYKEYS.dwFlags And SKF_LALTLATCHED Then chkLALTLATCHED.Value = 1
        If STICKYKEYS.dwFlags And SKF_LCTLLATCHED Then chkLCTLLATCHED.Value = 1
        If STICKYKEYS.dwFlags And SKF_LSHIFTLATCHED Then chkLSHIFTLATCHED.Value = 1
        If STICKYKEYS.dwFlags And SKF_RALTLATCHED Then chkRALTLATCHED.Value = 1
        If STICKYKEYS.dwFlags And SKF_RCTLLATCHED Then chkRCTLLATCHED.Value = 1
        If STICKYKEYS.dwFlags And SKF_RSHIFTLATCHED Then chkRSHIFTLATCHED.Value = 1
    
        'Locked
        If STICKYKEYS.dwFlags And SKF_LALTLOCKED Then chkLALTLOCKED.Value = 1
        If STICKYKEYS.dwFlags And SKF_LCTLLOCKED Then chkLCTLLOCKED.Value = 1
        If STICKYKEYS.dwFlags And SKF_LSHIFTLOCKED Then chkLSHIFTLOCKED.Value = 1
        If STICKYKEYS.dwFlags And SKF_RALTLOCKED Then chkRALTLOCKED.Value = 1
        If STICKYKEYS.dwFlags And SKF_RCTLLOCKED Then chkRCTLLOCKED.Value = 1
        If STICKYKEYS.dwFlags And SKF_RSHIFTLOCKED Then chkRSHIFTLOCKED.Value = 1
    
        'Winkeys
        If STICKYKEYS.dwFlags And SKF_LWINLATCHED Then chkLWINLATCHED.Value = 1
        If STICKYKEYS.dwFlags And SKF_RWINLATCHED Then chkRWINLATCHED.Value = 1
        If STICKYKEYS.dwFlags And SKF_LWINLOCKED Then chkLWINLOCKED.Value = 1
        If STICKYKEYS.dwFlags And SKF_RWINLOCKED Then chkRWINLOCKED.Value = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
