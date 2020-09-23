VERSION 5.00
Begin VB.Form frmWinInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows Info"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmWinInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtProdID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CheckBox chkDebuggerPresent 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkDBCS 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkDebug 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtBoot 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtPlatformID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblProdID 
      Caption         =   "Product ID"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblDebuggerPresent 
      Caption         =   "Debugger Present"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblDBCS 
      Caption         =   "DBCS Version"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblDebug 
      Caption         =   "Debug Version"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblBoot 
      Caption         =   "Boot Method"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblPlatformID 
      Caption         =   "Platform ID"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "frmWinInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Select Case GetSystemMetrics(SM_CLEANBOOT) 'Boot method
        Case 0: txtBoot.Text = "Normal Boot"
        Case 1: txtBoot.Text = "Fail-safe boot"
        Case 2: txtBoot.Text = "Fail-safe with network boot"
    End Select
    
    chkDBCS.Value = GetSystemMetrics(SM_DBCSENABLED)
    chkDebug.Value = GetSystemMetrics(SM_DEBUG)
    chkDebuggerPresent.Value = CInt(IsDebuggerPresent)
    
    '9x 2k use different reg settings
    If WinID = "WIN32_WINDOWS" Then '9x
        txtName.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "Version")
    Else 'Nt/2k
        txtName.Text = GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
    End If

    Dim tmpString As String
    tmpString = CStr(WinVer)
    txtVersion.Text = Left$(tmpString, 1) & "." & Mid$(tmpString, 3, 2) & "." & Right$(tmpString, 4)
    
    txtPlatformID.Text = WinID
    
    If WinID = "WIN32_WINDOWS" Then '9x
        txtProdID.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductId")
    Else 'Nt/2k
        txtProdID.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "ProductId")
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
