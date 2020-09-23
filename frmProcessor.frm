VERSION 5.00
Begin VB.Form frmProcessorInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Processor Info"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmProcessor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerCyclesElapsed 
      Interval        =   945
      Left            =   3480
      Top             =   1920
   End
   Begin VB.CheckBox chkSlow 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox txtSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtCyclesElapsed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtActiveProcessorMask 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtVendorIdentifier 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox txtIdentifier 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtProcessors 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txtProcessorName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblCyclesElapsed 
      Caption         =   "Cycles Elapsed"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblProcessors 
      Caption         =   "Number of Processors"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblSlow 
      Caption         =   "Slow Machine"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblProcessorName 
      Caption         =   "Processor Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblActiveProcessorMask 
      Caption         =   "Active Processor Mask"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblVendorIdentifier 
      Caption         =   "Vendor Identifier"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblIdentifier 
      Caption         =   "Identifier"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "frmProcessorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    GetSystemInfo SYSTEM_INFO
    
    Dim tmpbool As Boolean
    If WinID = "WIN32_NT" Then
        txtProcessorName.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Hardware\Description\System\CentralProcessor\0", "ProcessorNameString")
        txtSpeed.Text = GetSettingLong(HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "~MHz") & " MHz"
    Else
        lblProcessorName.Enabled = False 'Disable
        
        Dim tmpDbl As Double
        tmpDbl = cycles_elapsed
        Sleep 1000
        txtSpeed.Text = Round((cycles_elapsed - tmpDbl) / 1000000, 1) & " Mhz"
    End If

    txtActiveProcessorMask.Text = SYSTEM_INFO.dwActiveProcessorMask
    txtCyclesElapsed.Text = Format(cycles_elapsed, "###,###")
    txtIdentifier.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Hardware\Description\System\CentralProcessor\0", "Identifier")
    txtProcessors.Text = SYSTEM_INFO.dwNumberOrfProcessors
    chkSlow.Value = CInt(GetSystemMetrics(SM_SLOWMACHINE))
    txtVendorIdentifier.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Hardware\Description\System\CentralProcessor\0", "VendorIdentifier")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub timerCyclesElapsed_Timer()
    txtCyclesElapsed.Text = Format(cycles_elapsed, "###,###")
End Sub
