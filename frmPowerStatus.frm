VERSION 5.00
Begin VB.Form frmPowerStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Power Status"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmPowerStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPowerStatus 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   2835
      TabIndex        =   8
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Timer timerPowerStatus 
      Enabled         =   0   'False
      Interval        =   945
      Left            =   1680
      Top             =   480
   End
   Begin VB.TextBox txtBatteryFullLifeTime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtBatteryLifeTime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtBatteryFlag 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtACLineStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblBatteryFullLifeTime 
      Caption         =   "Battery Full Life"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblBatteryLifeTime 
      Caption         =   "Battery Life"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblBatteryLifePercent 
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   1380
      Width           =   375
   End
   Begin VB.Label lblBatteryFlag 
      Caption         =   "Battery Charge Status"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblACLineStatus 
      Caption         =   "AC Power Status"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmPowerStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    picPowerStatus.ScaleWidth = 100
    
    'Requires 95/2k
    Dim tmpBool As Boolean
    If WinID = "WIN32_NT" Then
        If WinVer > 5000000 Then tmpBool = True 'If 2k+
    Else
        tmpBool = True
    End If
    
    If tmpBool = True Then
        Call timerPowerStatus_Timer
    Else
        timerPowerStatus.Enabled = False
        
        'Disable
        lblACLineStatus.Enabled = False
        lblBatteryFlag.Enabled = False
        lblBatteryLifeTime.Enabled = False
        lblBatteryFullLifeTime.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerPowerStatus.Enabled = False
    Unload Me
End Sub

Private Sub timerPowerStatus_Timer()
    If GetSystemPowerStatus(SYSTEM_POWER_STATUS) = 0 Then
        timerPowerStatus.Enabled = False
        Failed "GetSystemPowerStatus"
    Else
        timerPowerStatus.Enabled = True
    End If
    
    'AC power status
    Select Case SYSTEM_POWER_STATUS.ACLineStatus
        Case 0: txtACLineStatus.Text = "Offline"
        Case 1: txtACLineStatus.Text = "Online"
        Case 255: txtACLineStatus.Text = "Unkown status"
    End Select
    
    'Battery charge status
    Select Case SYSTEM_POWER_STATUS.BatteryFlag
        Case 1: txtBatteryFlag.Text = "High"
        Case 2: txtBatteryFlag.Text = "Low"
        Case 4: txtBatteryFlag.Text = "Critical"
        Case 8: txtBatteryFlag.Text = "Charging"
        Case 128: txtBatteryFlag.Text = "No system battery"
        Case 255: txtBatteryFlag.Text = "Unknown status"
    End Select
    
    'BatteryLifePercent - if not error then continue
    If Not SYSTEM_POWER_STATUS.BatteryLifePercent = 255 Then
        picPowerStatus.Cls
        picPowerStatus.Line (0, 0)-(SYSTEM_POWER_STATUS.BatteryLifePercent, picPowerStatus.ScaleHeight), , BF
        
        lblBatteryLifePercent.Caption = SYSTEM_POWER_STATUS.BatteryLifePercent & "%"
    End If
    
    If SYSTEM_POWER_STATUS.BatteryLifeTime > -1 Then txtBatteryLifeTime.Text = SYSTEM_POWER_STATUS.BatteryLifeTime
    If SYSTEM_POWER_STATUS.BatteryFullLifeTime > -1 Then txtBatteryFullLifeTime.Text = SYSTEM_POWER_STATUS.BatteryFullLifeTime
End Sub
