VERSION 5.00
Begin VB.Form frmPerfMon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Performance Monitor"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmPerfMon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDiff 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.HScrollBar hsInterval 
      Height          =   135
      LargeChange     =   5
      Left            =   2640
      TabIndex        =   4
      Top             =   675
      Value           =   1000
      Width           =   1335
   End
   Begin VB.Timer timerPerfMon 
      Enabled         =   0   'False
      Left            =   2040
      Top             =   0
   End
   Begin VB.ComboBox cboObject 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblDiff 
      Caption         =   "Difference"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblData 
      Caption         =   "Data"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblInterval 
      Caption         =   "Update Interval"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblObject 
      Caption         =   "Object"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmPerfMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oldData As Long

Private Sub cboObject_Click()
    'Reset
    oldData = 0
    txtData.Text = ""
    txtDiff.Text = ""
    
    timerPerfMon.Interval = hsInterval.Value
    timerPerfMon.Enabled = True
    timerPerfMon_Timer
End Sub

Private Sub Form_Load()
    With cboObject
        .Clear
        
        Dim lngCount As Long
        Dim strValueName() As String
        Dim lngValueType As Long
        Dim tmpLong As Long
        
        If WinID = "WIN32_WINDOWS" Then 'Registry
            'Enumerate all objects
            EnumValue HKEY_DYN_DATA, "PerfStats\StatData", strValueName(), lngCount, lngValueType
            
            For tmpLong = 0 To lngCount - 2 'Cycle through
                .AddItem strValueName(tmpLong) 'Dump
            Next tmpLong
        Else 'PDH
            
        End If
    End With
    
    Dim srvValueName() As String
    Dim srvCount As Long
    EnumValue HKEY_DYN_DATA, "PerfStats\StartSrv", srvValueName(), srvCount, lngValueType
    
    For tmpLong = 0 To srvCount - 2
        'Start up monitoring services
        GetDataPerfMon HKEY_DYN_DATA, "PerfStats\StartSrv", srvValueName(tmpLong)
    Next tmpLong
    
    hsInterval.Value = PerfMonInterval
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerPerfMon.Enabled = True 'Disable timer
    
    Dim srvValueName() As String
    Dim srvCount As Long
    Dim lngValueType As Long
    EnumValue HKEY_DYN_DATA, "PerfStats\StopSrv", srvValueName(), srvCount, lngValueType
    
    Dim tmpLong As Long
    For tmpLong = 0 To srvCount - 2
        'Stop up monitoring services
        GetDataPerfMon HKEY_DYN_DATA, "PerfStats\StopSrv", srvValueName(tmpLong)
    Next tmpLong
    
    Unload Me
End Sub

Private Sub hsInterval_Change()
    txtInterval.Text = hsInterval.Value
    timerPerfMon.Interval = hsInterval.Value
    PerfMonInterval = hsInterval.Value
End Sub

Private Sub timerPerfMon_Timer()
    If txtData.Text <> "" Then oldData = CLng(txtData.Text)

    txtData.Text = GetDataPerfMon(HKEY_DYN_DATA, "PerfStats\StatData", cboObject.List(cboObject.ListIndex))
    txtDiff.Text = CLng(txtData.Text) - oldData
End Sub

Private Sub txtInterval_Change()
    On Error Resume Next
    
    If CInt(txtInterval.Text) < 0 Then txtInterval.Text = "0"
    hsInterval.Value = CInt(txtInterval.Text)
End Sub
