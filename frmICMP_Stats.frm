VERSION 5.00
Begin VB.Form frmICMP_Stats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ICMP Stats"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmICMP_Stats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerICMP_Stats 
      Enabled         =   0   'False
      Interval        =   945
      Left            =   3480
      Top             =   120
   End
   Begin VB.TextBox txtOutDestUnreachs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtInDestUnreachs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtOutTimestamps 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtOutTimestampReps 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtOutTimeExcds 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtOutSrcQuenchs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtOutRedirects 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtOutParmProbs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtOutMsgs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtOutErrors 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtOutEchos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtOutEchoReps 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtOutAddrMasks 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtOutAddrMaskReps 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtInTimestamps 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtInTimestampReps 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtInTimeExcds 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtInSrcQuenchs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtInRedirects 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtInParmProbs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtInMsgs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtInErrors 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtInEchos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtInEchoReps 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtInAddrMasks 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtInAddrMaskReps 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblOutTimestamps 
      Caption         =   "Time-Stamp Requests"
      Height          =   255
      Left            =   3960
      TabIndex        =   52
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label lblOutTimestampReps 
      Caption         =   "Time-Stamp Replies"
      Height          =   255
      Left            =   3960
      TabIndex        =   50
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblOutTimeExcds 
      Caption         =   "TTL Exceeded Messages"
      Height          =   255
      Left            =   3960
      TabIndex        =   36
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblOutSrcQuenchs 
      Caption         =   "Source Quench Messages"
      Height          =   255
      Left            =   3960
      TabIndex        =   34
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblOutRedirects 
      Caption         =   "Redirection Messages"
      Height          =   255
      Left            =   3960
      TabIndex        =   32
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblOutParmProbs 
      Caption         =   "Parameter Problem Messages"
      Height          =   255
      Left            =   3960
      TabIndex        =   30
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblOutMsgs 
      Caption         =   "Messages"
      Height          =   255
      Left            =   3960
      TabIndex        =   28
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblOutErrors 
      Caption         =   "Errors"
      Height          =   255
      Left            =   3960
      TabIndex        =   38
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblOutEchos 
      Caption         =   "Echo Requests"
      Height          =   255
      Left            =   3960
      TabIndex        =   48
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblOutEchoReps 
      Caption         =   "Echo Replies"
      Height          =   255
      Left            =   3960
      TabIndex        =   46
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblOutDestUnreachs 
      Caption         =   "Destination Unreachable messages"
      Height          =   255
      Left            =   3960
      TabIndex        =   40
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblOutAddrMasks 
      Caption         =   "Address Mask Requests"
      Height          =   255
      Left            =   3960
      TabIndex        =   44
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblOutAddrMaskReps 
      Caption         =   "Address Mask Replies"
      Height          =   255
      Left            =   3960
      TabIndex        =   42
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblInTimestamps 
      Caption         =   "Time-Stamp Requests"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label lblInTimestampReps 
      Caption         =   "Time-Stamp Replies"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblInTimeExcds 
      Caption         =   "TTL Exceeded Messages"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblInSrcQuenchs 
      Caption         =   "Source Quench Messages"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblInRedirects 
      Caption         =   "Redirection Messages"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblInParmProbs 
      Caption         =   "Parameter Problem Messages"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblInMsgs 
      Caption         =   "Messages"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblInErrors 
      Caption         =   "Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblInEchos 
      Caption         =   "Echo Requests"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblInEchoReps 
      Caption         =   "Echo Replies"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblInDestUnreachs 
      Caption         =   "Destination Unreachable messages"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblInAddrMasks 
      Caption         =   "Address Mask Requests"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblInAddrMaskReps 
      Caption         =   "Address Mask Replies"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblOut 
      Caption         =   "Out"
      Height          =   255
      Left            =   3960
      TabIndex        =   27
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblIn 
      Caption         =   "In"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmICMP_Stats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Requires 98/2k
    Dim tmpBool As Boolean
    If WinID = "WIN32_WINDOWS" Then
        If WinVer > 4010000 Then tmpBool = True 'If 98+
    Else
        If WinVer > 5000000 Then tmpBool = True 'If 2k+
    End If
    
    If tmpBool = True Then 'If go ahead
        Call timerICMP_Stats_Timer
        timerICMP_Stats.Enabled = True
    Else
        timerICMP_Stats.Enabled = False 'Disable timer
        
        'Disable everything
        lblInMsgs.Enabled = False
        lblInParmProbs.Enabled = False
        lblInRedirects.Enabled = False
        lblInSrcQuenchs.Enabled = False
        lblInTimeExcds.Enabled = False
        lblInErrors.Enabled = False
        lblInDestUnreachs.Enabled = False
        lblInAddrMaskReps.Enabled = False
        lblInAddrMasks.Enabled = False
        lblInEchoReps.Enabled = False
        lblInEchos.Enabled = False
        lblInTimestampReps.Enabled = False
        lblInTimestamps.Enabled = False
        'Second Half
        lblOutMsgs.Enabled = False
        lblOutParmProbs.Enabled = False
        lblOutRedirects.Enabled = False
        lblOutSrcQuenchs.Enabled = False
        lblOutTimeExcds.Enabled = False
        lblOutErrors.Enabled = False
        lblOutDestUnreachs.Enabled = False
        lblOutAddrMaskReps.Enabled = False
        lblOutAddrMasks.Enabled = False
        lblOutEchoReps.Enabled = False
        lblOutEchos.Enabled = False
        lblOutTimestampReps.Enabled = False
        lblOutTimestamps.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerICMP_Stats.Enabled = False
    Unload Me 'Just unloads this window not the whole program
End Sub

Private Sub timerICMP_Stats_Timer()
    If GetIcmpStatistics(MIBICMPINFO) <> 0 Then
        timerICMP_Stats.Enabled = False 'Disable timer
        Failed "GetIcmpStatistics"
        Exit Sub
    End If
    
    'In
    With MIBICMPINFO.icmpInStats
        txtInAddrMaskReps.Text = .dwAddrMaskReps
        txtInAddrMasks.Text = .dwAddrMasks
        txtInDestUnreachs.Text = .dwDestUnreachs
        txtInEchoReps.Text = .dwEchoReps
        txtInEchos.Text = .dwEchos
        txtInErrors.Text = .dwErrors
        txtInMsgs.Text = .dwMsgs
        txtInParmProbs.Text = .dwParmProbs
        txtInRedirects.Text = .dwRedirects
        txtInSrcQuenchs.Text = .dwSrcQuenchs
        txtInTimeExcds.Text = .dwTimeExcds
        txtInTimestampReps.Text = .dwTimestampReps
        txtInTimestamps.Text = .dwTimestamps
    End With
    
    'Out
    With MIBICMPINFO.icmpOutStats
        txtOutAddrMaskReps.Text = .dwAddrMaskReps
        txtOutAddrMasks.Text = .dwAddrMasks
        txtOutDestUnreachs.Text = .dwDestUnreachs
        txtOutEchoReps.Text = .dwEchoReps
        txtOutEchos.Text = .dwEchos
        txtOutErrors.Text = .dwErrors
        txtOutMsgs.Text = .dwMsgs
        txtOutParmProbs.Text = .dwParmProbs
        txtOutRedirects.Text = .dwRedirects
        txtOutSrcQuenchs.Text = .dwSrcQuenchs
        txtOutTimeExcds.Text = .dwTimeExcds
        txtOutTimestampReps.Text = .dwTimestampReps
        txtOutTimestamps.Text = .dwTimestamps
    End With
End Sub
