VERSION 5.00
Begin VB.Form frmTCP_Stats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TCP Stats"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmTCP_Stats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAttemptFails 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtRtoAlgorithm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtCurrEstab 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Timer timerTCP_Stats 
      Enabled         =   0   'False
      Interval        =   945
      Left            =   1440
      Top             =   960
   End
   Begin VB.TextBox txtRtoMin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtRtoMax 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtRetransSegs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtPassiveOpens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtOutSegs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtOutRsts 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtNumConns 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtMaxConn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtInSegs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtInErrs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtEstabResets 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtActiveOpens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblRtoMin 
      Caption         =   "Minimum Time-Out"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label lblRtoMax 
      Caption         =   "Maximum Time-Out"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblRtoAlgorithm 
      Caption         =   "Time-Out Algorithm"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblRetransSegs 
      Caption         =   "Segments Retransmitted"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblPassiveOpens 
      Caption         =   "Passive Opens"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblOutSegs 
      Caption         =   "Segments Sent"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblOutRsts 
      Caption         =   "Outgoing Resets"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblNumConns 
      Caption         =   "Cumulative Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblMaxConn 
      Caption         =   "Maximum Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblInSegs 
      Caption         =   "Segments Received"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblInErrs 
      Caption         =   "Incoming Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label lblEstabResets 
      Caption         =   "Established Connections Reset"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblCurrEstab 
      Caption         =   "Established Connections"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblAttemptFails 
      Caption         =   "Failed Attempts"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label lblActiveOpens 
      Caption         =   "Active Opens"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2415
   End
End
Attribute VB_Name = "frmTCP_Stats"
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
        Call timerTCP_Stats_Timer
        timerTCP_Stats.Enabled = True
    Else
        timerTCP_Stats.Enabled = False 'Disable timer

        'Disable All
        lblActiveOpens.Enabled = False
        lblAttemptFails.Enabled = False
        lblCurrEstab.Enabled = False
        lblEstabResets.Enabled = False
        lblInErrs.Enabled = False
        lblInSegs.Enabled = False
        lblMaxConn.Enabled = False
        lblNumConns.Enabled = False
        lblOutRsts.Enabled = False
        lblOutSegs.Enabled = False
        lblPassiveOpens.Enabled = False
        lblRetransSegs.Enabled = False
        lblRtoAlgorithm.Enabled = False
        lblRtoMax.Enabled = False
        lblRtoMin.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerTCP_Stats.Enabled = False
    Unload Me 'Just unloads this window not the whole program
End Sub

Private Sub timerTCP_Stats_Timer()
    If GetTcpStatistics(MIB_TCPSTATS) <> 0 Then
        timerTCP_Stats.Enabled = False 'Disable timer
        Failed "GetTcpStatistics"
        Exit Sub
    End If
    
    With MIB_TCPSTATS
        txtActiveOpens.Text = .dwActiveOpens
        txtAttemptFails.Text = .dwAttemptFails
        txtCurrEstab.Text = .dwCurrEstab
        txtEstabResets.Text = .dwEstabResets
        txtInErrs.Text = .dwInErrs
        txtInSegs.Text = .dwInSegs
        txtMaxConn.Text = .dwMaxConn
        txtNumConns.Text = .dwNumConns
        txtOutRsts.Text = .dwOutRsts
        txtOutSegs.Text = .dwOutSegs
        txtPassiveOpens.Text = .dwPassiveOpens
        txtRetransSegs.Text = .dwRetransSegs
        
        'Change into something usefull
        Select Case .dwRtoAlgorithm 'Time out algorithm
            Case 1: txtRtoAlgorithm.Text = "Other"
            Case 2: txtRtoAlgorithm.Text = "Constant Time-out"
            Case 3: txtRtoAlgorithm.Text = "MIL-STD-1778 Appendix B"
            Case 4: txtRtoAlgorithm.Text = "Van Jacobson's Algorithm"
        End Select
        
        txtRtoMax.Text = .dwRtoMax
        txtRtoMin.Text = .dwRtoMin
    End With
End Sub
