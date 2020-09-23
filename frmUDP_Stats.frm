VERSION 5.00
Begin VB.Form frmUDP_Stats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UDP Stats"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmUDP_Stats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNoPorts 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Timer timerUDP_Stats 
      Enabled         =   0   'False
      Interval        =   945
      Left            =   2160
      Top             =   240
   End
   Begin VB.TextBox txtOutDatagrams 
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
   Begin VB.TextBox txtNumAddrs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtInErrors 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtInDatagrams 
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
   Begin VB.Label lblOutDatagrams 
      Caption         =   "Sent Datagrams"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblNumAddrs 
      Caption         =   "Entries In UDP Listener Table"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblInErrors 
      Caption         =   "Errors On Received Datagrams"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblNoPorts 
      Caption         =   "Datagrams w/ No Port "
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblInDatagrams 
      Caption         =   "Received Datagrams"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmUDP_Stats"
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
        Call timerUDP_Stats_Timer
        timerUDP_Stats.Enabled = True
    Else
        timerUDP_Stats.Enabled = False 'Disable timer
        
        'Disable All
        lblInDatagrams.Enabled = False
        lblInErrors.Enabled = False
        lblNoPorts.Enabled = False
        lblNumAddrs.Enabled = False
        lblOutDatagrams.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerUDP_Stats.Enabled = False
    Unload Me 'Just unloads this window not the whole program
End Sub

Private Sub timerUDP_Stats_Timer()
    If GetUdpStatistics(MIB_UDPSTATS) <> 0 Then
        timerUDP_Stats.Enabled = False 'Disable timer
        Failed "GetUdpStatistics"
        Exit Sub
    End If
    
    With MIB_UDPSTATS
        txtInDatagrams.Text = .dwInDatagrams
        txtInErrors.Text = .dwInErrors
        txtNoPorts.Text = .dwNoPorts
        txtNumAddrs.Text = .dwNumAddrs
        txtOutDatagrams.Text = .dwOutDatagrams
    End With
End Sub
