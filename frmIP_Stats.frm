VERSION 5.00
Begin VB.Form frmIP_Stats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IP Stats"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmIP_Stats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDefaultTTL 
      Height          =   285
      Left            =   2640
      TabIndex        =   37
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   285
      Left            =   3480
      TabIndex        =   38
      Top             =   5160
      Width           =   975
   End
   Begin VB.Timer timerIP_Stats 
      Enabled         =   0   'False
      Interval        =   945
      Left            =   2160
      Top             =   360
   End
   Begin VB.TextBox txtInUnknownProtos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txtInHdrErrors 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txtInAddrErrors 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtNumIf 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtNumAddr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtRoutingDiscards 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtOutNoRoutes 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtNumRoutes 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtReasmReqds 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtReasmFails 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtReasmOks 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtFragFails 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtReasmTimeout 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtFragOks 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtFragCreates 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtOutDiscards 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtInDiscards 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox txtInDelivers 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtForwDatagrams 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtOutRequests 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtInReceives 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblReasmTimeout 
      Caption         =   "Datagrams Missing Frags"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label lblInHdrErrors 
      Caption         =   "Received Header Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label lblInAddrErrors 
      Caption         =   "Received Address Errors"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label lblInUnknownProtos 
      Caption         =   "Datagrams w/ Unknown Protocol"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label lblNumIf 
      Caption         =   "Interfaces On Computer"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label lblNumAddr 
      Caption         =   "IP Addresses On Computer"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lblDefaultTTL 
      Caption         =   "Default TTL"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label lblOutNoRoutes 
      Caption         =   "Datagrams w/ No Route"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label lblNumRoutes 
      Caption         =   "Routes In Routing Table"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblRoutingDiscards 
      Caption         =   "Datagram Routing Discards"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblReasmOks 
      Caption         =   "Successful Reassemblies"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblReasmFails 
      Caption         =   "Failed Reassemblies"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblReasmReqds 
      Caption         =   "Datagrams Requiring Reassembly"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label lblFragOks 
      Caption         =   "Successful Fragmentations"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblFragFails 
      Caption         =   "Failed Fragmentations"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label lblFragCreates 
      Caption         =   "Datagrams Fragmented"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblInDiscards 
      Caption         =   "Received Datagrams Discarded"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblInDelivers 
      Caption         =   "Received Datagrams Delivered"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblForwDatagrams 
      Caption         =   "Datagrams Forwarded"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblOutDiscards 
      Caption         =   "Sent Datagrams Discarded"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblOutRequests 
      Caption         =   "Datagram Sent"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblInReceives 
      Caption         =   "Datagrams Received"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmIP_Stats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    apiError = SetIpTTL(CInt(txtDefaultTTL.Text))
    If apiError <> 0 Then
        Errors.Errors apiError, "SetIpTTL"
    End If
End Sub

Private Sub Form_Load()
    'Requires 98/2k
    Dim tmpBool As Boolean
    If WinID = "WIN32_WINDOWS" Then
        If WinVer > 4010000 Then tmpBool = True 'If 98+
    Else
        If WinVer > 5000000 Then tmpBool = True 'If 2k+
    End If
    
    If tmpBool = True Then 'If go ahead
        If GetIpStatistics(MIB_IPSTATS) <> 0 Then
            Failed "GetIpStatistics"
        End If
        
        txtDefaultTTL.Text = MIB_IPSTATS.dwDefaultTTL 'Set it once
        Call timerIP_Stats_Timer
        
        timerIP_Stats.Enabled = True
        cmdApply.Enabled = True
    Else
        timerIP_Stats.Enabled = False 'Disable timer
    
        'Disable All
        lblDefaultTTL.Enabled = False
        txtDefaultTTL.Enabled = False
        lblForwDatagrams.Enabled = False
        lblFragCreates.Enabled = False
        lblFragFails.Enabled = False
        lblFragOks.Enabled = False
        lblInAddrErrors.Enabled = False
        lblInDelivers.Enabled = False
        lblInDiscards.Enabled = False
        lblInHdrErrors.Enabled = False
        lblInReceives.Enabled = False
        lblInUnknownProtos.Enabled = False
        lblNumAddr.Enabled = False
        lblNumIf.Enabled = False
        lblNumRoutes.Enabled = False
        lblOutDiscards.Enabled = False
        lblOutNoRoutes.Enabled = False
        lblOutRequests.Enabled = False
        lblReasmFails.Enabled = False
        lblReasmOks.Enabled = False
        lblReasmReqds.Enabled = False
        lblReasmTimeout.Enabled = False
        lblRoutingDiscards.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerIP_Stats.Enabled = False
    Unload Me
End Sub

Private Sub timerIP_Stats_Timer()
    If GetIpStatistics(MIB_IPSTATS) <> 0 Then
        timerIP_Stats.Enabled = False 'Disable timer
        Failed "GetIpStatistics"
        Exit Sub
    End If
    
    'Set textboxes with structure data
    'txtDefaultTTL.Text = MIB_IPSTATS.dwDefaultTTL
    With MIB_IPSTATS
        txtForwDatagrams.Text = .dwForwDatagrams
        txtFragCreates.Text = .dwFragCreates
        txtFragFails.Text = .dwFragFails
        txtFragOks.Text = .dwFragOks
        txtInAddrErrors.Text = .dwInAddrErrors
        txtInDelivers.Text = .dwInDelivers
        txtInDiscards.Text = .dwInDiscards
        txtInHdrErrors.Text = .dwInHdrErrors
        txtInReceives.Text = .dwInReceives
        txtInUnknownProtos.Text = .dwInUnknownProtos
        txtNumAddr.Text = .dwNumAddr
        txtNumIf.Text = .dwNumIf
        txtNumRoutes.Text = .dwNumRoutes
        txtOutDiscards.Text = .dwOutDiscards
        txtOutNoRoutes.Text = .dwOutNoRoutes
        txtOutRequests.Text = .dwOutRequests
        txtReasmFails.Text = .dwReasmFails
        txtReasmOks.Text = .dwReasmOks
        txtReasmReqds.Text = .dwReasmReqds
        txtReasmTimeout.Text = .dwReasmTimeout
        txtRoutingDiscards.Text = .dwRoutingDiscards
    End With
End Sub

