VERSION 5.00
Begin VB.Form frmNetworkInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Network Info"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmNetworkInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHostName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtDomainName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtLocalHostName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtLocalIP 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtWinsockSystemStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtWinsockDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CheckBox chkInetIsOffline 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkNetworkPresent 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkEnableDns 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtScopeId 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CheckBox chkNodeType 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chkEnableRouting 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox chkEnableProxy 
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblWinsockSystemStatus 
      Caption         =   "Winsock System Status"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblWinsockDescription 
      Caption         =   "Winsock Description"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblLocalIP 
      Caption         =   "Local IP"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblLocalHostName 
      Caption         =   "Local Host Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblInetIsOffline 
      Caption         =   "Inet Is Offline"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblNetworkPresent 
      Caption         =   "Network Present"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblEnableDns 
      Caption         =   "DNS Enabled"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblScopeId 
      Caption         =   "DHCP Scope Name"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblNodeType 
      Caption         =   "Use DHCP"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label lblEnableRouting 
      Caption         =   "Routing Enabled"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label lblEnableProxy 
      Caption         =   "ARP Proxy"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblHostName 
      Caption         =   "Host Name"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblDomainName 
      Caption         =   "Domain Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmNetworkInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Requires 98/2k
    Dim tmpbool As Boolean
    If WinID = "WIN32_WINDOWS" Then
        If WinVer > 4010000 Then tmpbool = True 'If 98+
    Else
        If WinVer > 5000000 Then tmpbool = True 'If 2k+
    End If
    
    If tmpbool = True Then 'If go ahead
        'Used 624 other wise it buffer overflowed me
        'Should be Len(FIXED_INFO), function returned me to use 624
        Dim FI As FIXED_INFO 'Cant use FIXED_INFO
        apiError = GetNetworkParams(FI, 624)
        If apiError <> 0 Then
            Errors.Errors apiError, "GetNetworkParams"
        End If
        
        'Send info to text boxes
        With FI
            txtDomainName.Text = .DomainName
            If .EnableDns > 0 Then chkEnableDns.Value = 1
            If .EnableProxy > 0 Then chkEnableProxy.Value = 1
            If .EnableRouting > 0 Then chkEnableRouting.Value = 1
            txtHostName.Text = .HostName
            If .NodeType > 0 Then chkNodeType.Value = 1
            txtScopeId.Text = .ScopeId
        End With
    Else
        'Disable All
        lblDomainName.Enabled = False
        lblEnableDns.Enabled = False
        lblEnableProxy.Enabled = False
        lblEnableRouting.Enabled = False
        lblHostName.Enabled = False
        lblNodeType.Enabled = False
        lblScopeId.Enabled = False
    End If
    
    chkInetIsOffline.Value = CInt(InetIsOffline(0))
    txtLocalHostName.Text = GetHostByIP(GetIPByHost(ComputerName))
    txtLocalIP.Text = GetIPByHost(ComputerName)
    chkNetworkPresent.Value = CInt(Right$(Asc2Bin(GetSystemMetrics(SM_NETWORK)), 1))
    txtWinsockDescription.Text = WinsockData.Description
    txtWinsockSystemStatus.Text = WinsockData.SystemStatus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
