VERSION 5.00
Begin VB.Form frmMonitorInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Monitor Info"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "frmMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVirtualScreenY 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox txtVirtualScreenX 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtScreenHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtScreenWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.CheckBox chkDisplayFormat 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtNumber 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblScreenWidth 
      Caption         =   "Screen Width"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblScreenHeight 
      Caption         =   "Screen Height"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblDisplayFormat 
      Caption         =   "Same Display Format"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblVirtualScreenY 
      Caption         =   "Virtual Screen Height"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblVirtualScreenX 
      Caption         =   "Virtual Screen Width"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblNumber 
      Caption         =   "Number of Monitors"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMonitorInfo"
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
        txtNumber.Text = GetSystemMetrics(SM_CMONITORS)
        If GetSystemMetrics(SM_CMONITORS) > 1 Then 'Need 2 monitors for this to matter
            chkDisplayFormat.Value = GetSystemMetrics(SM_SAMEDISPLAYFORMAT)
        Else
            lblDisplayFormat.Enabled = False
        End If
        txtVirtualScreenX.Text = GetSystemMetrics(SM_CXVIRTUALSCREEN)
        txtVirtualScreenY.Text = GetSystemMetrics(SM_CYVIRTUALSCREEN)
    Else
        'Disable All
        lblNumber.Enabled = False
        lblDisplayFormat.Enabled = False
        lblDisplayFormat.Enabled = False
        lblVirtualScreenX.Enabled = False
        lblVirtualScreenY.Enabled = False
    End If
    
    txtScreenWidth.Text = Screen.Width / Screen.TwipsPerPixelX
    txtScreenHeight.Text = Screen.Height / Screen.TwipsPerPixelY
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me 'Just unloads this window not the whole program
End Sub
