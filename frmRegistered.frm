VERSION 5.00
Begin VB.Form frmRegistered 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registration Info"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmRegistered.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOwner 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtOrg 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblOwner 
      Caption         =   "Owner"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblOrg 
      Caption         =   "Orginization"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmRegistered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    '9x use different reg settings
    If WinID = "WIN32_WINDOWS" Then '9x
        SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner", txtOwner.Text
        SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization", txtOrg.Text
    Else 'Nt/2k
        SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner", txtOwner.Text
        SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", txtOrg.Text
    End If
End Sub

Private Sub Form_Load()
    '9x 2k use different reg settings
    If WinID = "WIN32_WINDOWS" Then '9x
        txtOwner.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
        txtOrg.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "RegisteredOrganization")
    Else 'Nt/2k
        txtOwner.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
        txtOrg.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization")
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
