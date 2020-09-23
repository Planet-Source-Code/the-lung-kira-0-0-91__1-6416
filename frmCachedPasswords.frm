VERSION 5.00
Begin VB.Form frmCachedPasswords 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cached Passwords"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmCachedPasswords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCachedPasswords 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3975
   End
   Begin VB.CommandButton cmdGetData 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Get Data"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2880
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "frmCachedPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetData_Click()
    With lstCachedPasswords
        .Clear
        .AddItem "Type"
        .AddItem "Resource"
        .AddItem "Password"
        .AddItem ""
    End With
    
    If WNetEnumCachedPasswords("", 0, &HFF, AddressOf EnumCachedPasswordsProc, 0) <> 0 Then
        Failed "WNetEnumCachedPasswords"
    End If
End Sub

Private Sub Form_Load()
    '9x only
    If WinID = "WIN32_WINDOWS" Then
        If WinVer < 5000000 Then cmdGetData.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
