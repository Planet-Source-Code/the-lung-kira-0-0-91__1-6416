VERSION 5.00
Begin VB.Form frmCompUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Computer / User  Name"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmCompUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtComputer 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblUser 
      Caption         =   "User Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblComputer 
      Caption         =   "Computer Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmCompUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    Set_ComputerName txtComputer.Text
End Sub

Private Sub Form_Load()
    txtUser.Text = Get_UserName
    txtComputer.Text = Get_ComputerName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub txtComputer_Change()
    txtComputer.Text = Rem_NonStd_Chr(txtComputer.Text)
End Sub
