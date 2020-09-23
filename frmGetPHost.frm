VERSION 5.00
Begin VB.Form frmGetIPHost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get IP / Host"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmGetPHost.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetHost 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Get Host"
      Height          =   350
      Left            =   2880
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdGetIP 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Get IP"
      Height          =   350
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblIP 
      Caption         =   "IP"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblHost 
      Caption         =   "Host Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmGetIPHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetHost_Click()
    txtHost.Text = GetHostByIP(txtIP.Text)
End Sub

Private Sub cmdGetIP_Click()
    txtIP.Text = GetIPByHost(txtHost.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

