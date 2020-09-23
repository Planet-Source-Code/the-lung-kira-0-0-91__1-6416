VERSION 5.00
Begin VB.Form frmMouseWarp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mouse Warp"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "frmMouseWarp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblWarped 
      Caption         =   "Number of Warps"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMouseWarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtTotal.Text = MouseWarp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
