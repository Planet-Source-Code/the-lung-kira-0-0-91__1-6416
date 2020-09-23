VERSION 5.00
Begin VB.Form frmMouseMovement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mouse Movement"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2670
   Icon            =   "frmMouseMovement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblX 
      Caption         =   "Total X"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblY 
      Caption         =   "Total Y"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmMouseMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtTotal.Text = MouseMovX + MouseMovY
    txtX.Text = MouseMovX
    txtY.Text = MouseMovY

    'InchesX = (MouseMovX * (Screen.TwipsPerPixelX / 20)) / 72
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
