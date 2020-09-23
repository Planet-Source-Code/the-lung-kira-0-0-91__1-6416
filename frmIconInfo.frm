VERSION 5.00
Begin VB.Form frmIconInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Icon Info"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "frmIconInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSpacingHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtSmallHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtSmallWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtSpacingWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtDHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtDWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblSmallHeight 
      Caption         =   "Small Icon Height"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblSmallWidth 
      Caption         =   "Small Icon Width"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblSpacingWidth 
      Caption         =   "Spacing Width"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblSpacingHeight 
      Caption         =   "Spacing Height"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblDHeight 
      Caption         =   "Default Height"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblDWidth 
      Caption         =   "Default Width"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmIconInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Pull info
    txtDWidth.Text = GetSystemMetrics(SM_CXICON)
    txtDHeight.Text = GetSystemMetrics(SM_CYICON)
    txtSmallWidth.Text = GetSystemMetrics(SM_CXSMICON)
    txtSmallHeight.Text = GetSystemMetrics(SM_CYSMICON)
    txtSpacingWidth.Text = GetSystemMetrics(SM_CXICONSPACING)
    txtSpacingHeight.Text = GetSystemMetrics(SM_CYICONSPACING)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me 'Just unloads this window not the whole program
End Sub
