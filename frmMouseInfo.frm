VERSION 5.00
Begin VB.Form frmMouseInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mouse Info"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   Icon            =   "frmMouseInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMouseWheel 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   2640
      Width           =   255
   End
   Begin VB.TextBox txtDragDropWidth 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtDragDropHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtButtons 
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
   Begin VB.TextBox txtSwapButton 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtDblClickHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtDblClickWidth 
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
   Begin VB.TextBox txtCursorWidth 
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
   Begin VB.TextBox txtCursorHeight 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblDragDropWidth 
      Caption         =   "Drag Drop Width"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblDragDropHeight 
      Caption         =   "Drag Drop Height"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblButtons 
      Caption         =   "Buttons"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblMouseWheel 
      Caption         =   "Mouse Wheel"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblSwapButton 
      Caption         =   "Main Button"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblDblClickHeight 
      Caption         =   "Double Click Height"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblDblClickWidth 
      Caption         =   "Double Click Width"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblCursorWidth 
      Caption         =   "Cursor Width"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblCursorHeight 
      Caption         =   "Cursor Height"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmMouseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Pull info from system metrics
    txtButtons.Text = GetSystemMetrics(SM_CMOUSEBUTTONS)
    txtCursorHeight.Text = GetSystemMetrics(SM_CYCURSOR)
    txtCursorWidth.Text = GetSystemMetrics(SM_CXCURSOR)
    txtDblClickHeight.Text = GetSystemMetrics(SM_CYDOUBLECLK)
    txtDblClickWidth.Text = GetSystemMetrics(SM_CXDOUBLECLK)
    txtDragDropHeight.Text = GetSystemMetrics(SM_CXDRAG)
    txtDragDropWidth.Text = GetSystemMetrics(SM_CYDRAG)
    
    'Requires 98/NT4
    Dim tmpBool As Boolean
    If WinID = "WIN32_WINDOWS" Then
        If WinVer > 4010000 Then tmpBool = True 'If 98+
    Else
        If WinVer > 4000000 Then tmpBool = True 'If NT4+
    End If
    If tmpBool = True Then 'If go ahead
        chkMouseWheel.Value = GetSystemMetrics(SM_MOUSEWHEELPRESENT)
    Else
        lblMouseWheel.Enabled = False
    End If
    
    If GetSystemMetrics(SM_SWAPBUTTON) = True Then 'Right
        txtSwapButton.Text = "Right"
    Else 'Left
        txtSwapButton.Text = "Left"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me 'Just unloads this window not the whole program
End Sub
