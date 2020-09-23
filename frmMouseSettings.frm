VERSION 5.00
Begin VB.Form frmMouseSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mouse Settings"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmMouseSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDblClickTimeApply 
      Caption         =   "Apply"
      Height          =   285
      Left            =   4800
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtDblClickTime 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.HScrollBar hsDblClickTime 
      Height          =   255
      LargeChange     =   5
      Left            =   480
      Max             =   5000
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lbl0 
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lbl5000 
      Caption         =   "5000"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblDblClickTime 
      Caption         =   "Double Click Time"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMouseSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDblClickTimeApply_Click()
    If SetDoubleClickTime(hsDblClickTime.Value) = 0 Then
        Failed "SetDoubleClickTime"
    End If
End Sub

Private Sub Form_Load()
    'Gets double click time
    hsDblClickTime.Value = GetDoubleClickTime
    txtDblClickTime.Text = hsDblClickTime.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub hsDblClickTime_Change()
    txtDblClickTime.Text = hsDblClickTime.Value 'Gives value to text box once choosen
End Sub

Private Sub txtDblClickTime_Change()
    On Error Resume Next
    
    If CInt(txtDblClickTime.Text) < 0 Then txtDblClickTime.Text = "0" 'If less than 0 resets to min , also does error trapping
    If CInt(txtDblClickTime.Text) > 5000 Then txtDblClickTime.Text = "5000" 'If greater than 5000 resets to max
    
    hsDblClickTime.Value = CInt(txtDblClickTime.Text) 'Allows custom value to be set , by converting box to int sending it to slider
End Sub
