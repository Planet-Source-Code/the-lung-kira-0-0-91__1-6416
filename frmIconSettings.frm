VERSION 5.00
Begin VB.Form frmIconSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Icon Settings"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "frmIconSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHorzSpc 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtVertSpc 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox chkTitleWrap 
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblVertSpc 
      Caption         =   "Vertical Spacing"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblHorzSpc 
      Caption         =   "Horizontal Spacing"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblTitleWrap 
      Caption         =   "Title Wrap"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmIconSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    'Set values
    With ICONMETRICS
        .iHorzSpacing = CInt(txtHorzSpc.Text)
        .iVertSpacing = CInt(txtVertSpc.Text)
        .iTitleWrap = chkTitleWrap.Value
    End With
    
    If SystemParametersInfo(SPI_SETICONMETRICS, Len(ICONMETRICS), ICONMETRICS, 0) = 0 Then
        Failed "SystemParametersInfo"
    End If
End Sub

Private Sub Form_Load()
    ICONMETRICS.cbSize = Len(ICONMETRICS)
    
    If SystemParametersInfo(SPI_GETICONMETRICS, Len(ICONMETRICS), ICONMETRICS, 0) = 0 Then
        Failed "SystemParametersInfo"
    Else
        txtHorzSpc.Text = ICONMETRICS.iHorzSpacing
        txtVertSpc.Text = ICONMETRICS.iVertSpacing
        chkTitleWrap.Value = ICONMETRICS.iTitleWrap
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub txtHorzSpc_Change()
    On Error Resume Next
    If CInt(txtHorzSpc.Text) <= 0 Then txtHorzSpc.Text = "1" 'If less than 0 resets to min , also does error trapping
End Sub

Private Sub txtVertSpc_Change()
    On Error Resume Next
    If CInt(txtVertSpc.Text) <= 0 Then txtVertSpc.Text = "1" 'If less than 0 resets to min , also does error trapping
End Sub
