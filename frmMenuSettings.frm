VERSION 5.00
Begin VB.Form frmMenuSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Settings"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "frmMenuSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsShowDelay 
      Height          =   135
      LargeChange     =   5
      Left            =   120
      Max             =   999
      TabIndex        =   7
      Top             =   1275
      Width           =   1215
   End
   Begin VB.CheckBox chkMenuFade 
      Caption         =   "Check1"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton optDropAlignLeft 
      Alignment       =   1  'Right Justify
      Caption         =   "Left"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton optDropAlignRight 
      Caption         =   "  Right"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtShowDelay 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblMenuFade 
      Caption         =   "Menu Fade"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblDropAlignment 
      Caption         =   "Drop Alignment"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblShowDelay 
      Caption         =   "Show Delay"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmMenuSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim boolDropAlign As Boolean
Dim boolMenuFade As Boolean
Dim boolShowDelay As Boolean

Private Sub cmdApply_Click()
    If boolDropAlign = True Then
        Dim tmpbool As Boolean
        If optDropAlignRight.Value = True Then 'If right
            tmpbool = True
        Else 'If left
            tmpbool = False
        End If
        
        If SystemParametersInfo(SPI_GETMENUDROPALIGNMENT, tmpbool, 0, 0) = 0 Then
            Failed "SystemParametersInfo"
        End If
    End If
    If boolMenuFade = True Then
        tmpbool = chkMenuFade.Value
        
        If SystemParametersInfo(SPI_GETMENUFADE, tmpbool, 0, 0) = 0 Then
            Failed "SystemParametersInfo"
        End If
    End If
    If boolShowDelay = True Then 'If working then
        Dim tmpLong As Long
        tmpLong = hsShowDelay.Value

        If SystemParametersInfo(SPI_SETMENUSHOWDELAY, tmpLong, 0, 0) = 0 Then
            Failed "SystemParametersInfo"
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim tmpbool As Boolean
    Dim tmpLong As Long
    
    'Drop alignment
    If SystemParametersInfo(SPI_GETMENUDROPALIGNMENT, 0, tmpbool, 0) = 0 Then
        Failed "SystemParametersInfo"
        
        'Disable
        lblDropAlignment.Enabled = False
        optDropAlignLeft.Enabled = False
        optDropAlignRight.Enabled = False
    Else
        If tmpbool = True Then 'True = left
            optDropAlignLeft.Value = True
        Else 'False = right
            optDropAlignRight.Value = True
        End If
        boolDropAlign = True
    End If
    
    'Menu Fade
    If WinID = "WIN32_NT" Then 'NT
        If WinVer > 5000000 Then tmpbool = True '2k
    End If
    If tmpbool = True Then 'If go ahead
        If SystemParametersInfo(SPI_GETMENUFADE, 0, tmpbool, 0) = 0 Then
            Failed "SystemParametersInfo"
            
            'Disable all
            lblMenuFade.Enabled = False
            chkMenuFade.Enabled = False
        Else
            chkMenuFade.Enabled = tmpbool
            boolMenuFade = True
        End If
    Else
        'Disable all
        lblMenuFade.Enabled = False
        chkMenuFade.Enabled = False
    End If
    
    'Show delay
    If SystemParametersInfo(SPI_GETMENUSHOWDELAY, 0, tmpLong, 0) = 0 Then
        Failed "SystemParametersInfo"
        
        'Disables
        txtShowDelay.Enabled = False
        hsShowDelay.Enabled = False
    Else
        hsShowDelay.Value = tmpLong
        txtShowDelay.Text = tmpLong
        boolShowDelay = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub hsShowDelay_Change()
    txtShowDelay.Text = hsShowDelay.Value 'Gives value to text box once choosen
End Sub

Private Sub txtShowDelay_Change()
    On Error Resume Next
    
    If CInt(txtShowDelay.Text) < 0 Then txtShowDelay.Text = "0" 'If less than 0 resets to min , also does error trapping
    If CInt(txtShowDelay.Text) > 999 Then txtShowDelay.Text = "999" 'If greater than 999 resets to max
    
    hsShowDelay.Value = CInt(txtShowDelay.Text) 'Allows custom value to be set , by converting box to int sending it to slider
End Sub
