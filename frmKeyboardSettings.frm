VERSION 5.00
Begin VB.Form frmKeyboardSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Keyboard Settings"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmKeyboardSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkCues 
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3600
      TabIndex        =   17
      Top             =   3480
      Width           =   975
   End
   Begin VB.HScrollBar hsRepeatDelay 
      Height          =   255
      LargeChange     =   5
      Left            =   720
      Max             =   3
      TabIndex        =   11
      Top             =   2160
      Width           =   3255
   End
   Begin VB.HScrollBar hsRepeatRate 
      Height          =   255
      LargeChange     =   5
      Left            =   720
      Max             =   31
      TabIndex        =   15
      Top             =   3000
      Width           =   3255
   End
   Begin VB.CheckBox chkPref 
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   1320
      Width           =   255
   End
   Begin VB.HScrollBar hsBlinkRate 
      Height          =   255
      LargeChange     =   5
      Left            =   480
      Max             =   5000
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtBlinkRate 
      Height          =   285
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblCues 
      Caption         =   "Cues Underlined"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblPref 
      Caption         =   "Keyboard Preference"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblFast 
      Caption         =   "Fast"
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblSlow 
      Caption         =   "Slow"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblLong 
      Caption         =   "Long"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblShort 
      Caption         =   "Short"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label lblRepeatRate 
      Caption         =   "Repeat Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblRepeatDelay 
      Caption         =   "Repeat Delay"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblBlinkRate 
      Caption         =   "Caret Blink Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lbl0 
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lbl5000 
      Caption         =   "5000"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "frmKeyboardSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim boolCues As Boolean

Private Sub cmdApply_Click()
    Dim tmpLong As Long
    Dim tmpInt As Integer
    
    If boolCues = True Then
        If SystemParametersInfo(SPI_SETKEYBOARDCUES, 0, chkCues.Value, 0) = 0 Then
            Failed "SystemParametersInfo"
        End If
    End If

    If SystemParametersInfo(SPI_SETKEYBOARDPREF, 0, chkPref.Value, 0) = 0 Then
        Failed "SystemParametersInfo"
    End If
    
    If SetCaretBlinkTime(hsBlinkRate.Value) = 0 Then
        Failed "SetCaretBlinkTime"
    End If
    
    tmpInt = hsRepeatDelay.Value
    If SystemParametersInfo(SPI_SETKEYBOARDDELAY, tmpInt, 0, 0) = 0 Then
        Failed "SystemParametersInfo"
    End If

    tmpLong = hsRepeatRate.Value
    If SystemParametersInfo(SPI_SETKEYBOARDSPEED, tmpLong, 0, 0) = 0 Then
        Failed "SystemParametersInfo"
    End If
End Sub

Private Sub Form_Load()
    'Requires 98/2k
    Dim tmpBool As Boolean
    If WinID = "WIN32_WINDOWS" Then
        If WinVer > 4010000 Then tmpBool = True 'If 98+
    Else
        If WinVer > 5000000 Then tmpBool = True 'If 2k+
    End If
    
    If tmpBool = True Then 'If go ahead
        If SystemParametersInfo(SPI_GETKEYBOARDCUES, 0, tmpBool, 0) = 0 Then
            Failed "SystemParametersInfo"
        
            'Disables
            lblCues.Enabled = False
            chkCues.Enabled = False
        Else
            chkCues.Value = tmpBool
            boolCues = True
        End If
    Else 'Not avail
        'Disables
        lblCues.Enabled = False
        chkCues.Enabled = False
    End If
    
    hsBlinkRate.Value = GetCaretBlinkTime
    txtBlinkRate.Text = hsBlinkRate.Value
    
    Dim tmpInt As Integer
    If SystemParametersInfo(SPI_GETKEYBOARDDELAY, 0, tmpInt, 0) = 0 Then
        Failed "SystemParametersInfo"
    Else
        hsRepeatDelay.Value = tmpInt
    End If
    
    If WinID = "WIN32_NT" Then
        If WinVer > 5000000 Then 'If 2k
            If SystemParametersInfo(SPI_GETKEYBOARDPREF, 0, tmpBool, 0) = 0 Then
                Failed "SystemParametersInfo"
            Else
                chkPref.Value = tmpBool
            End If
        End If
    Else '9x
        If SystemParametersInfo(SPI_GETKEYBOARDPREF, 0, tmpBool, 0) = 0 Then
            Failed "SystemParametersInfo"
        Else
            chkPref.Value = tmpBool
        End If
    End If
    
    Dim tmpLong As Long
    If SystemParametersInfo(SPI_GETKEYBOARDSPEED, 0, tmpLong, 0) = 0 Then
        Failed "SystemParametersInfo"
    Else
        hsRepeatRate.Value = tmpLong
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub hsBlinkRate_Change()
    txtBlinkRate.Text = hsBlinkRate.Value
End Sub

Private Sub txtBlinkRate_Change()
    On Error Resume Next
    
    'Also does error trapping, if non number is entered then it resets to 0
    If CInt(txtBlinkRate.Text) < 0 Then txtBlinkRate.Text = "0" 'If less than 0 resets to min
    If CInt(txtBlinkRate.Text) > 5000 Then txtBlinkRate.Text = "5000" 'If greater than 5000 resets to max
    
    hsBlinkRate.Value = CInt(txtBlinkRate.Text) 'Allows custom value to be set , by converting box to int sending it to slider
End Sub
