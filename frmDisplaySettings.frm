VERSION 5.00
Begin VB.Form frmDisplaySettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Display Settings"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "frmDisplaySettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkGlobal 
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   2160
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.ListBox lstModes 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.ComboBox cboRate 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Text            =   "60"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   1920
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblGlobal 
      Caption         =   "Global Change"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblAvailable 
      Caption         =   "Modes Available"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblRate 
      Caption         =   "Refresh Rate"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frmDisplaySettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    With DEVMODE
        .dmSize = Len(DEVMODE)
        .dmBitsPerPel = CLng(Right(lstModes.List(lstModes.ListIndex), 2))
        .dmPelsWidth = CLng(Trim(Left(lstModes.List(lstModes.ListIndex), 8)))
        .dmPelsHeight = CLng(Trim(Mid(lstModes.List(lstModes.ListIndex), 8, 8)))
        .dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_DISPLAYFREQUENCY
        .dmDisplayFrequency = CLng(cboRate.Text)
        '.dmPosition 'Multimonitor
    End With
    
    'Test
    If ChangeDisplaySettings(DEVMODE, CDS_TEST) <> 0 Then
        MsgBox "Test failed.", vbExclamation, "Error"
        Exit Sub 'If error exit here
    End If
    
    apiError = chkGlobal.Value
    If apiError = 1 Then
        apiError = ChangeDisplaySettings(DEVMODE, CDS_UPDATEREGISTRY Or CDS_GLOBAL)
    Else
        apiError = ChangeDisplaySettings(DEVMODE, CDS_UPDATEREGISTRY)
    End If
    Select Case apiError 'What to do
        Case DISP_CHANGE_RESTART
            MsgBox "Must restart computer for changes to be implemented.", vbInformation, "Restart"
        Case DISP_CHANGE_BADFLAGS
            Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_BADPARAM
            Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_FAILED
            Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_BADMODE
            Failed "ChangeDisplaySettings"
        Case DISP_CHANGE_NOTUPDATED
            Failed "ChangeDisplaySettings"
    End Select
    
    'Change screenedges accordingly
    ScreenEdge.X = Screen.Width \ Screen.TwipsPerPixelX
    ScreenEdge.Y = Screen.Height \ Screen.TwipsPerPixelY
End Sub

Private Sub Form_Load()
    Dim tmpLong As Long
    DEVMODE.dmSize = Len(DEVMODE)
    
    Do 'Get all possible display modes
        apiError = EnumDisplaySettings(ByVal 0, tmpLong, DEVMODE)
        tmpLong = tmpLong + 1 'Increment
        
        lstModes.AddItem Left(DEVMODE.dmPelsWidth & Space(8), 8) & Left(DEVMODE.dmPelsHeight & Space(8), 8) & DEVMODE.dmBitsPerPel
    Loop While apiError > 0
    
    For tmpLong = 1 To 300
        cboRate.AddItem tmpLong
    Next tmpLong
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
