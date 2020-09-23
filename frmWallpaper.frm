VERSION 5.00
Begin VB.Form frmWallpaper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wallpaper"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmWallpaper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstNewWallpaper 
      Height          =   840
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5175
   End
   Begin VB.TextBox txtCurrent 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   350
      Left            =   4320
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3360
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblNewWallpaper 
      Caption         =   "New Wallpaper"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblCurrent 
      Caption         =   "Current"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmWallpaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    cmdApply.Enabled = False 'Reset
    Screen.MousePointer = vbArrowHourglass 'Sets mouse as hourglass
    
    Select Case lstNewWallpaper.List(lstNewWallpaper.ListIndex)
        Case "Default" 'Default = null
            apiError = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, 0&, 0)
            
            If apiError = 0 Then
                Failed "SystemParametersInfo"
            Else
                txtCurrent.Text = ""
            End If
        Case "None" 'None = ""
            apiError = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, "", 0)
            
            If apiError = 0 Then
                Failed "SystemParametersInfo"
            Else
                txtCurrent.Text = ""
            End If
        Case Else 'Selected wallpaper
            Dim tmpString As String
            tmpString = lstNewWallpaper.List(lstNewWallpaper.ListIndex)

            apiError = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, tmpString & Chr(0), SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

            If apiError = 0 Then
                Failed "SystemParametersInfo"
            Else
                txtCurrent.Text = tmpString
            End If
    End Select

    Screen.MousePointer = vbNormal 'Resets cursor so they can continue
    cmdApply.Enabled = True 'Reenable
End Sub

Private Sub cmdChoose_Click()
    Dim strFileName As String
    GetOpenName hwnd, "Open", strFileName
    
    cmdChoose.Enabled = False
    Screen.MousePointer = vbArrowHourglass 'Sets mouse as hourglass
    
    'Error checking
    If Not strFileName <> "" Then
        Exit Sub 'Dont worry just exit
    End If
    If Not FileLen(strFileName) > 0 Then 'If file len not greater than 0
        MsgBox "File size is 0.", vbExclamation, "Error"
        Exit Sub
    End If
    
    lstNewWallpaper.AddItem strFileName 'Add it to the list
    
    Screen.MousePointer = vbNormal 'Resets cursor so they can continue
    cmdChoose.Enabled = True
End Sub

Private Sub Form_Load()
    'Requires 2k
    If WinInfo.ID = "WIN32_NT" Then
        If WinInfo.Version > 5000000 Then
            Dim tmpString As String * MAX_PATH
        
            apiError = SystemParametersInfo(SPI_GETDESKWALLPAPER, MAX_PATH, tmpString, 0)
            If apiError = 0 Then Failed "SystemParametersInfo"
        
            txtCurrent.Text = tmpString
        End If
    End If
    
    lstNewWallpaper.AddItem "Default"
    lstNewWallpaper.AddItem "None"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub txtCurrent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'If enter key was pressed
        If txtCurrent.Text <> "" Then 'Must contain text
            lstNewWallpaper.AddItem txtCurrent.Text
        End If
    End If
End Sub
