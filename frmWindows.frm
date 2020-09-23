VERSION 5.00
Begin VB.Form frmWindows 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmWindows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNormalHeight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtNormalWidth 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox txtClientRectBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtClientRectRight 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtNormalPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtMaxPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtMinPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton cmdDestroy 
      Caption         =   "Destroy"
      Height          =   350
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "Minimize"
      Height          =   350
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdShowWindowApply 
      Caption         =   "Apply"
      Height          =   285
      Left            =   5640
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.ComboBox cboShowWindow 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1920
      TabIndex        =   10
      Top             =   2760
      Width           =   3495
   End
   Begin VB.CheckBox chkIsIconic 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   21
      Top             =   4560
      Width           =   255
   End
   Begin VB.CheckBox chkIsWindowUnicode 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox chkIsWindowVisible 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   19
      Top             =   6000
      Width           =   255
   End
   Begin VB.CheckBox chkIsZoomed 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   18
      Top             =   4320
      Width           =   255
   End
   Begin VB.TextBox txtWindows 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   2400
      Width           =   3495
   End
   Begin VB.CommandButton cmdWindowTitleApply 
      Caption         =   "Apply"
      Height          =   285
      Left            =   5640
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   5640
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.ListBox lstWindows 
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
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   6495
   End
   Begin VB.CommandButton cmdFlash 
      Caption         =   "Flash"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblNormalWidth 
      Caption         =   "Normal Width"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblNormalHeight 
      Caption         =   "Normal Height"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblNormalPosition 
      Caption         =   "Normal Position"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblMinPosition 
      Caption         =   "Minimized Position"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblMaxPosition 
      Caption         =   "Maximized Position"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label lblShowWindow 
      Caption         =   "Show Window"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblIsIconic 
      Caption         =   "Minimized"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblIsWindowUnicode 
      Caption         =   "Unicode"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lblIsWindowVisible 
      Caption         =   "Visible"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lblIsZoomed 
      Caption         =   "Maximized"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblClientRectBottom 
      Caption         =   "Height"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lblClientRectRight 
      Caption         =   "Width"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblWindowTitle 
      Caption         =   "Window Title"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblWindows 
      Caption         =   "Available Windows"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim w_hwnd As Long

Private Sub cmdDestroy_Click()
    If DestroyWindow(w_hwnd) = 0 Then Failed "CloseWindow"
End Sub

Private Sub cmdFlash_Click()
    apiError = FlashWindow(w_hwnd, True)
End Sub

Private Sub cmdMinimize_Click()
    If CloseWindow(w_hwnd) = 0 Then Failed "CloseWindow"
End Sub

Private Sub cmdRefresh_Click()
    'Clear
    lstWindows.Clear
    ReDim WindowListName(0)
    ReDim WindowListhWnd(0)
    WindowListNum = 0
    
    'Enumerate all the handles
    If EnumWindows(AddressOf EnumWindowsProc, 0&) = 0 Then
        Failed "EnumWindows"
    End If

    Dim tmpLong As Long
    For tmpLong = 0 To WindowListNum - 1 'Cycle through list
        lstWindows.AddItem Left(WindowListhWnd(tmpLong) & Space(8), 8) & WindowListName(tmpLong)
    Next tmpLong
End Sub

Private Sub cmdShowWindowApply_Click()
    Dim tmpInt As Integer
    
    Select Case cboShowWindow.List(cboShowWindow.ListIndex)
        Case "Force Minimize"
            'Requires 2k
            If WinID = "WIN32_NT" Then
                If WinVer > 5000000 Then
                    tmpInt = SW_FORCEMINIMIZE
                End If
            End If
        Case "Hide": tmpInt = SW_HIDE
        Case "Maximize": tmpInt = SW_MAXIMIZE
        Case "Minimize": tmpInt = SW_MINIMIZE
        Case "Restore": tmpInt = SW_RESTORE
        Case "Show": tmpInt = SW_SHOW
        Case "Show Default": tmpInt = SW_SHOWDEFAULT
        Case "Show Maximized": tmpInt = SW_SHOWMAXIMIZED
        Case "Show Minimized": tmpInt = SW_SHOWMINIMIZED
        Case "Show Minimized Not Activated": tmpInt = SW_SHOWMINNOACTIVE
        Case "Show NA": tmpInt = SW_SHOWNA
        Case "Show Not Activated": tmpInt = SW_SHOWNOACTIVATE
        Case "Show Normal": tmpInt = SW_SHOWNORMAL
    End Select
    
    apiError = ShowWindowAsync(w_hwnd, tmpInt)
End Sub

Private Sub cmdWindowTitleApply_Click()
    If SetWindowText(w_hwnd, txtWindows.Text & Chr(0)) = 0 Then
        Failed "SetWindowText"
    End If
End Sub

Private Sub Form_Load()
    Call cmdRefresh_Click
    
    With cboShowWindow
        .AddItem "Force Minimize"
        .AddItem "Hide"
        .AddItem "Maximize"
        .AddItem "Minimize"
        .AddItem "Restore"
        .AddItem "Show"
        .AddItem "Show Default"
        .AddItem "Show Maximized"
        .AddItem "Show Minimized"
        .AddItem "Show Minimized Not Activated"
        .AddItem "Show NA"
        .AddItem "Show Not Activated"
        .AddItem "Show Normal"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstWindows_Click()
    Dim tmpString As String
    
    tmpString = lstWindows.List(lstWindows.ListIndex)
    w_hwnd = CLng(Trim(Left$(tmpString, 8)))
    
    txtWindows.Text = Right$(tmpString, Len(tmpString) - 8)
    
    chkIsIconic.Value = CInt(IsIconic(w_hwnd))
    chkIsWindowUnicode.Value = CInt(IsWindowUnicode(w_hwnd))
    chkIsWindowVisible.Value = CInt(IsWindowVisible(w_hwnd))
    chkIsZoomed.Value = CInt(IsZoomed(w_hwnd))
    
    If GetClientRect(w_hwnd, RECT) = 0 Then Failed "GetClientRect"
    txtClientRectRight.Text = RECT.Right
    txtClientRectBottom.Text = RECT.Bottom
    
    
    WINDOWPLACEMENT.Length = Len(WINDOWPLACEMENT)
    If GetWindowPlacement(w_hwnd, WINDOWPLACEMENT) = 0 Then Failed "GetWindowPlacement"
    
    With WINDOWPLACEMENT
        txtMinPosition.Text = .ptMinPosition.X & "," & .ptMinPosition.Y
        txtMaxPosition.Text = .ptMaxPosition.X & "," & .ptMaxPosition.Y
        
        'Doesnt = width or height because this is the total height width of the window including borders n such
        txtNormalHeight.Text = .rcNormalPosition.Bottom - .rcNormalPosition.Top
        txtNormalWidth.Text = .rcNormalPosition.Right - .rcNormalPosition.Left
        
        txtNormalPosition.Text = .rcNormalPosition.Left & "," & .rcNormalPosition.Top
        cboShowWindow.ListIndex = .showCmd
    End With
End Sub
