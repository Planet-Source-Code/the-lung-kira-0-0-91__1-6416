VERSION 5.00
Begin VB.Form frmWindowKiller 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Window Killer"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmWindowKiller.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKillInterval 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   1215
   End
   Begin VB.HScrollBar hsKillInterval 
      Height          =   135
      LargeChange     =   5
      Left            =   120
      Max             =   30000
      TabIndex        =   7
      Top             =   4635
      Value           =   1
      Width           =   1215
   End
   Begin VB.CommandButton cmdImportList 
      Caption         =   "Import List"
      Height          =   350
      Left            =   5640
      TabIndex        =   12
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   350
      Left            =   4680
      TabIndex        =   11
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   350
      Left            =   5640
      TabIndex        =   10
      Top             =   4080
      Width           =   975
   End
   Begin VB.ListBox lstWindows 
      Height          =   1425
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   6495
   End
   Begin VB.TextBox txtKill 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   6495
   End
   Begin VB.ListBox lstKill 
      Height          =   1425
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   6495
   End
   Begin VB.CommandButton cmdRefreshWin 
      Caption         =   "Refresh Windows"
      Height          =   350
      Left            =   3120
      TabIndex        =   8
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   350
      Left            =   4680
      TabIndex        =   9
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblKillInterval 
      Caption         =   "Kill Interval"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblWindows 
      Caption         =   "Available Windows"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblWindowKiller 
      Caption         =   "Window Killer List"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "frmWindowKiller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    If txtKill.Text <> "" Then 'Cant enter a blank entry
        lstKill.AddItem txtKill.Text 'Add item to list box
        txtKill.Text = ""
        Call Refresh_Array
    End If
End Sub

Private Sub cmdClearAll_Click()
    lstKill.Clear
    Call Refresh_Array
End Sub

Private Sub cmdImportList_Click()
    Dim strFileName As String
    GetOpenName hwnd, "Open", strFileName
    
    'Error checking
    If Not strFileName <> "" Then Exit Sub 'Dont worry just exit
    If Not FileLen(strFileName) > 0 Then Exit Sub 'If file len not greater than 0
    
    Dim tmpString As String
    Open strFileName For Input As #1
        Do While Not EOF(1) 'Loop until end of file
            Line Input #1, tmpString 'Read line into variable
            If tmpString <> "" Then lstKill.AddItem tmpString
        Loop
    Close #1
    Call Refresh_Array
End Sub

Private Sub cmdRefreshWin_Click()
    'Clear
    lstWindows.Clear
    ReDim WindowListName(0)
    ReDim WindowListhWnd(0)
    WindowListNum = 0
    
    'Enumerate all the handles
    apiError = EnumWindows(AddressOf EnumWindowsProc, 0&)
    If apiError = 0 Then Failed "EnumWindows"

    Dim tmpLong As Long
    For tmpLong = 0 To WindowListNum - 1 'Cycle through list
        If WindowListName(tmpLong) <> "" Then 'If text then add
            lstWindows.AddItem WindowListName(tmpLong)
        End If
    Next tmpLong
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next 'Rather skip the error than call the function 2x
    lstKill.RemoveItem lstKill.ListIndex
    Call Refresh_Array
End Sub

Private Sub Form_Load()
    If WindowKillerNum > 0 Then 'Cant add nothing
        Dim tmpInt As Integer
    
        For tmpInt = 1 To WindowKillerNum 'Cycle through array
            lstKill.AddItem WindowKiller(tmpInt)
        Next tmpInt
    End If
    
    hsKillInterval.Value = frmMain.timerWindowKiller.Interval
    Call cmdRefreshWin_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub hsKillInterval_Change()
    txtKillInterval.Text = hsKillInterval.Value
    frmMain.timerWindowKiller.Interval = hsKillInterval.Value
End Sub

Private Sub lstKill_Click()
    txtKill.Text = lstKill.List(lstKill.ListIndex)
End Sub

Private Sub lstWindows_Click()
    txtKill.Text = lstWindows.List(lstWindows.ListIndex)
End Sub

Private Function Refresh_Array()
    If lstKill.ListCount = 0 Then
        Erase WindowKiller() 'Erase array if nothing
    Else
        Dim tmpInt As Integer
        
        For tmpInt = 0 To lstKill.ListCount - 1 'Cycle through list box
            ReDim Preserve WindowKiller(tmpInt + 1) 'Resize array without destroying
            WindowKiller(tmpInt + 1) = lstKill.List(tmpInt)
        Next tmpInt
        
        WindowKillerNum = tmpInt
    End If
End Function

Private Sub txtKillInterval_Change()
    On Error Resume Next
    
    If CInt(txtKillInterval.Text) <= 0 Then txtKillInterval.Text = "1" 'If less than 0 resets to min , also does error trapping
    hsKillInterval.Value = CInt(txtKillInterval.Text)
End Sub
