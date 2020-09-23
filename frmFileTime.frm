VERSION 5.00
Begin VB.Form frmFileTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Time"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "frmFileTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   350
      Left            =   4800
      TabIndex        =   37
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3840
      TabIndex        =   36
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtMillisecondLW 
      Height          =   285
      Left            =   4440
      TabIndex        =   34
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtMillisecondLA 
      Height          =   285
      Left            =   2880
      TabIndex        =   26
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtMillisecondCT 
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtSecondLW 
      Height          =   285
      Left            =   4440
      TabIndex        =   33
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtSecondLA 
      Height          =   285
      Left            =   2880
      TabIndex        =   25
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtSecondCT 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtMinuteLW 
      Height          =   285
      Left            =   4440
      TabIndex        =   32
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtMinuteLA 
      Height          =   285
      Left            =   2880
      TabIndex        =   24
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtMinuteCT 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtHourLW 
      Height          =   285
      Left            =   4440
      TabIndex        =   31
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtHourLA 
      Height          =   285
      Left            =   2880
      TabIndex        =   23
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtHourCT 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtDayLW 
      Height          =   285
      Left            =   4440
      TabIndex        =   30
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtDayLA 
      Height          =   285
      Left            =   2880
      TabIndex        =   22
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtDayCT 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtDayOfWeekLW 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtDayOfWeekLA 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtDayOfWeekCT 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtMonthLW 
      Height          =   285
      Left            =   4440
      TabIndex        =   28
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtMonthLA 
      Height          =   285
      Left            =   2880
      TabIndex        =   20
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtMonthCT 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtYearLW 
      Height          =   285
      Left            =   4440
      TabIndex        =   27
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtYearLA 
      Height          =   285
      Left            =   2880
      TabIndex        =   19
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtYearCT 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2880
      TabIndex        =   35
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblMillisecond 
      Caption         =   "Millisecond"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblSecond 
      Caption         =   "Second"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblMinute 
      Caption         =   "Minute"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblHour 
      Caption         =   "Hour"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblDay 
      Caption         =   "Day"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblDayOfWeek 
      Caption         =   "Day of Week"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblMonth 
      Caption         =   "Month"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblLastWriteTime 
      Caption         =   "Last Write Time"
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblLastAccessTime 
      Caption         =   "Last Access Time"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblCreationTime 
      Caption         =   "Creation Time"
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmFileTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFileName As String
Dim hFile As Long

'File Time
Dim ftCreationTime As FILETIME
Dim ftLastAccess As FILETIME
Dim ftLastWrite As FILETIME
'System Time
Dim stCreationTime As SYSTEMTIME
Dim stLastAccess As SYSTEMTIME
Dim stLastWrite As SYSTEMTIME

Private Sub cmdApply_Click()
    On Error Resume Next

    'Fill in structures
    With stCreationTime
        .wYear = CInt(txtYearCT.Text)
        .wMonth = CInt(txtMonthCT.Text)
        .wDayOfWeek = CInt(txtDayOfWeekCT.Text) 'Ignored
        .wDay = CInt(txtDayCT.Text)
        .wHour = CInt(txtHourCT.Text)
        .wMinute = CInt(txtMinuteCT.Text)
        .wSecond = CInt(txtSecondCT.Text)
        .wMilliseconds = CInt(txtMillisecondCT.Text)
    End With
    With stLastAccess
        .wYear = CInt(txtYearLA.Text)
        .wMonth = CInt(txtMonthLA.Text)
        .wDayOfWeek = CInt(txtDayOfWeekLA.Text) 'Ignored
        .wDay = CInt(txtDayLA.Text)
        .wHour = CInt(txtHourLA.Text)
        .wMinute = CInt(txtMinuteLA.Text)
        .wSecond = CInt(txtSecondLA.Text)
        .wMilliseconds = CInt(txtMillisecondLA.Text)
    End With
    With stLastWrite
        .wYear = CInt(txtYearLW.Text)
        .wMonth = CInt(txtMonthLW.Text)
        .wDayOfWeek = CInt(txtDayOfWeekLW.Text) 'Ignored
        .wDay = CInt(txtDayLW.Text)
        .wHour = CInt(txtHourLW.Text)
        .wMinute = CInt(txtMinuteLW.Text)
        .wSecond = CInt(txtSecondLW.Text)
        .wMilliseconds = CInt(txtMillisecondLW.Text)
    End With
    
    'Convert
    apiError = SystemTimeToFileTime(stCreationTime, ftCreationTime)
    If apiError = 0 Then Failed "SystemTimeToFileTime"
    apiError = SystemTimeToFileTime(stLastAccess, ftLastAccess)
    If apiError = 0 Then Failed "SystemTimeToFileTime"
    apiError = SystemTimeToFileTime(stLastWrite, ftLastWrite)
    If apiError = 0 Then Failed "SystemTimeToFileTime"
    
    If SetFileTime(hFile, ftCreationTime, ftLastAccess, ftLastWrite) = 0 Then Failed "SetFileTime"
End Sub

Private Sub cmdChoose_Click()
    GetOpenName hwnd, "Open", strFileName
    
    'Error checking
    If Not strFileName <> "" Then Exit Sub
    
    cmdChoose.Enabled = False
    Flush
    
    'Fill in security attributes
    SECURITY_ATTRIBUTES.nLength = Len(SECURITY_ATTRIBUTES)
    SECURITY_ATTRIBUTES.lpSecurityDescriptor = 0
    SECURITY_ATTRIBUTES.bInheritHandle = False
    
    hFile = CreateFile(strFileName & Chr(0), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, SECURITY_ATTRIBUTES, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    If hFile = -1 Then
        Failed "CreateFile"
        Exit Sub 'Exit here cant continue
    End If
    
    If GetFileTime(hFile, ftCreationTime, ftLastAccess, ftLastWrite) = 0 Then Failed "GetFileTime"
    
    'Convert
    apiError = FileTimeToSystemTime(ftCreationTime, stCreationTime)
    If apiError = 0 Then Failed "FiletimeToSystemTime"
    apiError = FileTimeToSystemTime(ftLastAccess, stLastAccess)
    If apiError = 0 Then Failed "FiletimeToSystemTime"
    apiError = FileTimeToSystemTime(ftLastWrite, stLastWrite)
    If apiError = 0 Then Failed "FiletimeToSystemTime"
    
    'Fill in text boxes
    With stCreationTime
        txtYearCT.Text = .wYear
        txtMonthCT.Text = .wMonth
        txtDayOfWeekCT.Text = .wDayOfWeek
        txtDayCT.Text = .wDay
        txtHourCT.Text = .wHour
        txtMinuteCT.Text = .wMinute
        txtSecondCT.Text = .wSecond
        txtMillisecondCT.Text = .wMilliseconds
    End With
    With stLastAccess
        txtYearLA.Text = .wYear
        txtMonthLA.Text = .wMonth
        txtDayOfWeekLA.Text = .wDayOfWeek
        txtDayLA.Text = .wDay
        txtHourLA.Text = .wHour
        txtMinuteLA.Text = .wMinute
        txtSecondLA.Text = .wSecond
        txtMillisecondLA.Text = .wMilliseconds
    End With
    With stLastWrite
        txtYearLW.Text = .wYear
        txtMonthLW.Text = .wMonth
        txtDayOfWeekLW.Text = .wDayOfWeek
        txtDayLW.Text = .wDay
        txtHourLW.Text = .wHour
        txtMinuteLW.Text = .wMinute
        txtSecondLW.Text = .wSecond
        txtMillisecondLW.Text = .wMilliseconds
    End With

    cmdApply.Enabled = True 'Allows applying after a file has been choosen and processed
    cmdClose.Enabled = True
End Sub

Private Sub cmdClose_Click()
    cmdApply.Enabled = False
    
    If CloseHandle(hFile) = 0 Then Failed "CloseHandle"

    Flush
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If file open the close it
    If cmdClose.Enabled = True Then Call cmdClose_Click
    
    Unload Me
End Sub

Private Sub Flush()
    'Clear all
    txtYearCT.Text = ""
    txtMonthCT.Text = ""
    txtDayOfWeekCT.Text = ""
    txtDayCT.Text = ""
    txtHourCT.Text = ""
    txtMinuteCT.Text = ""
    txtSecondCT.Text = ""
    txtMillisecondCT.Text = ""

    txtYearLA.Text = ""
    txtMonthLA.Text = ""
    txtDayOfWeekLA.Text = ""
    txtDayLA.Text = ""
    txtHourLA.Text = ""
    txtMinuteLA.Text = ""
    txtSecondLA.Text = ""
    txtMillisecondLA.Text = ""
    
    txtYearLW.Text = ""
    txtMonthLW.Text = ""
    txtDayOfWeekLW.Text = ""
    txtDayLW.Text = ""
    txtHourLW.Text = ""
    txtMinuteLW.Text = ""
    txtSecondLW.Text = ""
    txtMillisecondLW.Text = ""
End Sub
