VERSION 5.00
Begin VB.Form frmDiskSpace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disk Space"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "frmDiskSpace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDiskSpace 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   3195
      TabIndex        =   7
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Timer timerDiskSpace 
      Enabled         =   0   'False
      Interval        =   945
      Left            =   1800
      Top             =   720
   End
   Begin VB.ComboBox cboOutput 
      Height          =   315
      Left            =   2040
      TabIndex        =   12
      Text            =   "Gigabytes"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtTotalFreeSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.DriveListBox drvDiskSpace 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtTotalSpace 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtFreeSpaceAvail 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblDrive 
      Caption         =   "Drive"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblTotalNumberOfFreeBytesLabel 
      Caption         =   "Total Free Space"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblDiskSpacePerc 
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label lblFreeBytesAvailableLabel 
      Caption         =   "Free Space Available"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblTotalNumberOfBytesLabel 
      Caption         =   "Total Space"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblFreeSpace 
      Caption         =   "Free Space"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmDiskSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dims the nesccary variables needed for later
Dim FreeSpaceAvail As Double
Dim TotalSpace As Double
Dim TotalFreeSpace As Double
Dim OutputSize As Byte
Dim PercFreeSpace As Byte

'Extras
Dim liAvailable As LARGE_INTEGER 'See LARGE_INTEGER
Dim liTotal As LARGE_INTEGER
Dim liFree As LARGE_INTEGER

Private Sub cboOutput_Click()
    'Determine the Available Space, Total Size and Free Space of a drive
    apiError = GetDiskFreeSpaceEx(Left(drvDiskSpace.Drive, 2) & "\", liAvailable, liTotal, liFree)
    If apiError = 0 Then 'If error
        timerDiskSpace.Enabled = False 'Disable timer
        
        'Clear
        txtFreeSpaceAvail.Text = ""
        txtTotalFreeSpace.Text = ""
        txtTotalSpace.Text = ""
        picDiskSpace.Cls
        lblDiskSpacePerc = ""
        
        Failed "GetDiskFreeSpaceEx"
        Exit Sub
    Else
        timerDiskSpace.Enabled = True
    End If
    
    'Convert the return values from LARGE_INTEGER to doubles
    FreeSpaceAvail = CLargeInt(liAvailable.LowPart, liAvailable.HighPart)
    TotalSpace = CLargeInt(liTotal.LowPart, liTotal.HighPart)
    TotalFreeSpace = CLargeInt(liFree.LowPart, liFree.HighPart)
    PercFreeSpace = Round((FreeSpaceAvail / TotalSpace) * 100, 0)
    
    Dim Combo As Integer
    
    Combo = cboOutput.ListIndex 'Puts the index into the combo
    If Combo = -1 Then Combo = 4 'Default at GB

    Select Case Combo 'Allows selection
        Case 0 'bits
            '8bits = 1byte
            FreeSpaceAvail = FreeSpaceAvail * 8
            TotalSpace = TotalSpace * 8
            TotalFreeSpace = TotalFreeSpace * 8
            
            OutputSize = 0
        Case 1: OutputSize = 0  'bytes
        Case 2: OutputSize = 1  'kb
        Case 3: OutputSize = 2  'mb
        Case 4: OutputSize = 3  'gb
        Case 5: OutputSize = 4  'tb
    End Select
    
    'Divide
    FreeSpaceAvail = FreeSpaceAvail / (1024 ^ OutputSize)
    TotalSpace = TotalSpace / (1024 ^ OutputSize)
    TotalFreeSpace = TotalFreeSpace / (1024 ^ OutputSize)
    
    'Round to 6 places
    FreeSpaceAvail = Round(FreeSpaceAvail, 6)
    TotalSpace = Round(TotalSpace, 6)
    TotalFreeSpace = Round(TotalFreeSpace, 6)
    
    'Dump info to text boxes
    txtFreeSpaceAvail.Text = FreeSpaceAvail
    txtTotalFreeSpace.Text = TotalFreeSpace
    txtTotalSpace.Text = TotalSpace
    
    picDiskSpace.Cls
    
    If FreeSpaceAvail > 0 Then 'Cant divide with 0
        If TotalSpace > 0 Then 'Cant divide by 0
            picDiskSpace.Line (0, 0)-(PercFreeSpace, picDiskSpace.ScaleHeight), , BF
        Else
            picDiskSpace.Line (0, 0)-(0, picDiskSpace.ScaleHeight), , BF
        End If
    Else
        picDiskSpace.Line (0, 0)-(0, picDiskSpace.ScaleHeight), , BF
    End If
    lblDiskSpacePerc = PercFreeSpace & "%"
End Sub

Private Sub drvDiskSpace_Change()
    On Error Resume Next 'Just in case drive error
    Call cboOutput_Click
End Sub

Private Sub Form_Load()
    'Add items for choices
    cboOutput.AddItem "Bits"
    cboOutput.AddItem "Bytes"
    cboOutput.AddItem "Kilobytes"
    cboOutput.AddItem "Megabytes"
    cboOutput.AddItem "Gigabytes"
    cboOutput.AddItem "Terabytes"
    
    picDiskSpace.ScaleWidth = 100

    Call drvDiskSpace_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerDiskSpace.Enabled = False
    Unload Me
End Sub

Private Sub timerDiskSpace_Timer()
    Call cboOutput_Click
End Sub

