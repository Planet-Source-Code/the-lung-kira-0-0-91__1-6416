VERSION 5.00
Begin VB.Form frmMemoryStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Memory Status"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "frmMemoryStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMemoryLoad 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3240
      ScaleHeight     =   315
      ScaleWidth      =   2355
      TabIndex        =   22
      Top             =   1800
      Width           =   2415
   End
   Begin VB.PictureBox picPageFile 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   2355
      TabIndex        =   15
      Top             =   1800
      Width           =   2415
   End
   Begin VB.PictureBox picVirtual 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3240
      ScaleHeight     =   315
      ScaleWidth      =   2355
      TabIndex        =   8
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox picPhys 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Timer timerMemoryStatus 
      Interval        =   945
      Left            =   2760
      Top             =   1440
   End
   Begin VB.TextBox txtAvailPageFile 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtTotalPageFile 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtAvailVirtual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtTotalVirtual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtAvailPhys 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtTotalPhys 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox cboOptions 
      Height          =   315
      Left            =   4440
      TabIndex        =   25
      Text            =   "Megabytes"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output"
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Mem Load"
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblMemoryLoad 
      Height          =   255
      Left            =   5760
      TabIndex        =   23
      Top             =   1845
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Pagefile Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblPageFilePercentage 
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   1845
      Width           =   375
   End
   Begin VB.Label lblAvailPageFile 
      Caption         =   "Available Pagefile"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblTotalPageFile 
      Caption         =   "Total Pagefile"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Virtual Memory"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblTotalVirtual 
      Caption         =   "Total Virtual"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblAvailVirtual 
      Caption         =   "Available Virtual"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblVirtualPercentage 
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   405
      Width           =   375
   End
   Begin VB.Label lblPhys 
      Caption         =   "Physical Memory"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblTotalPhys 
      Caption         =   "Total Physical"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblAvailPhys 
      Caption         =   "Available Physical"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblPhysPercentage 
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   405
      Width           =   375
   End
End
Attribute VB_Name = "frmMemoryStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TotalPageFile As Double
Dim AvailPageFile As Double
Dim PercPageFile As Byte
Dim TotalPhys As Double
Dim AvailPhys As Double
Dim PercPhys As Byte
Dim TotalVirtual As Double
Dim AvailVirtual As Double
Dim PercVirtual As Byte

Private Sub cboOptions_Click()
    GlobalMemoryStatus MEMORYSTATUS 'Dumps the information to MemoryStatus
    
    With MEMORYSTATUS
        TotalPageFile = .dwTotalPageFile
        AvailPageFile = .dwAvailPageFile
        TotalPhys = .dwTotalPhys
        AvailPhys = .dwAvailPhys
        TotalVirtual = .dwTotalVirtual
        AvailVirtual = .dwAvailVirtual
    End With
    
    Dim Combo As Integer
    Dim OutputSize As Byte
    
    Combo = cboOptions.ListIndex 'Puts the index into the combo
    If Combo = -1 Then Combo = 3 'Defaults at MB
    
    Select Case Combo 'Allows selection
        Case 0  'bits
            '8bits = 1byte
            TotalPageFile = TotalPageFile * 8
            AvailPageFile = AvailPageFile * 8
            TotalPhys = TotalPhys * 8
            AvailPhys = AvailPhys * 8
            TotalVirtual = TotalVirtual * 8
            AvailVirtual = AvailVirtual * 8
            
            OutputSize = 0
        Case 1: OutputSize = 0  'bytes
        Case 2: OutputSize = 1  'kb
        Case 3: OutputSize = 2  'mb
        Case 4: OutputSize = 3  'gb
        Case 5: OutputSize = 4  'tb
    End Select
    
    'Divide
    TotalPageFile = TotalPageFile / (1024 ^ OutputSize)
    AvailPageFile = AvailPageFile / (1024 ^ OutputSize)
    TotalPhys = TotalPhys / (1024 ^ OutputSize)
    AvailPhys = AvailPhys / (1024 ^ OutputSize)
    TotalVirtual = TotalVirtual / (1024 ^ OutputSize)
    AvailVirtual = AvailVirtual / (1024 ^ OutputSize)
    
    'Round to 6 places
    txtTotalPageFile.Text = Round(TotalPageFile, 6)
    txtAvailPageFile.Text = Round(AvailPageFile, 6)
    txtTotalPhys.Text = Round(TotalPhys, 6)
    txtAvailPhys.Text = Round(AvailPhys, 6)
    txtTotalVirtual.Text = Round(TotalVirtual, 6)
    txtAvailVirtual.Text = Round(AvailVirtual, 6)
    
    'Get percentages
    PercPageFile = Round((AvailPageFile / TotalPageFile) * 100, 0)
    PercPhys = Round((AvailPhys / TotalPhys) * 100, 0)
    PercVirtual = Round((AvailVirtual / TotalVirtual) * 100, 0)
    
    'Sets all the bars max/value and labels %
    lblMemoryLoad.Caption = MEMORYSTATUS.dwMemoryLoad & "%"
    picMemoryLoad.Line (0, 0)-(MEMORYSTATUS.dwMemoryLoad, picMemoryLoad.ScaleHeight), , BF

    'Pagefile Mem
    picPageFile.Cls
    lblPageFilePercentage.Caption = PercPageFile & "%"
    picPageFile.Line (0, 0)-(PercPageFile, picPageFile.ScaleHeight), , BF
    
    'Physical Mem
    picPhys.Cls
    lblPhysPercentage.Caption = PercPhys & "%"
    picPhys.Line (0, 0)-(PercPhys, picPhys.ScaleHeight), , BF
    
    'Virtual Mem
    picVirtual.Cls
    lblVirtualPercentage.Caption = PercVirtual & "%"
    picVirtual.Line (0, 0)-(PercVirtual, picVirtual.ScaleHeight), , BF
End Sub

Private Sub Form_Load()
    'Adds needed items to combo box
    cboOptions.AddItem "Bits"
    cboOptions.AddItem "Bytes"
    cboOptions.AddItem "Kilobytes"
    cboOptions.AddItem "Megabytes"
    cboOptions.AddItem "Gigabytes"
    cboOptions.AddItem "Terabytes"
    
    picMemoryLoad.ScaleWidth = 100
    picPageFile.ScaleWidth = 100
    picPhys.ScaleWidth = 100
    picVirtual.ScaleWidth = 100
    
    Call cboOptions_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerMemoryStatus.Enabled = False
    Unload Me
End Sub

Private Sub timerMemoryStatus_Timer()
    Call cboOptions_Click
End Sub
