VERSION 5.00
Begin VB.Form frmUpTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Up Time"
   ClientHeight    =   1590
   ClientLeft      =   -10620
   ClientTop       =   795
   ClientWidth     =   1950
   Icon            =   "frmUpTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   1950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerUpTime 
      Interval        =   945
      Left            =   840
      Top             =   360
   End
   Begin VB.TextBox txtSeconds 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtMinutes 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtHours 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtDays 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblSeconds 
      Caption         =   "Seconds"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblMinutes 
      Caption         =   "Minutes"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblHours 
      Caption         =   "Hours"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblDays 
      Caption         =   "Days"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmUpTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Days As Integer 'Yeah whos windows is going to stay on for 32767 days
Dim Hours As Byte
Dim Minutes As Byte
Dim Seconds As Byte
Dim UpTime As Long

Private Sub Form_Load()
    Call timerUpTime_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timerUpTime.Enabled = False
    Unload Me
End Sub

Private Sub timerUpTime_Timer()
    UpTime = GetTickCount
    
    'Formats the data then subtracts the slack
    Days = Int(UpTime / 1000 / 60 ^ 2 / 24)
    Hours = (Int(UpTime / 1000 / 60 ^ 2) - (Int(UpTime / 1000 / 60 ^ 2 / 24) * 24))
    Minutes = (Int(UpTime / 1000 / 60) - (Int(UpTime / 1000 / 60 ^ 2) * 60))
    Seconds = (Int(UpTime / 1000) - (Int(UpTime / 1000 / 60) * 60))
    
    'Dumps to text box
    txtDays.Text = Days
    txtHours.Text = Hours
    txtMinutes.Text = Minutes
    txtSeconds.Text = Seconds
End Sub
