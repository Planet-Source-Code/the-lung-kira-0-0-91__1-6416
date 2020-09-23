VERSION 5.00
Begin VB.Form frmSharedFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shared Files"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmSharedFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLocation 
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   6495
   End
   Begin VB.ListBox lstLocation 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   350
      Left            =   5400
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "frmSharedFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
    If lstLocation.List(lstLocation.ListIndex) <> "" Then
        DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs", lstLocation.List(lstLocation.ListIndex)
    End If
End Sub

Private Sub Form_Load()
    Dim strValueName() As String
    Dim lngCount As Long
    Dim lngValueType As Long
    
    EnumValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs", strValueName(), lngCount, lngValueType
    
    Dim tmpLong As Long
    For tmpLong = 0 To lngCount - 1 'Cycle through the array
        'Do not add blank entries
        If Trim(strValueName(tmpLong)) <> "" Then
            lstLocation.AddItem strValueName(tmpLong)
        End If
    Next tmpLong
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstLocation_Click()
    txtLocation.Text = lstLocation.List(lstLocation.ListIndex)
End Sub
