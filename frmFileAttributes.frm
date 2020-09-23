VERSION 5.00
Begin VB.Form frmFileAttributes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Attributes"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   Icon            =   "frmFileAttributes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkTemporary 
      Height          =   255
      Left            =   2400
      TabIndex        =   26
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkSystem 
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkSparseFile 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   22
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkReparsePoint 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkReadOnly 
      Height          =   255
      Left            =   2400
      TabIndex        =   18
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkOffline 
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkNotContentIndexed 
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkNormal 
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkHidden 
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkEncrypted 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkDirectory 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkCompressed 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkArchive 
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   350
      Left            =   1680
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   350
      Left            =   720
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblTemporary 
      Caption         =   "Temporary"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblSystem 
      Caption         =   "System"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label lblSparseFile 
      Caption         =   "Sparse File"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblReparsePoint 
      Caption         =   "Reparse Point"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblReadOnly 
      Caption         =   "Read Only"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblOffline 
      Caption         =   "Offline"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label lblNotContentIndexed 
      Caption         =   "Not Content Indexed"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblNormal 
      Caption         =   "Normal"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblHidden 
      Caption         =   "Hidden"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblEncrypted 
      Caption         =   "Encrypted"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblDirectory 
      Caption         =   "Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblCompressed 
      Caption         =   "Compressed"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblArchive 
      Caption         =   "Archive"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmFileAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFileName As String
Dim FileAttributes As Long

Private Sub chkNormal_Click()
    'Normal is no attributes set
    chkArchive.Value = 0
    chkHidden.Value = 0
    chkNotContentIndexed.Value = 0
    chkOffline.Value = 0
    chkReadOnly.Value = 0
    chkSystem.Value = 0
    chkTemporary.Value = 0
End Sub

Private Sub cmdApply_Click()
    If Not strFileName <> "" Then Exit Sub
    
    FileAttributes = 0& 'Reset
    
    Dim Archive As Long
    Dim Hidden As Long
    Dim NotContentIndexed As Long
    Dim Offline As Long
    Dim ReadOnly As Long
    Dim System As Long
    Dim Temporary As Long
    
    If chkArchive.Value = 1 Then Archive = FILE_ATTRIBUTE_ARCHIVE
    If chkHidden.Value = 1 Then Hidden = FILE_ATTRIBUTE_HIDDEN
    If chkNotContentIndexed.Value = 1 Then NotContentIndexed = FILE_ATTRIBUTE_NOT_CONTENT_INDEXED
    If chkOffline.Value = 1 Then Offline = FILE_ATTRIBUTE_OFFLINE
    If chkReadOnly.Value = 1 Then ReadOnly = FILE_ATTRIBUTE_READONLY
    If chkSystem.Value = 1 Then System = FILE_ATTRIBUTE_SYSTEM
    If chkTemporary.Value = 1 Then Temporary = FILE_ATTRIBUTE_TEMPORARY
    
    If chkNormal.Value = 1 Then 'Overrides all
        FileAttributes = FILE_ATTRIBUTE_NORMAL
    Else
        FileAttributes = Archive Or Hidden Or NotContentIndexed Or Offline Or ReadOnly Or System Or Temporary
    End If

    If SetFileAttributes(strFileName, FileAttributes) = 0 Then
        Failed "SetFileAttributes"
    End If
End Sub

Private Sub cmdChoose_Click()
    'Clear
    strFileName = ""
    chkArchive.Value = 0
    chkCompressed.Value = 0
    chkDirectory.Value = 0
    chkEncrypted.Value = 0
    chkHidden.Value = 0
    chkNormal.Value = 0
    chkNotContentIndexed.Value = 0
    chkOffline.Value = 0
    chkReadOnly.Value = 0
    chkReparsePoint.Value = 0
    chkSparseFile.Value = 0
    chkSystem.Value = 0
    chkTemporary.Value = 0
    cmdApply.Enabled = False
    
    GetOpenName hwnd, "Open", strFileName
    
    'Error checking
    If Not strFileName <> "" Then Exit Sub

    FileAttributes = GetFileAttributes(strFileName)
    If FileAttributes = -1 Then
        Failed "GetFileAttributes"
        Exit Sub 'Exit here
    End If

    If FileAttributes And FILE_ATTRIBUTE_ARCHIVE Then chkArchive.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_COMPRESSED Then chkCompressed.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_DIRECTORY Then chkDirectory.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_ENCRYPTED Then chkEncrypted.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_HIDDEN Then chkHidden.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_NORMAL Then chkNormal.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_NOT_CONTENT_INDEXED Then chkNotContentIndexed.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_OFFLINE Then chkOffline.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_READONLY Then chkReadOnly.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_REPARSE_POINT Then chkReparsePoint.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_SPARSE_FILE Then chkSparseFile.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_SYSTEM Then chkSystem.Value = 1
    If FileAttributes And FILE_ATTRIBUTE_TEMPORARY Then chkTemporary.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
