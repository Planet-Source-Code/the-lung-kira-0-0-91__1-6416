VERSION 5.00
Begin VB.Form frmDiskVolume 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disk Volume Info"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmDiskVolume.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkVOLUME_QUOTAS 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkSUPPORTS_SPARSE_FILES 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   31
      Top             =   4200
      Width           =   255
   End
   Begin VB.CheckBox chkSUPPORTS_REPARSE_POINTS 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox chkSUPPORTS_OBJECT_IDS 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   25
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkSUPPORTS_ENCRYPTION 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkNAMED_STREAMS 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chkCASE_IS_PRESERVED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtVolumeLabel 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtSerialNumber 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtFileSystem 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.CheckBox chkVOL_IS_COMPRESSED 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chkFILE_COMPRESSION 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   2400
      Width           =   255
   End
   Begin VB.CheckBox chkPERSISTENT_ACLS 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   27
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox chkUNICODE_STORED_ON_DISK 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkCASE_SENSITIVE 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtComponentLength 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   2880
      TabIndex        =   33
      Top             =   4560
      Width           =   1095
   End
   Begin VB.DriveListBox drvVolumeInfo 
      Height          =   315
      Left            =   120
      TabIndex        =   32
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label lblVOLUME_QUOTAS 
      Caption         =   "Disk Quotas Supported"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label lblSUPPORTS_SPARSE_FILES 
      Caption         =   "Sparse Files Supported"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label lblSUPPORTS_REPARSE_POINTS 
      Caption         =   "Reparse Points Supported"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label lblSUPPORTS_OBJECT_IDS 
      Caption         =   "Object Identifiers Supported"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label lblSUPPORTS_ENCRYPTION 
      Caption         =   "Encrypted File System Supported"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label lblNAMED_STREAMS 
      Caption         =   "Named Streams Supported"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label lblCASE_IS_PRESERVED 
      Caption         =   "Preserve Case of Filename Supported"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label lblVOL_IS_COMPRESSED 
      Caption         =   "Compressed Volume"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblFILE_COMPRESSION 
      Caption         =   "File Based Compression Supported"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label lblPERSISTENT_ACLS 
      Caption         =   "Preserve and Enforce ACLs"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label lblUNICODE_STORED_ON_DISK 
      Caption         =   "Unicode In Filenames Supported"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblCASE_SENSITIVE 
      Caption         =   "Case Sensitive Filenames Supported"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblComponentLength 
      Caption         =   "Component Length"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblFileSystem 
      Caption         =   "File System"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblSerialNumber 
      Caption         =   "Serial Number"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblVolumeLabel 
      Caption         =   "Volume Label"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
End
Attribute VB_Name = "frmDiskVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    If SetVolumeLabel(Left(drvVolumeInfo, 2) & "\", txtVolumeLabel.Text) = 0 Then
        Failed "SetVolumeLabel"
    End If
End Sub

Private Sub drvVolumeInfo_Change()
    Call Form_Load 'Refreshes
End Sub

Private Sub Form_Load()
    'Dims nesccary values for volume info
    Dim VolumeNameBuffer As String * 11 'Can only be 11 bytes
    Dim VolumeSerialNumber As Long
    Dim MaximumComponentLength As Long
    Dim FileSystemFlags As Long 'Never used, need to be used in future
    Dim FileSystemNameBuffer As String
    
    'Padding
    VolumeNameBuffer = Space(11)
    FileSystemNameBuffer = Space(255)

    'Clear text boxes
    txtComponentLength.Text = ""
    txtFileSystem.Text = ""
    txtSerialNumber.Text = ""
    txtVolumeLabel.Text = ""
    'and checks
    chkCASE_IS_PRESERVED.Value = 0
    chkCASE_SENSITIVE.Value = 0
    chkUNICODE_STORED_ON_DISK.Value = 0
    chkPERSISTENT_ACLS.Value = 0
    chkFILE_COMPRESSION.Value = 0
    chkVOL_IS_COMPRESSED.Value = 0
    chkNAMED_STREAMS.Value = 0
    chkSUPPORTS_ENCRYPTION.Value = 0
    chkSUPPORTS_OBJECT_IDS.Value = 0
    chkSUPPORTS_REPARSE_POINTS.Value = 0
    chkSUPPORTS_SPARSE_FILES.Value = 0
    chkVOLUME_QUOTAS.Value = 0

    If GetVolumeInformation(Left(drvVolumeInfo, 2) & "\", VolumeNameBuffer, Len(VolumeNameBuffer), VolumeSerialNumber, MaximumComponentLength, FileSystemFlags, FileSystemNameBuffer, Len(FileSystemNameBuffer)) = 0 Then
        Failed "GetVolumeInformation"
        Exit Sub 'Exit from here
    End If
    
    'Dumps the info to the text boxes
    txtComponentLength.Text = MaximumComponentLength
    txtFileSystem.Text = Trim(FileSystemNameBuffer)
    txtSerialNumber.Text = Hex(VolumeSerialNumber) 'Does like windows does
    txtVolumeLabel.Text = Trim(VolumeNameBuffer)
    
    'First half
    If FileSystemFlags And FS_CASE_IS_PRESERVED Then chkCASE_IS_PRESERVED.Value = 1
    If FileSystemFlags And FS_CASE_SENSITIVE Then chkCASE_SENSITIVE.Value = 1
    If FileSystemFlags And FS_UNICODE_STORED_ON_DISK Then chkUNICODE_STORED_ON_DISK.Value = 1
    If FileSystemFlags And FS_PERSISTENT_ACLS Then chkPERSISTENT_ACLS.Value = 1
    If FileSystemFlags And FS_FILE_COMPRESSION Then chkFILE_COMPRESSION.Value = 1
    If FileSystemFlags And FS_VOL_IS_COMPRESSED Then chkVOL_IS_COMPRESSED.Value = 1
    'Second half
    If FileSystemFlags And FILE_NAMED_STREAMS Then chkNAMED_STREAMS.Value = 1
    If FileSystemFlags And FILE_SUPPORTS_ENCRYPTION Then chkSUPPORTS_ENCRYPTION.Value = 1
    If FileSystemFlags And FILE_SUPPORTS_OBJECT_IDS Then chkSUPPORTS_OBJECT_IDS.Value = 1
    If FileSystemFlags And FILE_SUPPORTS_REPARSE_POINTS Then chkSUPPORTS_REPARSE_POINTS.Value = 1
    If FileSystemFlags And FILE_SUPPORTS_SPARSE_FILES Then chkSUPPORTS_SPARSE_FILES.Value = 1
    If FileSystemFlags And FILE_VOLUME_QUOTAS Then chkVOLUME_QUOTAS.Value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub txtVolumeLabel_Change()
    'Removes all the characters not allowed in a volume name
    txtVolumeLabel.Text = Rem_NonFat_Chr(txtVolumeLabel.Text)
    
    'Maximum length for volume label is 11 char
    If Len(txtVolumeLabel.Text) > 11 Then txtVolumeLabel.Text = Left(txtVolumeLabel.Text, 11)
End Sub
