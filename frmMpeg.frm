VERSION 5.00
Begin VB.Form frmMpeg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MPEG"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "frmMpeg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEmphasis 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox chkCopyright 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   27
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtMode 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtSamplingFreq 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtBitrate 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox chkErrorProtection 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   19
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox chkOriginal 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   29
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox txtVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemoveTag 
      Caption         =   "Rem Tag"
      Enabled         =   0   'False
      Height          =   350
      Left            =   1080
      TabIndex        =   33
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdAddTag 
      Caption         =   "Add Tag"
      Enabled         =   0   'False
      Height          =   350
      Left            =   120
      TabIndex        =   32
      Top             =   2760
      Width           =   975
   End
   Begin VB.ComboBox cboGenre 
      Height          =   315
      Left            =   1320
      TabIndex        =   13
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtComments 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtAlbum 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtTag 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   350
      Left            =   6000
      TabIndex        =   35
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5040
      TabIndex        =   34
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblEmphasis 
      Caption         =   "Emphasis"
      Height          =   255
      Left            =   4560
      TabIndex        =   30
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblOriginal 
      Caption         =   "Original"
      Height          =   255
      Left            =   4560
      TabIndex        =   28
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright"
      Height          =   255
      Left            =   4560
      TabIndex        =   26
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblMode 
      Caption         =   "Mode"
      Height          =   255
      Left            =   4560
      TabIndex        =   24
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblSamplingFreq 
      Caption         =   "Sampling Freq"
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblBitrate 
      Caption         =   "Bitrate"
      Height          =   255
      Left            =   4560
      TabIndex        =   20
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblErrorProtection 
      Caption         =   "Error Protection"
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblLayer 
      Caption         =   "Layer"
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblGenre 
      Caption         =   "Genre"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblComments 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblYear 
      Caption         =   "Year"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblAlbum 
      Caption         =   "Album"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblArtist 
      Caption         =   "Artist"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Caption         =   "Title"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblTag 
      Caption         =   "Tag"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMpeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFileName As String

Private Sub cmdAddTag_Click()
    If Not strFileName <> "" Then Exit Sub
    
    cmdAddTag.Enabled = False 'Reset
    
    Dim tmpString As String
    
    Open strFileName For Append As #1 'Opens it for binary
        tmpString = Space(128) 'Add 128bytes of spaces
        Print #1, tmpString; 'Tack on to the end
    Close #1
    
    txtTag.Text = "True" 'Set to true cause we wrote one
    
    cmdApply.Enabled = True
    Call cmdApply_Click 'Moves back and writes over the tacked on 128 bytes with the tag
End Sub

Private Sub cmdApply_Click()
    If Not strFileName <> "" Then Exit Sub
    If Not txtTag.Text = "True" Then Exit Sub 'Some checking

    'Removes extra text based on limits
    txtTitle.Text = Left$(txtTitle.Text & Space(30), 30)
    txtArtist.Text = Left$(txtArtist.Text & Space(30), 30)
    txtAlbum.Text = Left$(txtAlbum.Text & Space(30), 30)
    txtYear.Text = Left$(txtYear.Text & Space(4), 4)
    txtComments.Text = Left$(txtComments.Text & Space(30), 30)
    
    'Set info from text boxes
    With MpgTag
        .Title = txtTitle.Text
        .Artist = txtArtist.Text
        .Album = txtAlbum.Text
        .Year = txtYear.Text
        .Comments = txtComments.Text
        
        If cboGenre.ListIndex = -1 Then
            .Genre = 80
        Else
            .Genre = cboGenre.ListIndex
        End If
    End With
    
    Write_MpgTag strFileName, MpgTag
End Sub

Private Sub cmdChoose_Click()
    Call Flush
    
    GetOpenName hwnd, "Open", strFileName
    
    'If nothing was returned
    If Not strFileName <> "" Then
        cmdAddTag.Enabled = False
        cmdApply.Enabled = False
        cmdRemoveTag.Enabled = False
        
        txtTag.Text = "False"
        
        Exit Sub 'Dont worry just exit
    End If
    
    Read_MpgInfo strFileName, MpgInfo
    
    With MpgInfo 'Dump mpginfo
        txtVersion.Text = .Version
        txtLayer.Text = .Layer
        chkErrorProtection.Value = CInt(.Error_Protection)
        txtBitrate.Text = .Bitrate_Index
        txtSamplingFreq.Text = .Sampling_Freq & "hz"
        txtMode.Text = .Mode
        chkCopyright.Value = .Copyright
        chkOriginal.Value = .Original
        txtEmphasis.Text = .Emphasis
    End With
    
    Read_MpgTag strFileName, MpgTag
    
    'Dump info back to text boxes
    If MpgTag.Tag = True Then
        txtTag.Text = "True"
        
        With MpgTag
            txtTitle.Text = .Title
            txtArtist.Text = .Artist
            txtAlbum.Text = .Album
            txtYear.Text = .Year
            txtComments.Text = .Comments
            cboGenre.Text = cboGenre.List(.Genre)
        End With
        
        'Set cmd buttons accordingly
        cmdAddTag.Enabled = False
        cmdRemoveTag.Enabled = True
        cmdApply.Enabled = True
    Else
        txtTag.Text = "False"
        
        'Set cmd buttons accordingly
        cmdAddTag.Enabled = True
        cmdRemoveTag.Enabled = False
    End If
End Sub

Private Sub cmdRemoveTag_Click()
    If Not strFileName <> "" Then Exit Sub
    
    Dim tmpString As String
    
    Open strFileName For Binary As #1 'Opens it for binary
        tmpString = Space(LOF(1)) 'Padd to length of file
        Get #1, , tmpString 'Dump contents of file to string
    Close #1
    
    Open strFileName For Output As #1
        'Send back out with out 128 bytes of tag
        Print #1, Left$(tmpString, Len(tmpString) - 128);
    Close #1
    
    'Clear tag area
    txtTitle.Text = ""
    txtArtist.Text = ""
    txtAlbum.Text = ""
    txtYear.Text = ""
    txtComments.Text = ""
    cboGenre.Text = ""
    
    txtTag.Text = "False" 'Set to false cause we deleted it
    cmdApply.Enabled = False
End Sub

Private Sub Form_Load()
    Dim tmpByte As Byte

    For tmpByte = 0 To 80 'Cycle through list
        cboGenre.AddItem MpgTagGenre(tmpByte) 'Adds needed items to combo box
    Next tmpByte
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Flush()
    'Clear variables and text boxes
    strFileName = ""

    txtVersion.Text = ""
    txtLayer.Text = ""
    chkErrorProtection.Value = 0
    txtBitrate.Text = ""
    txtSamplingFreq.Text = ""
    txtMode.Text = ""
    chkCopyright.Value = 0
    chkOriginal.Value = 0
    txtEmphasis.Text = ""
            
    txtTitle.Text = ""
    txtArtist.Text = ""
    txtAlbum.Text = ""
    txtYear.Text = ""
    txtComments.Text = ""
    cboGenre.Text = ""
End Sub
