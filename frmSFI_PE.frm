VERSION 5.00
Begin VB.Form frmSFI_PE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "String File Info Editor Win32/PE"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmSFI_PE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCompanyName 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox txtComments 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txtProductVersion 
      Height          =   285
      Left            =   1800
      TabIndex        =   23
      Top             =   4080
      Width           =   3735
   End
   Begin VB.TextBox txtProductName 
      Height          =   285
      Left            =   1800
      TabIndex        =   21
      Top             =   3720
      Width           =   3735
   End
   Begin VB.TextBox txtOriginalFilename 
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   3360
      Width           =   3735
   End
   Begin VB.TextBox txtLegalTrademarks 
      Height          =   285
      Left            =   1800
      TabIndex        =   17
      Top             =   3000
      Width           =   3735
   End
   Begin VB.TextBox txtLegalCopyright 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox txtInternalName 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox txtFileVersion 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtFileDescription 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      Height          =   350
      Left            =   4560
      TabIndex        =   25
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtFound 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox txtCharSet 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3600
      TabIndex        =   24
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lblComments 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblProductVersion 
      Caption         =   "Product Version"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lblProductName 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblOriginalFilename 
      Caption         =   "Original Filename"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblLegalTrademarks 
      Caption         =   "Legal Trademarks"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblLegalCopyright 
      Caption         =   "Legal Copyright"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblInternalName 
      Caption         =   "Internal Name"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblFileVersion 
      Caption         =   "File Version"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblFileDescription 
      Caption         =   "File Description"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblCompanyName 
      Caption         =   "Company Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblCharSet 
      Caption         =   "Character Set"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblFound 
      Caption         =   "SFI Found"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmSFI_PE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFile As String 'temp string for contents of file
Dim strFileName As String 'Name of file
Dim pos As Long 'Needed for replace on apply

'limitations, original, (start) position
Dim limComments As Integer, orgComments As String, posComments As Long
Dim limCompanyName As Integer, orgCompanyName As String, posCompanyName As Long
Dim limFileDescription As Integer, orgFileDescription As String, posFileDescription As Long
Dim limFileVersion As Integer, orgFileVersion As String, posFileVersion As Long
Dim limInternalName As Integer, orgInternalName As String, posInternalName As Long
Dim limLegalCopyright As Integer, orgLegalCopyright As String, posLegalCopyright As Long
Dim limLegalTrademarks As Integer, orgLegalTrademarks As String, posLegalTrademarks As Long
Dim limOriginalFilename As Integer, orgOriginalFilename As String, posOriginalFilename As Long
Dim limProductName As Integer, orgProductName As String, posProductName As Long
Dim limProductVersion As Integer, orgProductVersion As String, posProductVersion As Long
    
Private Sub cmdApply_Click()
    'If the string in text box is to long then it cuts it off based on size of the limitation
    If Len(txtComments.Text) > limComments Then txtComments.Text = Left$(txtComments.Text, limComments)
    If Len(txtCompanyName.Text) > limCompanyName Then txtCompanyName.Text = Left$(txtCompanyName.Text, limCompanyName)
    If Len(txtFileDescription.Text) > limFileDescription Then txtFileDescription.Text = Left$(txtFileDescription.Text, limFileDescription)
    If Len(txtFileVersion.Text) > limFileVersion Then txtFileVersion.Text = Left$(txtFileVersion.Text, limFileVersion)
    If Len(txtInternalName.Text) > limInternalName Then txtInternalName.Text = Left$(txtInternalName.Text, limInternalName)
    If Len(txtLegalCopyright.Text) > limLegalCopyright Then txtLegalCopyright.Text = Left$(txtLegalCopyright.Text, limLegalCopyright)
    If Len(txtLegalTrademarks.Text) > limLegalTrademarks Then txtLegalTrademarks.Text = Left$(txtLegalTrademarks.Text, limLegalTrademarks)
    If Len(txtOriginalFilename.Text) > limOriginalFilename Then txtOriginalFilename.Text = Left$(txtOriginalFilename.Text, limOriginalFilename)
    If Len(txtProductName.Text) > limProductName Then txtProductName.Text = Left$(txtProductName.Text, limProductName)
    If Len(txtProductVersion.Text) > limProductVersion Then txtProductVersion.Text = Left$(txtProductVersion.Text, limProductVersion)
    
    'If string in text box isnt the same size then it pads the rest of the string with spaces based on size by the difference between limitations and textbox len
    If Len(txtComments.Text) < limComments Then txtComments.Text = txtComments.Text & String(limComments - Len(txtComments.Text), " ")
    If Len(txtCompanyName.Text) < limCompanyName Then txtCompanyName.Text = txtCompanyName.Text & String(limCompanyName - Len(txtCompanyName.Text), " ")
    If Len(txtFileDescription.Text) < limFileDescription Then txtFileDescription.Text = txtFileDescription.Text & String(limFileDescription - Len(txtFileDescription.Text), " ")
    If Len(txtFileVersion.Text) < limFileVersion Then txtFileVersion.Text = txtFileVersion.Text & String(limFileVersion - Len(txtFileVersion.Text), " ")
    If Len(txtInternalName.Text) < limInternalName Then txtInternalName.Text = txtInternalName.Text & String(limInternalName - Len(txtInternalName.Text), " ")
    If Len(txtLegalCopyright.Text) < limLegalCopyright Then txtLegalCopyright.Text = txtLegalCopyright.Text & String(limLegalCopyright - Len(txtLegalCopyright.Text), " ")
    If Len(txtLegalTrademarks.Text) < limLegalTrademarks Then txtLegalTrademarks.Text = txtLegalTrademarks.Text & String(limLegalTrademarks - Len(txtLegalTrademarks.Text), " ")
    If Len(txtOriginalFilename.Text) < limOriginalFilename Then txtOriginalFilename.Text = txtOriginalFilename.Text & String(limOriginalFilename - Len(txtOriginalFilename.Text), " ")
    If Len(txtProductName.Text) < limProductName Then txtProductName.Text = txtProductName.Text & String(limProductName - Len(txtProductName.Text), " ")
    If Len(txtProductVersion.Text) < limProductVersion Then txtProductVersion.Text = txtProductVersion.Text & String(limProductVersion - Len(txtProductVersion.Text), " ")

    If Len(txtComments.Text) > 0 Then  'Cant replace nothing with nothing
        strFile = Left$(strFile, posComments - 1) & _
        Unicode_Padd(txtComments.Text) & _
        Right$(strFile, Len(strFile) - ((posComments - 1) + Len(Unicode_Padd(txtComments.Text))))
    End If
    
    If Len(txtCompanyName.Text) > 0 Then
        strFile = Left$(strFile, posCompanyName - 1) & _
        Unicode_Padd(txtCompanyName.Text) & _
        Right$(strFile, Len(strFile) - ((posCompanyName - 1) + Len(Unicode_Padd(txtCompanyName.Text))))
    End If

    If Len(txtFileDescription.Text) > 0 Then
        strFile = Left$(strFile, posFileDescription - 1) & _
        Unicode_Padd(txtFileDescription.Text) & _
        Right$(strFile, Len(strFile) - ((posFileDescription - 1) + Len(Unicode_Padd(txtFileDescription.Text))))
    End If

    If Len(txtFileVersion.Text) > 0 Then
        strFile = Left$(strFile, posFileVersion - 1) & _
        Unicode_Padd(txtFileVersion.Text) & _
        Right$(strFile, Len(strFile) - ((posFileVersion - 1) + Len(Unicode_Padd(txtFileVersion.Text))))
    End If
    
    If Len(txtInternalName.Text) > 0 Then
        strFile = Left$(strFile, posInternalName - 1) & _
        Unicode_Padd(txtInternalName.Text) & _
        Right$(strFile, Len(strFile) - ((posInternalName - 1) + Len(Unicode_Padd(txtInternalName.Text))))
    End If
    
    If Len(txtLegalCopyright.Text) > 0 Then
        strFile = Left$(strFile, posLegalCopyright - 1) & _
        Unicode_Padd(txtLegalCopyright.Text) & _
        Right$(strFile, Len(strFile) - ((posLegalCopyright - 1) + Len(Unicode_Padd(txtLegalCopyright.Text))))
    End If

    If Len(txtLegalTrademarks.Text) > 0 Then
        strFile = Left$(strFile, posLegalTrademarks - 1) & _
        Unicode_Padd(txtLegalTrademarks.Text) & _
        Right$(strFile, Len(strFile) - ((posLegalTrademarks - 1) + Len(Unicode_Padd(txtLegalTrademarks.Text))))
    End If
    
    If Len(txtOriginalFilename.Text) > 0 Then
        strFile = Left$(strFile, posOriginalFilename - 1) & _
        Unicode_Padd(txtOriginalFilename.Text) & _
        Right$(strFile, Len(strFile) - ((posOriginalFilename - 1) + Len(Unicode_Padd(txtOriginalFilename.Text))))
    End If
    
    If Len(txtProductName.Text) > 0 Then
        strFile = Left$(strFile, posProductName - 1) & _
        Unicode_Padd(txtProductName.Text) & _
        Right$(strFile, Len(strFile) - ((posProductName - 1) + Len(Unicode_Padd(txtProductName.Text))))
    End If
    
    If Len(txtProductVersion.Text) > 0 Then
        strFile = Left$(strFile, posProductVersion - 1) & _
        Unicode_Padd(txtProductVersion.Text) & _
        Right$(strFile, Len(strFile) - ((posProductVersion - 1) + Len(Unicode_Padd(txtProductVersion.Text))))
    End If

    GetSaveName hwnd, "Save", strFileName
    
    'Writes the string to a file
    If strFileName <> "" Then 'If strFileName is not empty
        Open strFileName For Output As #1
            Print #1, strFile;
        Close #1
    End If
End Sub

Private Sub cmdChoose_Click()
    Call Flush
    
    GetOpenName hwnd, "Open", strFileName
    
    'Error checking
    If Not strFileName <> "" Then
        cmdApply.Enabled = False
        Exit Sub 'Dont worry just exit
    End If
    If Not FileLen(strFileName) > 0 Then 'If file len not greater than 0
        cmdApply.Enabled = False
        MsgBox "File size is 0.", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Declare variables after checking is done
    Dim StartPos As Long, EndPos As Long
    
    Open strFileName For Binary As #1 'Opens it for binary
        strFile = Space(LOF(1)) 'Pads to length of string
        Get #1, , strFile 'Dumps contents of file to string
    Close #1
    
    pos = InStr(1, strFile, Unicode_Padd(Chr(1) & "StringFileInfo"))
    StartPos = pos
    
    If StartPos = 0 Then 'If nothing was found then
        txtFound.Text = "Not Found"
        
        Screen.MousePointer = vbNormal 'Resets cursor so they can continue
        cmdChoose.Enabled = True 'Enables so they can rehit
        
        Exit Sub
    End If

    'Dumps info to text boxes
    txtFound.Text = "Found"
    txtCharSet.Text = Replace(Mid$(strFile, StartPos + 38, 15), Chr(0), "", 1, -1) 'Gets string out
    
    'Looks for string, if found then puts it out in text box
    StartPos = InStr(pos, strFile, Unicode_Padd(Chr(1) & "Comments"))
    If StartPos > 0 Then 'If its even entered in
        posComments = StartPos + 20 'Manually figured out offsets
        EndPos = InStr(posComments, strFile, String(3, Chr(0))) 'Allwasy 3 nulls at end
        orgComments = Mid$(strFile, posComments, EndPos - (posComments))  'Pulls the original
        txtComments.Text = Replace(orgComments, Chr(0), "", 1, -1) 'Dumps non null filed string to text box
    End If
    
    StartPos = InStr(pos, strFile, Unicode_Padd(Chr(1) & "CompanyName"))
    If StartPos > 0 Then
        posCompanyName = StartPos + 28
        EndPos = InStr(posCompanyName, strFile, String(3, Chr(0)))
        orgCompanyName = Mid$(strFile, posCompanyName, EndPos - (posCompanyName))
        txtCompanyName.Text = Replace(orgCompanyName, Chr(0), "", 1, -1)
    End If
    
    StartPos = InStr(pos, strFile, Unicode_Padd(Chr(1) & "FileDescription"))
    If StartPos > 0 Then
        posFileDescription = StartPos + 36
        EndPos = InStr(posFileDescription, strFile, String(3, Chr(0)))
        orgFileDescription = Mid$(strFile, posFileDescription, EndPos - (posFileDescription))
        txtFileDescription.Text = Replace(orgFileDescription, Chr(0), "", 1, -1)
    End If
    
    StartPos = InStr(pos, strFile, Unicode_Padd(Chr(1) & "FileVersion"))
    If StartPos > 0 Then
        posFileVersion = StartPos + 28
        EndPos = InStr(posFileVersion, strFile, String(3, Chr(0)))
        orgFileVersion = Mid$(strFile, posFileVersion, EndPos - (posFileVersion))
        txtFileVersion.Text = Replace(orgFileVersion, Chr(0), "", 1, -1)
    End If
    
    StartPos = InStr(pos, strFile, Unicode_Padd(Chr(1) & "InternalName"))
    If StartPos > 0 Then
        posInternalName = StartPos + 28
        EndPos = InStr(posInternalName, strFile, String(3, Chr(0)))
        orgInternalName = Mid$(strFile, posInternalName, EndPos - (posInternalName))
        txtInternalName.Text = Replace(orgInternalName, Chr(0), "", 1, -1)
    End If
    
    StartPos = InStr(pos, strFile, Unicode_Padd(Chr(1) & "LegalCopyright"))
    If StartPos > 0 Then
        posLegalCopyright = StartPos + 32
        EndPos = InStr(posLegalCopyright, strFile, String(3, Chr(0)))
        orgLegalCopyright = Mid$(strFile, posLegalCopyright, EndPos - (posLegalCopyright))
        txtLegalCopyright.Text = Replace(orgLegalCopyright, Chr(0), "", 1, -1)
    End If
    
    StartPos = InStr(pos, strFile, Unicode_Padd(Chr(1) & "LegalTrademarks"))
    If StartPos > 0 Then
        posLegalTrademarks = StartPos + 36
        EndPos = InStr(posLegalTrademarks, strFile, String(3, Chr(0)))
        orgLegalTrademarks = Mid$(strFile, posLegalTrademarks, EndPos - (posLegalTrademarks))
        txtLegalTrademarks.Text = Replace(orgLegalTrademarks, Chr(0), "", 1, -1)
    End If
    
    StartPos = InStr(pos, strFile, Unicode_Padd(Chr(1) & "OriginalFilename"))
    If StartPos > 0 Then
        posOriginalFilename = StartPos + 36
        EndPos = InStr(posOriginalFilename, strFile, String(3, Chr(0)))
        orgOriginalFilename = Mid$(strFile, posOriginalFilename, EndPos - (posOriginalFilename))
        txtOriginalFilename.Text = Replace(orgOriginalFilename, Chr(0), "", 1, -1)
    End If
    
    StartPos = InStr(pos, strFile, Unicode_Padd(Chr(1) & "ProductName"))
    If StartPos > 0 Then
        posProductName = StartPos + 28
        EndPos = InStr(posProductName, strFile, String(3, Chr(0)))
        orgProductName = Mid$(strFile, posProductName, EndPos - (posProductName))
        txtProductName.Text = Replace(orgProductName, Chr(0), "", 1, -1)
    End If
    
    StartPos = InStr(pos, strFile, Unicode_Padd(Chr(1) & "ProductVersion"))
    If StartPos > 0 Then
        posProductVersion = StartPos + 32
        EndPos = InStr(posProductVersion, strFile, String(3, Chr(0)))
        orgProductVersion = Mid$(strFile, posProductVersion, EndPos - (posProductVersion))
        txtProductVersion.Text = Replace(orgProductVersion, Chr(0), "", 1, -1)
    End If
    
    'Do some error checking, based on chr(1) being leading character in identifier
    If InStr(1, txtComments.Text, Chr(1)) > 0 Then txtComments.Text = ""
    If InStr(1, txtCompanyName.Text, Chr(1)) > 0 Then txtCompanyName.Text = ""
    If InStr(1, txtFileDescription.Text, Chr(1)) > 0 Then txtFileDescription.Text = ""
    If InStr(1, txtFileVersion.Text, Chr(1)) > 0 Then txtFileVersion.Text = ""
    If InStr(1, txtInternalName.Text, Chr(1)) > 0 Then txtInternalName.Text = ""
    If InStr(1, txtLegalCopyright.Text, Chr(1)) > 0 Then txtLegalCopyright.Text = ""
    If InStr(1, txtLegalTrademarks.Text, Chr(1)) > 0 Then txtLegalTrademarks.Text = ""
    If InStr(1, txtOriginalFilename.Text, Chr(1)) > 0 Then txtOriginalFilename.Text = ""
    If InStr(1, txtProductName.Text, Chr(1)) > 0 Then txtProductName.Text = ""
    If InStr(1, txtProductVersion.Text, Chr(1)) > 0 Then txtProductVersion.Text = ""
    
    'Give limitations based on size
    limComments = Len(txtComments.Text)
    limCompanyName = Len(txtCompanyName.Text)
    limFileDescription = Len(txtFileDescription.Text)
    limFileVersion = Len(txtFileVersion.Text)
    limInternalName = Len(txtInternalName.Text)
    limLegalCopyright = Len(txtLegalCopyright.Text)
    limLegalTrademarks = Len(txtLegalTrademarks.Text)
    limOriginalFilename = Len(txtOriginalFilename.Text)
    limProductName = Len(txtProductName.Text)
    limProductVersion = Len(txtProductVersion.Text)

    cmdApply.Enabled = True 'Allows applying after a file has been choosen and processed
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Flush()
    'Reset filename
    strFileName = ""
    
    'Resets text boxes
    txtFound.Text = ""
    txtCharSet.Text = ""
    txtComments.Text = ""
    txtCompanyName.Text = ""
    txtFileDescription.Text = ""
    txtFileVersion.Text = ""
    txtInternalName.Text = ""
    txtLegalCopyright.Text = ""
    txtLegalTrademarks.Text = ""
    txtOriginalFilename.Text = ""
    txtProductName.Text = ""
    txtProductVersion.Text = ""
End Sub

