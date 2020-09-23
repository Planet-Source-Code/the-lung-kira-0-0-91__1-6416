VERSION 5.00
Begin VB.Form frmIEOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IE Extra Options"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmIEOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtProdID 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Width           =   3855
   End
   Begin VB.TextBox txtIEVer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.TextBox txtDefaultSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox txtDefaultURL 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox txtSearchPage 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox txtStartPage 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Top             =   2400
      Width           =   3855
   End
   Begin VB.CheckBox chkDelTempFile 
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chkEnDiskCache 
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   5040
      TabIndex        =   16
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblIEVer 
      Caption         =   "IE Version"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblProdID 
      Caption         =   "Product ID"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblDefaultSearch 
      Caption         =   "Default Search"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblDefaultURL 
      Caption         =   "Default URL"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblSearchPage 
      Caption         =   "Search Page"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblStartPage 
      Caption         =   "Start Page"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblDelTempFile 
      Caption         =   "Delete Temp Files On Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label lblEnDiskCache 
      Caption         =   "Enable Disk Cache"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2805
   End
End
Attribute VB_Name = "frmIEOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    'Save settings from options to reg
    If chkDelTempFile.Enabled = True Then
        If chkDelTempFile.Value = 0 Then 'If not choosen
            SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Delete_Temp_Files_On_Exit", "no"
        Else
            SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Delete_Temp_Files_On_Exit", "yes"
        End If
    End If
    If chkEnDiskCache.Enabled = True Then
        If chkEnDiskCache.Value = 0 Then 'If not choosen
            SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Enable_Disk_Cache", "no"
        Else
            SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Enable_Disk_Cache", "yes"
        End If
    End If
    
    'Save settings from text files to reg
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Page_URL", txtDefaultURL.Text
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Search_URL", txtDefaultSearch.Text
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Search Page", txtSearchPage.Text
    SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Start Page", txtStartPage.Text
End Sub

Private Sub Form_Load()
    Dim tmpString As String
    
    tmpString = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Delete_Temp_Files_On_Exit")
    Select Case tmpString
        Case "no": chkDelTempFile.Value = 0
        Case "yes": chkDelTempFile.Value = 1
        Case "": chkDelTempFile.Enabled = False 'If nothing thier then disable
    End Select
    
    tmpString = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Enable_Disk_Cache")
    Select Case tmpString
        Case "no": chkEnDiskCache.Value = 0
        Case "yes": chkEnDiskCache.Value = 1
        Case "": chkEnDiskCache.Enabled = False 'If nothing thier then disable
    End Select
    
    'Pulls settings from registry to text boxes
    txtDefaultURL.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Page_URL")
    txtDefaultSearch.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Default_Search_URL")
    txtIEVer.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer", "Version")
    'Thinking about removing this - as some version of IE do not place this info in the reg or in the same place
    txtProdID.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Registration", "ProductID")
    txtSearchPage.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Search Page")
    txtStartPage.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main", "Start Page")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
