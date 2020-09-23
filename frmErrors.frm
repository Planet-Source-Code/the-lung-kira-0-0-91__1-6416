VERSION 5.00
Begin VB.Form frmErrors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Errors"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmErrors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescription 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2760
      Width           =   6975
   End
   Begin VB.ListBox lstErrors 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim errDescription As String
    Dim errConst As String
    
    For apiError = 0 To 11031 'Cycle through all errors
        Errors.Errors apiError, "", errDescription, errConst, True
        
        'Must have data in errDescription or errConst
        If errConst <> "" Then
            lstErrors.AddItem Left(apiError & Space(8), 8) & errConst
        End If
        
        'Clear
        'errDescription = ""
        errConst = ""
    Next apiError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lstErrors.Clear 'Clear
    Unload Me
End Sub

Private Sub lstErrors_Click()
    Dim errDescription As String
    Dim errConst As String
    
    apiError = CLng(Left(lstErrors.List(lstErrors.ListIndex), 5))
    
    Errors.Errors apiError, "", errDescription, errConst, True
    txtDescription.Text = errDescription
End Sub
