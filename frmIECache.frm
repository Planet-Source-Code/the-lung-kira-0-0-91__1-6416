VERSION 5.00
Begin VB.Form frmIECache 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IE Cache Hit/Miss"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "frmIECache.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHitPerc 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtHit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtMiss 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtMissPerc 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtTotalPerc 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   340
      Left            =   1800
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblHits 
      Caption         =   "Hits"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblMiss 
      Caption         =   "Misses"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmIECache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
    'Clear text boxes
    txtMiss.Text = ""
    txtHit.Text = ""
    txtTotal.Text = ""
    
    If Not Dirs.Cache & "\Content.IE5\index.dat" <> "" Then 'No file
        MsgBox "Missing " & vbCrLf & Dirs.Cache & "\Content.IE5\index.dat", vbExclamation, "Error"
        Exit Sub
    End If
    
    Dim tmpString As String
    Dim Position As Long
    Dim CacheMiss As Long, CacheHit As Long

    Open Dirs.Cache & "\Content.IE5\index.dat" For Binary As #1  'Opens it for binary
        tmpString = Space(LOF(1)) 'Pads to length of string
        Get #1, , tmpString 'Dumps contents of file to string
    Close #1

    Do
        Position = InStr(Position + 1, tmpString, "X-Cache: ") 'Searches for hit miss string
        
        If Mid$(tmpString, Position + 9, 4) = "MISS" Then
            CacheMiss = CacheMiss + 1 'Pulls out miss and increments
            txtMiss.Text = CacheMiss
        End If
        If Mid$(tmpString, Position + 9, 4) = "HIT " Then
            CacheHit = CacheHit + 1 'Pulls out hit and increments
            txtHit.Text = CacheHit
        End If
        
        txtTotal.Text = (CacheHit + CacheMiss) 'Gives total to text box
        Interaction.DoEvents 'Yeilds to os
    Loop While Position > 0 'Do until it can't find string anymore
    
    'Gets percentage out
    If CacheMiss > 0 Then 'Prevents overflow
        txtMissPerc.Text = Round((CacheMiss / (CacheMiss + CacheHit) * 100), 0) & "%"
    Else
        txtMissPerc.Text = "0%"
    End If
    If CacheHit > 0 Then 'Prevents overflow
        txtHitPerc.Text = Round((CacheHit / (CacheMiss + CacheHit) * 100), 0) & "%"
    Else
        txtHitPerc.Text = "0%"
    End If
    
    txtTotalPerc.Text = "100%"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

