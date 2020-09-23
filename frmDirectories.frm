VERSION 5.00
Begin VB.Form frmDirectories 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Directories"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmDirectories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDirectories 
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
   Begin VB.ListBox lstDirectories 
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
End
Attribute VB_Name = "frmDirectories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With lstDirectories
        .AddItem Left("AdminTools" & Space(15), 15) & Dirs.AdminTools
        .AddItem Left("AppData" & Space(15), 15) & Dirs.AppData
        .AddItem Left("AppPath" & Space(15), 15) & Dirs.AppPath
        .AddItem Left("Cache" & Space(15), 15) & Dirs.Cache
        .AddItem Left("CommonFiles" & Space(15), 15) & Dirs.CommonFiles
        .AddItem Left("Cookies" & Space(15), 15) & Dirs.Cookies
        .AddItem Left("Current" & Space(15), 15) & Dirs.Current
        .AddItem Left("Desktop" & Space(15), 15) & Dirs.Desktop
        .AddItem Left("Favorites" & Space(15), 15) & Dirs.Favorites
        .AddItem Left("Fonts" & Space(15), 15) & Dirs.Fonts
        .AddItem Left("History" & Space(15), 15) & Dirs.History
        .AddItem Left("LocalAppData" & Space(15), 15) & Dirs.LocalAppData
        .AddItem Left("MediaPath" & Space(15), 15) & Dirs.MediaPath
        .AddItem Left("MyPictures" & Space(15), 15) & Dirs.MyPictures
        .AddItem Left("NetHood" & Space(15), 15) & Dirs.NetHood
        .AddItem Left("Personal" & Space(15), 15) & Dirs.Personal
        .AddItem Left("PrintHood" & Space(15), 15) & Dirs.PrintHood
        .AddItem Left("Programs" & Space(15), 15) & Dirs.Programs
        .AddItem Left("Recent" & Space(15), 15) & Dirs.Recent
        .AddItem Left("SendTo" & Space(15), 15) & Dirs.sendto
        .AddItem Left("StartMenu" & Space(15), 15) & Dirs.StartMenu
        .AddItem Left("Startup" & Space(15), 15) & Dirs.Startup
        .AddItem Left("System" & Space(15), 15) & Dirs.System
        .AddItem Left("Temp" & Space(15), 15) & Dirs.Temp
        .AddItem Left("Templates" & Space(15), 15) & Dirs.Templates
        .AddItem Left("Windows" & Space(15), 15) & Dirs.Windows
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstDirectories_Click()
    txtDirectories.Text = Right(lstDirectories.List(lstDirectories.ListIndex), Len(lstDirectories.List(lstDirectories.ListIndex)) - 15)
End Sub
