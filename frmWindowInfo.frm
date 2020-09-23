VERSION 5.00
Begin VB.Form frmWindowInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Window Info"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmWindowInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstWindowInfo 
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
      Width           =   5055
   End
End
Attribute VB_Name = "frmWindowInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With lstWindowInfo
        .AddItem Left("3D Border Width" & Space(30), 30) & GetSystemMetrics(SM_CXEDGE)
        .AddItem Left("3D Border Height" & Space(30), 30) & GetSystemMetrics(SM_CYEDGE)
        .AddItem Left("Border Width" & Space(30), 30) & GetSystemMetrics(SM_CXBORDER)
        .AddItem Left("Border Height" & Space(30), 30) & GetSystemMetrics(SM_CYBORDER)
        .AddItem Left("Default Maximum Width" & Space(30), 30) & GetSystemMetrics(SM_CXMAXTRACK)
        .AddItem Left("Default Maximum Height" & Space(30), 30) & GetSystemMetrics(SM_CYMAXTRACK)
        .AddItem Left("Default Maximized Width" & Space(30), 30) & GetSystemMetrics(SM_CXMAXIMIZED)
        .AddItem Left("Default Maximized Height" & Space(30), 30) & GetSystemMetrics(SM_CYMAXIMIZED)
        .AddItem Left("Dialog Border Width" & Space(30), 30) & GetSystemMetrics(SM_CXFIXEDFRAME)
        .AddItem Left("Dialog Border Height" & Space(30), 30) & GetSystemMetrics(SM_CYFIXEDFRAME)
        .AddItem Left("Full Screen Width" & Space(30), 30) & GetSystemMetrics(SM_CXFULLSCREEN)
        .AddItem Left("Full Screen Height" & Space(30), 30) & GetSystemMetrics(SM_CYFULLSCREEN)
        
        .AddItem ""
        .AddItem "Minimized Arranging - Starting Position"
        If GetSystemMetrics(SM_ARRANGE) And ARW_BOTTOMLEFT Then .AddItem "Bottom Left"
        If GetSystemMetrics(SM_ARRANGE) And ARW_BOTTOMRIGHT Then .AddItem "Bottom Right"
        If GetSystemMetrics(SM_ARRANGE) And ARW_HIDE Then .AddItem "Hide"
        If GetSystemMetrics(SM_ARRANGE) And ARW_TOPLEFT Then .AddItem "Top Left"
        If GetSystemMetrics(SM_ARRANGE) And ARW_TOPRIGHT Then .AddItem "Top Right"
        .AddItem "Minimized Arranging - Direction"
        If GetSystemMetrics(SM_ARRANGE) And ARW_DOWN Then .AddItem "Down"
        If GetSystemMetrics(SM_ARRANGE) And ARW_LEFT Then .AddItem "Left"
        If GetSystemMetrics(SM_ARRANGE) And ARW_RIGHT Then .AddItem "Right"
        If GetSystemMetrics(SM_ARRANGE) And ARW_UP Then .AddItem "Up"
        .AddItem ""
        
        .AddItem Left("Minimized GridSpace Width" & Space(30), 30) & GetSystemMetrics(SM_CXMINSPACING)
        .AddItem Left("Minimized GridSpace Height" & Space(30), 30) & GetSystemMetrics(SM_CYMINSPACING)
        .AddItem Left("Minimum Width" & Space(30), 30) & GetSystemMetrics(SM_CXMIN)
        .AddItem Left("Minimum Height" & Space(30), 30) & GetSystemMetrics(SM_CXMIN)
        .AddItem Left("Minimum Tracking Width" & Space(30), 30) & GetSystemMetrics(SM_CXMINTRACK)
        .AddItem Left("Minimum Tracking Height" & Space(30), 30) & GetSystemMetrics(SM_CYMINTRACK)
        .AddItem Left("Normal Minimized Width" & Space(30), 30) & GetSystemMetrics(SM_CXMINIMIZED)
        .AddItem Left("Normal Minimized Height" & Space(30), 30) & GetSystemMetrics(SM_CYMINIMIZED)
        .AddItem Left("Sizing Border Width" & Space(30), 30) & GetSystemMetrics(SM_CXSIZEFRAME)
        .AddItem Left("Sizing Border Height" & Space(30), 30) & GetSystemMetrics(SM_CYSIZEFRAME)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
