VERSION 5.00
Begin VB.Form frmMainAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmMainAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainAbout.frx":000C
   ScaleHeight     =   4830
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblPage 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "free.prohosting.com/~thelung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblPageLabel 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Web Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblContactLabel 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblContact 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "the_lung_@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblStarted 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "August 1, 1999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblMadeBy 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Professor Lung"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblStartedLabel 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Started on"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblMadeByLabel 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Made by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmMainAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
