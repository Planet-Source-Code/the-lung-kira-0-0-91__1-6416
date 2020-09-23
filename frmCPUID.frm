VERSION 5.00
Begin VB.Form frmCPUID 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CPUID"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmCPUID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkIA64 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   33
      Top             =   120
      Width           =   255
   End
   Begin VB.CheckBox chkSelfSnoop 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   53
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkSSE2 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   57
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkACPI 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkDTES 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   25
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chkCLFLSH 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtStepping 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtModel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtFamily 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CheckBox chkXMM 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   63
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox chkVME 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   61
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkTimeStampCounter 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   59
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chkPGE 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   45
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkMTRR 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   41
      Top             =   1080
      Width           =   255
   End
   Begin VB.CheckBox chkPageSizeExtensions 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   43
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox chkPhysicalAddressExtensions 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   47
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox chkPN 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   49
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkPSE36 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   51
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox chkSEP 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   55
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkMMX 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   39
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox chkMCA 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   37
      Top             =   600
      Width           =   255
   End
   Begin VB.CheckBox chkDebuggingExtensions 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   23
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkCMOV 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox chkCMPXCHG8B 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   21
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox chkFGPAT 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   27
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkFpuPresent 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   29
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox chkFXSR 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   31
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox chkMachineCheckException 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7080
      TabIndex        =   35
      Top             =   360
      Width           =   255
   End
   Begin VB.CheckBox chkAPICOnChip 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox chkCPUIDAvail 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblIA64 
      Caption         =   "IA64"
      Height          =   255
      Left            =   3840
      TabIndex        =   32
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblSelfSnoop 
      Caption         =   "Self Snoop"
      Height          =   255
      Left            =   3840
      TabIndex        =   52
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label lblSSE2 
      Caption         =   "SSE2"
      Height          =   255
      Left            =   3840
      TabIndex        =   56
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblACPI 
      Caption         =   "ACPI"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblDTES 
      Caption         =   "DTES"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lblCLFLSH 
      Caption         =   "CLFLSH"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblXMM 
      Caption         =   "XMM"
      Height          =   255
      Left            =   3840
      TabIndex        =   62
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label lblVME 
      Caption         =   "VME"
      Height          =   255
      Left            =   3840
      TabIndex        =   60
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label lblTimeStampCounter 
      Caption         =   "Time Stamp Counter"
      Height          =   255
      Left            =   3840
      TabIndex        =   58
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lblType 
      Caption         =   "Type"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblStepping 
      Caption         =   "Stepping"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblFamily 
      Caption         =   "Family"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblModel 
      Caption         =   "Model"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblPGE 
      Caption         =   "PGE"
      Height          =   255
      Left            =   3840
      TabIndex        =   44
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label lblMTRR 
      Caption         =   "MTRR"
      Height          =   255
      Left            =   3840
      TabIndex        =   40
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label lblPageSizeExtensions 
      Caption         =   "Page Size Extensions"
      Height          =   255
      Left            =   3840
      TabIndex        =   42
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label lblPhysicalAddressExtensions 
      Caption         =   "Physical Address Extensions"
      Height          =   255
      Left            =   3840
      TabIndex        =   46
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblPN 
      Caption         =   "PN"
      Height          =   255
      Left            =   3840
      TabIndex        =   48
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblPSE36 
      Caption         =   "PSE36"
      Height          =   255
      Left            =   3840
      TabIndex        =   50
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblSEP 
      Caption         =   "SEP"
      Height          =   255
      Left            =   3840
      TabIndex        =   54
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lblMMX 
      Caption         =   "MMX"
      Height          =   255
      Left            =   3840
      TabIndex        =   38
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label lblMCA 
      Caption         =   "MCA"
      Height          =   255
      Left            =   3840
      TabIndex        =   36
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label lblDebuggingExtensions 
      Caption         =   "Debugging Extensions"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblCMOV 
      Caption         =   "CMOV"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label lblCMPXCHG8B 
      Caption         =   "CMP XCHG 8B"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lblFGPAT 
      Caption         =   "FGPAT"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label lblFpuPresent 
      Caption         =   "Fpu Present"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label lblFXSR 
      Caption         =   "FXSR"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label lblMachineCheckException 
      Caption         =   "Machine Check Exception"
      Height          =   255
      Left            =   3840
      TabIndex        =   34
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblAPICOnChip 
      Caption         =   "APIC On Chip"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblCPUIDAvail 
      Caption         =   "CPUID Available"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   3015
   End
End
Attribute VB_Name = "frmCPUID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    txtDescription.Text = Get_CpuDesc(CStr(cpuid_Type) & CStr(cpuid_Family) & CStr(cpuid_Model) & CStr(cpuid_Stepping))
    txtFamily.Text = cpuid_Family
    txtModel.Text = cpuid_Model
    txtStepping.Text = cpuid_Stepping
    txtType.Text = cpuid_Type
    
    'Select Case cpuid_Type
    '    Case 0: txtType.Text = "Primary"
    '    Case 1: txtType.Text = "OverDrive"
    '    Case 2: txtType.Text = "Secondary"
    '    'Case 3: txtType.Text = "Reserved"
    'End Select
    
    chkAPICOnChip.Value = cpuid_APICOnChip
    chkCMOV.Value = cpuid_CMOV
    chkCMPXCHG8B.Value = cpuid_CMPXCHG8B
    chkCPUIDAvail.Value = cpuid_avail
    chkDebuggingExtensions.Value = cpuid_DebuggingExtensions
    chkFGPAT.Value = cpuid_FGPAT
    chkFpuPresent.Value = cpuid_FpuPresent
    chkFXSR.Value = cpuid_FXSR
    chkMachineCheckException.Value = cpuid_MachineCheckException
    chkMCA.Value = cpuid_MCA
    chkMMX.Value = cpuid_MMX
    chkMTRR.Value = cpuid_MTRR
    chkPageSizeExtensions.Value = cpuid_PageSizeExtensions
    chkPGE.Value = cpuid_PGE
    chkPhysicalAddressExtensions.Value = cpuid_PhysicalAddressExtensions
    chkPN.Value = cpuid_PN
    chkPSE36.Value = cpuid_PSE36
    chkSEP.Value = cpuid_SEP
    chkTimeStampCounter.Value = cpuid_TimeStampCounter
    chkVME.Value = cpuid_VME
    chkXMM.Value = cpuid_XMM
    
    'Reserved values
    Dim tmpString As String
    tmpString = Left$(CStr(cpuid_reserved3) & "0000", 4)
    
    chkCLFLSH.Value = CInt(Mid$(tmpString, 1, 1))
    chkDTES.Value = CInt(Mid$(tmpString, 3, 1))
    chkACPI.Value = CInt(Mid$(tmpString, 4, 1))
    
    tmpString = Left$(CStr(cpuid_reserved4) & "000000", 6)
    
    chkSSE2.Value = CInt(Mid$(tmpString, 1, 1))
    chkSelfSnoop.Value = CInt(Mid$(tmpString, 2, 1))
    chkIA64.Value = CInt(Mid$(tmpString, 5, 1))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
