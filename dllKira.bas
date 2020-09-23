Attribute VB_Name = "dllKira"
Option Explicit

'CPUID
'Dont blame me, blame interaction between lcc & vb & my c skills
Public Declare Sub cpu_id Lib "cpu.dll" Alias "_cpu_id" ()
Public Declare Function cpuid_APICOnChip Lib "cpu.dll" Alias "_cpuid_APICOnChip" () As Integer
Public Declare Function cpuid_CMOV Lib "cpu.dll" Alias "_cpuid_CMOV" () As Integer
Public Declare Function cpuid_CMPXCHG8B Lib "cpu.dll" Alias "_cpuid_CMPXCHG8B" () As Integer
Public Declare Function cpuid_DebuggingExtensions Lib "cpu.dll" Alias "_cpuid_DebuggingExtensions" () As Integer
Public Declare Function cpuid_Family Lib "cpu.dll" Alias "_cpuid_Family" () As Integer
Public Declare Function cpuid_FGPAT Lib "cpu.dll" Alias "_cpuid_FGPAT" () As Integer
Public Declare Function cpuid_FpuPresent Lib "cpu.dll" Alias "_cpuid_FpuPresent" () As Integer
Public Declare Function cpuid_FXSR Lib "cpu.dll" Alias "_cpuid_FXSR" () As Integer
Public Declare Function cpuid_MachineCheckException Lib "cpu.dll" Alias "_cpuid_MachineCheckException" () As Integer
Public Declare Function cpuid_MCA Lib "cpu.dll" Alias "_cpuid_MCA" () As Integer
Public Declare Function cpuid_MMX Lib "cpu.dll" Alias "_cpuid_MMX" () As Integer
Public Declare Function cpuid_Model Lib "cpu.dll" Alias "_cpuid_Model" () As Integer
Public Declare Function cpuid_MTRR Lib "cpu.dll" Alias "_cpuid_MTRR" () As Integer
Public Declare Function cpuid_PageSizeExtensions Lib "cpu.dll" Alias "_cpuid_PageSizeExtensions" () As Integer
Public Declare Function cpuid_PGE Lib "cpu.dll" Alias "_cpuid_PGE" () As Integer
Public Declare Function cpuid_PhysicalAddressExtensions Lib "cpu.dll" Alias "_cpuid_PhysicalAddressExtensions" () As Integer
Public Declare Function cpuid_PN Lib "cpu.dll" Alias "_cpuid_PN" () As Integer
Public Declare Function cpuid_PSE36 Lib "cpu.dll" Alias "_cpuid_PSE36" () As Integer
Public Declare Function cpuid_SEP Lib "cpu.dll" Alias "_cpuid_SEP" () As Integer
Public Declare Function cpuid_Stepping Lib "cpu.dll" Alias "_cpuid_Stepping" () As Integer
Public Declare Function cpuid_TimeStampCounter Lib "cpu.dll" Alias "_cpuid_TimeStampCounter" () As Integer
Public Declare Function cpuid_Type Lib "cpu.dll" Alias "_cpuid_Type" () As Integer
Public Declare Function cpuid_VME Lib "cpu.dll" Alias "_cpuid_VME" () As Integer
Public Declare Function cpuid_XMM Lib "cpu.dll" Alias "_cpuid_XMM" () As Integer

Public Declare Function cpuid_Reserved1 Lib "cpu.dll" Alias "_cpuid_Reserved1" () As Long
Public Declare Function cpuid_reserved2 Lib "cpu.dll" Alias "_cpuid_reserved2" () As Integer
Public Declare Function cpuid_reserved3 Lib "cpu.dll" Alias "_cpuid_reserved3" () As Integer
Public Declare Function cpuid_reserved4 Lib "cpu.dll" Alias "_cpuid_reserved4" () As Integer

Public Declare Function cycles_elapsed Lib "cpu.dll" Alias "_cycles_elapsed" () As Double
Public Declare Function cpuid_avail Lib "cpu.dll" Alias "_cpuid_avail" () As Boolean
Public Declare Sub HookCallbackInit Lib "kira.dll" (ByVal hwnd As Long, ByVal hHook As Long)


    Public mh_Instance As Long
    Public mh_ProcAddress As Long
    Public mh_Hook As Long


Public Sub InstallMouseHook()
    apiError = LoadLibrary(Dirs.System & "\kira.dll")
    If apiError = 0 Then
        Failed "LoadLibraryEx"
        Exit Sub 'Exit here
    End If
    mh_Instance = apiError 'Dump
    
    apiError = GetProcAddress(mh_Instance, "MouseHookProc")
    If apiError = 0 Then
        Failed "GetProcAddress"
    
        'Clean up
        apiError = FreeLibrary(mh_Instance)
        If apiError = 0 Then Failed "FreeLibrary"
        
        Exit Sub 'Exit here
    End If
    mh_ProcAddress = apiError 'Dump
    
    apiError = SetWindowsHookEx(WH_MOUSE, mh_ProcAddress, mh_Instance, 0)
    If apiError = 0 Then
        Failed "SetWindowsHookEx"
        Exit Sub 'Exit here
    End If
    mh_Hook = apiError 'Dump

    HookCallbackInit frmMain.picHM.hwnd, mh_Hook
End Sub
