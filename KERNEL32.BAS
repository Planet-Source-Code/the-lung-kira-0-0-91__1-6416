Attribute VB_Name = "kernel32"
Option Explicit

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OsVersionInfo) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function IsProcessorFeaturePresent Lib "kernel32" (ByVal ProcessorFeature As Long) As Boolean
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Public Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Public Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function Get_ComputerName() As String
    Dim tmpString As String
    
    tmpString = Space(1024) 'Padd
    
    apiError = GetComputerName(tmpString, 1024)
    If apiError = 0 Then
        Failed "GetComputerName"
    Else
        Get_ComputerName = Fix_NullTermStr(tmpString)
    End If
End Function

Public Function Get_CurrentDirectory() As String
    Dim tmpString As String
    
    tmpString = Space(1024) 'Padd
    apiError = GetCurrentDirectory(1024, tmpString)
    
    If apiError = 0 Then
        Failed "GetCurrentDirectory"
    Else
        Get_CurrentDirectory = Fix_Dir(Fix_NullTermStr(tmpString))
    End If
End Function

Public Function Get_SystemDirectory() As String
    Dim tmpString As String
    
    tmpString = Space(1024)
    apiError = GetSystemDirectory(tmpString, 1024)
    
    If apiError = 0 Then
        Failed "GetSystemDirectory"
    Else
        Get_SystemDirectory = Fix_Dir(Fix_NullTermStr(tmpString))
    End If
End Function

Public Function Get_TempPath() As String
    Dim tmpString As String
    
    tmpString = Space(1024) 'Padd
    apiError = GetTempPath(1024, tmpString)
    
    If apiError = 0 Then
        Failed "GetTempPath"
    Else
        Get_TempPath = Fix_Dir(Fix_NullTermStr(tmpString))
    End If
End Function

Public Function Get_WindowsDirectory() As String
    Dim tmpString As String
    
    tmpString = Space(1024) 'Padd
    apiError = GetWindowsDirectory(tmpString, 1024)
    
    If apiError = 0 Then
        Failed "GetWindowsDirectory"
    Else
        Get_WindowsDirectory = Fix_Dir(Fix_NullTermStr(tmpString))
    End If
End Function

Public Function Set_ComputerName(strName)
    If Len(strName) < 1 Then Exit Function 'Must be at least 1 in length
    
    'Guessed at maximum - if to large trim the string
    If Len(strName) > 1024 Then strName = Left(strName, 1024)
    
    apiError = SetComputerName(strName)
    If apiError = 0 Then Failed "SetComputerName"
End Function
