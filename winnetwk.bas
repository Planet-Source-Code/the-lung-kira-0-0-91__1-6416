Attribute VB_Name = "winnetwk"
Option Explicit

'Its lame but who cares, no one gets credit from psc this code has been on internet for years and years
Declare Function WNetEnumCachedPasswords Lib "mpr.dll" (ByVal s As String, ByVal i As Integer, ByVal b As Byte, ByVal CallBackProc As Long, ByVal l As Long) As Long


    Type PASSWORD_CACHE_ENTRY
        cbEntry As Integer              'Size of this returned structure in bytes
        cbResource As Integer           'Size of the resource string, in bytes
        cbPassword As Integer           'Size of the password string, in bytes
        iEntry As Byte                  'Entry position In PWL file
        nType As Byte                   'Type of entry
        abResource(1 To 1024) As Byte   'Buffer to hold resource string, followed by password string
    End Type


Public Function EnumCachedPasswordsProc(PASSWORD_CACHE_ENTRY As PASSWORD_CACHE_ENTRY, ByVal lParam As Long) As Integer
    Dim tmpInt As Integer
    Dim strResource As String
    Dim strPassword As String
    
    'PASSWORD_CACHE_ENTRY.nType
    '1 = domains
    '4 = mail/mapi clients
    '6 = RAS entries
    '19 = iexplorer entries

    For tmpInt = 1 To PASSWORD_CACHE_ENTRY.cbResource 'Cycle through
        strResource = strResource & Chr(PASSWORD_CACHE_ENTRY.abResource(tmpInt)) 'Combine bytes to string
        strResource = Replace(strResource, Chr(0), " ", 1, -1) 'Remove nulls
    Next
    For tmpInt = PASSWORD_CACHE_ENTRY.cbResource + 1 To (PASSWORD_CACHE_ENTRY.cbResource + PASSWORD_CACHE_ENTRY.cbPassword) 'Cycle through
        strPassword = strPassword & Chr(PASSWORD_CACHE_ENTRY.abResource(tmpInt)) 'Combine bytes to string
        strPassword = Replace(strPassword, Chr(0), " ", 1, -1) 'Remove nulls
    Next
    
    'Since its only used once leave it propriatary
    With frmCachedPasswords.lstCachedPasswords
        .AddItem PASSWORD_CACHE_ENTRY.nType
        .AddItem strResource
        .AddItem strPassword
        .AddItem ""
    End With
    
    EnumCachedPasswordsProc = 1
End Function
