Attribute VB_Name = "winnt"
Option Explicit


    Public Type LARGE_INTEGER
        LowPart As Long
        HighPart As Long
    End Type
    
    Public LUID As LUID
    Public Type LUID
        LowPart As Long
        HighPart As Long
    End Type
    
    Public LUID_AND_ATTRIBUTES As LUID_AND_ATTRIBUTES
    Public Type LUID_AND_ATTRIBUTES
        pLuid As LUID
        Attributes As Long
    End Type
    
    Public OsVersionInfo As OsVersionInfo
    Public Type OsVersionInfo
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
    End Type
    
    
    Public Const ANYSIZE_ARRAY = 1
    
    Public Const FILE_CASE_SENSITIVE_SEARCH = &H1
    Public Const FILE_CASE_PRESERVED_NAMES = &H2
    Public Const FILE_UNICODE_ON_DISK = &H4
    Public Const FILE_PERSISTENT_ACLS = &H8
    Public Const FILE_FILE_COMPRESSION = &H10
    Public Const FILE_VOLUME_QUOTAS = &H20
    Public Const FILE_SUPPORTS_SPARSE_FILES = &H40
    Public Const FILE_SUPPORTS_REPARSE_POINTS = &H80
    Public Const FILE_SUPPORTS_REMOTE_STORAGE = &H100
    Public Const FILE_VOLUME_IS_COMPRESSED = &H8000
    Public Const FILE_SUPPORTS_OBJECT_IDS = &H10000
    Public Const FILE_SUPPORTS_ENCRYPTION = &H20000
    Public Const FILE_NAMED_STREAMS = &H40000

    Public Const FILE_ATTRIBUTE_READONLY = &H1
    Public Const FILE_ATTRIBUTE_HIDDEN = &H2
    Public Const FILE_ATTRIBUTE_SYSTEM = &H4
    Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
    Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
    Public Const FILE_ATTRIBUTE_DEVICE = &H40
    Public Const FILE_ATTRIBUTE_NORMAL = &H80
    Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
    Public Const FILE_ATTRIBUTE_SPARSE_FILE = &H200
    Public Const FILE_ATTRIBUTE_REPARSE_POINT = &H400
    Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
    Public Const FILE_ATTRIBUTE_OFFLINE = &H1000
    Public Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
    Public Const FILE_ATTRIBUTE_ENCRYPTED = &H4000

    Public Const FILE_SHARE_READ = &H1
    Public Const FILE_SHARE_WRITE = &H2
    Public Const FILE_SHARE_DELETE = &H4

    Public Const GENERIC_READ = &H80000000
    Public Const GENERIC_WRITE = &H40000000
    Public Const GENERIC_EXECUTE = &H20000000
    Public Const GENERIC_ALL = &H10000000

    Public Const DELETE = &H10000
    Public Const READ_CONTROL = &H20000
    Public Const WRITE_DAC = &H40000
    Public Const WRITE_OWNER = &H80000
    Public Const SYNCHRONIZE = &H100000
    
    'These must be ahead of the KEY_'s to prevent forward referencing
    Public Const STANDARD_RIGHTS_ALL = &H1F0000
    Public Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
    Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
    Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
    Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

    Public Const KEY_CREATE_LINK = &H20
    Public Const KEY_CREATE_SUB_KEY = &H4
    Public Const KEY_ENUMERATE_SUB_KEYS = &H8
    Public Const KEY_EVENT = &H1     '  Event contains key event record
    Public Const KEY_NOTIFY = &H10
    Public Const KEY_QUERY_VALUE = &H1
    Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    Public Const KEY_SET_VALUE = &H2
    Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
    Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

    Public Const PROCESSOR_ARCHITECTURE_INTEL = 0
    Public Const PROCESSOR_ARCHITECTURE_MIPS = 1
    Public Const PROCESSOR_ARCHITECTURE_ALPHA = 2
    Public Const PROCESSOR_ARCHITECTURE_PPC = 3
    Public Const PROCESSOR_ARCHITECTURE_SHX = 4
    Public Const PROCESSOR_ARCHITECTURE_ARM = 5
    Public Const PROCESSOR_ARCHITECTURE_IA64 = 6
    Public Const PROCESSOR_ARCHITECTURE_ALPHA64 = 7
    Public Const PROCESSOR_ARCHITECTURE_MSIL = 8
    Public Const PROCESSOR_ARCHITECTURE_UNKNOWN = &HFFFF

    Public Const PROCESSOR_INTEL_386 = 386
    Public Const PROCESSOR_INTEL_486 = 486
    Public Const PROCESSOR_INTEL_PENTIUM = 586
    Public Const PROCESSOR_INTEL_IA64 = 2200
    Public Const PROCESSOR_MIPS_R4000 = 4000      'incl R4101 & R3910 for Windows CE
    Public Const PROCESSOR_ALPHA_21064 = 21064
    Public Const PROCESSOR_PPC_601 = 601
    Public Const PROCESSOR_PPC_603 = 603
    Public Const PROCESSOR_PPC_604 = 604
    Public Const PROCESSOR_PPC_620 = 620
    Public Const PROCESSOR_HITACHI_SH3 = 10003    'Windows CE
    Public Const PROCESSOR_HITACHI_SH3E = 10004   'Windows CE
    Public Const PROCESSOR_HITACHI_SH4 = 10005    'Windows CE
    Public Const PROCESSOR_MOTOROLA_821 = 821     'Windows CE
    Public Const PROCESSOR_SHx_SH3 = 103          'Windows CE
    Public Const PROCESSOR_SHx_SH4 = 104          'Windows CE
    Public Const PROCESSOR_STRONGARM = 2577       'Windows CE - &HA11
    Public Const PROCESSOR_ARM720 = 1824          'Windows CE - &H720
    Public Const PROCESSOR_ARM820 = 2080          'Windows CE - &H820
    Public Const PROCESSOR_ARM920 = 2336          'Windows CE - &H920
    Public Const PROCESSOR_ARM_7TDMI = 70001      'Windows CE
    Public Const PROCESSOR_OPTIL = &H494F         'MSIL

    Public Const REG_NONE = 0                       'No value type
    Public Const REG_SZ = 1                         'Unicode nul terminated string
    Public Const REG_EXPAND_SZ = 2                  'Unicode nul terminated string
    Public Const REG_BINARY = 3                     'Free form binary
    Public Const REG_DWORD = 4                      '32-bit number
    Public Const REG_DWORD_LITTLE_ENDIAN = 4        '32-bit number (same as REG_DWORD)
    Public Const REG_DWORD_BIG_ENDIAN = 5           '32-bit number
    Public Const REG_LINK = 6                       'Symbolic Link (unicode)
    Public Const REG_MULTI_SZ = 7                   'Multiple Unicode strings
    Public Const REG_RESOURCE_LIST = 8              'Resource list in the resource map
    Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9   'Resource list in the hardware description
    Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10
    Public Const REG_QWORD = 11                     '64-bit number
    Public Const REG_QWORD_LITTLE_ENDIAN = 11       '64-bit number (same as REG_QWORD)

    Public Const REG_OPTION_RESERVED = 0           'Parameter is reserved
    Public Const REG_OPTION_NON_VOLATILE = 0       'Key is preserved when system is rebooted
    Public Const REG_OPTION_VOLATILE = 1           'Key is not preserved when system is rebooted
    Public Const REG_OPTION_CREATE_LINK = 2        'Created key is a symbolic link
    Public Const REG_OPTION_BACKUP_RESTORE = 4     'open for backup or restore
    Public Const REG_OPTION_OPEN_LINK = &H8        'Open symbolic link
    
    Public Const VER_MINORVERSION = &H1
    Public Const VER_MAJORVERSION = &H2
    Public Const VER_BUILDNUMBER = &H4
    Public Const VER_PLATFORMID = &H8
    Public Const VER_SERVICEPACKMINOR = &H10
    Public Const VER_SERVICEPACKMAJOR = &H20
    Public Const VER_SUITENAME = &H40
    Public Const VER_PRODUCT_TYPE = &H80

    Public Const VER_NT_WORKSTATION = &H1
    Public Const VER_NT_DOMAIN_CONTROLLER = &H2
    Public Const VER_NT_SERVER = &H3
    
    Public Const VER_PLATFORM_WIN32s = 0
    Public Const VER_PLATFORM_WIN32_WINDOWS = 1
    Public Const VER_PLATFORM_WIN32_NT = 2
