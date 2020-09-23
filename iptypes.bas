Attribute VB_Name = "iptypes"
Option Explicit

    
    Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 128 'arb
    Public Const MAX_ADAPTER_NAME_LENGTH = 256        'arb
    Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8       'arb
    Public Const DEFAULT_MINIMUM_ENTITIES = 32        'arb
    Public Const MAX_HOSTNAME_LEN = 128               'arb
    Public Const MAX_DOMAIN_NAME_LEN = 128            'arb
    Public Const MAX_SCOPE_ID_LEN = 256               'arb
    

    Public IP_ADDRESS_STRING As IP_ADDRESS_STRING
    Public Type IP_ADDRESS_STRING
        String As String * 16 '4 x 4
    End Type

    Public IP_ADDR_STRING As IP_ADDR_STRING
    Public Type IP_ADDR_STRING
        IpAddress As IP_ADDRESS_STRING
        IpMask As IP_ADDRESS_STRING
        Context As Long
    End Type
    
    Public FIXED_INFO As FIXED_INFO
    Public Type FIXED_INFO
        HostName As String * 132            'MAX_HOSTNAME_LEN + 4
        DomainName As String * 132          'MAX_DOMAIN_NAME_LEN + 4
        CurrentDnsServer As IP_ADDR_STRING
        DnsServerList As IP_ADDR_STRING
        NodeType As Integer
        ScopeId As String * 260             'MAX_SCOPE_ID_LEN + 4
        EnableRouting As Integer
        EnableProxy As Integer
        EnableDns As Integer
    End Type
