Attribute VB_Name = "icmp"
Option Explicit


Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Public Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptions As IP_OPTION_INFORMATION, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
'Public Declare Function IcmpParseReplies Lib "icmp.dll" (ByVal ReplyBuffer As String, ByVal ReplySize As Long) As Long


    Public IP_OPTION_INFORMATION As IP_OPTION_INFORMATION
    Public Type IP_OPTION_INFORMATION
        TTL As Byte
        Tos As Byte
        flags As Byte
        OptionsSize As Byte
        OptionsData As Long
    End Type

    Public ICMP_ECHO_REPLY As ICMP_ECHO_REPLY
    Public Type ICMP_ECHO_REPLY
        Address As Long
        Status As Long
        RoundTripTime As Long
        DataSize As Integer
        Reserved As Integer
        DataPointer As Long
        Options As IP_OPTION_INFORMATION
        Data As String * 128
    End Type
