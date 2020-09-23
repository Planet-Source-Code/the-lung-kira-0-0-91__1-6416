Attribute VB_Name = "iphlpapi"
Option Explicit


Public Declare Function GetIpStatistics Lib "iphlpapi.dll" (pStats As MIB_IPSTATS) As Long
'It was MIB_ICMPSTATS instead of a extra structure to be passed I changed it to MIBICMPINFO
Public Declare Function GetIcmpStatistics Lib "iphlpapi.dll" (pStats As MIBICMPINFO) As Long
Public Declare Function GetNetworkParams Lib "iphlpapi.dll" (pFixedInfo As FIXED_INFO, pOutBufLen As Long) As Long
Public Declare Function GetTcpStatistics Lib "iphlpapi.dll" (pStats As MIB_TCPSTATS) As Long
Public Declare Function GetUdpStatistics Lib "iphlpapi.dll" (pStats As MIB_UDPSTATS) As Long
'Public Declare Function SetIpStatistics Lib "iphlpapi.dll" (pIpStats As MIB_IPSTATS) As Long
Public Declare Function SetIpTTL Lib "iphlpapi.dll" (ByVal nTTL As Integer) As Long
