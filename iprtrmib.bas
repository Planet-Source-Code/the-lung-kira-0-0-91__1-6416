Attribute VB_Name = "iprtrmib"
Option Explicit

    
    Public Const MIB_TCP_RTO_OTHER = 1
    Public Const MIB_TCP_RTO_CONSTANT = 2
    Public Const MIB_TCP_RTO_RSRE = 3
    Public Const MIB_TCP_RTO_VANJ = 4
    

    Public MIBICMPSTATS As MIBICMPSTATS
    Public Type MIBICMPSTATS        'The MIBICMPSTATS structure contains statistics for either incoming or outgoing Internet Control Message Protocol (ICMP) messages on a particular computer.
        dwMsgs As Long              'number of messages
        dwErrors As Long            'number of errors
        dwDestUnreachs As Long      'destination unreachable messages
        dwTimeExcds As Long         'time-to-live exceeded messages
        dwParmProbs As Long         'parameter problem messages
        dwSrcQuenchs As Long        'source quench messages
        dwRedirects As Long         'redirection messages
        dwEchos As Long             'echo requests
        dwEchoReps As Long          'echo replies
        dwTimestamps As Long        'time-stamp requests
        dwTimestampReps As Long     'time-stamp replies
        dwAddrMasks As Long         'address mask requests
        dwAddrMaskReps As Long      'address mask replies
    End Type

    Public MIBICMPINFO As MIBICMPINFO
    Public Type MIBICMPINFO             'The MIBICMPINFO structure contains Internet Control Message Protocol (ICMP) statistics for a particular computer.
        icmpInStats As MIBICMPSTATS     'stats for incoming messages
        icmpOutStats As MIBICMPSTATS    'stats for outgoing messages
    End Type

    Public MIB_ICMP As MIB_ICMP
    Public Type MIB_ICMP        'The MIB_ICMP structure contains the Internet Control Message Protocol (ICMP) statistics for a particular computer.
        stats As MIBICMPINFO    'contains ICMP stats
    End Type
    
    Public MIB_IPSTATS As MIB_IPSTATS
    Public Type MIB_IPSTATS         'The MIB_IPSTATS structure stores information about the IP protocol running on a particular computer.
        dwForwarding As Long        'IP forwarding enabled or disabled
        dwDefaultTTL As Long        'default time-to-live
        dwInReceives As Long        'datagrams received
        dwInHdrErrors As Long       'received header errors
        dwInAddrErrors As Long      'received address errors
        dwForwDatagrams As Long     'datagrams forwarded
        dwInUnknownProtos As Long   'datagrams with unknown protocol
        dwInDiscards As Long        'received datagrams discarded
        dwInDelivers As Long        'received datagrams delivered
        dwOutRequests As Long
        dwRoutingDiscards As Long
        dwOutDiscards As Long       'sent datagrams discarded
        dwOutNoRoutes As Long       'datagrams for which no route
        dwReasmTimeout As Long      'datagrams for which all frags didn't arrive
        dwReasmReqds As Long        'datagrams requiring reassembly
        dwReasmOks As Long          'successful reassemblies
        dwReasmFails As Long        'failed reassemblies
        dwFragOks As Long           'successful fragmentations
        dwFragFails As Long         'failed fragmentations
        dwFragCreates As Long       'datagrams fragmented
        dwNumIf As Long             'number of interfaces on computer
        dwNumAddr As Long           'number of IP address on computer
        dwNumRoutes As Long         'number of routes in routing table
    End Type
    
    Public MIB_TCPSTATS As MIB_TCPSTATS
    Public Type MIB_TCPSTATS    'The MIB_TCPSTATS structure contains statistics for the TCP protocol running on the local computer.
        dwRtoAlgorithm As Long  'time-out algorithm
        dwRtoMin As Long        'minimum time-out
        dwRtoMax As Long        'maximum time-out
        dwMaxConn As Long       'maximum connections
        dwActiveOpens As Long   'active opens
        dwPassiveOpens As Long  'passive opens
        dwAttemptFails As Long  'failed attempts
        dwEstabResets As Long   'established connections reset
        dwCurrEstab As Long     'established connections
        dwInSegs As Long        'segments received
        dwOutSegs As Long       'segment sent
        dwRetransSegs As Long   'segments retransmitted
        dwInErrs As Long        'incoming errors
        dwOutRsts As Long       'outgoing resets
        dwNumConns As Long      'cumulative connections
    End Type
    
    Public MIB_UDPSTATS As MIB_UDPSTATS
    Public Type MIB_UDPSTATS    'The MIB_UDPSTATS structure contains statistics for the User Datagram Protocol (UDP) running on the local computer.
        dwInDatagrams As Long   'received datagrams
        dwNoPorts As Long       'datagrams for which no port
        dwInErrors As Long      'errors on received datagrams
        dwOutDatagrams As Long  'sent datagrams
        dwNumAddrs As Long      'number of entries in UDP listener table
    End Type
