Attribute VB_Name = "baseNetwork"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "baseNetwork"

'=====================================================
'
'  baseNetwork by Ali Mousavi Kherad
'  alimousavikherad@gmail.com
'
'  WARNING ...
'    It's very important that you remove integer
'    Overflow option in your project properties
'    IP port provided in integer in ws2_32 dll
'    And i turn it to long but a don't write
'    Any code to cast it if you don't remove
'    This option integer overflow error will be
'    Occured when trying to pass ports to ws2_32.dll
'
' Thank you...
'=====================================================

'Constants and structures defined by the internet system,
'Per RFC 790, September 1981, taken from the BSD file netinet/in.h.

'Protocols
Public Const IPPROTO_UNKNOWN = -1               'Unknown protocol.
Public Const IPPROTO_IPV6HOPBYHOPOPTIONS = 0    'IPv6 Hop by Hop Options header.
Public Const IPPROTO_Unspecified = 0            'Unspecified protocol.
Public Const IPPROTO_IP = 0                     'dummy for IP | Internet Protocol.
Public Const IPPROTO_ICMP = 1                   'control message protocol | Internet Control Message Protocol.
Public Const IPPROTO_IGMP = 2                   'group management protocol | Internet Group Management Protocol.
Public Const IPPROTO_GGP = 3                    'gateway^2 (deprecated) | Gateway To Gateway Protocol.
Public Const IPPROTO_IPV4 = 4                   'Internet Protocol version 4.
Public Const IPPROTO_TCP = 6                    'Transmission Control Protocol.
Public Const IPPROTO_PUP = 12                   'PARC Universal Packet Protocol.
Public Const IPPROTO_UDP = 17                   'User Datagram Protocol.
Public Const IPPROTO_IDP = 22                   'xns idp | Internet Datagram Protocol.
Public Const IPPROTO_IPV6 = 41                  'Internet Protocol version 6 (IPv6).
Public Const IPPROTO_IPV6ROUTINGHEADER = 43     'IPv6 Routing header.
Public Const IPPROTO_IPV6FRAGMENTHEADER = 44    'IPv6 Fragment header.
Public Const IPPROTO_IPSECENCAPSULATINGSECURITYPAYLOAD = 50 'IPv6 Encapsulating Security Payload header.
Public Const IPPROTO_IPSECAUTHENTICATIONHEADER = 51 'IPv6 Authentication header. For details, see RFC 2292 section 2.2.1, available at http://www.ietf.org.
Public Const IPPROTO_ICMPV6 = 58                'Internet Control Message Protocol for IPv6.
Public Const IPPROTO_IPV6NONEXTHEADER = 59      'IPv6 No next header.
Public Const IPPROTO_IPV6DESTINATIONOPTIONS = 60 'IPv6 Destination Options header.
Public Const IPPROTO_ND = 77                    'UNOFFICIAL net disk proto | Net Disk Protocol (unofficial).
Public Const IPPROTO_RAW = 255                  'Raw IP packet protocol.
Public Const IPPROTO_IPX = 1000                 'Internet Packet Exchange Protocol.

Public Const IPPROTO_SPX = 1256                 'Sequenced Packet Exchange protocol.
Public Const IPPROTO_SPXII = 1257               'Sequenced Packet Exchange version 2 protocol.
Public Const IPPROTO_MAX = 256


'Port/socket numbers: network standard functions.
Public Const IPPORT_ECHO = 7
Public Const IPPORT_DISCARD = 9
Public Const IPPORT_SYSTAT = 11
Public Const IPPORT_DAYTIME = 13
Public Const IPPORT_NETSTAT = 15
Public Const IPPORT_FTP = 21
Public Const IPPORT_TELNET = 23
Public Const IPPORT_SMTP = 25
Public Const IPPORT_TIMESERVER = 37
Public Const IPPORT_NAMESERVER = 42
Public Const IPPORT_WHOIS = 43
Public Const IPPORT_MTP = 57

'Port/socket numbers: host specific functions.
Public Const IPPORT_TFTP = 69
Public Const IPPORT_RJE = 77
Public Const IPPORT_FINGER = 79
Public Const IPPORT_TTYLINK = 87
Public Const IPPORT_SUPDUP = 95
'UNIX TCP sockets.
Public Const IPPORT_EXECSERVER = 512
Public Const IPPORT_LOGINSERVER = 513
Public Const IPPORT_CMDSERVER = 514
Public Const IPPORT_EFSSERVER = 520
'UNIX UDP sockets.
Public Const IPPORT_BIFFUDP = 512
Public Const IPPORT_WHOSERVER = 513
Public Const IPPORT_ROUTESERVER = 520
' 520+1 also used.
'Ports < IPPORT_RESERVED are reserved for
' privileged processes (e.g. root).
Public Const IPPORT_RESERVED = 1024

'Options for use with [gs]etsockopt at the IP level.
Public Const IP_OPTIONS = 1                   'set/get IP per-packet options
Public Const IP_MULTICAST_IF = 2              'set/get IP multicast interface
Public Const IP_MULTICAST_TTL = 3             'set/get IP multicast timetolive
Public Const IP_MULTICAST_LOOP = 4            'set/get IP multicast loopback
Public Const IP_ADD_MEMBERSHIP = 5            'add  an IP group membership
Public Const IP_DROP_MEMBERSHIP = 6           'drop an IP group membership
Public Const IP_TTL = 7                       'set/get IP Time To Live
Public Const IP_TOS = 8                       'set/get IP Type Of Service
Public Const IP_DONTFRAGMENT = 9              'set/get IP Don't Fragment flag
'
Public Const IP_DEFAULT_MULTICAST_TTL = 1     'normally limit m'casts to 1 hop
Public Const IP_DEFAULT_MULTICAST_LOOP = 1    'normally hear sends if a member
Public Const IP_MAX_MEMBERSHIPS = 20          'per socket; must fit in one mbuf



'Link numbers
Public Const IMPLINK_IP = 155
Public Const IMPLINK_LOWEXPER = 156
Public Const IMPLINK_HIGHEXPER = 158

'This is used instead of -1, since the
' SOCKET type is unsigned.
Public Const INVALID_SOCKET = -1
Public Const SOCKET_ERROR = -1

'Types
Public Const SOCK_UNKNOWN = -1           'Specifies an unknown Socket type.
Public Const SOCK_STREAM = 1             'stream socket
'SOCK_STREAM:
'   Supports reliable, two-way, connection-based byte streams without the duplication
'   of data and without preservation of boundaries. A Socket of this type communicates
'   with a single peer and requires a remote host connection before communication
'   can begin. System.Net.Sockets.SocketType.Stream uses the Transmission Control
'   Protocol (System.Net.Sockets.ProtocolType.Tcp) System.Net.Sockets.ProtocolType
'   and the InterNetworkSystem.Net.Sockets.AddressFamily.
Public Const SOCK_DGRAM = 2              'datagram socket
'SOCK_DGRAM:
'   Supports datagrams, which are connectionless, unreliable messages of a fixed
'   (typically small) maximum length. Messages might be lost or duplicated and
'   might arrive out of order. A System.Net.Sockets.Socket of type System.Net.Sockets.SocketType.Dgram
'   requires no connection prior to sending and receiving data, and can communicate
'   with multiple peers. System.Net.Sockets.SocketType.Dgram uses the Datagram
'   Protocol (System.Net.Sockets.ProtocolType.Udp) and the System.Net.Sockets.AddressFamily.InterNetworkSystem.Net.Sockets.AddressFamily.
Public Const SOCK_RAW = 3                'raw-protocol interface
'SOCK_RAW:
'   Supports access to the underlying transport protocol. Using the System.Net.Sockets.SocketTypeSystem.Net.Sockets.SocketType.Raw,
'   you can communicate using protocols like Internet Control Message Protocol
'   (System.Net.Sockets.ProtocolType.Icmp) and Internet Group Management Protocol
'   (System.Net.Sockets.ProtocolType.Igmp). Your application must provide a complete
'   IP header when sending. Received datagrams return with the IP header and
'   options intact.
Public Const SOCK_RDM = 4                'reliably-delivered message
'SOCK_RDM:
'   Supports connectionless, message-oriented, reliably delivered messages, and
'   preserves message boundaries in data. Rdm (Reliably Delivered Messages) messages
'   arrive unduplicated and in order. Furthermore, the sender is notified if
'   messages are lost. If you initialize a Socket using System.Net.Sockets.SocketType.Rdm,
'   you do not require a remote host connection before sending and receiving
'   data. With System.Net.Sockets.SocketType.Rdm, you can communicate with multiple
'   peers.
Public Const SOCK_SEQPACKET = 5          'sequenced packet stream
'SOCK_SEQPACKET:
'   Provides connection-oriented and reliable two-way transfer of ordered byte
'   streams across a network. System.Net.Sockets.SocketType.Seqpacket does not
'   duplicate data, and it preserves boundaries within the data stream. A Socket
'   of type System.Net.Sockets.SocketType.Seqpacket communicates with a single
'   peer and requires a remote host connection before communication can begin.

'Option flags per-socket.
Public Const SO_DEBUG = &H1              'turn on debugging info recording
Public Const SO_ACCEPTCONN = &H2         'socket has had listen()
Public Const SO_REUSEADDR = &H4          'allow local address reuse
Public Const SO_KEEPALIVE = &H8          'keep connections alive
Public Const SO_DONTROUTE = &H10         'just use interface addresses
Public Const SO_BROADCAST = &H20         'permit sending of broadcast msgs
Public Const SO_USELOOPBACK = &H40       'bypass hardware when possible
Public Const SO_LINGER = &H80            'linger on close if data present
Public Const SO_OOBINLINE = &H100        'leave received OOB data in line
'Public Const  SO_DONTLINGER   (u_int)(~SO_LINGER)

'Additional options.
Public Const SO_SNDBUF = &H1001          'send buffer size
Public Const SO_RCVBUF = &H1002          'receive buffer size
Public Const SO_SNDLOWAT = &H1003        'send low-water mark
Public Const SO_RCVLOWAT = &H1004        'receive low-water mark
Public Const SO_SNDTIMEO = &H1005        'send timeout
Public Const SO_RCVTIMEO = &H1006        'receive timeout
Public Const SO_ERROR = &H1007           'get error status and clear
Public Const SO_TYPE = &H1008            'get socket type

'Options for connect and disconnect data and options.  Used only by
' non-TCP/IP transports such as DECNet, OSI TP4, etc.
Public Const SO_CONNDATA = &H7000
Public Const SO_CONNOPT = &H7001
Public Const SO_DISCDATA = &H7002
Public Const SO_DISCOPT = &H7003
Public Const SO_CONNDATALEN = &H7004
Public Const SO_CONNOPTLEN = &H7005
Public Const SO_DISCDATALEN = &H7006
Public Const SO_DISCOPTLEN = &H7007

'Option for opening sockets for synchronous access.
Public Const SO_OPENTYPE = &H7008
Public Const SO_SYNCHRONOUS_ALERT = &H10
Public Const SO_SYNCHRONOUS_NONALERT = &H20

'Other NT-specific options.
Public Const SO_MAXDG = &H7009
Public Const SO_MAXPATHDG = &H700A
Public Const SO_UPDATE_ACCEPT_CONTEXT = &H700B
Public Const SO_CONNECT_TIME = &H700C

'Socket Flags.
Public Const SFLAG_NONE = 0                     'Use no flags for this call.
Public Const SFLAG_OOUTOFBAND = 1                'Process out-of-band data.
Public Const SFLAG_PEEK = 2                     'Peek at the incoming message.
Public Const SFLAG_DONTROUTE = 4                'Send without using routing tables.
Public Const SFLAG_MAXIOVECTORLENGTH = 16       'Provides a standard value for the number of WSABUF structures that are used to send and receive data.
Public Const SFLAG_TRUNCATED = 256              'The message was too large to fit into the specified buffer and was truncated.
Public Const SFLAG_CONTROLDATATRUNCATED = 512   'Indicates that the control data did not fit into an internal 64-KB buffer and was truncated.
Public Const SFLAG_BROADCAST = 1024             'Indicates a broadcast packet.
Public Const SFLAG_MULTICAST = 2048             'Indicates a multicast packet.
Public Const SFLAG_PARTIAL = 32768              'Partial send or receive for message.

'TCP options.
Public Const TCP_NODELAY = &H1
Public Const TCP_BSDURGENT = &H7000

'Address families.
Public Const AF_UNKNOWN = -1                'Unknown address family.
Public Const AF_UNSPEC = 0                  'Unknown address family.
Public Const AF_UNIX = 1                    'local to host (pipes, portals) | Unix local to host address.
Public Const AF_INET = 2                    'internetwork: UDP, TCP, etc. | Address for IP version 4.
Public Const AF_INTERNETWORK = AF_INET
Public Const AF_IMPLINK = 3                 'arpanet imp addresses | ARPANET IMP address.
Public Const AF_PUP = 4                     'pup protocols: e.g. BSP | Address for PUP protocols.
Public Const AF_CHAOS = 5                   'mit CHAOS protocols | Address for MIT CHAOS protocols.
Public Const AF_IPX = 6                     'IPX and SPX | IPX or SPX address.
Public Const AF_NS = 6                      'XEROX NS protocols | Address for Xerox NS protocols.
Public Const AF_ISO = 7                     'ISO protocols | Address for ISO protocols.
Public Const AF_OSI = AF_ISO                'OSI is ISO | Address for ISO protocols.
Public Const AF_ECMA = 8                    'european computer manufacturers | European Computer Manufacturers Association (ECMA) address.
Public Const AF_DATAKIT = 9                 'datakit protocols | Address for Datakit protocols.
Public Const AF_CCITT = 10                  'CCITT protocols, X.25 etc | Addresses for CCITT protocols, such as X.25.
Public Const AF_SNA = 11                    'IBM SNA | IBM SNA address.
Public Const AF_DECnet = 12                 'DECnet | DECnet address.
Public Const AF_DLI = 13                    'Direct data link interface | Direct data-link interface address.
Public Const AF_DATALINK = AF_DLI
Public Const AF_LAT = 14                    'LAT | LAT address.
Public Const AF_HYLINK = 15                 'NSC Hyperchannel | NSC Hyperchannel address.
Public Const AF_HYPERCHANNEL = AF_HYLINK
Public Const AF_APPLETALK = 16              'AppleTalk | AppleTalk address.
Public Const AF_NETBIOS = 17                'NetBios-style addresses | NetBios address.
Public Const AF_VOICEVIEW = 18              'VoiceView | VoiceView address.
Public Const AF_FIREFOX = 19                'FireFox | FireFox address.
Public Const AF_UNKNOWN1 = 20               'Somebody is using this!
Public Const AF_BAN = 21                    'Banyan
Public Const AF_BANYAN = AF_BAN
Public Const AF_ATM = 22                    'Native ATM services address.
Public Const AF_INTERNETWORKV6 = 23         'Address for IP version 6.
Public Const AF_CLUSTER = 24                'Address for Microsoft cluster products.
Public Const AF_IEEE12844 = 25              'IEEE 1284.4 workgroup address.
Public Const AF_IRDA = 26                   'IrDA address.
Public Const AF_NETWORKDESIGNERS = 28       'Address for Network Designers OSI gateway-enabled protocols.
Public Const AF_MAX = 29                    'MAX address.

'Level number for (get/set)sockopt() to apply to socket itself.
Public Const SOL_SOCKET = &HFFFF 'options for socket level

'Maximum queue length specifiable by listen.
Public Const SOMAXCONN = 5

Public Const MSG_OOB = &H1                    'process out-of-band data
Public Const MSG_PEEK = &H2                   'peek at incoming message
Public Const MSG_DONTROUTE = &H4              'send without using routing tables

Public Const MSG_MAXIOVLEN = 16

Public Const MSG_PARTIAL = &H8000             'partial send or recv for message xport

'Define constant based on rfc883, used by gethostbyxxxx() calls.
Public Const MAXGETHOSTSTRUCT = 1024

'Define flags to be used with the WSAAsyncSelect() call.
Public Const FD_READ = &H1
Public Const FD_WRITE = &H2
Public Const FD_OOB = &H4
Public Const FD_ACCEPT = &H8
Public Const FD_CONNECT = &H10
Public Const FD_CLOSE = &H20


Public Const INADDR_ANY As Long = &H0
Public Const INADDR_LOOPBACK As Long = &H7F000001
Public Const INADDR_BROADCAST As Long = -1
Public Const INADDR_NONE As Long = -1

Public Const MAX_PORT As Long = 65535
Public Const MIN_PORT As Long = 1

'Private Declare Function API_WSAAddressToString Lib "ws2_32.dll" Alias "WSAAddressToStringA" (ByRef lpsaAddress As sockaddr, ByVal dwAddressLength As Long, ByRef lpProtocolInfo As WSAPROTOCOL_INFOA, ByVal lpszAddressString As String, ByRef lpdwAddressStringLength As Long) As Long
Private Declare Function API_WSACleanup Lib "ws2_32" Alias "WSACleanup" () As Long
'Private Declare Function API_WSAGetLastError Lib "ws2_32.dll" () As Long
'Private Declare Function API_WSAGetQOSByName Lib "ws2_32.dll" (ByVal s As Long, ByRef lpQOSName As WSABUF, ByRef lpQOS As QOS) As Long
'Private Declare Function API_WSAIoctl Lib "ws2_32.dll" (ByVal s As Long, ByVal dwIoControlCode As Long, lpvInBuffer As Any, ByVal cbInBuffer As Long, lpvOutBuffer As Any, ByVal cbOutBuffer As Long, ByRef lpcbBytesReturned As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByRef lpCompletionRoutine As WSAOVERLAPPED_COMPLETION_ROUTINE) As Long
'Private Declare Function API_WSAIsBlocking Lib "ws2_32.dll" () As Long
'Private Declare Function API_WSARecvFrom Lib "ws2_32.dll" (ByVal s As Long, ByRef lpBuffers As WSABUF, ByVal dwBufferCount As Long, ByRef lpNumberOfBytesRecvd As Long, ByRef lpFlags As Long, ByRef lpFrom As sockaddr, ByRef lpFromlen As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByRef lpCompletionRoutine As WSAOVERLAPPED_COMPLETION_ROUTINE) As Long
'Private Declare Function API_WSARecvEx Lib "mswsock.dll" (ByVal s As Long, ByVal buf As String, ByVal lLen As Long, ByRef flags As Long) As Long
'Private Declare Function API_WSARecvDisconnect Lib "ws2_32.dll" (ByVal s As Long, ByRef lpInboundDisconnectData As WSABUF) As Long
'Private Declare Function API_WSARemoveServiceClass Lib "ws2_32.dll" (ByVal lpServiceClassId As Long) As Long
'Private Declare Function API_WSAResetEvent Lib "ws2_32.dll" (ByRef hEvent As WSAEVENT) As Long
'Private Declare Function API_WSASend Lib "ws2_32.dll" (ByVal s As Long, ByRef lpBuffers As WSABUF, ByVal dwBufferCount As Long, ByRef lpNumberOfBytesSent As Long, ByVal dwFlags As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByRef lpCompletionRoutine As WSAOVERLAPPED_COMPLETION_ROUTINE) As Long
'Private Declare Function API_WSASendDisconnect Lib "ws2_32.dll" (ByVal s As Long, ByRef lpOutboundDisconnectData As WSABUF) As Long
'Private Declare Function API_WSASendTo Lib "ws2_32.dll" (ByVal s As Long, ByRef lpBuffers As WSABUF, ByVal dwBufferCount As Long, ByRef lpNumberOfBytesSent As Long, ByVal dwFlags As Long, ByRef lpTo As sockaddr, ByVal iTolen As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByRef lpCompletionRoutine As WSAOVERLAPPED_COMPLETION_ROUTINE) As Long
'Private Declare Function API_WSASetBlockingHook Lib "ws2_32.dll" (ByVal lpBlockFunc As Long) As Long
'Private Declare Function API_WSASetEvent Lib "ws2_32.dll" (ByRef hEvent As WSAEVENT) As Long
'Private Declare Function API_WSASetService Lib "ws2_32.dll" (ByRef lpqsRegInfo As WSAQUERYSETA, ByVal essoperation As Struct_MembersOf_WSAESETSERVICEOP, ByVal dwControlFlags As Long) As Long
'Private Declare Function API_WSASocket Lib "ws2_32.dll" (ByVal af As Long, ByVal lType As Long, ByVal protocol As Long, ByRef lpProtocolInfo As WSAPROTOCOL_INFOA, ByRef g As Group, ByVal dwFlags As Long) As Long
Public Declare Function API_WSAStartup Lib "ws2_32" Alias "WSAStartup" (ByVal wVersionRequired As Integer, ByRef lpWSAData As API_WSADATA) As Long
'Private Declare Function API_WSAStringToAddress Lib "ws2_32.dll" (ByVal AddressString As String, ByVal AddressFamily As Long, ByRef lpProtocolInfo As WSAPROTOCOL_INFOA, ByRef lpAddress As sockaddr, ByRef lpAddressLength As Long) As Long
'Private Declare Function API_WSAUnhookBlockingHook Lib "ws2_32.dll" () As Long
'Private Declare Function API_WSAWaitForMultipleEvents Lib "ws2_32.dll" (ByVal cEvents As Long, ByRef lphEvents As WSAEVENT, ByVal fWaitAll As Long, ByVal dwTimeout As Long, ByVal fAlertable As Long) As Long
'Private Declare Function API_WSARecv Lib "ws2_32.dll" (ByVal s As Long, ByRef lpBuffers As WSABUF, ByVal dwBufferCount As Long, ByRef lpNumberOfBytesRecvd As Long, ByRef lpFlags As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByRef lpCompletionRoutine As WSAOVERLAPPED_COMPLETION_ROUTINE) As Long
'Private Declare Function API_WSAProviderConfigChange Lib "ws2_32.dll" (ByRef lpNotificationHandle As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByRef lpCompletionRoutine As WSAOVERLAPPED_COMPLETION_ROUTINE) As Long
'Private Declare Function API_WSANtohs Lib "ws2_32.dll" (ByVal s As Long, ByVal netshort As Integer, ByRef lphostshort As Integer) As Long
'Private Declare Function API_WSANtohl Lib "ws2_32.dll" (ByVal s As Long, ByVal netlong As Long, ByRef lphostlong As Long) As Long
'Private Declare Function API_WSALookupServiceNext Lib "ws2_32.dll" Alias "WSALookupServiceNextA" (ByVal hLookup As Long, ByVal dwControlFlags As Long, ByRef lpdwBufferLength As Long, ByRef lpqsResults As WSAQUERYSETA) As Long
'Private Declare Function API_WSALookupServiceEnd Lib "ws2_32.dll" (ByVal hLookup As Long) As Long
'Private Declare Function API_WSALookupServiceBegin Lib "ws2_32.dll" Alias "WSALookupServiceBeginA" (ByRef lpqsRestrictions As WSAQUERYSETA, ByVal dwControlFlags As Long, ByRef lphLookup As Long) As Long
'Private Declare Function API_WSAJoinLeaf Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr, ByVal namelen As Long, ByRef lpCallerData As WSABUF, ByRef lpCalleeData As WSABUF, ByRef lpSQOS As QOS, ByRef lpGQOS As QOS, ByVal dwFlags As Long) As Long
'Private Declare Function API_WSAInstallServiceClass Lib "ws2_32.dll" Alias "WSAInstallServiceClassA" (ByRef lpServiceClassInfo As WSASERVICECLASSINFOA) As Long
'Private Declare Function API_WSAHtons Lib "ws2_32.dll" (ByVal s As Long, ByVal hostshort As Integer, ByRef lpnetshort As Integer) As Long
'Private Declare Function API_WSAHtonl Lib "ws2_32.dll" (ByVal s As Long, ByVal hostlong As Long, ByRef lpnetlong As Long) As Long
'Private Declare Function API_WSAGetServiceClassNameByClassId Lib "ws2_32.dll" Alias "WSAGetServiceClassNameByClassIdA" (ByVal lpServiceClassId As Long, ByVal lpszServiceClassName As String, ByRef lpdwBufferLength As Long) As Long
'Private Declare Function API_WSAGetServiceClassInfo Lib "ws2_32.dll" Alias "WSAGetServiceClassInfoA" (ByVal lpProviderId As Long, ByVal lpServiceClassId As Long, ByRef lpdwBufSize As Long, ByRef lpServiceClassInfo As WSASERVICECLASSINFOA) As Long
'Private Declare Function API_WSAGetOverlappedResult Lib "ws2_32.dll" (ByVal s As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByRef lpcbTransfer As Long, ByVal fWait As Long, ByRef lpdwFlags As Long) As Long
'Private Declare Function API_WSAFDIsSet Lib "ws2_32.dll" (ByVal socket As Long, ByRef TFd_set As fd_set) As Long
'Private Declare Function API_WSAEventSelect Lib "ws2_32.dll" (ByVal s As Long, ByRef hEventObject As WSAEVENT, ByVal lNetworkEvents As Long) As Long
'Private Declare Function API_WSAEnumProtocols Lib "ws2_32.dll" Alias "WSAEnumProtocolsA" (ByRef lpiProtocols As Long, ByRef lpProtocolBuffer As WSAPROTOCOL_INFOA, ByRef lpdwBufferLength As Long) As Long
'Private Declare Function API_WSAEnumNetworkEvents Lib "ws2_32.dll" (ByVal s As Long, ByRef hEventObject As WSAEVENT, ByRef lpNetworkEvents As WSANETWORKEVENTS) As Long
'Private Declare Function API_WSAEnumNameSpaceProviders Lib "ws2_32.dll" Alias "WSAEnumNameSpaceProvidersA" (ByRef lpdwBufferLength As Long, ByRef lpnspBuffer As WSANAMESPACE_INFOA) As Long
'Private Declare Function API_WSADuplicateSocket Lib "ws2_32.dll" Alias "WSADuplicateSocketA" (ByVal s As Long, ByVal dwProcessId As Long, ByRef lpProtocolInfo As WSAPROTOCOL_INFOA) As Long
'Private Declare Function API_WSAConnect Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr, ByVal namelen As Long, ByRef lpCallerData As WSABUF, ByRef lpCalleeData As WSABUF, ByRef lpSQOS As QOS, ByRef lpGQOS As QOS) As Long
'Private Declare Function API_WSACloseEvent Lib "ws2_32.dll" (ByRef hEvent As WSAEVENT) As Long
'Private Declare Function API_WSACancelBlockingCall Lib "ws2_32.dll" () As Long
'Private Declare Function API_WSACancelAsyncRequest Lib "ws2_32.dll" (ByVal hAsyncTaskHandle As Long) As Long
'Private Declare Function API_WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
'Private Declare Function API_WSAAsyncGetServByPort Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal port As Long, ByVal proto As String, ByVal buf As String, ByVal buflen As Long) As Long
'Private Declare Function API_WSAAsyncGetServByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal name As String, ByVal proto As String, ByVal buf As String, ByVal buflen As Long) As Long
'Private Declare Function API_WSAAsyncGetProtoByNumber Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal number As Long, ByVal buf As String, ByVal buflen As Long) As Long
'Private Declare Function API_WSAAsyncGetProtoByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal name As String, ByVal buf As String, ByVal buflen As Long) As Long
'Private Declare Function API_WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal name As String, ByVal buf As String, ByVal buflen As Long) As Long
'Private Declare Function API_WSAAsyncGetHostByAddr Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal addr As String, ByVal lLen As Long, ByVal lType As Long, ByVal buf As String, ByVal buflen As Long) As Long
'Private Declare Function API_WSAAccept Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As sockaddr, ByRef addrlen As Long, ByRef lpfnCondition As CONDITIONPROC, ByVal dwCallbackData As Long) As Long

Private Declare Function API_baseNetwork_Socket Lib "ws2_32" Alias "socket" (ByVal AddressFamily As Long, ByVal sockType As Long, ByVal Protocol As Long) As Long
Private Declare Function API_baseNetwork_Bind Lib "ws2_32" Alias "bind" (ByVal Sock As Long, ByRef Addr As API_SOCKADDR, ByVal NameLen As Long) As Long
Private Declare Function API_baseNetwork_Accept Lib "ws2_32" Alias "accept" (ByVal Sock As Long, ByRef Addr As API_SOCKADDR, ByRef AddrLen As Long) As Long
Private Declare Function API_baseNetwork_Connect Lib "ws2_32" Alias "connect" (ByVal Sock As Long, ByRef Name As API_SOCKADDR, ByVal NameLen As Long) As Long
Private Declare Function API_baseNetwork_Listen Lib "ws2_32" Alias "listen" (ByVal Sock As Long, ByVal BackLog As Long) As Long

Private Declare Function API_baseNetwork_Send Lib "ws2_32" Alias "send" (ByVal Sock As Long, buff As Any, ByVal lLen As Long, ByVal Flags As Long) As Long
Private Declare Function API_baseNetwork_Recieve Lib "ws2_32" Alias "recv" (ByVal Sock As Long, buff As Any, ByVal lLen As Long, ByVal Flags As Long) As Long
Private Declare Function API_baseNetwork_SendTo Lib "ws2_32" Alias "sendto" (ByVal Sock As Long, buff As Any, ByVal lLen As Long, ByVal Flags As Long, ByRef toClient As API_SOCKADDR, ByVal toLen As Long) As Long
Private Declare Function API_baseNetwork_RecieveFrom Lib "ws2_32" Alias "recvfrom" (ByVal Sock As Long, buff As Any, ByVal lLen As Long, ByVal Flags As Long, ByRef From As API_SOCKADDR_IN, ByRef FromLen As Long) As Long

Private Declare Function API_baseNetwork_INet_addr Lib "ws2_32" Alias "inet_addr" (ByVal CP As String) As Long
Private Declare Function API_baseNetwork_INet_ntoa Lib "ws2_32" Alias "inet_ntoa" (ByRef lIn As Long) As String

Private Declare Function API_baseNetwork_htonl Lib "ws2_32" Alias "htonl" (ByVal HostLong As Long) As Long
Private Declare Function API_baseNetwork_htons Lib "ws2_32" Alias "htons" (ByVal HostShort As Integer) As Integer
Private Declare Function API_baseNetwork_ntohs Lib "ws2_32" Alias "ntohs" (ByVal netShort As Integer) As Integer
Private Declare Function API_baseNetwork_ntohl Lib "ws2_32" Alias "ntohl" (ByVal netLong As Long) As Long
Private Declare Function API_baseNetwork_GetPeerName Lib "ws2_32" Alias "getpeername" (ByVal Sock As Long, ByRef Name As API_SOCKADDR, ByRef NameLen As Long) As Long
Private Declare Function API_baseNetwork_GetHostName Lib "ws2_32" Alias "gethostname" (ByVal Name As String, ByVal NameLen As Long) As Long
Private Declare Function API_baseNetwork_GetHostByName Lib "ws2_32" Alias "gethostbyname" (Name As String) As API_HostEnt  'API_HostEnt 'ALSO Name is char far*

Private Declare Function API_baseNetwork_SetSocketOption Lib "ws2_32" Alias "setsockopt" (ByVal Sock As Long, ByVal Level As Long, ByVal OptName As Long, ByVal OptVal As String, ByVal OptLen As Long) As Long

Private Declare Function API_baseNetwork_Shutdown Lib "ws2_32" Alias "shutdown" (ByVal Sock As Long, ByVal How As Long) As Long
Private Declare Function API_baseNetwork_CloseSocket Lib "ws2_32" Alias "closesocket" (ByVal Sock As Long) As Long

Const WSADESCRIPTION_LEN As Long = 256
Const WSASYS_STATUS_LEN   As Long = 128

Public Type API_HostEnt
    h_name As String ': char far*
    h_aliases() As String 'char far* far* : means: Array<String>
    h_addrtype As Integer
    h_length As Integer
    h_addr_list() As String 'char far* far* : means: Array<String>
End Type
Public Type API_EndPoint
    IP As String
    Port As Long
End Type
Public Type API_WSADATA
  wVersion As Integer
  wHighVersion As Integer
  szDescription As String * WSADESCRIPTION_LEN
  szSystemStatus As String * WSASYS_STATUS_LEN
  iMaxSockets As Integer
  iMaxUdpDg As Integer
  lpVendorInfo As Long
End Type


Public Enum API_AddressFamily
    API_afUnknown = AF_UNKNOWN
    API_afUnSpec = AF_UNSPEC
    API_afUnspecified = API_afUnSpec
    API_afUnix = AF_UNIX
    API_afINet = AF_INET
    API_afInterNetwork = AF_INET
    API_afImpLink = AF_IMPLINK
    API_afPUP = AF_PUP
    API_afChaos = AF_CHAOS
    API_afIpx = AF_IPX
    API_afNS = AF_NS
    API_afIso = AF_ISO
    API_afOsi = AF_OSI
    API_afEcma = AF_ECMA
    API_afDataKit = AF_DATAKIT
    API_afCCitt = AF_CCITT
    API_afSna = AF_SNA
    API_afDecNet = AF_DECnet
    API_afDLI = AF_DLI
    API_afLAT = AF_LAT
    API_afHyLink = AF_HYLINK
    API_afHyperChannel = AF_HYLINK
    API_afAppleTalk = AF_APPLETALK
    API_afNetBios = AF_NETBIOS
    API_afVoiceView = AF_VOICEVIEW
    API_afFirefox = AF_FIREFOX
    API_afUnknown1 = AF_UNKNOWN1 'Somebody is using this!
    API_afBAN = AF_BAN
    API_afBanyan = AF_BAN
    API_afAtm = AF_ATM
    API_afInterNetworkV6 = AF_INTERNETWORKV6
    API_afCluster = AF_CLUSTER
    API_afIeee12844 = AF_IEEE12844
    API_afIrda = AF_IRDA
    API_afNetworkDesigners = AF_NETWORKDESIGNERS

    API_afMAX = AF_MAX
End Enum
Public Enum API_SocketType
    API_stUnknown = SOCK_UNKNOWN
    API_stStream = SOCK_STREAM                  'stream socket
    API_stDataGram = SOCK_DGRAM                 'datagram socket
    API_stRAW = SOCK_RAW                        'raw-protocol interface
    API_stReliablyDeliveredMessage = SOCK_RDM   'reliably-delivered message
    API_stSequencedPacket = SOCK_SEQPACKET      'sequenced packet stream
End Enum
Public Enum API_Protocol
    API_pUNKNOWN = IPPROTO_UNKNOWN
    API_pIPV6HOPBYHOPOPTIONS = IPPROTO_IPV6HOPBYHOPOPTIONS
    API_pUnspecified = IPPROTO_Unspecified
    API_pIP = IPPROTO_IP
    API_pICMP = IPPROTO_ICMP
    API_pIGMP = IPPROTO_IGMP
    API_pGGP = IPPROTO_GGP
    API_pIPV4 = IPPROTO_IPV4
    API_pTCP = IPPROTO_TCP
    API_pPUP = IPPROTO_PUP
    API_pUDP = IPPROTO_UDP
    API_pIDP = IPPROTO_IDP
    API_pIPV6 = IPPROTO_IPV6
    API_pIPV6ROUTINGHEADER = IPPROTO_IPV6ROUTINGHEADER
    API_pIPV6FRAGMENTHEADER = IPPROTO_IPV6FRAGMENTHEADER
    API_pIPSECENCAPSULATINGSECURITYPAYLOAD = IPPROTO_IPSECENCAPSULATINGSECURITYPAYLOAD
    API_pIPSECAUTHENTICATIONHEADER = IPPROTO_IPSECAUTHENTICATIONHEADER
    API_pICMPV6 = IPPROTO_ICMPV6
    API_pIPV6NONEXTHEADER = IPPROTO_IPV6NONEXTHEADER
    API_pIPV6DESTINATIONOPTIONS = IPPROTO_IPV6DESTINATIONOPTIONS
    API_pND = IPPROTO_ND
    API_pRAW = IPPROTO_RAW
    API_pIPX = IPPROTO_IPX
    
    API_pSPX = IPPROTO_SPX
    API_pSPXII = IPPROTO_SPXII
    API_pMAX = IPPROTO_MAX
End Enum

Public Enum API_SocketDirection
    API_sdInput = 0
    API_sdOutput = 1
    API_sdBoth = 2
End Enum

'Structure used for manipulating linger option.
Public Type API_LINGER
    l_onoff As Integer ' option on/off
    l_linger As Integer ' linger time
End Type
Public Type API_S_UN_B
    s_b1 As Byte
    s_b2 As Byte
    s_b3 As Byte
    s_b4 As Byte
End Type
Public Type API_S_UN_W
    s_w1 As Integer
    s_w2 As Integer
End Type
Public Type API_INTERNET_ADDR
'    S_un_b As API_S_UN_B
'    S_un_w  As API_S_UN_W
    S_addr As Long
End Type
Public Type API_SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As API_INTERNET_ADDR
    sin_zero As String * 8
End Type
Public Type API_SOCKADDR
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type


Public Type API_IPv4Address
    IP As Long '4Bytes
End Type 'Total 4 Bytes
Public Type API_IPv6Address
    P1 As Long 'High 8Bytes
    P2 As Long
    P3 As Long
    P4 As Long 'Low  8Bytes
End Type 'Total 16 Bytes

Public Type API_Socket
    Handle As Long
    AddressFamily As API_AddressFamily
    SocketType As API_SocketType
    Protocol As API_Protocol
    Port As Long
    EndPointAddress As API_SOCKADDR
    EndPointPort As Long
End Type

Dim inited As Boolean

Public Sub Initialize()
    If inited Then Exit Sub
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    inited = False
End Sub

Public Sub InitializeWinSock(lpVersionRequiredMajor As Byte, lpVersionRequiredMinor As Byte)
    Dim WSA As API_WSADATA
    If API_WSAStartup((lpVersionRequiredMajor * &HFF) + lpVersionRequiredMinor, WSA) <> 0 Then _
        throw SystemCallFailureException
End Sub
Public Sub DisposeWinSock()
    Call API_WSACleanup
End Sub

Public Function Create(AddressFamily As API_AddressFamily, SocketType As API_SocketType, Protocol As API_Protocol) As API_Socket
    Dim Sock_retVal As Long
    Sock_retVal = API_baseNetwork_Socket(AddressFamily, SocketType, Protocol)
    If Sock_retVal = INVALID_SOCKET Then _
        throw SystemCallFailureException("An error occured when trying to create socket in ws2_32.dll.")
    
    Create.Handle = Sock_retVal
    
    Create.AddressFamily = AddressFamily
    Create.SocketType = SocketType
    Create.Protocol = Protocol
    Create.Port = 0
End Function
Public Sub Destroy(Sock As API_Socket)
    If Sock.Handle = 0 Or Sock.Handle = INVALID_SOCKET Then Exit Sub
    Call API_baseNetwork_CloseSocket(Sock.Handle)
End Sub
Public Sub Shutdown(Sock As API_Socket, Dir As API_SocketDirection)
    If (Dir < 0) Or (Dir > 2) Then throw OutOfRangeException
    Call API_baseNetwork_Shutdown(Sock.Handle, Dir)
End Sub

Public Function SockADDRtoSockADDR_IN(SockAddr As API_SOCKADDR) As API_SOCKADDR_IN
    Call API_CopyMemory(ByVal VarPtr(SockADDRtoSockADDR_IN), ByVal VarPtr(SockAddr), Len(SockAddr))
End Function
Public Function SockADDR_INtoSockADDR(SOCKADDR_IN As API_SOCKADDR_IN) As API_SOCKADDR
    Call API_CopyMemory(ByVal VarPtr(SockADDR_INtoSockADDR), ByVal VarPtr(SOCKADDR_IN), Len(SOCKADDR_IN))
End Function
Public Function Internet_ADDRtoIntegers(Internet_ADDR As API_INTERNET_ADDR) As API_S_UN_W
    Call API_CopyMemory(ByVal VarPtr(Internet_ADDRtoIntegers), ByVal VarPtr(Internet_ADDR), Len(Internet_ADDR))
End Function
Public Function Internet_ADDRtoBytes(Internet_ADDR As API_INTERNET_ADDR) As API_S_UN_B
    Call API_CopyMemory(ByVal VarPtr(Internet_ADDRtoBytes), ByVal VarPtr(Internet_ADDR), Len(Internet_ADDR))
End Function



Public Function htons(HostShort As Integer) As Integer
    htons = API_baseNetwork_htons(HostShort)
End Function
Public Function htonl(HostLong As Long) As Long
    htonl = API_baseNetwork_htonl(HostLong)
End Function
Public Function ntohs(NetworkShort As Integer) As Integer
    ntohs = API_baseNetwork_ntohs(NetworkShort)
End Function
Public Function ntohl(HostLong As Long) As Long
    ntohl = API_baseNetwork_ntohl(HostLong)
End Function

Public Function NetAddressToAddress(NetAddress As Long) As String
    NetAddressToAddress = API_baseNetwork_INet_ntoa(NetAddress)
End Function
Public Function AddressToNetAddress(Address As String) As Long
    AddressToNetAddress = API_baseNetwork_INet_addr(Address)
End Function

Public Function Bind(Sock As API_Socket, SockAddress As API_SOCKADDR) As Boolean
    If Sock.Handle = 0 Or Sock.Handle = INVALID_SOCKET Then throw InvalidStatusException("Invalid socket.")
    Dim retVal As Long
    retVal = API_baseNetwork_Bind(Sock.Handle, SockAddress, Len(SockAddress))
    If retVal = INADDR_NONE Then
        Bind = False
        'throw SystemCallFailureException("Unable to bind socket, An error occured in ws2_32.dll.")
    Else
        Bind = True
        Sock.Port = SockAddress.sin_port
        Sock.AddressFamily = SockAddress.sin_family
    End If
End Function
Public Sub Listen(Sock As API_Socket, BackLog As Long)
    If Sock.Handle = 0 Or Sock.Handle = INVALID_SOCKET Then throw InvalidStatusException("Invalid socket.")
    If API_baseNetwork_Listen(Sock.Handle, BackLog) = SOCKET_ERROR Then _
        throw SystemCallFailureException("An error occured in ws2_32.dll while trying to call listen().")
End Sub
Public Function Accept(Sock As API_Socket) As API_Socket
    Dim retVal As Long, outAddr As API_SOCKADDR
    retVal = API_baseNetwork_Accept(Sock.Handle, outAddr, Len(outAddr))
    
    If retVal = SOCKET_ERROR Then _
        throw SystemCallFailureException("An error occured in ws2_32.dll while trying to call accept().")
        
    Accept.Handle = retVal
    Accept.AddressFamily = outAddr.sin_family
    Accept.SocketType = Sock.SocketType
    Accept.Protocol = Sock.Protocol
End Function
Public Sub Connect(Sock As API_Socket, SockAddress As API_SOCKADDR)
    If Sock.Handle = 0 Or Sock.Handle = INVALID_SOCKET Then throw InvalidStatusException("Invalid socket.")
    If API_baseNetwork_Connect(Sock.Handle, SockAddress, Len(SockAddress)) = SOCKET_ERROR Then _
        throw SystemCallFailureException("Unable to connect to host, An error occured in ws2_32.dll while trying to call connect().")
End Sub


Public Sub Ding(Sock As API_Socket, DingValue As Long, Flags As Long)
    If Sock.Handle = 0 Or Sock.Handle = INVALID_SOCKET Then throw InvalidStatusException("Invalid socket.")
    If API_baseNetwork_Send(Sock.Handle, htonl(DingValue), 4, Flags) = SOCKET_ERROR Then _
        throw SystemCallFailureException("Unable to send data through socket, An error occured in ws2_32.dll.")
End Sub
Public Sub Send(Sock As API_Socket, baData() As Byte, Flags As Long)
    If Sock.Handle = 0 Or Sock.Handle = INVALID_SOCKET Then throw InvalidStatusException("Invalid socket.")
    Dim baLen As Long
    baLen = ArraySize(baData)
    If baLen <= 0 Then Exit Sub
    If API_baseNetwork_Send(Sock.Handle, baData(0), baLen, Flags) = SOCKET_ERROR Then
        throw SystemCallFailureException("Unable to send data through socket, An error occured in ws2_32.dll.")
    End If
End Sub
Public Function ReadDing(Sock As API_Socket, DingValue As Long, Flags As Long) As Long
    If Sock.Handle = 0 Or Sock.Handle = INVALID_SOCKET Then throw InvalidStatusException("Invalid socket.")
    If API_baseNetwork_Recieve(Sock.Handle, ReadDing, 4, Flags) = SOCKET_ERROR Then _
        throw SystemCallFailureException("Unable to recieve data from socket, An error occured in ws2_32.dll.")
    ReadDing = ntohl(ReadDing)
End Function
Public Function Recieve(Sock As API_Socket, BufferSize As Long, Flags As Long) As Byte()
    If Sock.Handle = 0 Or Sock.Handle = INVALID_SOCKET Then throw InvalidStatusException("Invalid socket.")
    If BufferSize = 0 Then Exit Function
    If BufferSize < 0 Then throw NegativeArgumentException
    If BufferSize > (50 * 1024) Then throw ValueIsTooHightException("BufferSize is too large, MintAPI dont let socket recieve buffer length more than 51200.")
    Dim baData() As Byte
    ReDim baData(BufferSize - 1)
    If API_baseNetwork_Recieve(Sock.Handle, baData(0), BufferSize, Flags) = SOCKET_ERROR Then
        throw SystemCallFailureException("Unable to recieve data from socket, An error occured in ws2_32.dll.")
    End If
End Function

Public Function ValidateSockAddress(IPAddress As String, SocketPort As Long) As Boolean
    If (SocketPort < MIN_PORT) Or (SocketPort > MAX_PORT) Then Exit Function
    Dim retVal As Long
    retVal = API_baseNetwork_INet_addr(IPAddress)
    
    If retVal <> SOCKET_ERROR Then
        ValidateSockAddress = True
    Else
        ValidateSockAddress = False
    End If
End Function
Public Function ValidateIPAddress(IPAddress As String) As Boolean
    Dim retVal As Long
    retVal = API_baseNetwork_INet_addr(IPAddress)
    
    If retVal <> SOCKET_ERROR Then
        ValidateIPAddress = True
    Else
        ValidateIPAddress = False
    End If
End Function

Public Function EndPointToAddress(EndPoint As API_EndPoint) As API_SOCKADDR
    
End Function

Public Function CreateEndPoint_IPV4(IPAddress As String, SocketPort As Long) As API_EndPoint
    If Not ValidateSockAddress(IPAddress, SocketPort) Then _
        throw InvalidArgumentValueException("Invalid IP/Port.")
    CreateEndPoint_IPV4.IP = IPAddress
    CreateEndPoint_IPV4.Port = SocketPort
End Function
Public Function CreateEndPoint_IPV6(IPAddress As String, SocketPort As Long) As API_EndPoint
    If Not ValidateSockAddress(IPAddress, SocketPort) Then _
        throw InvalidArgumentValueException("Invalid IP/Port.")
    CreateEndPoint_IPV6.IP = IPAddress
    CreateEndPoint_IPV6.Port = SocketPort
End Function
