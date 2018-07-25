VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Socket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

'Ideas from Me , Qt , .net Framework , WinSock2

Option Base 0
Option Explicit
Const CLASSID As String = "Socket"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


''<class>
''  <summary>Provides socket connection both serverside and clientside.</summary>
''  <using>ByteArray</using>
''</class>

Const SIGNAL_SENDING As String = "sending"
Const SIGNAL_SEND As String = "send"
Const SIGNAL_RECIEVE As String = "recieve"
Const SIGNAL_LISTENING As String = "listening"
Const SIGNAL_CONNECT As String = "connect"
Const SIGNAL_ANSWER As String = "answer"
Const SIGNAL_CREATE As String = "create"
Const SIGNAL_SHUTDOWN As String = "shutdown"
Const SIGNAL_INPUTSHUTDOWN As String = "inputshutdown"
Const SIGNAL_OUTPUTSHUTDOWN As String = "outputshutdown"
Const SIGNAL_DISCONNECT As String = "disconnect"
Const SIGNAL_CLEAR As String = "clear"
Const SIGNAL_WSAINITIALIZE As String = "wsainitialize"
Const SIGNAL_CLOSE As String = "close"
Const SIGNAL_INITIALIZE As String = "initialize"
Const SIGNAL_DATARECIEVED As String = "datarecieved"
Const SIGNAL_CONNECTIONREQUEST As String = "connectionrequest"

Const SLOT_FLUSH As String = "flush"
Const SLOT_DISCONNECT As String = "disconnect"
Const SLOT_CLEAR As String = "clear"


Implements ITargetStream
Implements ICloneable
Implements ITexable
Implements ISupportSignal
Implements ISupportSlot
Implements IAliasable

Public Event DataRecieved(Socket As Socket)
Public Event ConnectionRequest(Socket As Socket)
Public Event Disconnect(Socket As Socket)
Public Event Shutdown(Socket As Socket)
Public Event Create(Socket As Socket)
Public Event SocketClose(Socket As Socket)
Public Event Connect(Socket As Socket)
Public Event Send(Socket As Socket)
Public Event Recieve(Socket As Socket)
Public Event Listening(Socket As Socket)
Public Event WSAInitialize(Socket As Socket)


Const DEFAULTSENDBUFFER As Long = 2048
Const DEFAULTRECIEVEBUFFER As Long = 1024


'=====================================================
'
'  by Ali Mousavi Kherad
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
Private Const IPPROTO_UNKNOWN = -1               'Unknown protocol.
Private Const IPPROTO_IPV6HOPBYHOPOPTIONS = 0    'IPv6 Hop by Hop Options header.
Private Const IPPROTO_Unspecified = 0            'Unspecified protocol.
Private Const IPPROTO_IP = 0                     'dummy for IP | Internet Protocol.
Private Const IPPROTO_ICMP = 1                   'control message protocol | Internet Control Message Protocol.
Private Const IPPROTO_IGMP = 2                   'group management protocol | Internet Group Management Protocol.
Private Const IPPROTO_GGP = 3                    'gateway^2 (deprecated) | Gateway To Gateway Protocol.
Private Const IPPROTO_IPV4 = 4                   'Internet Protocol version 4.
Private Const IPPROTO_TCP = 6                    'Transmission Control Protocol.
Private Const IPPROTO_PUP = 12                   'PARC Universal Packet Protocol.
Private Const IPPROTO_UDP = 17                   'User Datagram Protocol.
Private Const IPPROTO_IDP = 22                   'xns idp | Internet Datagram Protocol.
Private Const IPPROTO_IPV6 = 41                  'Internet Protocol version 6 (IPv6).
Private Const IPPROTO_IPV6ROUTINGHEADER = 43     'IPv6 Routing header.
Private Const IPPROTO_IPV6FRAGMENTHEADER = 44    'IPv6 Fragment header.
Private Const IPPROTO_IPSECENCAPSULATINGSECURITYPAYLOAD = 50 'IPv6 Encapsulating Security Payload header.
Private Const IPPROTO_IPSECAUTHENTICATIONHEADER = 51 'IPv6 Authentication header. For details, see RFC 2292 section 2.2.1, available at http://www.ietf.org.
Private Const IPPROTO_ICMPV6 = 58                'Internet Control Message Protocol for IPv6.
Private Const IPPROTO_IPV6NONEXTHEADER = 59      'IPv6 No next header.
Private Const IPPROTO_IPV6DESTINATIONOPTIONS = 60 'IPv6 Destination Options header.
Private Const IPPROTO_ND = 77                    'UNOFFICIAL net disk proto | Net Disk Protocol (unofficial).
Private Const IPPROTO_RAW = 255                  'Raw IP packet protocol.
Private Const IPPROTO_IPX = 1000                 'Internet Packet Exchange Protocol.

Private Const IPPROTO_SPX = 1256                 'Sequenced Packet Exchange protocol.
Private Const IPPROTO_SPXII = 1257               'Sequenced Packet Exchange version 2 protocol.
Private Const IPPROTO_MAX = 256


'Port/socket numbers: network standard functions.
Private Const IPPORT_ECHO = 7
Private Const IPPORT_DISCARD = 9
Private Const IPPORT_SYSTAT = 11
Private Const IPPORT_DAYTIME = 13
Private Const IPPORT_NETSTAT = 15
Private Const IPPORT_FTP = 21
Private Const IPPORT_TELNET = 23
Private Const IPPORT_SMTP = 25
Private Const IPPORT_TIMESERVER = 37
Private Const IPPORT_NAMESERVER = 42
Private Const IPPORT_WHOIS = 43
Private Const IPPORT_MTP = 57

'Port/socket numbers: host specific functions.
Private Const IPPORT_TFTP = 69
Private Const IPPORT_RJE = 77
Private Const IPPORT_FINGER = 79
Private Const IPPORT_TTYLINK = 87
Private Const IPPORT_SUPDUP = 95
'UNIX TCP sockets.
Private Const IPPORT_EXECSERVER = 512
Private Const IPPORT_LOGINSERVER = 513
Private Const IPPORT_CMDSERVER = 514
Private Const IPPORT_EFSSERVER = 520
'UNIX UDP sockets.
Private Const IPPORT_BIFFUDP = 512
Private Const IPPORT_WHOSERVER = 513
Private Const IPPORT_ROUTESERVER = 520
' 520+1 also used.
'Ports < IPPORT_RESERVED are reserved for
' privileged processes (e.g. root).
Private Const IPPORT_RESERVED = 1024

'Options for use with [gs]etsockopt at the IP level.
Private Const IP_OPTIONS = 1                   'set/get IP per-packet options
Private Const IP_MULTICAST_IF = 2              'set/get IP multicast interface
Private Const IP_MULTICAST_TTL = 3             'set/get IP multicast timetolive
Private Const IP_MULTICAST_LOOP = 4            'set/get IP multicast loopback
Private Const IP_ADD_MEMBERSHIP = 5            'add  an IP group membership
Private Const IP_DROP_MEMBERSHIP = 6           'drop an IP group membership
Private Const IP_TTL = 7                       'set/get IP Time To Live
Private Const IP_TOS = 8                       'set/get IP Type Of Service
Private Const IP_DONTFRAGMENT = 9              'set/get IP Don't Fragment flag
'
Private Const IP_DEFAULT_MULTICAST_TTL = 1     'normally limit m'casts to 1 hop
Private Const IP_DEFAULT_MULTICAST_LOOP = 1    'normally hear sends if a member
Private Const IP_MAX_MEMBERSHIPS = 20          'per socket; must fit in one mbuf



'Link numbers
Private Const IMPLINK_IP = 155
Private Const IMPLINK_LOWEXPER = 156
Private Const IMPLINK_HIGHEXPER = 158

'This is used instead of -1, since the
' SOCKET type is unsigned.
Private Const INVALID_SOCKET = -1
Private Const SOCKET_ERROR = -1

'Types
Private Const SOCK_UNKNOWN = -1           'Specifies an unknown Socket type.
Private Const SOCK_STREAM = 1             'stream socket
'SOCK_STREAM:
'   Supports reliable, two-way, connection-based byte streams without the duplication
'   of data and without preservation of boundaries. A Socket of this type communicates
'   with a single peer and requires a remote host connection before communication
'   can begin. System.Net.Sockets.SocketType.Stream uses the Transmission Control
'   Protocol (System.Net.Sockets.ProtocolType.Tcp) System.Net.Sockets.ProtocolType
'   and the InterNetworkSystem.Net.Sockets.AddressFamily.
Private Const SOCK_DGRAM = 2              'datagram socket
'SOCK_DGRAM:
'   Supports datagrams, which are connectionless, unreliable messages of a fixed
'   (typically small) maximum length. Messages might be lost or duplicated and
'   might arrive out of order. A System.Net.Sockets.Socket of type System.Net.Sockets.SocketType.Dgram
'   requires no connection prior to sending and receiving data, and can communicate
'   with multiple peers. System.Net.Sockets.SocketType.Dgram uses the Datagram
'   Protocol (System.Net.Sockets.ProtocolType.Udp) and the System.Net.Sockets.AddressFamily.InterNetworkSystem.Net.Sockets.AddressFamily.
Private Const SOCK_RAW = 3                'raw-protocol interface
'SOCK_RAW:
'   Supports access to the underlying transport protocol. Using the System.Net.Sockets.SocketTypeSystem.Net.Sockets.SocketType.Raw,
'   you can communicate using protocols like Internet Control Message Protocol
'   (System.Net.Sockets.ProtocolType.Icmp) and Internet Group Management Protocol
'   (System.Net.Sockets.ProtocolType.Igmp). Your application must provide a complete
'   IP header when sending. Received datagrams return with the IP header and
'   options intact.
Private Const SOCK_RDM = 4                'reliably-delivered message
'SOCK_RDM:
'   Supports connectionless, message-oriented, reliably delivered messages, and
'   preserves message boundaries in data. Rdm (Reliably Delivered Messages) messages
'   arrive unduplicated and in order. Furthermore, the sender is notified if
'   messages are lost. If you initialize a Socket using System.Net.Sockets.SocketType.Rdm,
'   you do not require a remote host connection before sending and receiving
'   data. With System.Net.Sockets.SocketType.Rdm, you can communicate with multiple
'   peers.
Private Const SOCK_SEQPACKET = 5          'sequenced packet stream
'SOCK_SEQPACKET:
'   Provides connection-oriented and reliable two-way transfer of ordered byte
'   streams across a network. System.Net.Sockets.SocketType.Seqpacket does not
'   duplicate data, and it preserves boundaries within the data stream. A Socket
'   of type System.Net.Sockets.SocketType.Seqpacket communicates with a single
'   peer and requires a remote host connection before communication can begin.

'Option flags per-socket.
Private Const SO_DEBUG = &H1              'turn on debugging info recording
Private Const SO_ACCEPTCONN = &H2         'socket has had listen()
Private Const SO_REUSEADDR = &H4          'allow local address reuse
Private Const SO_KEEPALIVE = &H8          'keep connections alive
Private Const SO_DONTROUTE = &H10         'just use interface addresses
Private Const SO_BROADCAST = &H20         'permit sending of broadcast msgs
Private Const SO_USELOOPBACK = &H40       'bypass hardware when possible
Private Const SO_LINGER = &H80            'linger on close if data present
Private Const SO_OOBINLINE = &H100        'leave received OOB data in line
'Private Const  SO_DONTLINGER   (u_int)(~SO_LINGER)

'Additional options.
Private Const SO_SNDBUF = &H1001          'send buffer size
Private Const SO_RCVBUF = &H1002          'receive buffer size
Private Const SO_SNDLOWAT = &H1003        'send low-water mark
Private Const SO_RCVLOWAT = &H1004        'receive low-water mark
Private Const SO_SNDTIMEO = &H1005        'send timeout
Private Const SO_RCVTIMEO = &H1006        'receive timeout
Private Const SO_ERROR = &H1007           'get error status and clear
Private Const SO_TYPE = &H1008            'get socket type

'Options for connect and disconnect data and options.  Used only by
' non-TCP/IP transports such as DECNet, OSI TP4, etc.
Private Const SO_CONNDATA = &H7000
Private Const SO_CONNOPT = &H7001
Private Const SO_DISCDATA = &H7002
Private Const SO_DISCOPT = &H7003
Private Const SO_CONNDATALEN = &H7004
Private Const SO_CONNOPTLEN = &H7005
Private Const SO_DISCDATALEN = &H7006
Private Const SO_DISCOPTLEN = &H7007

'Option for opening sockets for synchronous access.
Private Const SO_OPENTYPE = &H7008
Private Const SO_SYNCHRONOUS_ALERT = &H10
Private Const SO_SYNCHRONOUS_NONALERT = &H20

'Other NT-specific options.
Private Const SO_MAXDG = &H7009
Private Const SO_MAXPATHDG = &H700A
Private Const SO_UPDATE_ACCEPT_CONTEXT = &H700B
Private Const SO_CONNECT_TIME = &H700C

'Socket Flags.
Private Const SFLAG_NONE = 0                     'Use no flags for this call.
Private Const SFLAG_OOUTOFBAND = 1                'Process out-of-band data.
Private Const SFLAG_PEEK = 2                     'Peek at the incoming message.
Private Const SFLAG_DONTROUTE = 4                'Send without using routing tables.
Private Const SFLAG_MAXIOVECTORLENGTH = 16       'Provides a standard value for the number of WSABUF structures that are used to send and receive data.
Private Const SFLAG_TRUNCATED = 256              'The message was too large to fit into the specified buffer and was truncated.
Private Const SFLAG_CONTROLDATATRUNCATED = 512   'Indicates that the control data did not fit into an internal 64-KB buffer and was truncated.
Private Const SFLAG_BROADCAST = 1024             'Indicates a broadcast packet.
Private Const SFLAG_MULTICAST = 2048             'Indicates a multicast packet.
Private Const SFLAG_PARTIAL = 32768              'Partial send or receive for message.

'TCP options.
Private Const TCP_NODELAY = &H1
Private Const TCP_BSDURGENT = &H7000

'Address families.
Private Const AF_UNKNOWN = -1                'Unknown address family.
Private Const AF_UNSPEC = 0                  'Unknown address family.
Private Const AF_UNIX = 1                    'local to host (pipes, portals) | Unix local to host address.
Private Const AF_INET = 2                    'internetwork: UDP, TCP, etc. | Address for IP version 4.
Private Const AF_INTERNETWORK = AF_INET
Private Const AF_IMPLINK = 3                 'arpanet imp addresses | ARPANET IMP address.
Private Const AF_PUP = 4                     'pup protocols: e.g. BSP | Address for PUP protocols.
Private Const AF_CHAOS = 5                   'mit CHAOS protocols | Address for MIT CHAOS protocols.
Private Const AF_IPX = 6                     'IPX and SPX | IPX or SPX address.
Private Const AF_NS = 6                      'XEROX NS protocols | Address for Xerox NS protocols.
Private Const AF_ISO = 7                     'ISO protocols | Address for ISO protocols.
Private Const AF_OSI = AF_ISO                'OSI is ISO | Address for ISO protocols.
Private Const AF_ECMA = 8                    'european computer manufacturers | European Computer Manufacturers Association (ECMA) address.
Private Const AF_DATAKIT = 9                 'datakit protocols | Address for Datakit protocols.
Private Const AF_CCITT = 10                  'CCITT protocols, X.25 etc | Addresses for CCITT protocols, such as X.25.
Private Const AF_SNA = 11                    'IBM SNA | IBM SNA address.
Private Const AF_DECnet = 12                 'DECnet | DECnet address.
Private Const AF_DLI = 13                    'Direct data link interface | Direct data-link interface address.
Private Const AF_DATALINK = AF_DLI
Private Const AF_LAT = 14                    'LAT | LAT address.
Private Const AF_HYLINK = 15                 'NSC Hyperchannel | NSC Hyperchannel address.
Private Const AF_HYPERCHANNEL = AF_HYLINK
Private Const AF_APPLETALK = 16              'AppleTalk | AppleTalk address.
Private Const AF_NETBIOS = 17                'NetBios-style addresses | NetBios address.
Private Const AF_VOICEVIEW = 18              'VoiceView | VoiceView address.
Private Const AF_FIREFOX = 19                'FireFox | FireFox address.
Private Const AF_UNKNOWN1 = 20               'Somebody is using this!
Private Const AF_BAN = 21                    'Banyan
Private Const AF_BANYAN = AF_BAN
Private Const AF_ATM = 22                    'Native ATM services address.
Private Const AF_INTERNETWORKV6 = 23         'Address for IP version 6.
Private Const AF_CLUSTER = 24                'Address for Microsoft cluster products.
Private Const AF_IEEE12844 = 25              'IEEE 1284.4 workgroup address.
Private Const AF_IRDA = 26                   'IrDA address.
Private Const AF_NETWORKDESIGNERS = 28       'Address for Network Designers OSI gateway-enabled protocols.
Private Const AF_MAX = 29                    'MAX address.

'Level number for (get/set)sockopt() to apply to socket itself.
Private Const SOL_SOCKET = &HFFFF 'options for socket level

'Maximum queue length specifiable by listen.
Private Const SOMAXCONN = 5

Private Const MSG_OOB = &H1                    'process out-of-band data
Private Const MSG_PEEK = &H2                   'peek at incoming message
Private Const MSG_DONTROUTE = &H4              'send without using routing tables

Private Const MSG_MAXIOVLEN = 16

Private Const MSG_PARTIAL = &H8000             'partial send or recv for message xport

'Define constant based on rfc883, used by gethostbyxxxx() calls.
Private Const MAXGETHOSTSTRUCT = 1024

'Define flags to be used with the WSAAsyncSelect() call.
Private Const FD_READ = &H1
Private Const FD_WRITE = &H2
Private Const FD_OOB = &H4
Private Const FD_ACCEPT = &H8
Private Const FD_CONNECT = &H10
Private Const FD_CLOSE = &H20


Private Const INADDR_ANY As Long = &H0
Private Const INADDR_LOOPBACK As Long = &H7F000001
Private Const INADDR_BROADCAST As Long = -1
Private Const INADDR_NONE As Long = -1

Private Const MAX_PORT As Long = 65535
Private Const MIN_PORT As Long = 1

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
Private Declare Function API_WSAStartup Lib "ws2_32" Alias "WSAStartup" (ByVal wVersionRequired As Integer, lpWSAData As API_WSADATA) As Long
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

Private Declare Function API_Socket Lib "ws2_32" Alias "socket" (ByVal AddressFamily As Long, ByVal SockType As Long, ByVal Protocol As Long) As Long
Private Declare Function API_Bind Lib "ws2_32" Alias "bind" (ByVal Sock As Long, ByRef Addr As API_SOCKADDR, ByVal NameLen As Long) As Long
Private Declare Function API_Accept Lib "ws2_32" Alias "accept" (ByVal Sock As Long, ByRef Addr As API_SOCKADDR, ByRef AddrLen As Long) As Long
Private Declare Function API_Connect Lib "ws2_32" Alias "connect" (ByVal Sock As Long, ByRef Name As API_SOCKADDR, ByVal NameLen As Long) As Long
Private Declare Function API_Listen Lib "ws2_32" Alias "listen" (ByVal Sock As Long, ByVal BackLog As Long) As Long

Private Declare Function API_Send Lib "ws2_32" Alias "send" (ByVal Sock As Long, buff As Any, ByVal lLen As Long, ByVal Flags As Long) As Long
Private Declare Function API_Recieve Lib "ws2_32" Alias "recv" (ByVal Sock As Long, buff As Any, ByVal lLen As Long, ByVal Flags As Long) As Long
Private Declare Function API_SendTo Lib "ws2_32" Alias "sendto" (ByVal Sock As Long, buff As Any, ByVal lLen As Long, ByVal Flags As Long, ByRef toClient As API_SOCKADDR, ByVal toLen As Long) As Long
Private Declare Function API_RecieveFrom Lib "ws2_32" Alias "recvfrom" (ByVal Sock As Long, buff As Any, ByVal lLen As Long, ByVal Flags As Long, ByRef From As API_SOCKADDR_IN, ByRef FromLen As Long) As Long

Private Declare Function API_INet_addr Lib "ws2_32" Alias "inet_addr" (ByVal CP As String) As Long
Private Declare Function API_INet_ntoa Lib "ws2_32" Alias "inet_ntoa" (ByRef lIn As Long) As String

Private Declare Function API_htonl Lib "ws2_32" Alias "htonl" (ByVal HostLong As Long) As Long
Private Declare Function API_htons Lib "ws2_32" Alias "htons" (ByVal HostShort As Integer) As Integer
Private Declare Function API_ntohs Lib "ws2_32" Alias "ntohs" (ByVal netShort As Integer) As Integer
Private Declare Function API_ntohl Lib "ws2_32" Alias "ntohl" (ByVal netLong As Long) As Long
Private Declare Function API_GetPeerName Lib "ws2_32" Alias "getpeername" (ByVal Sock As Long, ByRef Name As API_SOCKADDR, ByRef NameLen As Long) As Long
Private Declare Function API_GetHostName Lib "ws2_32" Alias "gethostname" (ByVal Name As String, ByVal NameLen As Long) As Long
Private Declare Function API_GetHostByName Lib "ws2_32" Alias "gethostbyname" (Name As String) As API_HostEnt  'API_HostEnt 'ALSO Name is char far*
Private Declare Function API_GetHostByNamePTR Lib "ws2_32" Alias "gethostbyname" (Name As String) As Long  'API_HostEnt 'ALSO Name is char far*

Private Declare Function API_IOCtlSocket Lib "ws2_32" Alias "ioctlsocket" (ByVal S As Long, ByVal Cmd As Long, ByRef Argp As Long) As Long

Private Declare Function API_SetSocketOption Lib "ws2_32" Alias "setsockopt" (ByVal Sock As Long, ByVal Level As Long, ByVal OptName As Long, ByVal OptVal As String, ByVal OptLen As Long) As Long

Private Declare Function API_Shutdown Lib "ws2_32" Alias "shutdown" (ByVal Sock As Long, ByVal How As Long) As Long
Private Declare Function API_CloseSocket Lib "ws2_32" Alias "closesocket" (ByVal Sock As Long) As Long

Public Type IPAddress
    IPVersion As Long
    IPAddress As Long
End Type
Public Type SocketAddress
    IPVersion As Long
    IPAddress As Long
    SocketPort As Long
End Type
Public Type EndPointTag
    Name As String
    AddressType As AddressFamily
    AliasListCount As Long
    AliasList() As String
    AddressListCount As Long
    AddressList() As String
    
    DCTag As Long 'daisy-chain tags
End Type
Public Type EndPoint
    AddressFamily As AddressFamily
    Tag As Long
    SocketPort As Long
    IPAddress As String
End Type

Public Enum AddressFamily
    afUnknown = AF_UNKNOWN
    afUnSpec = AF_UNSPEC
    afUnspecified = afUnSpec
    afUnix = AF_UNIX
    afINet = AF_INET
    afInterNetwork = AF_INET
    afImpLink = AF_IMPLINK
    afPUP = AF_PUP
    afChaos = AF_CHAOS
    afIpx = AF_IPX
    afNS = AF_NS
    afIso = AF_ISO
    afOsi = AF_OSI
    afEcma = AF_ECMA
    afDataKit = AF_DATAKIT
    afCCitt = AF_CCITT
    afSna = AF_SNA
    afDecNet = AF_DECnet
    afDLI = AF_DLI
    afLAT = AF_LAT
    afHyLink = AF_HYLINK
    afHyperChannel = AF_HYLINK
    afAppleTalk = AF_APPLETALK
    afNetBios = AF_NETBIOS
    afVoiceView = AF_VOICEVIEW
    afFirefox = AF_FIREFOX
    afUnknown1 = AF_UNKNOWN1 'Somebody is using this!
    afBAN = AF_BAN
    afBanyan = AF_BAN
    afAtm = AF_ATM
    afInterNetworkV6 = AF_INTERNETWORKV6
    afCluster = AF_CLUSTER
    afIeee12844 = AF_IEEE12844
    afIrda = AF_IRDA
    afNetworkDesigners = AF_NETWORKDESIGNERS

    afMAX = AF_MAX
End Enum
Public Enum SocketType
    stUnknown = SOCK_UNKNOWN
    stStream = SOCK_STREAM                  'stream socket
    stDataGram = SOCK_DGRAM                 'datagram socket
    stRAW = SOCK_RAW                        'raw-protocol interface
    stReliablyDeliveredMessage = SOCK_RDM   'reliably-delivered message
    stSequencedPacket = SOCK_SEQPACKET      'sequenced packet stream
End Enum
Public Enum Protocol
    pUNKNOWN = IPPROTO_UNKNOWN
    pIPV6HOPBYHOPOPTIONS = IPPROTO_IPV6HOPBYHOPOPTIONS
    pUnspecified = IPPROTO_Unspecified
    pIP = IPPROTO_IP
    pICMP = IPPROTO_ICMP
    pIGMP = IPPROTO_IGMP
    pGGP = IPPROTO_GGP
    pIPV4 = IPPROTO_IPV4
    pTCP = IPPROTO_TCP
    pPUP = IPPROTO_PUP
    pUDP = IPPROTO_UDP
    pIDP = IPPROTO_IDP
    pIPV6 = IPPROTO_IPV6
    pIPV6ROUTINGHEADER = IPPROTO_IPV6ROUTINGHEADER
    pIPV6FRAGMENTHEADER = IPPROTO_IPV6FRAGMENTHEADER
    pIPSECENCAPSULATINGSECURITYPAYLOAD = IPPROTO_IPSECENCAPSULATINGSECURITYPAYLOAD
    pIPSECAUTHENTICATIONHEADER = IPPROTO_IPSECAUTHENTICATIONHEADER
    pICMPV6 = IPPROTO_ICMPV6
    pIPV6NONEXTHEADER = IPPROTO_IPV6NONEXTHEADER
    pIPV6DESTINATIONOPTIONS = IPPROTO_IPV6DESTINATIONOPTIONS
    pND = IPPROTO_ND
    pRAW = IPPROTO_RAW
    pIPX = IPPROTO_IPX
    
    pSPX = IPPROTO_SPX
    pSPXII = IPPROTO_SPXII
    pMAX = IPPROTO_MAX
End Enum
Public Enum SocketFlags
    sflagNone = SFLAG_NONE
    sflagOOutOfBand = SFLAG_OOUTOFBAND
    sflagPeek = SFLAG_PEEK
    sflagDontRoute = SFLAG_DONTROUTE
    sflagMaxIOVectorLength = SFLAG_MAXIOVECTORLENGTH
    sflagTruncated = SFLAG_TRUNCATED
    sflagControlDataTruncated = SFLAG_CONTROLDATATRUNCATED
    sflagBroadcast = SFLAG_BROADCAST
    sflagMulticast = SFLAG_MULTICAST
    sflagPartial = SFLAG_PARTIAL
End Enum
Public Enum SocketPorts
    'Port/socket numbers: network standard functions.
    spECHO = 7
    spDISCARD = 9
    spSYSTAT = 11
    spDAYTIME = 13
    spNETSTAT = 15
    spFTP = 21
    spTELNET = 23
    spSMTP = 25
    spTIMESERVER = 37
    spNAMESERVER = 42
    spWHOIS = 43
    spMTP = 57
    
    'Port/socket numbers: host specific functions.
    spTFTP = 69
    spRJE = 77
    spFINGER = 79
    spHTTP = 80
    spTTYLINK = 87
    spSUPDUP = 95
    'UNIX TCP sockets.
    spEXECSERVER = 512
    spLOGINSERVER = 513
    spCMDSERVER = 514
    spEFSSERVER = 520
    'UNIX UDP sockets.
    spBIFFUDP = 512
    spWHOSERVER = 513
    spROUTESERVER = 520
    ' 520+1 also used.
    'Ports < spRESERVED are reserved for
    ' privileged processes (e.g. root).
    spRESERVED = 1024
End Enum
Public Enum SocketStates
    ssIdle = 0
    ssConnect = 1
    ssListen = 2
    ssAnswered = 3
    ssSend = 4
    ssRecieve = 5
    ssClose = 6
End Enum

Private Type API_Socket
    Handle As Long
    AddressFamily As AddressFamily
    SocketType As SocketType
    Protocol As Protocol
    Port As Long
    EndPointAddress As API_SOCKADDR
End Type

Public Enum SocketDirection
    skdInput = 0
    skdOutput = 1
    skdBoth = 2
End Enum

Public Type SocketInfo
    Port As String
    IP As String
    SocketType As SocketType
End Type

Dim sSock As API_Socket

Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String

Dim buffSendLength As Long
Dim buffRecieveLength As Long

Dim recieveBufferEmptyRaiseError As Boolean

Dim mySignalEmitter As New SignalEmitter

Private Sub Class_Initialize()
    recieveBufferEmptyRaiseError = True
    Call mySignalEmitter.Initialize(Me)
    buffSendLength = DEFAULTSENDBUFFER
    buffRecieveLength = DEFAULTRECIEVEBUFFER
End Sub

Private Sub InitializeWSA()
    If mint_api_winsock_instances = 0 Then
        If mint_api_winsock_versionrequired_buffer = 0 Then _
            mint_api_winsock_versionrequired_buffer = &H101
        Dim WSA As API_WSADATA
        If API_WSAStartup(mint_api_winsock_versionrequired_buffer, WSA) <> 0 Then _
            throw SystemCallFailureException("WSAStartup failed, while trying to call it in ws2_32.dll")
        
        Call mySignalEmitter.Emit(SIGNAL_WSAINITIALIZE)
    End If
End Sub

Friend Sub Constructor0(Optional AddressFamily As AddressFamily = afINet, Optional SocketType As SocketType = stStream, Optional Protocol As Protocol = pTCP)
    If sSock.Handle <> 0 Then throw InvalidStatusException("Socket already initialized.")
    sSock = Create(AddressFamily, SocketType, Protocol)
    Call mySignalEmitter.Emit(SIGNAL_INITIALIZE)
    Call mySignalEmitter.Emit(SIGNAL_CREATE)
End Sub
Private Function Create(AddressFamily As AddressFamily, SocketType As SocketType, Protocol As Protocol) As API_Socket
    Dim Sock_retVal As Long
    Call InitializeWSA
    
    Sock_retVal = API_Socket(AddressFamily, SocketType, Protocol)
    If Sock_retVal = INVALID_SOCKET Then _
        throw SystemCallFailureException("An error occured when trying to create socket in ws2_32.dll.")
    
    Create.Handle = Sock_retVal
    
    Create.AddressFamily = AddressFamily
    Create.SocketType = SocketType
    Create.Protocol = Protocol
    Create.Port = 0
End Function
Friend Sub SetHandle(pSock As API_Socket)
    sSock.Handle = pSock.Handle
    sSock.AddressFamily = pSock.AddressFamily
    sSock.SocketType = pSock.SocketType
    sSock.Protocol = pSock.Protocol
    sSock.Port = pSock.Port
    sSock.EndPointAddress = pSock.EndPointAddress
End Sub

Private Function GetIPAddressOfEndPoint(EndPoint As EndPoint) As Long
    Dim Addr As Long, HostEntry As API_HostEnt, HostEntryPTR As Long
    Addr = API_INet_addr(EndPoint.IPAddress)
    If (Addr = INADDR_NONE) Then
        '' if not, assume it's a name, resolve it
        HostEntryPTR = API_GetHostByNamePTR(EndPoint.IPAddress)
        If (HostEntryPTR = 0) Then
            Exit Function
        End If
        
        Call API_CopyMemory(ByVal VarPtr(HostEntry), ByVal HostEntryPTR, Len(HostEntry))
        GetIPAddressOfEndPoint = API_VarPtrArray(HostEntry.h_addr_list)
    Else
        '' just return the address
        GetIPAddressOfEndPoint = Addr
    End If
End Function

Public Function Connect(EndPoint As EndPoint) As Socket
    If sSock.Handle = 0 Or sSock.Handle = INVALID_SOCKET Then throw InvalidStatusException("Socket has not been initialized.")
    If (EndPoint.SocketPort < MIN_PORT) Or (EndPoint.SocketPort > MAX_PORT) Then throw OutOfRangeException("Port number must be in range of 0 to 65535.")
    
    Dim SockAddress As API_SOCKADDR
    SockAddress.sin_port = API_htons(EndPoint.SocketPort)
    SockAddress.sin_family = EndPoint.AddressFamily
    SockAddress.sin_addr = GetIPAddressOfEndPoint(EndPoint)
    
    If API_Connect(sSock.Handle, SockAddress, Len(SockAddress)) = SOCKET_ERROR Then _
        throw SystemCallFailureException("Unable to connect to host, An error occured in ws2_32.dll while trying to call connect().")
    'Call baseNetwork.Connect(sSock, baseNetwork.EndPointToAddress(EP))
    
    sSock.EndPointAddress = SockAddress
    sSock.Port = SockAddress.sin_port
    'sSock.AddressFamily = SockAddress.sin_family
    
    Call mySignalEmitter.Emit(SIGNAL_CONNECT)
    Set Connect = Me
End Function

Public Sub Destroy(Optional DestroyHandle As Long = False)
    Call Shutdown(skdBoth)
    If DestroyHandle Then _
        Call CloseSocket
End Sub
Public Sub Shutdown(direction As SocketDirection)
    Call API_Shutdown(sSock.Handle, direction)
    If (direction = skdInput) Or (direction = skdBoth) Then _
        Call mySignalEmitter.Emit(SIGNAL_INPUTSHUTDOWN)
    If (direction = skdOutput) Or (direction = skdBoth) Then _
        Call mySignalEmitter.Emit(SIGNAL_OUTPUTSHUTDOWN)
    Call mySignalEmitter.Emit(SIGNAL_SHUTDOWN)
End Sub
Public Sub CloseSocket()
    Call API_CloseSocket(sSock.Handle)
    Call mySignalEmitter.Emit(SIGNAL_CLOSE)
End Sub
Public Sub Disconnect(Disconnect As Boolean)
    If Not Disconnect Then Exit Sub
    Call Shutdown(skdBoth)
    Call mySignalEmitter.Emit(SIGNAL_DISCONNECT)
End Sub

Public Property Get Handle() As Long
    Handle = sSock.Handle
End Property

Public Property Get Available() As Long
    
End Property

Public Sub Clear()
    
End Sub

Public Property Get IsOpen() As Boolean
    IsOpen = ((sSock.Handle <> 0))
End Property
Public Property Get IsConnect() As Boolean
    IsConnect = IsOpen
End Property

Public Property Get State() As SocketStates
    
End Property


Public Function Bind(EndPoint As EndPoint) As Boolean
    'Bind = baseNetwork.Bind(sSock, baseNetwork.EndPointToAddress(EP))
    Dim Addr As API_SOCKADDR
    Addr.sin_family = EndPoint.AddressFamily
    Addr.sin_port = API_htons(EndPoint.SocketPort)
    
    Select Case LCase$(EndPoint.IPAddress)
        Case "", "any", "inaddr_any", "local", "localhost", "loopback"
            Addr.sin_addr = INADDR_ANY
        Case Else
            Addr.sin_addr = GetIPAddressOfEndPoint(EndPoint)
    End Select
    
    Call rLastError
    Dim retVal As Long
    retVal = API_Bind(sSock.Handle, Addr, Len(Addr))
    
    If retVal = SOCKET_ERROR Then
        'throw SystemCallFailureException("An error occured when try to call bind() in ws2_32.dll.")
        Bind = False
    Else
        sSock.Port = Addr.sin_port
        sSock.EndPointAddress = Addr
        Bind = True
    End If
End Function
Public Sub Listen(Optional BackLog As Long = 10) 'doesnt block process
    Dim retVal As Long
    retVal = API_Listen(sSock.Handle, BackLog)
    
    If retVal = SOCKET_ERROR Then _
        throw SystemCallFailureException("An error occured when trying to call listen() in ws2_32.dll.")
End Sub
'Public Function ConcurrentListen() As Thread
'
'End Function

Friend Sub Accepted()
    
End Sub
Public Function Accept() As Socket
    Dim iSock As API_Socket
    Dim Addr As API_SOCKADDR
    'Sock = baseNetwork.Accept(sSock)
    
    Dim retVal As Long 'It is new socket handle.
    retVal = API_Accept(sSock.Handle, Addr, Len(Addr))
    
    If retVal = SOCKET_ERROR Then
        If IOBlockingState Then
            throw SystemCallFailureException("An error occured when trying to call accept() in ws2_32.dll.")
        Else
            Set Accept = Nothing
            Exit Function
        End If
    End If
    
    iSock.Handle = retVal
    iSock.AddressFamily = Addr.sin_family
    iSock.EndPointAddress = Addr
    iSock.Port = API_ntohs(Addr.sin_port)
    iSock.Protocol = sSock.Protocol
    iSock.SocketType = sSock.SocketType
    
    Dim S As New Socket
    Call S.SetHandle(iSock)
    Call S.Accepted
    Call mySignalEmitter.Emit(SIGNAL_ANSWER)
    Set Accept = S
End Function


Public Property Get IOBlockingState() As Boolean
    
End Property
Public Property Let IOBlockingState(Value As Boolean)
    
End Property



Public Sub Send(Data, Optional Flags As SocketFlags = sflagNone)
    If Not IsOpen Then throw InvalidStatusException("Socket is not connected.")
    Dim BA() As Byte
    BA = mint_get_byte_array_of(Data)
    If ArraySize(BA) = 0 Then Exit Sub
    
    If (API_Send(sSock.Handle, BA(0), ArraySize(BA), Flags) = SOCKET_ERROR) Then _
        throw SystemCallFailureException("An error occured when trying to call send() in ws2_32.dll.")
End Sub

Public Function PeekData(Optional Flags As SocketFlags = sflagNone) As ByteArray
    Set PeekData = ReadData(Flags Or sflagPeek)
End Function
Public Function ReadData(Optional Flags As SocketFlags = sflagNone) As ByteArray
    Dim buff() As Byte, readBytes As Long
    ReDim buff(buffRecieveLength - 1)
    readBytes = API_Recieve(sSock.Handle, buff(0), buffRecieveLength, Flags)
    If (readBytes <= 0) Then
        If recieveBufferEmptyRaiseError Then
            throw BufferIsEmptyException("Socket recieve buffer is empty.")
        Else
            Set ReadData = ByteArray()
            Exit Function
        End If
    End If
    Set ReadData = ByteArray(buff, readBytes)
End Function

Public Property Get RecieveBufferLength() As Long
    RecieveBufferLength = buffRecieveLength
End Property
Public Property Let RecieveBufferLength(Value As Long)
    If Value <= 0 Then throw ZeroNegativeArgumentException("Recieve buffer length can't be zero or negative.")
    buffRecieveLength = Value
End Property
Public Property Get SendBufferLength() As Long
    SendBufferLength = buffSendLength
End Property
Public Property Let SendBufferLength(Value As Long)
    If Value <= 0 Then throw ZeroNegativeArgumentException("Send buffer length can't be zero or negative.")
    buffSendLength = Value
End Property

Public Property Get EndPoint() As EndPoint
    Dim Addr As API_SOCKADDR
    Dim retVal As Long
    
    retVal = API_GetPeerName(sSock.Handle, Addr, Len(Addr))
    If retVal = SOCKET_ERROR Then _
        throw SystemCallFailureException("An error occured when trying to get end point info using getpeername() in ws2_32.dll.")
    
    EndPoint.AddressFamily = Addr.sin_family
    EndPoint.IPAddress = API_INet_ntoa(Addr.sin_addr)
    EndPoint.SocketPort = Addr.sin_port
    EndPoint.Tag = 0
End Property
Public Property Let EndPoint(Value As EndPoint)
    
End Property
Public Property Get IP() As IPAddress
    IP.IPAddress = EndPoint.IPAddress
End Property
Public Property Get Port() As Long
    Port = EndPoint.SocketPort
End Property
Public Property Get SockAddress() As SocketAddress
    Dim EP As EndPoint
    EP = EndPoint
    SockAddress.IPAddress = EP.IPAddress
    SockAddress.SocketPort = EP.SocketPort
End Property


Public Property Get Fragmentation() As Boolean
    
End Property
Public Property Let Fragmentation(Value As Boolean)
    
End Property
Public Property Get TimeToLive() As Long
    
End Property
Public Property Let TimeToLive(Value As Long)
    
End Property





Public Function Clone() As Socket
    
End Function



Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function

Private Sub IClassExceptionManager_setAllExceptionsState(Optional States As Boolean = True)

End Sub

Private Sub IClassExceptionManager_setExceptionState(exp As ExceptionType, Optional State As Boolean = True)

End Sub

Private Sub ITexable_fromByteArray(Value() As Byte)

End Sub

Private Sub ITexable_fromString(Value As String)

End Sub
Private Property Get ITexable_texHeader(texType As mint_TexingType) As String

End Property
Private Function ITexable_toByteArray() As Byte()
    
End Function

Private Function ITexable_toString() As String
    
End Function

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter.ExternEmitter
End Function

Private Property Get ISupportSlot_AutoManage() As Boolean
    ISupportSlot_AutoManage = False
End Property
Private Sub ISupportSlot_ManualManageCall(SlotName As String)
    
End Sub
Private Function ISupportSlot_Methods() As String
    ISupportSlot_Methods = _
        "0" & SLOT_FLUSH & _
        " 0" & SLOT_DISCONNECT & _
        " 0" & SLOT_FLUSH
End Function

Private Function ITargetStream_Available(direction As StreamDirection) As Long
    Dim val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        val = ((Available > 0) And inState)
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        val = val Or (outState)
    End If
    ITargetStream_Available = val
End Function
Private Sub ITargetStream_CloseStream(Optional direction As StreamDirection = 3&)
    If (direction And sdInStream) = sdInStream Then
        inState = False
    End If
    If (direction And sdOutStream) = sdOutStream Then
        outState = False
    End If
    If Not (inState Or outState) Then Call Disconnect(True)
End Sub
Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean
    Dim val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        val = inState
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        val = val Or outState
    End If
    ITargetStream_getState = val
End Function
Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = 3&)
    If (direction And sdInStream) = sdInStream Then
        inState = True
    End If
    If (direction And sdOutStream) = sdOutStream Then
        outState = True
    End If
End Sub
Private Sub ITargetStream_outStream(Data As Variant, Optional Length As Long = -1&)
    If Not outState Then throw InvalidStatusException
    If IsObject(Data) Then
        If TypeOf Data Is ObjectBuffer Then
            Dim objBuffer As ObjectBuffer
            Set objBuffer = Data
            If objBuffer.IsStreamingValue Then
                Call objBuffer.thrownotargs(1)
                If objBuffer.ArgType(0) = vbString Then
                    Dim strFirst As String
                    strFirst = objBuffer.First
                    If strFirst = OBJECTBUFFER_VALUES_STREAMING_CLEAR Then
                        Call Me.Clear
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_NEWLINE Then
                        Call Me.Send(vbCrLf)
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_FLUSH Then
                        'Call Me.Flush
                        Exit Sub
                    Else
                        throw InvalidArgumentTypeException("Streaming Command '" & strFirst & "' Is Not Available For Socket."), Me
                    End If
                Else
                    throw InvalidArgumentTypeException("Streaming Type Is Not Available For Socket.")
                End If
            End If
        End If
    End If
    Call Send(mint_get_byte_array_of(Data, Length))
End Sub
Private Sub ITargetStream_inStream(Data As Variant, Optional Length As Long = -1&)
    
End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    
End Function

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, inState, inState_LOCK, Reserved)
End Sub
Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, outState, outState_LOCK, Reserved)
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
