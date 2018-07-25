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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Socket"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Const SIGNAL_SENDING As String = "sending"
Const SIGNAL_SEND As String = "send"
Const SIGNAL_RECIEVE As String = "recieve"
Const SIGNAL_LISTENING As String = "listening"
Const SIGNAL_CONNECT As String = "connect"
Const SIGNAL_ANSWER As String = "answer"
Const SIGNAL_CREATE As String = "create"
Const SIGNAL_CLOSE As String = "close"

Const SLOT_FLUSH As String = "flush"
Const SLOT_DISCONNECT As String = "disconnect"
Const SLOT_CLEAR As String = "clear"


Implements ITargetStream
Implements IClassClonable
Implements IClassTexer
Implements ISupportSignal
Implements ISupportSlot
Implements IAliasable

Public Enum AddressFamily
    afUnknown = API_afUnknown
    afUnSpec = API_afUnSpec
    afUnspecified = API_afUnspecified
    afUnix = API_afUnix
    afINet = API_afINet
    afInterNetwork = API_afInterNetwork
    afImpLink = API_afImpLink
    afPUP = API_afPUP
    afChaos = API_afChaos
    afIpx = API_afIpx
    afNS = API_afNS
    afIso = API_afIso
    afOsi = API_afOsi
    afEcma = API_afEcma
    afDataKit = API_afDataKit
    afCCitt = API_afCCitt
    afSna = API_afSna
    afDecNet = API_afDecNet
    afDLI = API_afDLI
    afLAT = API_afLAT
    afHyLink = API_afHyLink
    afHyperChannel = API_afHyperChannel
    afAppleTalk = API_afAppleTalk
    afNetBios = API_afNetBios
    afVoiceView = API_afVoiceView
    afFirefox = API_afFirefox
    afUnknown1 = API_afUnknown1 'Somebody is using this!
    afBAN = API_afBAN
    afBanyan = API_afBanyan
    afAtm = API_afAtm
    afInterNetworkV6 = API_afInterNetworkV6
    afCluster = API_afCluster
    afIeee12844 = API_afIeee12844
    afIrda = API_afIrda
    afNetworkDesigners = API_afNetworkDesigners
    
    afMAX = AF_MAX
End Enum
Public Enum SocketType
    stUnknown = API_stUnknown
    
    stStream = API_stStream
    stTCP = API_stStream
    
    stDataGram = API_stDataGram
    stUDP = stDataGram
    
    stRAW = API_stRAW
    stReliablyDeliveredMessage = API_stReliablyDeliveredMessage
    stSequencedPacket = API_stSequencedPacket
End Enum
Public Enum Protocol
    pUNKNOWN = API_pUNKNOWN
    pIPV6HOPBYHOPOPTIONS = API_pIPV6HOPBYHOPOPTIONS
    pUnspecified = API_pUnspecified
    pIP = API_pIP
    pICMP = API_pICMP
    pIGMP = API_pIGMP
    pGGP = API_pGGP
    pIPV4 = API_pIPV4
    pTCP = API_pTCP
    pPUP = API_pPUP
    pUDP = API_pUDP
    pIDP = API_pIDP
    pIPV6 = API_pIPV6
    pIPV6ROUTINGHEADER = API_pIPV6ROUTINGHEADER
    pIPV6FRAGMENTHEADER = API_pIPV6FRAGMENTHEADER
    pIPSECENCAPSULATINGSECURITYPAYLOAD = API_pIPSECENCAPSULATINGSECURITYPAYLOAD
    pIPSECAUTHENTICATIONHEADER = API_pIPSECAUTHENTICATIONHEADER
    pICMPV6 = API_pICMPV6
    pIPV6NONEXTHEADER = API_pIPV6NONEXTHEADER
    pIPV6DESTINATIONOPTIONS = API_pIPV6DESTINATIONOPTIONS
    pND = API_pND
    pRAW = API_pRAW
    pIPX = API_pIPX
    
    pSPX = API_pSPX
    pSPXII = API_pSPXII
    pMAX = API_pMAX
End Enum

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

Dim mySignalEmitter As New SignalEmitter

Private Sub Class_Initialize()
    Call mySignalEmitter.Initialize(Me)
End Sub

Private Function GetAPIEndPoint(EndPoint As EndPoint) As API_EndPoint
    Call API_CopyMemory(ByVal VarPtr(GetAPIEndPoint), ByVal VarPtr(EndPoint), Len(GetAPIEndPoint))
End Function

Friend Sub Initialize(Optional AddressFamily As AddressFamily = afUnspecified, Optional SocketType As SocketType = stStream, Optional Protocol As Protocol = pUnspecified)
    If sSock.Handle <> 0 Then throw InvalidStatusException("Socket already initialized.")
    sSock = baseNetwork.Create(AddressFamily, SocketType, Protocol)
End Sub
Public Sub Create(Optional AddressFamily As AddressFamily = afUnspecified, Optional SocketType As SocketType = stStream, Optional Protocol As Protocol = pUnspecified)
    If sSock.Handle <> 0 Then throw InvalidStatusException("Socket already initialized.")
    sSock = baseNetwork.Create(AddressFamily, SocketType, Protocol)
End Sub
Public Sub Destroy(Optional DestroyHandle As Long = False)
    
End Sub

Public Property Get Available() As Long
    
End Property

Public Sub Clear()
    
End Sub

Public Function Connect(EndPoint As EndPoint) As Socket
    Dim EP As API_EndPoint
    EP = GetAPIEndPoint(EndPoint)
    Call baseNetwork.Connect(sSock, baseNetwork.EndPointToAddress(EP))
    Set Connect = Me
End Function
Public Sub Disconnect(Disconnect As Boolean)
    If Not Disconnect Then Exit Sub
    
End Sub

Public Property Get IsOpen() As Boolean
    
End Property
Public Property Get IsConnect() As Boolean
    IsConnect = IsOpen
End Property


Public Function Bind(EndPoint As EndPoint) As Boolean
    Dim EP As API_EndPoint
    EP = GetAPIEndPoint(EndPoint)
    Bind = baseNetwork.Bind(sSock, baseNetwork.EndPointToAddress(EP))
End Function
Public Function Listen(BackLog As Long) As EndPoint
    Call baseNetwork.Listen(sSock, BackLog)
End Function
Public Function ConcurrentListen() As Thread
    
End Function

Friend Sub Asnwered()
    
End Sub
Public Function Answer() As Socket
    Dim Sock As API_Socket
    Sock = baseNetwork.Accept(sSock)
    Dim s As New Socket
    Call s.Initialize(Sock.AddressFamily, Sock.SocketType, Sock.Protocol)
    Call s.Asnwered
    Set Answer = s
End Function

Public Sub Send(Data)
    If Not IsOpen Then throw InvalidStatusException("Socket is not connected.")
    Dim ba() As Byte
    ba = mint_get_byte_array_of(Data)
    
    Call baseNetwork.Send(sSock, ba, 0)
End Sub

Public Function PeekData() As ByteArray
    
End Function
Public Function ReadData() As ByteArray
    
End Function


Public Property Get EndPoint() As EndPoint
    
End Property
Public Property Let EndPoint(Value As EndPoint)
    
End Property

Public Property Get IP() As IPAddress
    'IP.IPAddress = sSock.EndPointAddress
End Property
Public Property Get Port() As Long
    Port = sSock.Port
End Property
Public Property Get SockAddress() As SocketAddress
    'SockAddress.IPAddress = sSock.EndPointAddress
    SockAddress.SocketPort = sSock.Port
End Property


Public Property Get Fragmentation() As Boolean
    
End Property
Public Property Let Fragmentation(Value As Boolean)
    
End Property
Public Property Get TimeToLive() As Long
    
End Property
Public Property Let TimeToLive(Value As Long)
    
End Property


Public Function ValidateSocketAddress(SockAddress As SocketAddress) As Boolean
    ValidateSocketAddress = baseNetwork.ValidateSockAddress(baseNetwork.NetAddressToAddress(SockAddress.IPAddress), SockAddress.SocketPort)
End Function
Public Function ValidateIPAddress(IPAddressV As IPAddress) As Boolean
    ValidateIPAddress = baseNetwork.ValidateIPAddress(baseNetwork.NetAddressToAddress(IPAddressV.IPAddress))
End Function



Private Function IClassClonable_Clone() As Object

End Function

Private Sub IClassExceptionManager_setAllExceptionsState(Optional States As Boolean = True)

End Sub

Private Sub IClassExceptionManager_setExceptionState(exp As ExceptionType, Optional State As Boolean = True)

End Sub

Private Sub IClassTexer_fromByteArray(Value() As Byte)

End Sub

Private Sub IClassTexer_fromString(Value As String)

End Sub
Private Property Get IClassTexer_texHeader(texType As mint_TexingType) As String

End Property
Private Function IClassTexer_toByteArray() As Byte()
    
End Function

Private Function IClassTexer_toString() As String
    
End Function

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter
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
    
End Function

Private Sub ITargetStream_CloseStream(Optional direction As StreamDirection = 3&)
    Call Disconnect(True)
End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean
    
End Function


Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = 3&)
    
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
                        throw InvalidArgumentTypeException("Streaming Command '" & strFirst & "' Is Not Available For File."), Me
                    End If
                Else
                    throw InvalidArgumentTypeException("Streaming Type Is Not Available For File.")
                End If
            End If
        End If
    End If
    Call Me.Send(mint_get_byte_array_of(Data, Length))
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
