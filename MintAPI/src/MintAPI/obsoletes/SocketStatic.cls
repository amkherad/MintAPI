VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SocketStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "SocketStatic"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

Private Const INADDR_NONE As Long = -1

Private Declare Function API_GetHostByNamePTR Lib "ws2_32" Alias "gethostbyname" (Name As String) As Long  'API_HostEnt 'ALSO Name is char far*
Private Declare Function API_GetHostByName Lib "ws2_32" Alias "gethostbyname" (Name As String) As API_HostEnt 'ALSO Name is char far*
Private Declare Function API_INet_addr Lib "ws2_32" Alias "inet_addr" (ByVal CP As String) As Long
'Private Declare Function API_INet_ntoa Lib "ws2_32" Alias "inet_ntoa" (ByRef lIn As Long) As String


Public Function Create(Optional AddressFamily As AddressFamily = afINet, Optional SocketType As SocketType = stStream, Optional Protocol As Protocol = pUnspecified) As Socket
    Dim S As New Socket
    Call S.Constructor0(AddressFamily, SocketType, Protocol)
    Set Create = S
End Function

Public Function NewLine() As String
    NewLine = vbCrLf
End Function

Public Function ValidateSocketAddress(SockAddress As SocketAddress) As Boolean
    'ValidateSocketAddress = baseNetwork.ValidateSockAddress(baseNetwork.NetAddressToAddress(SockAddress.IPAddress), SockAddress.SocketPort)
End Function
Public Function ValidateIPAddress(IPAddressV As IPAddress) As Boolean
    'ValidateIPAddress = baseNetwork.ValidateIPAddress(baseNetwork.NetAddressToAddress(IPAddressV.IPAddress))
End Function

Public Function GetEndPointByDomain(HostName As String) As EndPoint
    Dim Addr As Long, HostEntry As API_HostEnt, HostEntryPTR As Long
    Dim EP As EndPoint
    
    If (HostName Like "*.*.*.*:*") Then '127.0.0.1:80
        EP.IPAddress = Left(HostName, InStr(1, HostName, ":") - 1)
        EP.SocketPort = val(Right(HostName, Len(HostName) - InStrRev(HostName, ":")))
        Addr = API_INet_addr(EP.IPAddress)
    Else
        EP.IPAddress = HostName
        Addr = API_INet_addr(EP.IPAddress)
    End If
    
    If (Addr = INADDR_NONE) Then
        '' if not, assume it's a name, resolve it
        'If (HostName Like "*:*") Then 'localhost:80
        '    EP.IPAddress = Left(HostName, InStr(1, HostName, ":") - 1)
        '    EP.SocketPort = val(Right(HostName, Len(HostName) - InStrRev(HostName, ":")))
        '    HostEntryPTR = API_GetHostByNamePTR(EP.IPAddress)
        'Else
            HostEntryPTR = API_GetHostByNamePTR(HostName)
        'End If
        If (HostEntryPTR = 0) Then _
            throw SystemCallFailureException("Unable to translate domain by gethostbyname() in ws2_32.dll.")
        
        Call API_CopyMemory(ByVal VarPtr(HostEntry), ByVal HostEntryPTR, Len(HostEntry))
        
        EP.IPAddress = HostEntry.h_addr_list(0)
        EP.AddressFamily = HostEntry.h_addrtype
        Dim EPT As EndPointTag
        
        EPT.AddressType = HostEntry.h_addrtype
        EPT.Name = HostEntry.h_name
        
        EPT.AliasListCount = ArraySize(HostEntry.h_aliases)
        EPT.AliasList = HostEntry.h_aliases
        
        EPT.AddressListCount = ArraySize(HostEntry.h_addr_list)
        EPT.AddressList = HostEntry.h_addr_list
        
        EPT.DCTag = 0
        
        EP.Tag = VarPtr(EPT)
    'Else
    '    EP.IPAddress = API_INet_ntoa(Addr)
    End If
    GetEndPointByDomain = EP
End Function

Public Function GetEndPointTag(EndPoint As EndPoint) As EndPointTag
    Dim EPT As EndPointTag
    If EndPoint.Tag = 0 Then
        throw InvalidStatusException("End point tag is null.")
    ElseIf API_IsBadCodePtr(EndPoint.Tag) Then
        throw Exception("Bad code ptr.")
    End If
    Call API_CopyMemory(ByVal VarPtr(EPT), ByVal EndPoint.Tag, Len(EPT))
End Function
Public Sub SetEndPointTag(EndPoint As EndPoint, EndPointTag As EndPointTag)
    'must reserve dynamic memory, because EndPointTag itself is in the stack and cannot be pointed forever.
    'EndPoint.Tag = VarPtr(EndPointTag)
End Sub

