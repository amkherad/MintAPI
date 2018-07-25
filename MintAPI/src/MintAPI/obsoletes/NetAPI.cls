VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "NetAPI"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable

Private Const API_RAS95_MaxEntryName = 256
Private Const API_RAS95_MaxDeviceType = 16
Private Const API_RAS95_MaxDeviceName = 32
Private Const API_NET_CONNECTED_VALUE = &H2000



Private Declare Function API_RasEnumConnections Lib "RasApi32" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function API_RasGetConnectStatus Lib "RasApi32" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long


Private Type API_RASCONN95
    dwSize As Long
    hRasCon As Long
    szEntryName(API_RAS95_MaxEntryName) As Byte
    szDeviceType(API_RAS95_MaxDeviceType) As Byte
    szDeviceName(API_RAS95_MaxDeviceName) As Byte
End Type
Private Type API_RASCONNSTATUS95
    dwSize As Long
    RasConnState As Long
    dwError As Long
    szDeviceType(API_RAS95_MaxDeviceType) As Byte
    szDeviceName(API_RAS95_MaxDeviceName) As Byte
End Type

Friend Sub Initialize(Optional Arguments)
    
End Sub


Public Function CreateRemoteAPI(Optional ServerName As String) As API
    If IsMissing(ServerName) Then
        Set CreateRemoteAPI = CreateObject(APPLICATIONID & ".API")
    Else
        Set CreateRemoteAPI = CreateObject(APPLICATIONID & ".API", ServerName)
    End If
End Function


Public Function InternetConnectionStatus() As Boolean
    Dim TRasCon(255) As API_RASCONN95
    Dim lg As Long
    Dim lpcon As Long
    Dim retVal As Long
    Dim Tstatus As API_RASCONNSTATUS95
    '
    TRasCon(0).dwSize = 412
    lg = 256 * TRasCon(0).dwSize
    '
    retVal = API_RasEnumConnections(TRasCon(0), lg, lpcon)

    If retVal <> 0 Then
        throw Exps.SystemCallFailureException("Error In Method RasEnumConnectionsA() In RasApi32.dll.")
        Exit Function
    End If

    Tstatus.dwSize = 160
    retVal = API_RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)

    InternetConnectionStatus = (Tstatus.RasConnState = API_NET_CONNECTED_VALUE)
End Function
Public Function CheckLoopbackState() As Boolean
    
End Function
Public Function Ping(DNS As String) As Boolean
    
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
