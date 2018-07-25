VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ep As MintAPI.EndPoint
Dim ept As MintAPI.EndPointTag

Public Sub Initialize(IPAddress As String, Port As Long, Optional AddressFamily As AddressFamily = afINet)
    ep.IPAddress = IPAddress
    ep.SocketPort = Port
    ep.AddressFamily = AddressFamily
    ep.Tag = 0
End Sub
Public Sub InitializeCP(cp As CodeParser)
    
End Sub

Public Function GetEndPoint() As MintAPI.EndPoint
    GetEndPoint = ep
End Function
Public Function GetEndPointTag() As MintAPI.EndPointTag
    GetEndPointTag = ept
End Function
