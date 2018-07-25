VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Headers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum RequestType
    rtGET
    rtCONNECT
    rtDELETE
    rtAPPEND
End Enum
Public Enum RequestProtocol
    rpHTTP
    rpHTTPS
    rpFTP
End Enum

Friend Sub Initialize(s As netSock)
    
End Sub

Public Sub Send(Optional Force As Boolean = False)
    
End Sub

Public Sub Clear()
    
End Sub
Public Sub Add()
    
End Sub

Public Property Get Request() As String
    
End Property

'Public Property Get Length() As Long
'
'End Property
Public Sub addHeader(strHeaderName As String, strHeaderValue As String)
    
End Sub
Public Sub removeHeader(strHeaderName As String)
    
End Sub

Public Property Get Header(strHeaderName As String) As String
    
End Property
Public Property Let Header(strHeaderName As String, strHeaderValue As String)
    
End Property

Public Sub setHeader(strHeaderName As String, strHeaderValue As String)
    
End Sub
Public Function Header(strHeaderName As String) As String
    
End Function

Public Function Headers() As Vector
    
End Function


Public Property Get RequestType() As RequestType
    
End Property
Public Property Let RequestType(Value As RequestType)
    
End Property
Public Property Get RequestTypeString() As String
    
End Property
Public Property Let RequestTypeString(Value As String)
    
End Property

Public Property Get RequestProtocol() As RequestProtocol
    
End Property
Public Property Let RequestProtocol(Value As RequestProtocol)
    
End Property
Public Property Get RequestProtocolString() As String
    
End Property
Public Property Let RequerstProtocolString(Value As String)
    
End Property
