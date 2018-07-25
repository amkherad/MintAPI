VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim cp As CodeParser

Friend Sub Initialize(cp1 As CodeParser)
    Set cp = cp1
End Sub

'socket [AddressFamily] [SocketType] [Protocol]
Public Function funcSocket(Execute As Boolean)
    If Not Execute Then Exit Function
    Dim arg0 As String, arg1 As String, arg2 As String, arg3 As String
    Call cp.ensurege(1, 4)
    Dim args As Long
    arg0 = cp.GetArg(0)
    args = cp.CountArguments
    
    Dim AF As AddressFamily
    Dim ST As SocketType
    Dim P As Protocol
    
    AF = afINet
    ST = stStream
    P = pTCP
    
    If args >= 3 Then _
        arg2 = outQuotPath(cp.GetArg(2))
    If args >= 2 Then _
        arg1 = outQuotPath(cp.GetArg(1))
    If args >= 1 Then _
        arg0 = outQuotPath(cp.GetArg(0))
    
    Dim Sock As Socket
    
    Set Sock = Socket(AF, ST, P)
    
    If Not CheckVar(arg0) Then Exit Function
        
    Set stack(arg0) = Sock
    
    funcSocket = True
End Function
'sockVar, EndPoint
Public Function funcConnect(Execute As Boolean)
    If Not Execute Then Exit Function
    Dim arg0 As String, arg1 As String
    Call cp.ensure(2)
    arg0 = outQuotPath(cp.GetArg(0))
    arg1 = outQuotPath(cp.GetArg(1))
    
    Dim Sock As Socket, ep As EndPoint
    
    If Not stack.Exists(arg0) Then Exit Function
    
    If Not TypeOf stack(arg0) Is Socket Then _
        throw Exceptions.InvalidArgumentTypeException("Argument Socket is not a Socket class.")
    Set Sock = stack(arg0)
    
    If Not stack.Exists(arg1) Then Exit Function
    If Not TypeOf stack(arg1) Is Socket Then _
        throw Exceptions.InvalidArgumentTypeException("Argument EndPoint is not a EndPoint class.")
    Set ep = stack(arg1)
    
    Call Sock.Connect(ep.GetEndPoint)
    
    funcConnect = True
End Function
'sockVar,[AddressFamily],[asdf]
Public Function funcClose(Execute As Boolean)
    If Not Execute Then Exit Function
    Call cp.ensure(1)
    
    
    funcClose = True
End Function
'sockVar,[AddressFamily],[asdf]
Public Function funcShutdown(Execute As Boolean)
    If Not Execute Then Exit Function
    
    funcShutdown = True
End Function
Public Function funcSend(Execute As Boolean)
    If Not Execute Then Exit Function
    
    Dim Sock As Socket
    
    funcSend = True
End Function
Public Function funcRecieve(Execute As Boolean)
    If Not Execute Then Exit Function
    
    funcRecieve = True
End Function
Public Function funcBind(Execute As Boolean)
    If Not Execute Then Exit Function
    
    funcBind = True
End Function
Public Function funcListen(Execute As Boolean)
    If Not Execute Then Exit Function
    
    funcListen = True
End Function
Public Function funcDisconnect(Execute As Boolean)
    If Not Execute Then Exit Function
    
    funcDisconnect = True
End Function
