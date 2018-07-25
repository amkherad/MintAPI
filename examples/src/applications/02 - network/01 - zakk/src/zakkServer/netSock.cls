VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "netSock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISupportSignal

Dim sock As Socket

Private Sub Class_Initialize()
    Set sock = Socket(afINet, stStream, pTCP)
End Sub

Friend Sub Initialize()
    
End Sub

Private Function ISupportSignal_Emitter() As MintAPI.SignalEmitter
    Dim ISS As ISupportSignal
    Set ISS = sock
    Set ISupportSignal_Emitter = ISS.Emitter.ExternEmitter
End Function
