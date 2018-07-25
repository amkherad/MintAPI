VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zakkLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const SIGNAL_LOG As String = "log"
Const SIGNAL_NOTIFY As String = "notify"

Implements ISupportSignal


Dim mySignalEmitter As New SignalEmitter

Private Sub Class_Initialize()
    Call mySignalEmitter.SetParent(Me)
End Sub

Public Sub Log(Message As String)
    
    Call mySignalEmitter.Emit(SIGNAL_LOG)
End Sub
Public Sub Notify(Arg)
    
    Call mySignalEmitter.Emit(SIGNAL_NOTIFY)
End Sub


Public Property Get Format() As String
    Connect
End Property
Public Property Let Format(Value As String)
    
End Property

Public Sub Connect(Signal, Slot)
    Call MintAPI.Connect(Signal, Me, Slot)
End Sub
Public Sub Disconnect(Signal, Slot)
    Call MintAPI.Disconnect(Signal, Me, Slot)
End Sub

Public Sub Emit(Signal)
    Call mySignalEmitter.Emit(Signal)
End Sub
Public Function SignalEmitter() As SignalEmitter
    Set SignalEmitter = mySignalEmitter.ExternEmitter
End Function

Public Sub appendLogTargets(logTarget As ITargetStream)
    
End Sub
Public Sub removeLogTargets(logTarget As ITargetStream)
    
End Sub
Public Function getLogTargets(Index As Long) As ITargetStream
    
End Function
Public Property Get CountLogTargets() As Long
    
End Property


Private Function ISupportSignal_Emitter() As MintAPI.SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter.ExternEmitter
End Function
