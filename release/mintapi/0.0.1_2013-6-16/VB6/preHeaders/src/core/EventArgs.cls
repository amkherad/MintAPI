VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IEventArgs

Dim obj As Object

Friend Sub Initialize(targetObject As Object)
    If Not obj Is Nothing Then throw InvalidCallException("EventArgs initialized before.")
    Set obj = targetObject
End Sub

Public Property Get Sender() As Object: Set Sender = obj: End Property
Private Property Get IEventArgs_Sender() As Object: Set IEventArgs_Sender = obj: End Property

