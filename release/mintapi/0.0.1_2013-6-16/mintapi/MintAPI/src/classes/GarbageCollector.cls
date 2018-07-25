VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GarbageCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Const CLASSID As String = "GarbageCollector"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Public Enum GarbageCollectorStates
    gcCollecting
    gcActive
    gcWaiting
    gcDead
End Enum


Friend Sub Initialize()
    
End Sub


Public Sub Collect(Optional TreatAsIdleTime As Boolean = False)
    Call mint_garbagec_CleanOperations
    Call mint_garbagec_CleanBoundTR
End Sub

Public Sub Register()
    
End Sub
Public Sub Unregister()
    
End Sub

Public Function Percentage() As Long
    
End Function

Public Property Get State() As GarbageCollectorStates
    
End Property

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
