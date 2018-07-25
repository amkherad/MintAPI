VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Operation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Const CLASSID As String = "Operation"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

Const SIGNAL_STATECHANGED As String = "statechanged"

'Implements IAliasable
'Implements ISupportSignal

Dim myKey As String
Dim Args As ArgumentList
Dim stt As Boolean

Dim read_only_state As Boolean
Dim is_abstract As Boolean

Dim mySignalEmitter As New SignalEmitter


Friend Sub Initialize(Key As String, State As Boolean, Optional Arguments As ArgumentList = Nothing, Optional ReadOnly As Boolean = False)
    myKey = Key
    If Arguments Is Nothing Then
        Set Args = New ArgumentList
    Else
        Set Args = Arguments
    End If
    stt = State
    read_only_state = ReadOnly
End Sub
Friend Sub MakeAbstractOperation()
'this will mark operation as a non existing operation that
'will be registered with the first modification.
    is_abstract = True
End Sub

Private Sub CheckAbstraction()
    If is_abstract Then
        Call mod_application.AddOperation(Key, stt, Args, read_only_state)
        is_abstract = False
    End If
End Sub


Public Property Get Key() As String
    Key = myKey
End Property

Public Sub Cancel()
    If read_only_state Then throw InvalidStatusException("Operation is readonly.")
    stt = False
    Call CheckAbstraction
    Call mySignalEmitter.Emit(SIGNAL_STATECHANGED)
End Sub
Public Sub Continue()
    If read_only_state Then throw InvalidStatusException("Operation is readonly.")
    stt = True
    Call CheckAbstraction
    Call mySignalEmitter.Emit(SIGNAL_STATECHANGED)
End Sub

Public Property Get State() As Boolean
    State = stt
End Property
Public Property Let State(Value As Boolean)
    If read_only_state Then _
        If Value <> stt Then _
            throw InvalidStatusException("Operation is readonly.")
    stt = Value
    Call CheckAbstraction
    Call mySignalEmitter.Emit(SIGNAL_STATECHANGED)
End Property



Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter
End Function
