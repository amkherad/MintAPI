VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "ObjectBuffer"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Const SIGNAL_CALLMETHOD As String = "call"
Const SIGNAL_APPEND As String = "append"
Const SIGNAL_RENAME As String = "rename"


Implements ICloneable
'Implements ISupportSignal
'Implements IAliasable


Private Type ObjectBufferArg
    Name As String
    Value As Variant
End Type

Dim Key As String
Dim Arg() As ObjectBufferArg
Dim argCount As Long
Dim argsBuffer As ArgumentList

Dim iSignalEmitter As New SignalEmitter

'Public Enum BufferType
'    bt_EventNamesList
'    bt_hdc
'    bt_Handle
'End Enum

Private Function mySignalEmitter() As SignalEmitter
    If iSignalEmitter Is Nothing Then _
        Set iSignalEmitter = New SignalEmitter
    'Call iSignalEmitter.Initialize(Me)
End Function

Friend Sub Initialize(ByVal Name As String, ParamArray Args() As Variant)
    Dim mx As Long
    On Error GoTo zeroLength
    mx = UBound(Args) + 1
    ReDim Arg(mx - 1)
zeroLength:
    Dim i As Long
    For i = 0 To mx - 1
        If VarType(Args(i)) = VBObject Then
            Set Arg(i).Value = Args(i)
        Else
                Arg(i).Value = Args(i)
        End If
    Next
    argCount = mx
    Key = Name
End Sub
Friend Sub InitializeW(ByVal Name As String, Args() As Variant)
    Dim mx As Long
    On Error GoTo zeroLength
    mx = UBound(Args) + 1
    ReDim Arg(mx - 1)
zeroLength:
    Dim i As Long
    For i = 0 To mx - 1
        If VarType(Args(i)) = VBObject Then
            Set Arg(i).Value = Args(i)
        Else
                Arg(i).Value = Args(i)
        End If
    Next
    argCount = mx
    Key = Name
End Sub
Friend Sub InitializeCopy(Name As String, Args() As ObjectBufferArg)
    Arg = Args
    On Error GoTo zeroLength
    argCount = UBound(Args) - LBound(Args) + 1
zeroLength:
    Key = Name
End Sub


Public Sub CallMethod(ByVal Arguments As ArgumentList)
    If Not iSignalEmitter Is Nothing Then _
        Call mySignalEmitter.Emit(SIGNAL_CALLMETHOD)
End Sub
Public Property Get IsExtraArgument() As Boolean
    IsExtraArgument = (Not (argsBuffer Is Nothing))
End Property
Public Property Get ExtraArguments() As ArgumentList
    Set ExtraArguments = argsBuffer
End Property
Public Property Let ExtraArguments(ByVal Value As ArgumentList)
    Set argsBuffer = Value
End Property
Public Property Set ExtraArguments(ByVal Value As ArgumentList)
    Set argsBuffer = Value
End Property


Friend Function Append(Value) As ObjectBuffer
    ReDim Preserve Arg(argCount)
    If IsObject(Value) Then
        Set Arg(argCount).Value = Value
    Else
            Arg(argCount).Value = Value
    End If
    argCount = argCount + 1
    Set Append = Me
    If Not iSignalEmitter Is Nothing Then _
        Call mySignalEmitter.Emit(SIGNAL_APPEND)
End Function

Friend Sub Clear()
    Erase Arg
    argCount = 0
End Sub
Public Property Get Count() As Long
    Count = argCount
End Property

Friend Sub Rename(ByVal NewName As String)
    Key = NewName
    If Not iSignalEmitter Is Nothing Then _
        Call mySignalEmitter.Emit(SIGNAL_RENAME)
End Sub

Public Property Get Name() As String
    Name = UCase$(Key)
End Property

Friend Sub ThrowNot(Name As String, Count As Long)
    If (Count < argCount) Then throw Exps.InvalidArgumentException
    If (Count > argCount) Then throw Exps.InvalidArgumentException
    If (Name <> Key) Then throw Exps.Exception("Unknown type.")
End Sub
Friend Sub thrownotargs(Count As Long)
    If (Count < argCount) Then throw Exps.InvalidArgumentException
    If (Count > argCount) Then throw Exps.InvalidArgumentException
End Sub

Friend Function ArgType(Index As Long) As VbVarType
    ArgType = VarType(Arg(Index).Value)
End Function
Friend Property Get Item(Index As Long)
    If VarType(Arg(Index).Value) = VBObject Then
        Set Item = Arg(Index).Value
    Else
            Item = Arg(Index).Value
    End If
End Property
Friend Property Let Item(Index As Long, Value)
    If VarType(Value) = VBObject Then
        Set Arg(Index).Value = Value
    Else
            Arg(Index).Value = Value
    End If
End Property
Friend Property Set Item(Index As Long, Value As Object)
    Set Arg(Index).Value = Value
End Property

Friend Function First() As Variant
    If IsObject(Arg(0).Value) Then
        Set First = Arg(0).Value
    Else
            First = Arg(0).Value
    End If
End Function
Friend Function Last() As Variant
    If VarType(Arg(argCount - 1).Value) = VBObject Then
        Set Last = Arg(argCount - 1).Value
    Else
            Last = Arg(argCount - 1).Value
    End If
End Function

'Friend Function SubArguments(StartAt As Long, Optional Length As Long = -1) As ObjectBuffer
'
'End Function

Public Function IsEventNamesList() As Boolean
    IsEventNamesList = (Key = OBJECTBUFFER_EVENTNAMES_LIST)
End Function
Public Function IsHandle() As Boolean
    IsHandle = (Key = OBJECTBUFFER_HANDLE)
End Function
Public Function IsHDC() As Boolean
    IsHDC = (Key = OBJECTBUFFER_DISPLAY_CONTEXT)
End Function
Public Function IsStreamingValue() As Boolean
    IsStreamingValue = (Key = OBJECTBUFFER_STREAMINGVALUE)
End Function
Public Function IsTexedObject() As Boolean
    IsTexedObject = (Key = OBJECTBUFFER_TEXEDOBJECT)
End Function
Public Function IsInherit() As Boolean
    IsInherit = (Key = OBJECTBUFFER_INHERIT)
End Function
Public Function IsRenderTarget() As Boolean
    IsRenderTarget = (Key = OBJECTBUFFER_RENDERTARGET)
End Function
Public Function IsMintAPILocalType() As Boolean
    IsMintAPILocalType = (Key = OBJECTBUFFER_MINTLOCALTYPE)
End Function

Public Function IsSignal() As Boolean
    IsSignal = (Key = OBJECTBUFFER_SIGNAL)
End Function
Public Function IsSlot() As Boolean
    IsSlot = (Key = OBJECTBUFFER_SLOT)
End Function

Public Function Clone() As ObjectBuffer
    Dim cOB As New ObjectBuffer
    Call cOB.InitializeCopy(Key, Arg)
    Set Clone = cOB
End Function

Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter.ExternEmitter
End Function
