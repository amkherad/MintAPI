VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"


'Ideas from Me , Qt , Java

Option Base 0
Option Explicit
Const CLASSID As String = "fclass"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

Private Declare Function API_IsBadCodePtr Lib "kernel32" Alias "IsBadCodePtr" (ByVal lpfn As Long) As Long


Implements IAliasable


Dim Obj As Object

Friend Sub Initialize(Optional Object As Object)
    Set Obj = Object
End Sub
Friend Sub InitializeFromHandle(Optional ObjHandle As Long = 0)
    If ObjHandle <> 0 Then
        Dim retObj As Object
        Call API_CopyMemory(ByVal ObjPtr(retObj), ByVal ObjHandle, 4)
        
        If API_IsBadCodePtr(ObjHandle) Then throw InvalidHandleException
        On Error GoTo nonObject
        Set Obj = retObj
nonObject:
    End If
End Sub

Public Sub SetObject(Object As Object)
    Set Obj = Object
End Sub

Private Sub CheckObject()
    If Obj Is Nothing Then throw InvalidStatusException("Object Not Set.")
End Sub

Public Property Get Name() As String
CheckObject
    If Not Obj Is Nothing Then Name = TypeName(Obj)
End Property

Public Function supporttexing() As Boolean
CheckObject
    If Obj Is Nothing Then Exit Function
    supporttexing = (TypeOf Obj Is ITexable)
End Function

Public Function constData() As ByteArray
CheckObject
    If TypeOf Obj Is ITexable Then
        Dim objTexer As ITexable
        Set objTexer = Obj
        Set constData = ByteArray(objTexer.toByteArray)
    Else
        throw Exception("Unknown Class Type.")
    End If
End Function
Public Function ToString() As String
CheckObject
    If TypeOf Obj Is ITexable Then
        Dim objTexer As ITexable
        Set objTexer = Obj
        ToString = objTexer.ToString
    Else
        throw Exception("Unknown Class Type.")
    End If
End Function
Public Function toByteArray() As ByteArray
CheckObject
    Set toByteArray = constData()
End Function
Public Function getclass() As Object
CheckObject
    Set getclass = Obj
End Function

Public Sub out(Data, Optional Length As Long = -1)
CheckObject
    If TypeOf Obj Is ITargetStream Then
        mint_application.out Obj, Data, Length
    Else
        throw Exception("Unknown Class Type.")
    End If
End Sub
Public Sub inp(Data, Optional Length As Long = -1)
CheckObject
    If TypeOf Obj Is ITargetStream Then
        mint_application.inp Obj, Data, Length
    Else
        throw Exception("Unknown Class Type.")
    End If
End Sub
Public Sub lout(ParamArray Data() As Variant)
CheckObject
    Dim objTargetStream As ITargetStream
    If Not TypeOf Obj Is ITargetStream Then throw InvalidArgumentTypeException
    Set objTargetStream = Obj
    If Not objTargetStream.getState(sdOutStream) Then throw InvalidStatusException
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Call objTargetStream.outStream(Data(i))
    Next
zeroLength:
End Sub
Public Sub linp(ParamArray Data() As Variant)
CheckObject
    Dim objTargetStream As ITargetStream
    If Not TypeOf Obj Is ITargetStream Then throw InvalidArgumentTypeException
    Set objTargetStream = Obj
    If Not objTargetStream.getState(sdInStream) Then throw InvalidStatusException
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Call objTargetStream.inStream(Data(i))
    Next
zeroLength:
End Sub

Public Sub SetStreamState(Dir As StreamDirection, Optional State As Boolean = True)
CheckObject
    If TypeOf Obj Is ITargetStream Then
        Dim Stream As ITargetStream
        Set Stream = Obj
        If (Dir And sdInStream) = sdInStream Then
            Call Stream.setInState(True)
        End If
        If (Dir And sdOutStream) = sdOutStream Then
            Call Stream.setOutState(True)
        End If
    Else
        throw Exception("Unknown Class Type.")
    End If
End Sub

Public Function getStreamState(Dir As StreamDirection) As Boolean
CheckObject
    If TypeOf Obj Is ITargetStream Then
        Dim Stream As ITargetStream
        Set Stream = Obj
        getStreamState = Stream.getState(Dir)
    Else
        throw Exception("Unknown Class Type.")
    End If
End Function

Public Sub registerObject()
CheckObject
    
End Sub
Public Sub unregisterObject()
CheckObject
    
End Sub

Public Sub boundTranslation(targetProperty As String, Translation As String)
CheckObject
    If Obj Is Nothing Then throw InvalidStatusException
    Call boundtranslationto(Translation, Obj, targetProperty, True)
End Sub
Public Sub unboundTranslation(targetProperty As String)
CheckObject
    If Obj Is Nothing Then throw InvalidStatusException
    Call unboundtranslationto(Obj, targetProperty)
End Sub

Public Sub WriteToFile(File As File)
CheckObject
    If TypeOf Obj Is ITargetStream Or TypeOf Obj Is ITexable Or TypeOf Obj Is IData Then
        mint_application.out File.EnsureOpen, Obj
    Else
        throw Exception("Unknown Class Type.")
    End If
End Sub
Public Sub WriteToStream(targetStream As ITargetStream)
CheckObject
    If TypeOf Obj Is ITargetStream Or TypeOf Obj Is ITexable Or TypeOf Obj Is IData Then
        mint_application.out File.EnsureOpen, Obj
    Else
        throw Exception("Unknown Class Type.")
    End If
End Sub
Public Function ReadFromFile(File As File) As Object
CheckObject
    If TypeOf Obj Is ITargetStream Or TypeOf Obj Is ITexable Or TypeOf Obj Is IData Then
        mint_application.inp File.EnsureOpen, Obj
    Else
        throw Exception("Unknown Class Type.")
    End If
End Function
Public Function ReadFromStream(targetStream As ITargetStream) As Object
CheckObject
    If TypeOf Obj Is ITargetStream Or TypeOf Obj Is ITexable Or TypeOf Obj Is IData Then
        mint_application.inp File.EnsureOpen, Obj
    Else
        throw Exception("Unknown Class Type.")
    End If
End Function

Public Function haveSlot() As Boolean
CheckObject
    If TypeOf Obj Is ISupportSlot Then _
        haveSlot = True
End Function
Public Function haveSignal() As Boolean
CheckObject
    If TypeOf Obj Is ISupportSignal Then _
        haveSignal = True
End Function
Public Sub Emit(Signal)
CheckObject
    If Not TypeOf Obj Is ISupportSignal Then _
        throw InvalidVarTypeException
    Dim SignalObject As ISupportSignal
    Set SignalObject = Obj
    Call SignalObject.Emitter.Emit(Signal)
End Sub

Public Function Clone() As Object
CheckObject
    If TypeOf Obj Is ICloneable Then
        Dim tclonableObj As ICloneable
        Set tclonableObj = Obj
        Set Clone = tclonableObj.Clone
    Else
        throw Exception("Unknown Class Type.")
    End If
End Function
Public Function clonable() As Boolean
CheckObject
    clonable = (TypeOf Obj Is ICloneable)
End Function
Public Function streamable() As Boolean
CheckObject
    streamable = (TypeOf Obj Is ITargetStream)
End Function

Public Function objdefault(objName As String) As Object 'Only MintAPI classes are accepted.
CheckObject
    Set objdefault = CreateObject(APPLICATIONID & "." & objName)
End Function
Public Function Default() As Object 'Only MintAPI classes are accepted.
CheckObject
    If Obj Is Nothing Then throw InvalidCallException
    Set Default = CreateObject(APPLICATIONID & "." & TypeName(Obj))
End Function

Public Sub BlockSignals()
CheckObject
    If Not TypeOf Obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = Obj
    Call ISS.Emitter.BlockSignals
End Sub
Public Sub UnblockSignals()
CheckObject
    If Not TypeOf Obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = Obj
    Call ISS.Emitter.UnblockSignals
End Sub
Public Sub BlockSignal(SignalName As String)
CheckObject
    If Not TypeOf Obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = Obj
    Call ISS.Emitter.BlockSignal(SignalName)
End Sub
Public Sub UnblockSignal(SignalName As String)
CheckObject
    If Not TypeOf Obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = Obj
    Call ISS.Emitter.UnblockSignal(SignalName)
End Sub
Public Sub SetSignalState(SignalName As String, State As Boolean)
CheckObject
    If Not TypeOf Obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = Obj
    Call ISS.Emitter.SetSignalState(SignalName, State)
End Sub
Friend Function SignalEmitterState() As Boolean
CheckObject
    If Not TypeOf Obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = Obj
    Call ISS.Emitter.SignalEmitterState
End Function
Friend Function SignalState(SignalName As String) As Boolean
CheckObject
    If Not TypeOf Obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = Obj
    Call ISS.Emitter.SignalState(SignalName)
End Function

Public Property Get Alias() As String
CheckObject
    If Not TypeOf Obj Is IAliasable Then throw InvalidStatusException
    Dim IA As IAliasable
    Set IA = Obj
    Alias = IA.Alias
End Property
Public Function aliasable() As Boolean
CheckObject
    If Not TypeOf Obj Is IAliasable Then Exit Function
    aliasable = True
End Function

'availableData bool
'availableDataLength
'SetExceptionState void
'haveCallbacks bool
'appendCallback void
'removeCallback bool
'clearCallbacks void
'Thread Thread
'setID void
'getID
'
'tex Object   'it returns internal object which keeps object texed values.


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
