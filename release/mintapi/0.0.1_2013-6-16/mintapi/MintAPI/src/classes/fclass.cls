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
Option Base 0
Option Explicit
Const CLASSID As String = "fclass"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

Private Declare Function API_IsBadCodePtr Lib "kernel32" Alias "IsBadCodePtr" (ByVal lpfn As Long) As Long


Implements IAliasable


Dim obj As Object
Dim txObject As texedObject

Friend Sub Initialize(Optional Object As Object)
    Set obj = Object
End Sub
Friend Sub InitializeFromHandle(Optional ObjHandle As Long = 0)
    If ObjHandle <> 0 Then
        Dim retObj As Object
        Call API_CopyMemory(ByVal ObjPtr(retObj), ByVal ObjHandle, 4)
        
        If API_IsBadCodePtr(ObjHandle) Then throw InvalidHandleException
        On Error GoTo nonObject
        Set obj = ObjHandle
nonObject:
    End If
End Sub

Public Sub SetObject(Object As Object)
    Set obj = Object
End Sub

Private Sub CheckObject()
    If obj Is Nothing Then throw InvalidStatusException("Object Not Set.")
End Sub

Public Property Get Name() As String
CheckObject
    If Not obj Is Nothing Then Name = TypeName(obj)
End Property

Public Function supporttexing() As Boolean
CheckObject
    If obj Is Nothing Then Exit Function
    supporttexing = (TypeOf obj Is IClassTexer)
End Function

Public Function constData() As ByteArray
CheckObject
    If TypeOf obj Is IClassTexer Then
        Dim objTexer As IClassTexer
        Set objTexer = obj
        Set constData = ByteArray(objTexer.toByteArray)
    Else
        throw Exception("Unknown Class Type.")
    End If
End Function
Public Function toString() As String
CheckObject
    If TypeOf obj Is IClassTexer Then
        Dim objTexer As IClassTexer
        Set objTexer = obj
        toString = objTexer.toString
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
    Set getclass = obj
End Function

Public Sub out(Data, Optional Length As Long = -1)
CheckObject
    If TypeOf obj Is ITargetStream Then
        mod_application.out obj, Data, Length
    Else
        throw Exception("Unknown Class Type.")
    End If
End Sub
Public Sub inp(Data, Optional Length As Long = -1)
CheckObject
    If TypeOf obj Is ITargetStream Then
        mod_application.inp obj, Data, Length
    Else
        throw Exception("Unknown Class Type.")
    End If
End Sub
Public Sub lout(ParamArray Data() As Variant)
CheckObject
    Dim objTargetStream As ITargetStream
    If Not TypeOf obj Is ITargetStream Then throw InvalidArgumentTypeException
    Set objTargetStream = obj
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
    If Not TypeOf obj Is ITargetStream Then throw InvalidArgumentTypeException
    Set objTargetStream = obj
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

Public Sub setStreamState(Dir As StreamDirection, Optional State As Boolean = True)
CheckObject
    If TypeOf obj Is ITargetStream Then
        Dim Stream As ITargetStream
        Set Stream = obj
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
    If TypeOf obj Is ITargetStream Then
        Dim Stream As ITargetStream
        Set Stream = obj
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

Public Property Get TexingFormat() As TexingFormat
    TexingFormat = tfAuto_Editable
End Property
Public Property Let TexingFormat(Value As TexingFormat)
    
End Property

Public Function tex() As texedObject
    If txObject Is Nothing Then Call CreateTexObject
    Set tex = txObject
End Function
Public Function GetTexObject() As texedObject
    If txObject Is Nothing Then Call CreateTexObject
    Set GetTexObject = txObject
End Function
Public Function CreateTexObject() As texedObject
CheckObject
    If Not TypeOf obj Is IClassTexer Then throw InvalidStatusException("")
    Dim txBuff As texedObject
    Set txBuff = New texedObject
    Call txBuff.Initialize(obj)
    Set txObject = txBuff
    Set CreateTexObject = txBuff
End Function
Public Sub DestroyTexObject()
    
End Sub

Public Sub boundTranslation(targetProperty As String, Translation As String)
CheckObject
    If obj Is Nothing Then throw InvalidStatusException
    Call boundtranslationto(Translation, obj, targetProperty, True)
End Sub
Public Sub unboundTranslation(targetProperty As String)
CheckObject
    If obj Is Nothing Then throw InvalidStatusException
    Call unboundtranslationto(obj, targetProperty)
End Sub

Public Sub WriteToFile(File As File)
CheckObject
    If TypeOf obj Is ITargetStream Or TypeOf obj Is IClassTexer Or TypeOf obj Is IData Then
        mod_application.out File.EnsureOpen, obj
    Else
        throw Exception("Unknown Class Type.")
    End If
End Sub
Public Sub WriteToStream(targetStream As ITargetStream)
CheckObject
    If TypeOf obj Is ITargetStream Or TypeOf obj Is IClassTexer Or TypeOf obj Is IData Then
        mod_application.out File.EnsureOpen, obj
    Else
        throw Exception("Unknown Class Type.")
    End If
End Sub
Public Function ReadFromFile(File As File) As Object
CheckObject
    If TypeOf obj Is ITargetStream Or TypeOf obj Is IClassTexer Or TypeOf obj Is IData Then
        mod_application.inp File.EnsureOpen, obj
    Else
        throw Exception("Unknown Class Type.")
    End If
End Function
Public Function ReadFromStream(targetStream As ITargetStream) As Object
CheckObject
    If TypeOf obj Is ITargetStream Or TypeOf obj Is IClassTexer Or TypeOf obj Is IData Then
        mod_application.inp File.EnsureOpen, obj
    Else
        throw Exception("Unknown Class Type.")
    End If
End Function

Public Function haveSlot() As Boolean
CheckObject
    If TypeOf obj Is ISupportSlot Then _
        haveSlot = True
End Function
Public Function haveSignal() As Boolean
CheckObject
    If TypeOf obj Is ISupportSignal Then _
        haveSignal = True
End Function
Public Sub Emit(Signal)
CheckObject
    If Not TypeOf obj Is ISupportSignal Then _
        throw InvalidVarTypeException
    Dim SignalObject As ISupportSignal
    Set SignalObject = obj
    Call SignalObject.Emitter.Emit(Signal)
End Sub

Public Function Clone() As Object
CheckObject
    If TypeOf obj Is IClassClonable Then
        Dim tclonableObj As IClassClonable
        Set tclonableObj = obj
        Set Clone = tclonableObj.Clone
    Else
        throw Exception("Unknown Class Type.")
    End If
End Function
Public Function clonable() As Boolean
CheckObject
    clonable = (TypeOf obj Is IClassClonable)
End Function
Public Function streamable() As Boolean
CheckObject
    streamable = (TypeOf obj Is ITargetStream)
End Function

Public Function objdefault(objName As String) As Object 'Only MintAPI classes are accepted.
CheckObject
    Set objdefault = CreateObject(APPLICATIONID & "." & objName)
End Function
Public Function Default() As Object 'Only MintAPI classes are accepted.
CheckObject
    If obj Is Nothing Then throw InvalidCallException
    Set Default = CreateObject(APPLICATIONID & "." & TypeName(obj))
End Function

Public Sub BlockSignals()
CheckObject
    If Not TypeOf obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = obj
    Call ISS.Emitter.BlockSignals
End Sub
Public Sub UnblockSignals()
CheckObject
    If Not TypeOf obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = obj
    Call ISS.Emitter.UnblockSignals
End Sub
Public Sub BlockSignal(SignalName As String)
CheckObject
    If Not TypeOf obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = obj
    Call ISS.Emitter.BlockSignal(SignalName)
End Sub
Public Sub UnblockSignal(SignalName As String)
CheckObject
    If Not TypeOf obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = obj
    Call ISS.Emitter.UnblockSignal(SignalName)
End Sub
Public Sub SetSignalState(SignalName As String, State As Boolean)
CheckObject
    If Not TypeOf obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = obj
    Call ISS.Emitter.SetSignalState(SignalName, State)
End Sub
Friend Function SignalEmitterState() As Boolean
CheckObject
    If Not TypeOf obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = obj
    Call ISS.Emitter.SignalEmitterState
End Function
Friend Function SignalState(SignalName As String) As Boolean
CheckObject
    If Not TypeOf obj Is ISupportSignal Then throw InvalidStatusException("Object does not support signals.")
    Dim ISS As ISupportSignal
    Set ISS = obj
    Call ISS.Emitter.SignalState(SignalName)
End Function

Public Property Get Alias() As String
CheckObject
    If Not TypeOf obj Is IAliasable Then throw InvalidStatusException
    Dim IA As IAliasable
    Set IA = obj
    Alias = IA.Alias
End Property
Public Function aliasable() As Boolean
CheckObject
    If Not TypeOf obj Is IAliasable Then Exit Function
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
