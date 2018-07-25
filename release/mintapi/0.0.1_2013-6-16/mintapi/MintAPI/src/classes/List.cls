VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "List"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Const SIGNAL_ADD As String = "add"
Const SIGNAL_REMOVE As String = "remove"
Const SIGNAL_CLEAR As String = "clear"
Const SIGNAL_SWAP As String = "swap"
Const SIGNAL_CHANGED As String = "changed"
Const SIGNAL_SORT As String = "sort"
Const SIGNAL_LEFTTRIM As String = "lefttrim"
Const SIGNAL_RIGHTTRIM As String = "righttrim"
Const SIGNAL_TRIM As String = "trim"

Const SLOT_CLEAR As String = "clear"
Const SLOT_ACCENDINGSORT As String = "accendingsort"
Const SLOT_DECENDINGSORT As String = "decendingsort"
Const SLOT_POP As String = "pop"


'Implements ITargetStream
Implements IClassClonable
'Implements IClassTexer
'Implements IClassExceptionManager
Implements ISupportSignal
Implements ISupportSlot
Implements IAliasable

Dim obj()
Dim objCount As Long
Dim tp As VbVarType
Dim sttControlTypes As Boolean
Dim nam As String

Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String

Dim mySignalEmitter As New SignalEmitter

Private Sub Class_Initialize()
    Call mySignalEmitter.Initialize(Me)
End Sub
Private Sub Class_Terminate()
    Erase obj
    objCount = 0
End Sub


Friend Sub Initialize(Content() As Variant)
    Dim i As Long
    For i = 0 To ArraySize(Content) - 1
        Call Add(Content(i))
    Next
End Sub


Public Property Get Count() As Long
    Count = objCount
End Property
Public Property Get Item(ByVal Index As Long)
Attribute Item.VB_UserMemId = 0
    If Index < 0 Or Index >= objCount Then throw OutOfRangeException
    If VBA.IsObject(obj(Index)) Then
        Set Item = obj(Index)
    Else
            Item = obj(Index)
    End If
End Property
Public Property Let Item(ByVal Index As Long, Value)
    If Index < 0 Or Index >= objCount Then throw OutOfRangeException
    If sttControlTypes Then _
        If VarType(Value) <> tp Then throw Exception("Value type is not same as assigned type.")
    If VBA.IsObject(Value) Then
        Set obj(Index) = Value
    Else
            obj(Index) = Value
    End If
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Property
Public Property Set Item(ByVal Index As Long, Value As Object)
    If Index < 0 Or Index >= objCount Then throw OutOfRangeException
    If sttControlTypes Then _
        If VarType(Value) <> tp Then throw Exception("Value type is not same as assigned type.")
    Set obj(Index) = Value
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Property
Public Function getTypeOf(ByVal Index As Long) As VbVarType
    If Index < 0 Or Index >= objCount Then throw OutOfRangeException
    getTypeOf = VarType(obj(Index))
End Function
Public Function IsObject(ByVal Index As Long) As VbVarType
    If Index < 0 Or Index >= objCount Then throw OutOfRangeException
    IsObject = VBA.IsObject(obj(Index))
End Function
Public Function IsArray(ByVal Index As Long) As VbVarType
    If Index < 0 Or Index >= objCount Then throw OutOfRangeException
    IsArray = VBA.IsArray(obj(Index))
End Function

Public Sub Append(Value)
    Call Add(Value)
End Sub
Public Sub Add(Value)
Dim vt As VbVarType
vt = VarType(Value)
    If sttControlTypes Then
        If objCount > 0 Then
            If tp <> vbVariant Then
                If vt <> tp Then
                    throw Exception("Value type is not same as assigned type.")
                End If
            End If
        Else
            tp = vt
        End If
    End If
    ReDim Preserve obj(objCount)
    If vt = VBObject Then
        Set obj(objCount) = Value
    Else
            obj(objCount) = Value
    End If
    objCount = objCount + 1
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_ADD)
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub Push(Value)
    Call Add(Value)
End Sub
Public Function Pop() As Variant
    Pop = obj(objCount - 1)
    Call RemoveByIndex(objCount - 1)
End Function

Public Sub RemoveByIndex(ByVal ind As Long)
Dim vt As VbVarType
Dim i As Long
If ind >= 0 And ind <= objCount - 2 Then
    Dim xb()
    ReDim xb(objCount - 2)
    For i = 0 To objCount - 1
        vt = VarType(obj(i))
        If i < ind Then
            If vt = VBObject Then
                Set xb(i) = obj(i)
            Else
                    xb(i) = obj(i)
            End If
        ElseIf i > ind Then
            If vt = VBObject Then
                Set xb(i - 1) = obj(i)
            Else
                xb(i - 1) = obj(i)
            End If
        End If
    Next
    objCount = objCount - 1
    ReDim obj(objCount - 1)
    For i = 0 To objCount - 1
        If VBA.IsObject(obj(i)) Then
            Set obj(i) = xb(i)
        Else
                obj(i) = xb(i)
        End If
    Next
ElseIf ind = objCount - 1 Then
    If objCount > 1 Then
        Dim xb1()
        objCount = objCount - 1
        ReDim xb1(objCount - 1)
        For i = 0 To objCount - 1
            If VBA.IsObject(obj(i)) Then
                Set xb1(i) = obj(i)
            Else
                    xb1(i) = obj(i)
            End If
        Next
        ReDim obj(objCount - 1)
        For i = 0 To objCount - 1
            If VBA.IsObject(xb1(i)) Then
                Set obj(i) = xb1(i)
            Else
                    obj(i) = xb1(i)
            End If
        Next
    Else
        Call Clear
    End If
Else
    throw OutOfRangeException()
End If
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_REMOVE)
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub Remove(Value)
Dim i As Long
    If VBA.IsObject(Value) Then
        For i = 0 To objCount - 1
            If Not VBA.IsObject(obj(i)) Then
                If obj(i) = Value Then
                    Call RemoveByIndex(i)
                    Exit Sub
                End If
            End If
        Next
    Else
        For i = 0 To objCount - 1
            If VBA.IsObject(obj(i)) Then
                If obj(i) Is Value Then
                    Call RemoveByIndex(i)
                    Exit Sub
                End If
            End If
        Next
    End If
End Sub
Public Sub Clear()
    Erase obj()
    objCount = 0
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CLEAR)
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub Swap(ByVal FirstItem As Long, ByVal SecondItem As Long)
    Dim X
    If VBA.IsObject(obj(FirstItem)) Then
        Set X = obj(FirstItem)
    Else
            X = obj(FirstItem)
    End If
    If VBA.IsObject(obj(SecondItem)) Then
        Set obj(FirstItem) = obj(SecondItem)
    Else
            obj(FirstItem) = obj(SecondItem)
    End If
    If VBA.IsObject(X) Then
        Set obj(SecondItem) = X
    Else
            obj(SecondItem) = X
    End If
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_SWAP)
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Property Get First() As Variant
    If objCount <= 0 Then throw Exception("List is empty.")
    If VBA.IsObject(obj(0)) Then
        Set First = obj(0)
    Else
            First = obj(0)
    End If
End Property
Public Property Let First(Value As Variant)
    If objCount <= 0 Then throw Exception("List is empty.")
    If sttControlTypes Then _
        If VarType(Value) <> tp Then throw Exception("Value type is not same as assigned type.")
    If VBA.IsObject(Value) Then
        Set obj(0) = Value
    Else
            obj(0) = Value
    End If
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Property
Public Property Set First(Value As Variant)
    If objCount <= 0 Then throw Exception("List is empty.")
    If sttControlTypes Then _
        If VarType(Value) <> tp Then throw Exception("Value type is not same as assigned type.")
    Set obj(0) = Value
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Property
Public Property Get Last() As Variant
    If objCount <= 0 Then throw Exception("List is empty.")
    If VBA.IsObject(obj(objCount - 1)) Then
        Set First = obj(objCount - 1)
    Else
            First = obj(objCount - 1)
    End If
End Property
Public Property Let Last(Value As Variant)
    If objCount <= 0 Then throw Exception("List is empty.")
    If sttControlTypes Then _
        If VarType(Value) <> tp Then throw Exception("Value type is not same as assigned type.")
    If VBA.IsObject(Value) Then
        Set obj(objCount - 1) = Value
    Else
            obj(objCount - 1) = Value
    End If
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Property
Public Property Set Last(Value As Object)
    If objCount <= 0 Then throw Exception("List is empty.")
    If sttControlTypes Then _
        If VarType(Value) <> tp Then throw Exception("Value type is not same as assigned type.")
    Set obj(objCount - 1) = Value
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Property
Public Function IsExist(Value) As Boolean
    Dim i As Long
    If VBA.IsObject(Value) Then
        For i = 0 To objCount - 1
            If VBA.IsObject(obj(i)) Then
                If obj(i) Is Value Then
                    IsExist = True
                    Exit Function
                End If
            End If
        Next
    Else
        For i = 0 To objCount - 1
            If Not VBA.IsObject(obj(i)) Then
                If obj(i) = Value Then
                    IsExist = True
                    Exit Function
                End If
            End If
        Next
    End If
End Function
Public Function IndexOf(Value) As Long
    Dim i As Long
    If VBA.IsObject(Value) Then
        For i = 0 To objCount - 1
            If VBA.IsObject(obj(i)) Then
                If obj(i) Is Value Then
                    IndexOf = i
                    Exit Function
                End If
            End If
        Next
    Else
         For i = 0 To objCount - 1
            If Not VBA.IsObject(obj(i)) Then
                If obj(i) = Value Then
                    IndexOf = i
                    Exit Function
                End If
            End If
        Next
    End If
    IndexOf = -1
End Function
Public Function LastIndexOf(Value) As Long
    Dim i As Long
    If VBA.IsObject(Value) Then
        For i = objCount - 1 To 0 Step -1
            If VarType(obj(i)) = VBObject Then
                If obj(i) Is Value Then
                    LastIndexOf = i
                    Exit Function
                End If
            End If
        Next
    Else
        For i = objCount - 1 To 0 Step -1
            If Not VBA.IsObject(obj(i)) Then
                If obj(i) = Value Then
                    LastIndexOf = i
                    Exit Function
                End If
            End If
        Next
    End If
    LastIndexOf = -1
End Function
Public Function HowIn(Value) As Long
    Dim i As Long, cc As Long
    If VBA.IsObject(Value) Then
        For i = 0 To objCount - 1
            If VBA.IsObject(obj(i)) Then
                If obj(i) Is Value Then
                    cc = cc + 1
                End If
            End If
        Next
    Else
        For i = 0 To objCount - 1
            If Not VBA.IsObject(obj(i)) Then
                If obj(i) = Value Then
                    cc = cc + 1
                End If
            End If
        Next
    End If
    HowIn = cc
End Function
Public Sub LeftTrim(Value)
If VarType(Value) = VBObject Then
    While objCount > 0 And VBA.IsObject(obj(0)) And obj(0) = Value
        Call RemoveByIndex(0)
    Wend
Else
    While objCount > 0 And obj(0) = Value
        Call RemoveByIndex(0)
    Wend
End If
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_LEFTTRIM)
End Sub
Public Sub RightTrim(Value)
If VarType(Value) = VBObject Then
    While objCount > 0 And VBA.IsObject(obj(objCount - 1)) And obj(objCount - 1) Is Value
        Call RemoveByIndex(objCount - 1)
    Wend
Else
    While objCount > 0 And obj(objCount - 1) = Value
        Call RemoveByIndex(objCount - 1)
    Wend
End If
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_RIGHTTRIM)
End Sub
Public Sub Trim(Value)
    Call LeftTrim(Value)
    Call RightTrim(Value)
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_TRIM)
End Sub

Public Sub Sort(Optional ByVal Accending As Boolean = True)
    '
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_SORT)
End Sub

Public Sub fromArray(targetArray() As Variant)
On Error GoTo Err
objCount = UBound(targetArray) - LBound(targetArray)
    ReDim obj(objCount)
    Dim i As Long, j As Long
    For i = LBound(targetArray) To UBound(targetArray)
        If VBA.IsObject(targetArray(i)) Then
            Set obj(j) = targetArray(i)
        Else
                obj(j) = targetArray(i)
        End If
        j = j + 1
    Next
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
Err:
End Sub
Public Function toArray() As Variant()
    Dim targetArray() As Variant
    ReDim targetArray(objCount - 1)
    Dim i As Long
    For i = 0 To objCount - 1
        If VBA.IsObject(obj(i)) Then
            Set targetArray(i) = obj(i)
        Else
                targetArray(i) = obj(i)
        End If
    Next
    toArray = targetArray()
End Function
Public Sub fromSTDCollection(targetCollection As VBA.Collection)
objCount = targetCollection.Count
    ReDim obj(objCount - 1)
    Dim i As Long
    For i = 0 To objCount - 1
        If VBA.IsObject(targetCollection(i)) Then
            Set obj(i) = targetCollection(i)
        Else
                obj(i) = targetCollection(i)
        End If
    Next
    On Error GoTo 0
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Function toSTDCollection(targetCollection) As VBA.Collection
    Dim collectionBuff As New VBA.Collection
    Dim i As Long
    For i = 0 To objCount - 1
        Call collectionBuff.Add(obj(i))
    Next
    Set toSTDCollection = collectionBuff
End Function


Public Function Clone() As List
    Dim lst As New List
    Call lst.Initialize(obj)
    Set Clone = lst
End Function

Public Function toVector() As Vector
    
End Function
Public Function toCollection() As Collection
    
End Function



Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property

Private Function IClassClonable_Clone() As Object
    Set IClassClonable_Clone = Clone
End Function

'Private Sub IClassExceptionManager_setAllExceptionsState(Optional States As Boolean = True)
'
'End Sub
'Private Sub IClassExceptionManager_setExceptionState(exp As ExceptionType, Optional State As Boolean = True)
'
'End Sub
'
'Private Sub IClassTexer_fromByteArray(Value() As Byte)
'
'End Sub
'Private Sub IClassTexer_fromString(Value As String)
'
'End Sub
'Private Function IClassTexer_toByteArray() As Byte()
'
'End Function
'Private Function IClassTexer_toString() As String
'
'End Function
'Private Property Get IClassTexer_texHeader(texType As mint_TexingType) As String
'
'End Property

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter
End Function
Private Property Get ISupportSlot_AutoManage() As Boolean
    ISupportSlot_AutoManage = False
End Property
Private Sub ISupportSlot_ManualManageCall(SlotName As String)
    Select Case SlotName
        Case SLOT_CLEAR
            Call Clear
        Case SLOT_ACCENDINGSORT
            Call Sort(True)
        Case SLOT_DECENDINGSORT
            Call Sort(False)
        Case SLOT_POP
            Call Pop
    End Select
End Sub
Private Function ISupportSlot_Methods() As String
    ISupportSlot_Methods = _
        "0" & SLOT_CLEAR & _
        " 0" & SLOT_ACCENDINGSORT & _
        " 0" & SLOT_DECENDINGSORT & _
        " 0" & SLOT_POP
End Function

'Private Function ITargetStream_Available(dir As StreamDirection) As Long
'    Dim val As Boolean
'    If (dir And sdInStream) = sdInStream Then
'        val = ((in_hndl <> INVALID_HANDLE_VALUE) And inState)
'    End If
'    If (dir And sdOutStream) = sdOutStream Then
'        val = val Or ((out_hndl <> INVALID_HANDLE_VALUE) And outState)
'    End If
'    ITargetStream_Available = val
'End Function
'Private Sub ITargetStream_CloseStream()
'    inState = False
'    outState = False
'End Sub
'Private Function ITargetStream_getState(dir As StreamDirection) As Boolean
'    Dim val As Boolean
'    If (dir And sdInStream) = sdInStream Then
'        val = inState
'    End If
'    If (dir And sdOutStream) = sdOutStream Then
'        val = val Or outState
'    End If
'    ITargetStream_getState = val
'End Function
'Private Sub ITargetStream_inStream(Data As Variant, Optional Length As Long = -1&)
'
'End Sub
'Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
'    throw Exception("Unable to read bytes from list.")
'End Function
'Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = 3&)
'    If (direction And sdInStream) = sdInStream Then
'        inState = True
'    End If
'    If (direction And sdOutStream) = sdOutStream Then
'        outState = True
'    End If
'End Sub
'Private Sub ITargetStream_outStream(Data As Variant, Optional Length As Long = -1&)
'
'End Sub
'Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
'
'End Sub
'Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
'
'End Sub
