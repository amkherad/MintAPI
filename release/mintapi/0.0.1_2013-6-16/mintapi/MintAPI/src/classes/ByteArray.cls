VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ByteArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "ByteArray"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Const SIGNAL_CHANGED As String = "changed"
Const SIGNAL_CLEARING As String = "clearing"
Const SIGNAL_CLEARED As String = "cleared"
Const SIGNAL_BUFFERBOUNDED As String = "bufferbounded"
Const SIGNAL_BUFFERUNBOUNDED As String = "bufferunbounded"

Implements ITargetStream
Implements IClassClonable
Implements IClassExceptionManager
Implements IClassTexer
Implements ISupportSignal
Implements IAliasable

Dim dlen As Long
Dim d() As Byte
Dim IsBounded As Boolean
Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String

Dim mySignalEmitter As New SignalEmitter

Friend Sub Initialize(Optional target)
    If Not IsMissing(target) Then Call setBuffer(target)
End Sub

Private Sub Class_Initialize()
    Call mySignalEmitter.Initialize(Me)
    inState = True
    outState = True
End Sub

Public Property Let Length(Value As Long)
    Call reserve(Value, True)
End Property
Public Property Get Length() As Long
    Length = dlen
End Property
Public Property Get IsEmpty() As Boolean
    IsEmpty = (dlen = 0)
End Property
Public Sub Clear()
    Call mySignalEmitter.Emit(SIGNAL_CLEARING)
    Erase d
    dlen = 0
    Call mySignalEmitter.Emit(SIGNAL_CLEARED)
End Sub
Public Sub reserve(Length As Long, Optional keepLoadedData As Boolean = False)
    If Length < 0 Then throw InvalidArrayLBoundException("Length Can't Be Zero Or Negative.")
    If keepLoadedData Then
        ReDim Preserve d(Length - 1)
    Else
        ReDim d(Length - 1)
    End If
    dlen = Length
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub setBuffer(targetData)
    Dim bArray() As Byte
    bArray = mint_get_byte_array_of(targetData)
    'gets all types of data into byte array.
    Dim tdlen As Long
    tdlen = ArraySize(bArray)
    If tdlen <= 0 Then
        Call Clear
        Exit Sub
    End If
    If dlen <> tdlen Then Call reserve(tdlen)
    Dim i As Long
    For i = 0 To dlen - 1
        d(i) = bArray(i)
    Next
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub setBoundedBuffer(targetByteArray() As Byte)
    Dim blen As Long
    blen = ArraySize(targetByteArray)
    If blen <= 0 Then throw InvalidArgumentValueException("targetByteArray Is Null.")
    If LBound(targetByteArray) <> 0 Then throw InvalidArrayLBoundException("Invalid targetByteArray LBound.")
    Erase d
    dlen = 0
    Call API_CopyMemory(ByVal API_VarPtrArray(d), ByVal API_VarPtrArray(targetByteArray), 4)
    dlen = blen
    Call mySignalEmitter.Emit(SIGNAL_BUFFERBOUNDED)
End Sub

Public Function IsArrayBounded() As Boolean
    IsArrayBounded = IsBounded
End Function
Public Function MakeBufferInternal() As Boolean
    If Not IsBounded Then GoTo exitfunctionimmediately
    d = d
exitfunctionimmediately:
    MakeBufferInternal = IsBounded
    Call mySignalEmitter.Emit(SIGNAL_BUFFERUNBOUNDED)
End Function

Public Sub Append(target)
    Dim targetArray() As Byte
    Dim blen As Long
    targetArray = mint_get_byte_array_of(target)
    blen = ArraySize(targetArray)
    If blen = 0 Then Exit Sub
    '=======================
    'Call AppendArrayToArray(d, targetArray)
    Dim bIndex As Long, i As Long
    blen = (dlen + blen)
    ReDim Preserve d(blen - 1)
    For i = dlen To blen - 1
        d(i) = targetArray(bIndex)
        bIndex = bIndex + 1
    Next
    '=======================
    dlen = ArraySize(d)
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub CopyByteArray(targetByteArray() As Byte)
    targetByteArray = d
End Sub
Public Function GetByteArray() As Byte()
    GetByteArray = d
    'Call API_CopyMemory(ByVal API_VarPtrArray(GetByteArray), ByVal API_VarPtrArray(d), 4)
End Function
Public Function constData() As Byte()
    constData = d
    'Call API_CopyMemory(ByVal API_VarPtrArray(retVal), ByVal API_VarPtrArray(d), 4)
End Function
Public Sub SetByteArray(B() As Byte)
    Dim tdlen As Long
    tdlen = ArraySize(B)
    If dlen > tdlen Then Call Clear
    If dlen < tdlen Then Call reserve(tdlen)

    Dim i As Long
    For i = 0 To tdlen - 1
        d(i) = B(i)
    Next
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub

Public Function GetByteAt(Index As Long) As Byte
    If (Index < 0) Or (Index >= dlen) Then throw OutOfRangeException, Me
    GetByteAt = d(Index)
End Function
Public Sub SetByteAt(Index As Long, Value As Byte)
    If (Index < 0) Or (Index >= dlen) Then throw OutOfRangeException, Me
    If d(Index) = Value Then Exit Sub
    d(Index) = Value
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub

Public Property Get ByteAt(Index As Long) As Byte
    If (Index < 0) Or (Index >= dlen) Then throw OutOfRangeException, Me
    ByteAt = d(Index)
End Property
Public Property Let ByteAt(Index As Long, Value As Byte)
    If (Index < 0) Or (Index >= dlen) Then throw OutOfRangeException, Me
    If d(Index) = Value Then Exit Property
    d(Index) = Value
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Property

Public Property Get First() As Byte
    First = d(0)
End Property
Public Property Let First(Value As Byte)
    If d(0) = Value Then Exit Property
    d(0) = Value
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Property
Public Property Get Last() As Byte
    If dlen <= 0 Then throw OutOfRangeException, Me
    Last = d(dlen - 1)
End Property
Public Property Let Last(Value As Byte)
    If dlen <= 0 Then throw OutOfRangeException, Me
    If d(dlen - 1) = Value Then Exit Property
    d(dlen - 1) = Value
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Property

Public Function SubArray(StartIndex As Long, Optional Length As Long = -1) As ByteArray
    If ArraySize(d) = 0 Then throw InvalidCallException("ByteArray Buffer Is Empty.")
    Dim ba As New ByteArray
    Call ba.setBuffer(GetSubByteArray(d, StartIndex, Length))
    Set SubArray = ba
End Function

Public Sub DropLast(Length As Long)
    Dim rtLength As Long
    rtLength = dlen - Length
    If rtLength = 0 Then Exit Sub
    If rtLength < 0 Then throw OutOfRangeException
    ReDim Preserve d(rtLength - 1)
    dlen = ArraySize(d)
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub

Public Sub Replace(btaWhat As ByteArray, btaWith As ByteArray, Optional StartIndex As Long = 1, Optional Count As Long = -1)

End Sub
Public Function IndexOf(btaWhat As ByteArray, Optional StartIndex As Long = 0) As Long
    IndexOf = IndexOfByteArray(d, btaWhat.constData, StartIndex)
End Function
Public Function LastIndexOf(btaWhat As ByteArray, Optional StartIndex As Long = -1) As Long
    LastIndexOf = LastIndexOfByteArray(d, btaWhat.constData, StartIndex)
End Function
Public Function HowManyOf(btaWhat As ByteArray, Optional StartIndex As Long = 1) As Long

End Function
Public Sub Trim()

End Sub
Public Sub TrimLeft()

End Sub
Public Sub TrimRight()

End Sub
Public Sub TrimArr(trimBytes() As ByteArray)

End Sub
Public Sub TrimLeftArr(trimBytes() As ByteArray)

End Sub
Public Sub TrimRightArr(trimBytes() As ByteArray)

End Sub

Private Function IsWhiteSpace(Optional OptionalByteArray) As Boolean
    'space: ,tab: ,unbreakablespace: ,enter:

End Function
Public Function IsNullOrEmpty(Optional OptionalByteArray) As Boolean

End Function
Public Function IsNullOrWhiteSpace(Optional OptionalByteArray) As Boolean

End Function

Public Sub Join(targetArray, Optional Delimiter)

End Sub
Public Function Split(Optional Delimiter, Optional Limit As Long = -1) As ByteArray()

End Function
Public Sub Concat(btaConcat As ByteArray, Optional Separator As String = "")

End Sub
Public Sub Format(btaFormat As ByteArray)

End Sub
Public Function Compare(btaWhat As ByteArray) As Boolean

End Function
Public Sub Reverse(Optional Length As Long = -1)
    If dlen = 0 Then Exit Sub
    If Length = 0 Then Exit Sub
    If Length < -1 Then throw OutOfRangeException
    If Length > dlen Then throw OutOfRangeException
    Dim rtLength As Long
    rtLength = dlen
    If Length >= 0 Then
        rtLength = Length
    End If
    Dim i As Long, chBuff As Byte, sum_X As Long
    For i = 0 To (rtLength / 2) - 1
        sum_X = rtLength - i - 1
        chBuff = d(i)
        d(i) = d(sum_X)
        d(sum_X) = chBuff
    Next
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub Remove(StartIndex As Long, Optional Length As Long = -1)
    If StartIndex < 0 Then throw OutOfRangeException
    If StartIndex >= dlen Then throw OutOfRangeException
    If Length > dlen - StartIndex Then throw OutOfRangeException
    If Length <= 0 Then Exit Sub
    Dim i As Long, rtLength As Long
    If Length = -1 Then
        rtLength = dlen - StartIndex
    Else
        rtLength = Length
    End If
    For i = StartIndex To dlen - rtLength - 1
        d(i) = d(i + rtLength)
    Next
    dlen = dlen - rtLength
    ReDim Preserve d(dlen - 1)
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub InsertAt(Index As Long, ByteValue As Byte)
    Dim i As Long
    ReDim Preserve d(dlen)
    For i = dlen To Index + 1 Step -1
        d(i) = d(i - 1)
    Next
    d(Index) = ByteValue

    dlen = dlen + 1
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub Insert(Index As Long, targetBA As ByteArray, Optional Length As Long = -1)
    If targetBA Is Nothing Then throw ArgumentNullException
    Call InsertArrayIndexArray(d, Index, targetBA.constData, Length)
    dlen = ArraySize(d)
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub PadLeft(TotalLength As Long, FillChars As ByteArray)

End Sub
Public Sub PadRight(TotalLength As Long, FillChars As ByteArray)

End Sub
Public Function Contains(SearchWhat As ByteArray, Optional StartIndex As Long = 0) As Boolean

End Function
Public Function StartsWith(btaWhat As ByteArray, Optional StartIndex As Long = 0) As Boolean

End Function
Public Function EndsWith(btaWhat As ByteArray, Optional LastStartIndex As Long = -1) As Boolean

End Function
Public Sub ShiftLeft(StartIndex As Long, TimesToShift As Long, Optional Length As Long = -1)
    If StartIndex < 0 Then throw OutOfRangeException
    If StartIndex >= dlen Then throw OutOfRangeException
    If Length > dlen - StartIndex Then throw OutOfRangeException
    If Length <= 0 Then Exit Sub
    Dim i As Long, rtLength As Long
    If Length = -1 Then
        rtLength = dlen - StartIndex
    ElseIf Length < dlen Then
        rtLength = Length
    Else
        rtLength = dlen - StartIndex
    End If
    For i = StartIndex To dlen - 1
        If i + TimesToShift < dlen Then
            d(i) = d(i + TimesToShift)
        Else
            d(i) = 0
        End If
    Next
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub ShiftRight(StartIndex As Long, TimesToShift As Long, Optional Length As Long = -1)
    If StartIndex < 0 Then throw OutOfRangeException
    If StartIndex >= dlen Then throw OutOfRangeException
    If StartIndex - Length < 0 Then throw OutOfRangeException
    If Length <= 0 Then Exit Sub
    Dim i As Long, rtTo As Long
    If Length = -1 Then
        rtTo = 0
    Else
        rtTo = StartIndex - Length + 1
    End If
    For i = StartIndex To rtTo Step -1
        If i - TimesToShift >= 0 Then
            d(i) = d(i - TimesToShift)
        Else
            d(i) = 0
        End If
    Next
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub

Public Sub toLower()
    Dim i As Long
    For i = 0 To dlen - 1
        If d(i) >= 65 And d(i) <= 90 Then
            d(i) = d(i) + 32
        End If
    Next
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub
Public Sub toUpper()
    Dim i As Long
    For i = 0 To dlen - 1
        If d(i) >= 97 And d(i) <= 122 Then
            d(i) = d(i) - 32
        End If
    Next
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
End Sub


Public Function toBigNumber() As BigNumber
    If dlen = 0 Then Exit Function

End Function
Public Function toInt() As Integer
    If dlen = 0 Then Exit Function
    If dlen <= 2 Then
        toInt = ByteArrayToInteger(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toBoolean() As Boolean
    If dlen = 0 Then Exit Function
    If dlen <= 2 Then
        toBoolean = ByteArrayToBoolean(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toString() As String
    If dlen = 0 Then Exit Function
    toString = ByteArrayToString(d)
End Function
Public Function toLPSTR() As String
    If dlen = 0 Then Exit Function
    toLPSTR = ByteArrayToSafeString(d)
End Function
Public Function toLong() As Long
    If dlen = 0 Then Exit Function
    If dlen <= 4 Then
        toLong = ByteArrayToLong(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toDouble() As Double
    If dlen = 0 Then Exit Function
    If dlen <= 8 Then
        toDouble = ByteArrayToDouble(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toFloat() As Single
    If dlen = 0 Then Exit Function
    If dlen <= 4 Then
        toFloat = ByteArrayToSingle(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toSingle() As Single
    If dlen = 0 Then Exit Function
    If dlen <= 4 Then
        toSingle = ByteArrayToSingle(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toDate() As Date
    If dlen = 0 Then Exit Function
    If dlen <= 8 Then
        toDate = ByteArrayToDate(d)
    Else
        throw OutOfRangeException
    End If
End Function
Public Function toCurrency() As Currency
    If dlen = 0 Then Exit Function
    If dlen <= 8 Then
        toCurrency = ByteArrayToCurrency(d)
    Else
        throw OutOfRangeException
    End If
End Function


Public Function fromBigNumber(Value As BigNumber) As ByteArray
    Set fromBigNumber = New ByteArray
    Call fromBigNumber.setBuffer(Value)
End Function
Public Function fromInt(Value As Integer) As ByteArray
    Set fromInt = New ByteArray
    Call fromInt.setBuffer(Value)
End Function
Public Function fromBoolean(Value As Boolean) As ByteArray
    Set fromBoolean = New ByteArray
    Call fromBoolean.setBuffer(Value)
End Function
Public Function fromString(Value As String) As ByteArray
    Set fromString = New ByteArray
    Call fromString.setBuffer(Value)
End Function
Public Function fromLong(Value As Long) As ByteArray
    Set fromLong = New ByteArray
    Call fromLong.setBuffer(Value)
End Function
Public Function fromDouble(Value As Double) As ByteArray
    Set fromDouble = New ByteArray
    Call fromDouble.setBuffer(Value)
End Function
Public Function fromFloat(Value As Single) As ByteArray
    Set fromFloat = New ByteArray
    Call fromFloat.setBuffer(Value)
End Function
Public Function fromSingle(Value As Single) As ByteArray
    Set fromSingle = New ByteArray
    Call fromSingle.setBuffer(Value)
End Function
Public Function fromDate(Value As Date) As ByteArray
    Set fromDate = New ByteArray
    Call fromDate.setBuffer(Value)
End Function
Public Function fromCurrency(Value As Currency) As ByteArray
    Set fromCurrency = New ByteArray
    Call fromCurrency.setBuffer(Value)
End Function


Public Function toHex() As ByteArray
    'toHex = Hex$(d)
End Function
Public Function toOctal() As ByteArray
    'toOctal = Oct$(d)
End Function
Public Function toBinary() As ByteArray ' binary number 1010001011101
    '
End Function
Public Function fromHex(Value As ByteArray) As ByteArray
    '
End Function
Public Function fromOctal(Value As ByteArray) As ByteArray
    '
End Function
Public Function fromBinary(Value As ByteArray) As ByteArray ' binary number 1010001011101
    '
End Function
Public Function IsNumeric() As Boolean
    IsNumeric = IsByteArrayNumeric(d)
End Function
Public Function IsAlphabetic() As Boolean
    IsAlphabetic = IsByteArrayAlphabetic(d)
End Function
Public Function IsLike(Value2 As ByteArray) As Boolean
    IsLike = IsByteArrayLikeAnother(d, Value2.GetByteArray)
End Function
Public Function IsEqualTo(Value2 As ByteArray, Optional ByVal LengthToCompare As Long = -1) As Boolean
    IsEqualTo = ArrayCompare(d, Value2.constData, LengthToCompare)
End Function


'$ByteArray$08datetime3###
Private Function IsByteArraySpecifiedData(B() As Byte) As Boolean
On Error GoTo Err
    Dim header() As Byte
    header = GetSubByteArray(B, 0, 11)

    If ArrayCompare(header, StringToByteArray("$ByteArray$")) Then
        IsByteArraySpecifiedData = True
    Else
        IsByteArraySpecifiedData = False
    End If
Err:
End Function
Private Function checkifspc(B() As Byte, ClassName As String) As Boolean
    If Not IsByteArraySpecifiedData(B) Then
        checkifspc = False
        Exit Function
    End If
    If GetByteArrayClassTypeName(B) <> ClassName Then
        checkifspc = False
        Exit Function
    End If
End Function
Private Function GetByteArrayClassTypeName(B() As Byte) As String
On Error GoTo returnNA
    If Not IsByteArraySpecifiedData(B) Then GoTo returnNA
    Dim bSize As Long
    bSize = ArraySize(B)
    If bSize <= 0 Then GoTo returnNA

    Dim ln() As Byte, afterCLSNameLength As Long
    ln = GetSubByteArray(B, 11, 2)

    afterCLSNameLength = ByteArray(B).toInt
    If afterCLSNameLength <= 0 Then GoTo returnNA

    ln = GetSubByteArray(B, 13, afterCLSNameLength)
    If ArraySize(ln) <= 0 Then GoTo returnNA

    GetByteArrayClassTypeName = ByteArrayToString(ln)

    Exit Function
returnNA:
    GetByteArrayClassTypeName = "N/A"
End Function
Private Function GetByteArrayInnerData(B() As Byte) As Byte()
    On Error GoTo returnN
    Dim bSize As Long
    bSize = ArraySize(B)
    If bSize <= 0 Then GoTo returnN

    Dim retVal() As Byte

    Dim ln() As Byte, afterCLSNameLength As Long, innerDataLength As Long
    ln = GetSubByteArray(B, 11, 2)

    afterCLSNameLength = ByteArray(ln).toInt
    If afterCLSNameLength <= 0 Then GoTo returnN

    ln = GetSubByteArray(B, afterCLSNameLength + 11, 2)
    innerDataLength = ByteArray(ln).toInt

        Call API_CopyMemory(ByVal API_VarPtrArray(GetByteArrayInnerData), ByVal API_VarPtrArray(GetSubByteArray(B, afterCLSNameLength + 12, innerDataLength)), 4)

    Exit Function
returnN:
End Function

Public Function Clone() As ByteArray
    Dim clonedByteArray As New ByteArray
    Call clonedByteArray.Initialize(d)
    Set Clone = clonedByteArray
End Function


Private Sub IData_Append(target As Variant): Call Append(target): End Sub

Private Sub IData_Clear(): Call Clear: End Sub
Private Sub IData_CopyByteArray(targetByteArray() As Byte): Call CopyByteArray(targetByteArray): End Sub

Private Function IData_GetByteArray() As Byte()
    Call API_CopyMemory(ByVal API_VarPtrArray(IData_GetByteArray), ByVal API_VarPtrArray(d), 4)
End Function
Private Property Get IData_Length() As Long: IData_Length = dlen: End Property

Private Function IClassClonable_Clone() As Object: Set IClassClonable_Clone = Clone: End Function

Private Sub IClassExceptionManager_setAllExceptionsState(Optional States As Boolean = True)

End Sub

Private Sub IClassExceptionManager_setExceptionState(exp As ExceptionType, Optional State As Boolean = True)

End Sub

Private Sub IClassTexer_fromByteArray(Value() As Byte)
    d = Value
    dlen = ArraySize(d)
End Sub
Private Sub IClassTexer_fromString(Value As String)
    d = StringToByteArray(Value)
    dlen = ArraySize(d)
End Sub
Private Property Get IClassTexer_texHeader(texType As mint_TexingType) As String

End Property
Private Function IClassTexer_toByteArray() As Byte()
    IClassTexer_toByteArray = d
End Function

Private Function IClassTexer_toString() As String
    IClassTexer_toString = ByteArrayToString(d)
End Function

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter
End Function

'Private Property Get ISupportSlot_AutoManage() As Boolean
'    ISupportSlot_AutoManage = True
'End Property
'Private Function ISupportSlot_Methods() As String
'    ISupportSlot_Methods = _
'        "0" & SLOT_CHANGED
'End Function
'Private Sub ISupportSlot_ManualManageCall(SlotName As String)
'
'End Sub

Private Function ITargetStream_Available(Dir As StreamDirection) As Long
    Dim val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        val = ((dlen > 0) And inState)
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        val = val Or (outState)
    End If
    ITargetStream_Available = val
End Function

Private Sub ITargetStream_CloseStream(Optional direction As StreamDirection = StreamDirection.sdBoth)
    If (direction And sdInStream) = sdInStream Then
        inState = False
    End If
    If (direction And sdOutStream) = sdOutStream Then
        outState = False
    End If
    If direction = sdBoth Then Call Clear
End Sub
Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = StreamDirection.sdBoth)
    If (direction And sdInStream) = sdInStream Then
        inState = True
    End If
    If (direction And sdOutStream) = sdOutStream Then
        outState = True
    End If
End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean
    Dim val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        val = inState
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        val = val Or outState
    End If
    ITargetStream_getState = val
End Function

Private Sub ITargetStream_inStream(Data, Optional Length As Long = -1)
    If Not inState Then throw InvalidStatusException
    'Call mint_put_byte_array_to(Data, d, Length)
End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    ITargetStream_readBytes = GetSubByteArray(d, StartAt, Length)
End Function
Private Sub ITargetStream_outStream(Data, Optional Length As Long = -1)
    If Not outState Then throw InvalidStatusException
    If VarType(Data) = VBObject Then
        If TypeOf Data Is ObjectBuffer Then
            Dim objBuffer As ObjectBuffer
            Set objBuffer = Data
            If objBuffer.IsStreamingValue Then
                Call objBuffer.thrownotargs(1)
                If objBuffer.ArgType(0) = vbString Then
                    Dim strFirst As String
                    strFirst = objBuffer.First
                    If strFirst = OBJECTBUFFER_VALUES_STREAMING_CLEAR Then
                        Call Clear
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_NEWLINE Then
                        Call Append(vbCrLf)
                        Exit Sub
                    Else
                        throw InvalidArgumentTypeException("Streaming Command '" & strFirst & "' Is Not Available For ByteArray.")
                    End If
                Else
                    throw InvalidArgumentTypeException("Streaming Type Is Not Available For ByteArray.")
                End If
            End If
        End If
    End If
    Call Append(mint_get_byte_array_of(Data, Length))
End Sub

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, inState, inState_LOCK, Reserved)
End Sub
Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, outState, outState_LOCK, Reserved)
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
