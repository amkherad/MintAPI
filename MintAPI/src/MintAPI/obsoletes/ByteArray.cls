VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
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


'Ideas from Me , Qt

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

'Implements ITargetStream
Implements ICloneable
Implements IEnumerable

Dim oaLen As Long, pExtraBuffer As Long
Dim dlen As Long
Dim d() As Byte
Dim IsBounded As Boolean
Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String

Dim CModify As Long

Friend Sub Constructor0(ByVal Length As Long)
    Call Reserve(Length)
End Sub
Friend Sub Constructor1(Optional target, Optional ByVal Length As Long = -1)
    If Not IsMissing(target) Then
        Call SetBuffer(target, Length)
    Else
        If Length > 0 Then
            Call Reserve(Length)
        End If
    End If
End Sub
Friend Sub Constructor2(ByRef arr As Variant)
    d = arr
    dlen = ArraySize(d)
End Sub
Friend Sub Constructor3(ByRef arr As Variant)
    
End Sub

Private Sub Class_Initialize()
    inState = True
    outState = True
End Sub

Public Property Let Length(Value As Long)
    Call Reserve(Value, True)
End Property
Public Property Get Length() As Long
    Length = dlen
End Property
Public Property Get Handle() As Long
    Handle = API_VarPtrArray(d)
End Property
Public Property Get DataHandle() As Long
    DataHandle = ArrayDataPtr(API_VarPtrArray(d))
End Property
Public Property Get IsEmpty() As Boolean
    IsEmpty = (dlen = 0)
End Property
Public Sub Clear()
    Call Signal.Emit(Me, SIGNAL_CLEARING)
    Erase d
    dlen = 0
    Call Signal.Emit(Me, SIGNAL_CLEARED)
End Sub
Public Sub Dispose()
    Call Clear
End Sub
Public Sub Reserve(Length As Long, Optional keepLoadedData As Boolean = False)
    If Length < 0 Then throw Exps.InvalidArrayLBoundException("Length Can't Be Zero Or Negative.")
    Dim d1() As Byte
    d1 = Arrays.CreateSafeByteArray(Length)
    If keepLoadedData Then
        If dlen > 0 Then
            Call memcpy(ByVal ArrayDataPtr(d1), ByVal ArrayDataPtr(d), dlen)
            Call memzero(ByVal API_VarPtrArray(d), VLEN_PTR)
        End If
    End If
    Call memcpy(ByVal API_VarPtrArray(d), ByVal API_VarPtrArray(d1), VLEN_PTR)
    Call memzero(ByVal API_VarPtrArray(d1), VLEN_PTR)
    dlen = Length
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub
Public Sub SetBuffer(targetData, Optional Length As Long = -1)
    Dim bArray() As Byte
    bArray = mint_get_byte_array_of(targetData, Length)
    'gets all types of data into byte array.
    Dim tdlen As Long
    tdlen = ArraySize(bArray)
    If tdlen <= 0 Then
        Call Clear
        Exit Sub
    End If
    If dlen <> tdlen Then Call Reserve(tdlen)
    Dim i As Long
    For i = 0 To dlen - 1
        d(i) = bArray(i)
    Next
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub
Public Sub SetBoundedBuffer(targetByteArray() As Byte)
    Dim blen As Long
    blen = ArraySize(targetByteArray)
    If blen <= 0 Then throw Exps.InvalidArgumentValueException("targetByteArray Is Null.")
    If LBound(targetByteArray) <> 0 Then throw Exps.InvalidArrayLBoundException("Invalid targetByteArray LBound.")
    Erase d
    dlen = 0
    Call API_CopyMemory(ByVal API_VarPtrArray(d), ByVal API_VarPtrArray(targetByteArray), 4)
    dlen = blen
    Call Signal.Emit(Me, SIGNAL_BUFFERBOUNDED)
End Sub

Public Function IsArrayBounded() As Boolean
    IsArrayBounded = IsBounded
End Function
Public Function MakeBufferInternal() As Boolean
    If Not IsBounded Then GoTo exitfunctionimmediately
    d = d
exitfunctionimmediately:
    MakeBufferInternal = IsBounded
    Call Signal.Emit(Me, SIGNAL_BUFFERUNBOUNDED)
End Function

Public Sub Append(target, Optional Length As Long = -1)
    Dim targetArray() As Byte
    Dim blen As Long
    targetArray = mint_get_byte_array_of(target, Length)
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
    Call Signal.Emit(Me, SIGNAL_CHANGED)
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
    If dlen < tdlen Then Call Reserve(tdlen)

    Dim i As Long
    For i = 0 To tdlen - 1
        d(i) = B(i)
    Next
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub

Public Function GetByteAt(Index As Long) As Byte
    If (Index < 0) Or (Index >= dlen) Then throw Exps.OutOfRangeException, Me
    GetByteAt = d(Index)
End Function
Public Sub SetByteAt(Index As Long, Value As Byte)
    If (Index < 0) Or (Index >= dlen) Then throw Exps.OutOfRangeException, Me
    If d(Index) = Value Then Exit Sub
    d(Index) = Value
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub

Public Property Get ByteAt(Index As Long) As Byte
    If (Index < 0) Or (Index >= dlen) Then throw Exps.OutOfRangeException, Me
    ByteAt = d(Index)
End Property
Public Property Let ByteAt(Index As Long, Value As Byte)
    If (Index < 0) Or (Index >= dlen) Then throw Exps.OutOfRangeException, Me
    If d(Index) = Value Then Exit Property
    d(Index) = Value
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Property

Public Property Get First() As Byte
    First = d(0)
End Property
Public Property Let First(Value As Byte)
    If d(0) = Value Then Exit Property
    d(0) = Value
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Property
Public Property Get Last() As Byte
    If dlen <= 0 Then throw Exps.OutOfRangeException, Me
    Last = d(dlen - 1)
End Property
Public Property Let Last(Value As Byte)
    If dlen <= 0 Then throw Exps.OutOfRangeException, Me
    If d(dlen - 1) = Value Then Exit Property
    d(dlen - 1) = Value
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Property

Public Property Get ExtraBuffer() As Long
    ExtraBuffer = pExtraBuffer
End Property
Public Property Let ExtraBuffer(ByVal Value As Long)
    pExtraBuffer = Value
End Property

Public Function SubArray(StartIndex As Long, Optional Length As Long = -1) As ByteArray
    If ArraySize(d) = 0 Then throw Exps.InvalidCallException("ByteArray Buffer Is Empty.")
    Dim BA As New ByteArray
    Call BA.SetBuffer(GetSubByteArray(d, StartIndex, Length))
    Set SubArray = BA
End Function

Public Sub DropLast(Length As Long)
    Dim rtLength As Long
    rtLength = dlen - Length
    If rtLength = 0 Then Exit Sub
    If rtLength < 0 Then throw Exps.OutOfRangeException
    ReDim Preserve d(rtLength - 1)
    dlen = ArraySize(d)
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub

Public Sub Replace(ByVal btaWhat As ByteArray, ByVal btaWith As ByteArray, Optional ByVal StartIndex As Long = 1, Optional ByVal Count As Long = -1)
    
End Sub
Public Function IndexOf(ByVal btaWhat As ByteArray, Optional ByVal StartIndex As Long = 0) As Long
    IndexOf = IndexOfByteArray(d, btaWhat.constData, StartIndex)
End Function
Public Function LastIndexOf(ByVal btaWhat As ByteArray, Optional ByVal StartIndex As Long = -1) As Long
    LastIndexOf = LastIndexOfByteArray(d, btaWhat.constData, StartIndex)
End Function
Public Function HowManyOf(ByVal btaWhat As ByteArray, Optional ByVal StartIndex As Long = 1) As Long

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

Public Sub Join(targetArray, Optional Delimiter)
    
End Sub
Public Function Split(Optional Delimiter, Optional ByVal Limit As Long = -1) As ByteArray()

End Function
Public Sub Concat(ByVal btaConcat As ByteArray, Optional ByVal Separator As String = "")

End Sub
Public Sub Format(ByVal btaFormat As ByteArray)

End Sub
Public Function Compare(ByVal btaWhat As ByteArray) As Boolean

End Function
Public Sub Reverse(Optional Length As Long = -1)
    If dlen = 0 Then Exit Sub
    If Length = 0 Then Exit Sub
    If Length < -1 Then throw Exps.OutOfRangeException
    If Length > dlen Then throw Exps.OutOfRangeException
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
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub
Public Sub Remove(StartIndex As Long, Optional Length As Long = -1)
    If StartIndex < 0 Then throw Exps.OutOfRangeException
    If StartIndex >= dlen Then throw Exps.OutOfRangeException
    If Length > dlen - StartIndex Then throw Exps.OutOfRangeException
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
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub
Public Sub InsertAt(Index As Long, ByteValue As Byte)
    Dim i As Long
    ReDim Preserve d(dlen)
    For i = dlen To Index + 1 Step -1
        d(i) = d(i - 1)
    Next
    d(Index) = ByteValue

    dlen = dlen + 1
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub
Public Sub Insert(Index As Long, targetBA As ByteArray, Optional Length As Long = -1)
    If targetBA Is Nothing Then throw Exps.ArgumentNullException
    Call InsertArrayIndexArray(d, Index, targetBA.constData, Length)
    dlen = ArraySize(d)
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub
Public Sub PadLeft(ByVal TotalLength As Long, ByVal FillChars As ByteArray)

End Sub
Public Sub PadRight(ByVal TotalLength As Long, ByVal FillChars As ByteArray)

End Sub
Public Function Contains(ByVal SearchWhat As ByteArray, Optional ByVal StartIndex As Long = 0) As Boolean

End Function
Public Function StartsWith(ByVal btaWhat As ByteArray, Optional ByVal StartIndex As Long = 0) As Boolean

End Function
Public Function EndsWith(ByVal btaWhat As ByteArray, Optional ByVal LastStartIndex As Long = -1) As Boolean

End Function
Public Sub ShiftLeft(ByVal StartIndex As Long, ByVal TimesToShift As Long, Optional ByVal Length As Long = -1)
    If StartIndex < 0 Then throw Exps.OutOfRangeException
    If StartIndex >= dlen Then throw Exps.OutOfRangeException
    If Length > dlen - StartIndex Then throw Exps.OutOfRangeException
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
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub
Public Sub ShiftRight(ByVal StartIndex As Long, ByVal TimesToShift As Long, Optional ByVal Length As Long = -1)
    If StartIndex < 0 Then throw Exps.OutOfRangeException
    If StartIndex >= dlen Then throw Exps.OutOfRangeException
    If StartIndex - Length < 0 Then throw Exps.OutOfRangeException
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
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub

Public Sub toLower()
    Dim i As Long
    For i = 0 To dlen - 1
        If d(i) >= 65 And d(i) <= 90 Then
            d(i) = d(i) + 32
        End If
    Next
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub
Public Sub toUpper()
    Dim i As Long
    For i = 0 To dlen - 1
        If d(i) >= 97 And d(i) <= 122 Then
            d(i) = d(i) - 32
        End If
    Next
    Call Signal.Emit(Me, SIGNAL_CHANGED)
End Sub


Public Function ToBigNumber() As BigNumber
    If dlen = 0 Then Exit Function

End Function
Public Function ToInt() As Integer
    If dlen = 0 Then Exit Function
    If dlen <= 2 Then
        Call ByteArrayToMemory(API_VarPtr(ToInt), d, VLEN_INTEGER)
    Else
        throw Exps.OutOfRangeException
    End If
End Function
Public Function ToBoolean() As Boolean
    If dlen = 0 Then Exit Function
    If dlen <= 2 Then
        Call ByteArrayToMemory(API_VarPtr(ToBoolean), d, VLEN_BOOLEAN)
    Else
        throw Exps.OutOfRangeException
    End If
End Function
Public Function ToString() As String
    If dlen = 0 Then Exit Function
    ToString = ByteArrayToString(d)
End Function
Public Function ToLPSTR() As String
    If dlen = 0 Then Exit Function
    ToLPSTR = ByteArrayToSafeString(d)
End Function
Public Function ToLong() As Long
    If dlen = 0 Then Exit Function
    If dlen <= 4 Then
        Call ByteArrayToMemory(API_VarPtr(ToLong), d, VLEN_LONG)
    Else
        throw Exps.OutOfRangeException
    End If
End Function
Public Function ToDouble() As Double
    If dlen = 0 Then Exit Function
    If dlen <= 8 Then
        Call ByteArrayToMemory(API_VarPtr(ToDouble), d, VLEN_DOUBLE)
    Else
        throw Exps.OutOfRangeException
    End If
End Function
Public Function ToFloat() As Single
    If dlen = 0 Then Exit Function
    If dlen <= 4 Then
        Call ByteArrayToMemory(API_VarPtr(ToFloat), d, VLEN_SINGLE)
    Else
        throw Exps.OutOfRangeException
    End If
End Function
Public Function ToSingle() As Single
    If dlen = 0 Then Exit Function
    If dlen <= 4 Then
        Call ByteArrayToMemory(API_VarPtr(ToSingle), d, VLEN_SINGLE)
    Else
        throw Exps.OutOfRangeException
    End If
End Function
Public Function ToDate() As Date
    If dlen = 0 Then Exit Function
    If dlen <= 8 Then
        Call ByteArrayToMemory(API_VarPtr(ToDate), d, VLEN_DATE)
    Else
        throw Exps.OutOfRangeException
    End If
End Function
Public Function ToCurrency() As Currency
    If dlen = 0 Then Exit Function
    If dlen <= 8 Then
        Call ByteArrayToMemory(API_VarPtr(ToCurrency), d, VLEN_CURRENCY)
    Else
        throw Exps.OutOfRangeException
    End If
End Function

Public Function FromInt(Value As Integer) As ByteArray
    Set FromInt = New ByteArray
    Call FromInt.SetBuffer(Value)
End Function
Public Function FromBoolean(Value As Boolean) As ByteArray
    Set FromBoolean = New ByteArray
    Call FromBoolean.SetBuffer(Value)
End Function
Public Function FromString(Value As String) As ByteArray
    Set FromString = New ByteArray
    Call FromString.SetBuffer(Value)
End Function
Public Function FromLong(Value As Long) As ByteArray
    Set FromLong = New ByteArray
    Call FromLong.SetBuffer(Value)
End Function
Public Function FromDouble(Value As Double) As ByteArray
    Set FromDouble = New ByteArray
    Call FromDouble.SetBuffer(Value)
End Function
Public Function FromFloat(Value As Single) As ByteArray
    Set FromFloat = New ByteArray
    Call FromFloat.SetBuffer(Value)
End Function
Public Function FromSingle(Value As Single) As ByteArray
    Set FromSingle = New ByteArray
    Call FromSingle.SetBuffer(Value)
End Function
Public Function FromDate(Value As Date) As ByteArray
    Set FromDate = New ByteArray
    Call FromDate.SetBuffer(Value)
End Function
Public Function FromCurrency(Value As Currency) As ByteArray
    Set FromCurrency = New ByteArray
    Call FromCurrency.SetBuffer(Value)
End Function

Public Function DataMemory() As Memory
    'Set DataMemory = dAG.DataMemory
End Function

Public Function ToHex() As ByteArray
    'toHex = Hex$(d)
End Function
Public Function ToOctal() As ByteArray
    'toOctal = Oct$(d)
End Function
Public Function ToBinary() As ByteArray ' binary number 1010001011101
    '
End Function
Public Function FromHex(Value As ByteArray) As ByteArray
    '
End Function
Public Function FromOctal(Value As ByteArray) As ByteArray
    '
End Function
Public Function FromBinary(Value As ByteArray) As ByteArray ' binary number 1010001011101
    '
End Function
Public Function IsNumeric() As Boolean
    IsNumeric = IsByteArrayNumeric(d)
End Function
Public Function IsAlphabetic() As Boolean
    IsAlphabetic = IsByteArrayAlphabetic(d)
End Function
Public Function IsLike(ByVal Value2 As ByteArray) As Boolean
    IsLike = IsByteArrayLikeAnother(d, Value2.GetByteArray)
End Function
Public Function IsEqualTo(ByVal Value2 As ByteArray, Optional ByVal LengthToCompare As Long = -1) As Boolean
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
Private Function checkifspc(B() As Byte, ByVal ClassName As String) As Boolean
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

    afterCLSNameLength = ByteArray(B).ToInt
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

    afterCLSNameLength = ByteArray(ln).ToInt
    If afterCLSNameLength <= 0 Then GoTo returnN

    ln = GetSubByteArray(B, afterCLSNameLength + 11, 2)
    innerDataLength = ByteArray(ln).ToInt

        Call API_CopyMemory(ByVal API_VarPtrArray(GetByteArrayInnerData), ByVal API_VarPtrArray(GetSubByteArray(B, afterCLSNameLength + 12, innerDataLength)), 4)

    Exit Function
returnN:
End Function

Public Function Clone() As ByteArray
    Dim clonedByteArray As New ByteArray
    Call clonedByteArray.Constructor3(d)
    Set Clone = clonedByteArray
End Function


Private Sub IData_Append(target As Variant): Call Append(target): End Sub

Private Sub IData_Clear(): Call Clear: End Sub
Private Sub IData_CopyByteArray(targetByteArray() As Byte): Call CopyByteArray(targetByteArray): End Sub

Private Function IData_GetByteArray() As Byte()
    Call API_CopyMemory(ByVal API_VarPtrArray(IData_GetByteArray), ByVal API_VarPtrArray(d), 4)
End Function
Private Property Get IData_Length() As Long: IData_Length = dlen: End Property

Private Sub Class_Terminate()
    Call API_ZeroMemory(ByVal API_VarPtrArray(d), 4)
End Sub

Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function


Private Function GetEnumerator() As IEnumerator
    Dim BAEnum As New ByteArrayEnumerator
    Call BAEnum.Constructor0(Me)
    Set GetEnumerator = BAEnum
End Function
Private Function IEnumerable_GetEnumerator() As IEnumerator
    Set IEnumerable_GetEnumerator = GetEnumerator
End Function
Public Function NewEnum() As IUnknown
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

Private Function IEnumerable_NewEnum() As stdole.IUnknown
    Set IEnumerable_NewEnum = CreateEnumerator(GetEnumerator)
End Function

Private Function ITargetStream_Available(Dir As StreamDirection) As Long
    Dim Val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        Val = ((dlen > 0) And inState)
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        Val = Val Or (outState)
    End If
    ITargetStream_Available = Val
End Function

Private Sub ITargetStream_CloseStream(Optional Direction As StreamDirection = StreamDirection.sdBoth)
    If (Direction And sdInStream) = sdInStream Then
        inState = False
    End If
    If (Direction And sdOutStream) = sdOutStream Then
        outState = False
    End If
    If Direction = sdBoth Then Call Clear
End Sub
Private Sub ITargetStream_OpenStream(Optional Direction As StreamDirection = StreamDirection.sdBoth)
    If (Direction And sdInStream) = sdInStream Then
        inState = True
    End If
    If (Direction And sdOutStream) = sdOutStream Then
        outState = True
    End If
End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean
    Dim Val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        Val = inState
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        Val = Val Or outState
    End If
    ITargetStream_getState = Val
End Function

Private Sub ITargetStream_inStream(Data, Optional Length As Long = -1)
    If Not inState Then throw Exps.InvalidStatusException
    'Call mint_put_byte_array_to(Data, d, Length)
End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    ITargetStream_readBytes = GetSubByteArray(d, StartAt, Length)
End Function
Private Sub ITargetStream_outStream(Data, Optional Length As Long = -1)
    If Not outState Then throw Exps.InvalidStatusException
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
                        throw Exps.InvalidArgumentTypeException("Streaming Command '" & strFirst & "' Is Not Available For ByteArray.")
                    End If
                Else
                    throw Exps.InvalidArgumentTypeException("Streaming Type Is Not Available For ByteArray.")
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
