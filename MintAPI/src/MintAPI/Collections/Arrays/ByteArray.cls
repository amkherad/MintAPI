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

''<class name="ByteArray" tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
Implements IMemoryBlock
Implements ICloneable
Implements IDisposable
Implements IEnumerable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ByteArray" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"

Private Const ReservedBufferGapLength As Long = 50 'Bytes
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables/Events>
'
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Private p_Mem As IMemoryBlock 'UnderlyingMemory
Private p_Capacity As Long
Private p_DataLength As Long
Private p_DoNotUseGap As Boolean
Private p_Disposed As Boolean
Private p_EnumVersion As Long
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Constructors>
'
''<summary>
''  Basic constructor to use by MintAPI class provider,
''  When creating new instance of this class.
''</summary>
''<params>
''  <param name="Arguments">
''      A list of construction arguments.
''  </param>
''</params>
''<cargs order="0" count="1">
''  <carg type="long" name="ID"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal InitialCapacity As Long)
    Capacity = InitialCapacity
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(Optional Expression, Optional ByVal ConvertToBinary As Boolean = True, Optional ByVal Length As Long = -1)
    Call Append(Expression)
    
    This_Constructed = True
End Sub
Friend Sub Constructor2(ByRef Arr As Variant)
    
    This_Constructed = True
End Sub
Friend Sub Constructor3(ByRef Arr As Variant)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As ByteArray)
'-------------------------
Locker:   LockObj Me
    On Error GoTo Unlocker
'=========================
    Dim Length As Long, Ptr As Long
    Length = That.Capacity
    ReserveGapBuffer = That.ReserveGapBuffer
    If Length <> NO_VALUE Then
        Ptr = That.DataHandle
        If Ptr <> vbNullPtr Then
            Capacity = Length
            Call memcpy(ByVal p_Mem.Reference, _
                        ByVal Ptr, Length)
        End If
    End If
    p_DataLength = Length
    
    This_Constructed = True
'=========================
ExitLock: EndLockObj Me
    Exit Sub
Unlocker: EndLockObj Me
    rethrow
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

Public Function GetEnumerator() As IEnumerator
THROW_IfNull
    Dim BAEnumerator As New ByteArrayEnumerator
    Call BAEnumerator.Constructor0(Me)
    Set GetEnumerator = BAEnumerator
End Function
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
THROW_IfNull
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

Friend Property Get UnderlyingMemoryBlock() As IMemoryBlock
THROW_IfNull
    Set UnderlyingMemoryBlock = p_Mem
End Property

Public Property Get Capacity() As Long
THROW_IfNull
    Capacity = p_Capacity
End Property
Public Property Let Capacity(ByVal Length As Long)
THROW_IfNull
    If Length < 0 Then throw Exps.InvalidArgumentException("Length")
    If Length = 0 Then throw Exps.InvalidArgumentException("Length")
    
    If p_Mem Is Nothing Then
        p_Capacity = IIf(p_DoNotUseGap, Length, Length + ReservedBufferGapLength)
        Set p_Mem = Memory.Allocate(p_Capacity)
    Else
        If Length <= p_Capacity Then Exit Property
        If Not p_DoNotUseGap Then
            If (p_Capacity - Length) <= ReservedBufferGapLength Then Exit Property
            p_Capacity = Length + ReservedBufferGapLength
        Else
            p_Capacity = Length
        End If
'-------------------------
Locker:   LockObj Me
    On Error GoTo Unlocker
'=========================
        Call p_Mem.Resize(p_Capacity)
'=========================
ExitLock: EndLockObj Me
        GoTo Continue
Unlocker: EndLockObj Me
        rethrow
    End If
Continue:
    p_EnumVersion = p_EnumVersion + 1
End Property

Public Property Get ReserveGapBuffer() As Boolean
THROW_IfNull
    ReserveGapBuffer = Not p_DoNotUseGap
End Property
Public Property Let ReserveGapBuffer(ByVal Value As Boolean)
THROW_IfNull
    p_DoNotUseGap = Not Value
End Property

Public Property Get DefaultBufferGapLength() As Long
    DefaultBufferGapLength = ReservedBufferGapLength
End Property

Public Property Get Length() As Long
THROW_IfNull
    Length = p_DataLength
End Property

Public Property Get IsEmpty() As Boolean
THROW_IfNull
    IsEmpty = (p_DataLength = 0)
End Property

Public Sub Clear()
THROW_IfNull
    Dim Disposable As IDisposable
    Set Disposable = p_Mem
    Call Disposable.Dispose
    Set p_Mem = Memory.Allocate(p_Capacity)
    p_DataLength = 0
End Sub

Public Property Get ByteAt(ByVal Index As Long) As Byte
    
End Property
Public Property Let ByteAt(ByVal Index As Long, ByVal Value As Byte)
    
End Property

Public Property Get WordAt(ByVal Index As Long) As Integer
    
End Property
Public Property Let WordAt(ByVal Index As Long, ByVal Value As Integer)
    
End Property

Public Property Get LongAt(ByVal Index As Long) As Long
    
End Property
Public Property Let LongAt(ByVal Index As Long, ByVal Value As Long)
    
End Property

Public Function Append(ByRef Expression As Variant, Optional ByVal StartIndex As Long = 0, Optional ByVal Length As Long = -1) As ByteArray
THROW_IfNull
    If Length < -1 Then throw Exps.IndexOutOfRangeException
    If StartIndex < 0 Then throw Exps.IndexOutOfRangeException
    
    Set Append = Me 'Only Append() returns current instance of the class.
    
    If Length = 0 Then Exit Function
    If VBA.IsEmpty(Expression) Then Exit Function
    If VBA.IsNull(Expression) Then Exit Function
    
    
    Dim BA() As Byte, Size As Long
    'BA = mint_get_byte_array_of(Expression)
    Dim SourcePtr As Long
    
    If IsObject(Expression) Then
        Dim Obj As Object
        Set Obj = Expression
        If TypeOf Obj Is ByteArray Then
            Dim ByteA As ByteArray
            Set ByteA = Obj
            BA = ByteA.ToConstData
        Else
            BA = mint_binaryserialize(Expression)
        End If
        Size = ArraySize(BA)
        SourcePtr = Arrays.GetDataPointerOf(BA)
    Else
        Size = Info.GetArgumentLength(Expression)
        SourcePtr = VariantAbsoluteDataPtr(Expression)
    End If
    
    If SourcePtr = vbNullPtr Then throw Exps.InvalidOperationException
    
    If Length = -1 Then Length = Size - StartIndex
    If (StartIndex + Length) > Size Then throw Exps.IndexOutOfRangeException
    
'-------------------------
Locker:   LockObj Me
    On Error GoTo Unlocker
'=========================
    Capacity = (p_DataLength + Length)
    
    p_EnumVersion = p_EnumVersion + 1
    
    Call memcpy(ByVal (p_Mem.Reference + p_DataLength), _
                ByVal (StartIndex + SourcePtr), Length)
    
    p_DataLength = p_DataLength + Length
    
'=========================
ExitLock: EndLockObj Me
    Exit Function
Unlocker: EndLockObj Me
    rethrow
End Function
Public Function AppendLine() As ByteArray
    Set AppendLine = Append(sNewLine)
End Function

Public Function Insert(ByVal Index As Long, ByVal TargetBA As ByteArray, Optional ByVal Length As Long = -1) As ByteArray
THROW_IfNull
    
End Function

Public Sub Replace(ByVal btaWhat As ByteArray, ByVal btaWith As ByteArray, Optional ByVal StartIndex As Long = 1, Optional ByVal Count As Long = -1)
THROW_IfNull
    
End Sub

Public Function SubArray(ByVal Start As Long, Optional ByVal Length As Long = -1) As ByteArray
THROW_IfNull
    If (Start < 0) Or (Length < -1) Or (Start >= p_DataLength) Then throw Exps.IndexOutOfRangeException
    If Length <> -1 Then
        If (Start + Length) >= p_DataLength Then throw Exps.IndexOutOfRangeException
    End If
    Set SubArray = New ByteArray
    If Length = 0 Then Exit Function
    
    Dim outLen As Long
    If Length = -1 Then
        outLen = p_DataLength - Start
    Else
        outLen = Length
    End If
    SubArray.Capacity = outLen
    
    Call memcpy(ByVal SubArray.UnderlyingMemoryBlock.Reference, _
                ByVal p_Mem.Reference + Start, outLen)
End Function

Public Function IndexOf(ByVal btaWhat As ByteArray, Optional ByVal StartIndex As Long = 0) As Long
THROW_IfNull
    
End Function
Public Function LastIndexOf(ByVal btaWhat As ByteArray, Optional ByVal StartIndex As Long = -1) As Long
THROW_IfNull
    
End Function
Public Function HowManyOf(ByVal btaWhat As ByteArray, Optional ByVal StartIndex As Long = 1) As Long
THROW_IfNull
    
End Function
Public Function Split(Optional Delimiter, Optional ByVal Limit As Long = -1) As ByteArray()
THROW_IfNull
    
End Function
Public Function Compare(ByVal btaWhat As ByteArray) As Boolean
THROW_IfNull
    
End Function
Public Function Contains(ByVal SearchWhat As ByteArray, Optional ByVal StartIndex As Long = 0) As Boolean
THROW_IfNull
    
End Function
Public Function StartsWith(ByVal btaWhat As ByteArray, Optional ByVal StartIndex As Long = 0) As Boolean
THROW_IfNull
    
End Function
Public Function EndsWith(ByVal btaWhat As ByteArray, Optional ByVal LastStartIndex As Long = -1) As Boolean
THROW_IfNull
    
End Function

Public Function Trim() As ByteArray
THROW_IfNull
    
End Function
Public Function TrimLeft() As ByteArray
THROW_IfNull
    
End Function
Public Function TrimRight() As ByteArray
THROW_IfNull
    
End Function
Public Function TrimArr(ByRef TrimBytes() As ByteArray) As ByteArray
THROW_IfNull
    
End Function
Public Function TrimLeftArr(ByRef TrimBytes() As ByteArray) As ByteArray
THROW_IfNull
    
End Function
Public Function TrimRightArr(ByRef TrimBytes() As ByteArray) As ByteArray
THROW_IfNull
    
End Function

Public Function Join(ByVal BA As ByteArray, Optional ByVal Delimiter As ByteArray) As ByteArray
THROW_IfNull
    
End Function
Public Function Concat(ByVal Value As ByteArray, Optional ByVal Separator As String = "") As ByteArray
THROW_IfNull
    
End Function
Public Function Format(ByVal Pattern As ByteArray) As ByteArray
THROW_IfNull
    
End Function
Public Function Remove(ByVal StartIndex As Long, Optional ByVal Length As Long = -1) As ByteArray
THROW_IfNull
    
End Function
Public Function InsertAt(ByVal Index As Long, ByVal ByteValue As Byte) As ByteArray
THROW_IfNull
    
End Function
Public Function PadLeft(ByVal TotalLength As Long, ByVal FillChars As ByteArray) As ByteArray
THROW_IfNull
    
End Function
Public Function PadRight(ByVal TotalLength As Long, ByVal FillChars As ByteArray) As ByteArray
THROW_IfNull
    
End Function
Public Function ShiftLeft(ByVal StartIndex As Long, ByVal TimesToShift As Long, Optional ByVal Length As Long = -1) As ByteArray
THROW_IfNull
    
End Function
Public Function ShiftRight(ByVal StartIndex As Long, ByVal TimesToShift As Long, Optional ByVal Length As Long = -1) As ByteArray
THROW_IfNull
    
End Function

'Public Function ToLower() As ByteArray
'THROW_IfNull
'
'End Function
'Public Function ToUpper() As ByteArray
'THROW_IfNull
'
'End Function


'Public Function ToInt() As Integer
'THROW_IfNull
'    If p_DataLength = 0 Then Exit Function
'    If p_DataLength <= VLEN_INTEGER Then
'        Call memcpy(VarPtr(ToInt), ByVal p_Mem.Reference, p_DataLength)
'    Else
'        throw Exps.InvalidOperationException
'    End If
'End Function
'Public Function ToBoolean() As Boolean
'THROW_IfNull
'    If p_DataLength = 0 Then Exit Function
'    If p_DataLength <= VLEN_BOOLEAN Then
'        Call memcpy(VarPtr(ToBoolean), ByVal p_Mem.Reference, p_DataLength)
'    Else
'        throw Exps.InvalidOperationException
'    End If
'End Function
'Public Function ToLPSTR(Optional ByVal TreatAsUnicode As Boolean = False) As String
'THROW_IfNull
'    If p_DataLength = 0 Then Exit Function
'
'    If TreatAsUnicode Then
'
'    Else
'
'    End If
'End Function
'Public Function ToLong() As Long
'THROW_IfNull
'    If p_DataLength = 0 Then Exit Function
'    If p_DataLength <= VLEN_LONG Then
'        Call memcpy(API_VarPtr(ToLong), ByVal p_Mem.Reference, p_DataLength)
'    Else
'        throw Exps.InvalidOperationException
'    End If
'End Function
'Public Function ToDouble() As Double
'THROW_IfNull
'    If p_DataLength = 0 Then Exit Function
'    If p_DataLength = VLEN_DOUBLE Then
'        Call memcpy(API_VarPtr(ToDouble), ByVal p_Mem.Reference, p_DataLength)
'    Else
'        throw Exps.InvalidOperationException
'    End If
'End Function
'Public Function ToFloat() As Single
'THROW_IfNull
'    If p_DataLength = 0 Then Exit Function
'    If p_DataLength = VLEN_SINGLE Then
'        Call memcpy(API_VarPtr(ToFloat), ByVal p_Mem.Reference, p_DataLength)
'    Else
'        throw Exps.InvalidOperationException
'    End If
'End Function
'Public Function ToSingle() As Single
'THROW_IfNull
'    If p_DataLength = 0 Then Exit Function
'    If p_DataLength = VLEN_SINGLE Then
'        Call memcpy(API_VarPtr(ToSingle), ByVal p_Mem.Reference, p_DataLength)
'    Else
'        throw Exps.InvalidOperationException
'    End If
'End Function
'Public Function ToDate() As Date
'THROW_IfNull
'    If p_DataLength = 0 Then Exit Function
'    If p_DataLength <= VLEN_DATE Then
'        Call memcpy(API_VarPtr(ToDate), ByVal p_Mem.Reference, p_DataLength)
'    Else
'        throw Exps.InvalidOperationException
'    End If
'End Function
'Public Function ToCurrency() As Currency
'THROW_IfNull
'    If p_DataLength = 0 Then Exit Function
'    If p_DataLength <= VLEN_CURRENCY Then
'        Call memcpy(API_VarPtr(ToCurrency), ByVal p_Mem.Reference, p_DataLength)
'    Else
'        throw Exps.InvalidOperationException
'    End If
'End Function

Public Function LockToMemory() As IMemoryBlock
THROW_IfNull
'-------------------------
Locker:   LockObj Me
    On Error GoTo Unlocker
'=========================
    
    Set LockToMemory = Memory.Allocate(p_DataLength)
    Call memcpy(ByVal LockToMemory.Reference, _
                ByVal p_Mem.Reference, p_DataLength)
    
'=========================
ExitLock: EndLockObj Me
    Exit Function
Unlocker: EndLockObj Me
    rethrow
End Function
Public Function ToConstData() As Byte()
THROW_IfNull
'-------------------------
Locker:   LockObj Me
    On Error GoTo Unlocker
'=========================
    
    ToConstData = Arrays.CreateSafeByteArray(p_DataLength)
    Call memcpy(ByVal Arrays.GetDataPointerOf(ToConstData), _
                ByVal p_Mem.Reference, p_DataLength)
    
'=========================
ExitLock: EndLockObj Me
    Exit Function
Unlocker: EndLockObj Me
    rethrow
End Function
Friend Function OriginalDataDelegatorArray() As Byte()
THROW_IfNull
    
End Function

Friend Property Get DataHandle() As Long
THROW_IfNull
    If p_Mem Is Nothing Then
        DataHandle = vbNullPtr
    Else
        DataHandle = p_Mem.Reference
    End If
End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Basic Members>
'
Public Function Equals(Value As Variant) As Boolean
    Equals = Objects.Equals(Me, Value)
End Function
Public Function ToString() As String
THROW_IfNull
'-------------------------
Locker:   LockObj Me
    On Error GoTo Unlocker
'=========================
    
    ToString = mString.Create(p_DataLength)
    Call memcpy(ByVal StrPtr(ToString), _
                ByVal p_Mem.Reference, p_DataLength)
    
'=========================
ExitLock: EndLockObj Me
    Exit Function
Unlocker: EndLockObj Me
    rethrow
End Function
Public Function GetHashCode() As Long
    GetHashCode = Objects.GetHashCode(Me)
End Function
Public Function MetaObject() As MetaObject
    If myMetaObject Is Nothing Then
        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    End If
    Set MetaObject = myMetaObject
End Function

Public Function Clone() As ByteArray
THROW_IfNull
    Set Clone = New ByteArray
    Call Clone.ConstructorCopy(Me)
End Function

Public Sub Dispose()
    If Not p_Mem Is Nothing Then
        Dim IDisp As IDisposable
        Set IDisp = p_Mem
        Call IDisp.Dispose
    End If
    p_Capacity = 0
    p_DataLength = 0
    Set p_Mem = Nothing
    p_Disposed = True
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
'Private Sub Class_Initialize()
'    ...
'End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    Call Dispose
End Sub

Private Sub THROW_IfNull()
    If p_Disposed Then throw Exps.ObjectDisposedException
End Sub

Friend Function GetEnumerableVersion() As Long
    GetEnumerableVersion = p_EnumVersion
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo): Call Constructor(CInfo): End Sub
Private Function IObject_Equals(Value As Variant) As Boolean: IObject_Equals = Equals(Value): End Function
Private Function IObject_GetHashCode() As Long: IObject_GetHashCode = GetHashCode: End Function
Private Function IObject_MetaObject() As MetaObject: Set IObject_MetaObject = MetaObject: End Function
Private Function IObject_ToString() As String: IObject_ToString = ToString: End Function
Private Function IObject_Stub(ByVal ProxyInfo As ObjectProxyInfo) As Boolean
    
    IObject_Stub = False
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IDisposable Implementations>
'
Private Sub IDisposable_Dispose(): Call Dispose: End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IDisposable Implementations>
'
Private Function IEnumerable_GetEnumerator() As IEnumerator: Set IEnumerable_GetEnumerator = GetEnumerator: End Function
Private Function IEnumerable_NewEnum() As stdole.IUnknown: Set IEnumerable_NewEnum = NewEnum: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IMemoryBlock Implementations>
'
Private Property Get IMemoryBlock_Access() As MemoryAccessPermissions: IMemoryBlock_Access = mapReadWrite: End Property
Private Property Get IMemoryBlock_IsLimited() As Boolean: IMemoryBlock_IsLimited = True: End Property
Private Property Get IMemoryBlock_IsResizable() As Boolean: IMemoryBlock_IsResizable = False: End Property
Private Property Get IMemoryBlock_Limit() As Long: IMemoryBlock_Limit = Length: End Property
Private Sub IMemoryBlock_Resize(ByVal NewSize As Long): throw Exps.InvalidOperationException: End Sub
Private Property Get IMemoryBlock_Reference() As Long
    If p_Mem Is Nothing Then
        throw Exps.InvalidOperationException
    Else
        IMemoryBlock_Reference = p_Mem.Reference
    End If
End Property
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Buffer Implementations>
'
Private Property Get Buffer_AvailableBytes() As Long: Buffer_AvailableBytes = Length: End Property
Private Property Get Buffer_IsReadOnly() As Boolean: Buffer_IsReadOnly = False: End Property
Private Function Buffer_Peek(Optional ByVal Length As Long = -1&) As Buffer: Set Buffer_Pop = SubArray(0, Length): End Function
Private Function Buffer_Pop(Optional ByVal Length As Long = -1&) As Buffer: Set Buffer_Pop = SubArray(0, Length): End Function
Private Sub Buffer_Push(ByVal Buffer As Buffer): Call Append(Buffer): End Sub
Private Property Get Buffer_ReliabilityState() As StreamReliability: Buffer_ReliabilityState = srConstant: End Property
Private Function Buffer_ToConstData() As Byte(): Buffer_ToConstData = ToConstData: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
