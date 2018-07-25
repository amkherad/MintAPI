VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Memory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "Memory"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements ITargetStream
Implements IClassClonable
Implements IClassTexer
Implements IClassExceptionManager
Implements IAliasable

Private Declare Function API_HeapAlloc Lib "kernel32" Alias "HeapAlloc" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function API_HeapCompact Lib "kernel32" Alias "HeapCompact" (ByVal hHeap As Long, ByVal dwFlags As Long) As Long
Private Declare Function API_HeapCreate Lib "kernel32" Alias "HeapCreate" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Private Declare Function API_HeapDestroy Lib "kernel32" Alias "HeapDestroy" (ByVal hHeap As Long) As Long
Private Declare Function API_HeapFree Lib "kernel32" Alias "HeapFree" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function API_HeapLock Lib "kernel32" Alias "HeapLock" (ByVal hHeap As Long) As Long
Private Declare Function API_HeapReAlloc Lib "kernel32" Alias "HeapReAlloc" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long, ByVal dwBytes As Long) As Long
Private Declare Function API_HeapSize Lib "kernel32" Alias "HeapSize" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function API_HeapUnlock Lib "kernel32" Alias "HeapUnlock" (ByVal hHeap As Long) As Long
Private Declare Function API_HeapValidate Lib "kernel32" Alias "HeapValidate" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
'Private Declare Function API_HeapWalk Lib "Kernel32" Alias "HeapWalk" (ByVal hHeap As Long, ByRef lpEntry As PROCESS_HEAP_ENTRY) As Long
'Private Declare Function API_Heap32Next Lib "Kernel32" Alias "Heap32Next" (ByRef lphe As LPHEAPENTRY32) As Long
'Private Declare Function API_Heap32ListNext Lib "Kernel32" Alias "Heap32ListNext" (ByVal hSnapshot As Long, ByRef lphl As LPHEAPLIST32) As Long
'Private Declare Function API_Heap32ListFirst Lib "Kernel32" Alias "Heap32ListFirst" (ByVal hSnapshot As Long, ByRef lphl As LPHEAPLIST32) As Long
'Private Declare Function API_Heap32First Lib "Kernel32" Alias "Heap32First" (ByRef lphe As LPHEAPENTRY32, ByVal th32ProcessID As Long, ByRef th32HeapID As ULONG_PTR) As Long


Dim mhndl As Long 'memory heap handle.
Dim mslen As Long 'memory real segment length.
Dim mlen As Long 'memory evaluated length.
Dim is_dynS As Boolean
Dim allow_SubSegment As Boolean


Friend Sub Initialize(Optional target)
    
End Sub

Public Property Get Length() As Long
    Length = mlen
End Property
Public Property Let Length(Value As Long)
    Call reserve(Value)
End Property
Public Property Get SegmentLength() As Long
    SegmentLength = mslen
End Property

Public Property Get AllowSubSegment() As Boolean
    AllowSubSegment = allow_SubSegment
End Property
Public Property Let AllowSubSegment(Value As Boolean)
    If Not Value Then
        If allow_SubSegment Then
            
        End If
    Else
        If Not allow_SubSegment Then
            
        End If
    End If
    allow_SubSegment = Value
End Property

Public Property Get IsDynamicSegment() As Boolean
    IsDynamicSegment = is_dynS
End Property
Public Property Let IsDynamicSegment(Value As Boolean)
    is_dynS = Value
End Property

'Clear the memory data and turn them to zeros.
Public Sub Clear()
    'is_dynS
End Sub
'Dispose all memory data and reserve nothing.
Public Sub Dispose()
    'is_dynS
End Sub
'Reserves a static memory segment in heap.
Public Sub reserve(Length As Long, Optional keepLoadedData As Boolean = False)
    If Length < 0 Then throw NegativeArgumentException("Length can't be negative.")
    'is_dynS
End Sub
'Reserves a memory segment in heap, _
 if existing segment have been bigger than requested length _
 existing segment will be used and remained size will fill with zeros.
Public Sub reserveSubSegment(Length As Long, Optional keepLoadedData As Boolean = False)
    If (Not allow_SubSegment) Or (Length > mslen) Then
        Call reserve(Length, keepLoadedData)
        Exit Sub
    End If
    
    
End Sub

Public Property Get IsEmpty() As Boolean
    IsEmpty = (mlen <= 0)
End Property
Public Property Get IsFilled() As Boolean
    IsFilled = (mslen = mlen)
End Property
Public Property Get IsReserved() As Boolean
    IsReserved = ((mhndl <> 0) And (mslen > 0))
End Property
Public Function CanGrowTo(Length As Long) As Boolean
    If Length < 0 Then CanGrowTo = False: Exit Function
    If is_dynS Then _
        CanGrowTo = True: Exit Function
    If allow_SubSegment Then _
        CanGrowTo = (Length <= mslen): Exit Function
    CanGrowTo = False
End Function
Public Function CanShrinkTo(Length As Long) As Boolean
    If Length < 0 Then CanShrinkTo = False: Exit Function
    If (is_dynS Or allow_SubSegment) Then _
        CanShrinkTo = True
End Function

Public Sub Append(target)
    If Not allow_SubSegment Then throw Exception("Subsegment state is not activated.")
    
End Sub
'Evaluate the reserved buffer.
Public Sub setBuffer(targetData, Optional UseExistedBuffer As Boolean = False)
    Dim bArray() As Byte
    bArray = mint_get_byte_array_of(targetData)
    'gets all types of data into byte array.
    
End Sub
Public Function constData() As Long
    
End Function
Public Function GetByteArray() As Byte()
    
End Function
Public Sub CopyByteArray(targetByteArray() As Byte)
    
End Sub
Public Sub SetByteArray(B() As Byte)
    
End Sub
Public Function toByteArray() As ByteArray
    
End Function

Public Function GetByteAt(Index As Long) As Byte
    If (Index < 0) Or (Index >= Length) Then throw OutOfRangeException, Me
    
End Function
Public Sub SetByteAt(Index As Long, Value As Byte)
    If (Index < 0) Or (Index >= Length) Then throw OutOfRangeException, Me
    
End Sub
Public Property Get ByteAt(Index As Long) As Byte
    If (Index < 0) Or (Index >= Length) Then throw OutOfRangeException, Me
    
End Property
Public Property Let ByteAt(Index As Long, Value As Byte)
    If (Index < 0) Or (Index >= Length) Then throw OutOfRangeException, Me
    
End Property


Private Function IClassClonable_Clone() As Object
    
End Function

Private Sub IClassExceptionManager_setAllExceptionsState(Optional States As Boolean = True)

End Sub
Private Sub IClassExceptionManager_setExceptionState(exp As ExceptionType, Optional State As Boolean = True)

End Sub
Private Sub IClassTexer_fromByteArray(Value() As Byte)

End Sub
Private Sub IClassTexer_fromString(Value As String)

End Sub
Private Property Get IClassTexer_texHeader(texType As mint_TexingType) As String

End Property
Private Function IClassTexer_toByteArray() As Byte()

End Function
Private Function IClassTexer_toString() As String

End Function

Private Function ITargetStream_Available(direction As StreamDirection) As Long

End Function

Private Sub ITargetStream_CloseStream(Optional direction As StreamDirection = 3&)
    
End Sub
Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean

End Function
Private Sub ITargetStream_inStream(Data As Variant, Optional Length As Long = -1&)

End Sub
Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = 3&)

End Sub
Private Sub ITargetStream_outStream(Data As Variant, Optional Length As Long = -1&)

End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()

End Function
Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)

End Sub
Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)

End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
