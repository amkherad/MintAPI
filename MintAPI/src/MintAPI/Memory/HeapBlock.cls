VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HeapBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="HeapBlock" tags="">
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
Implements ICloneable
Implements IDisposable
Implements IMemoryBlock

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "HeapBlock" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"
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
Public Event BlockMoved(ByVal E As EventArgs)

Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Private WithEvents p_Heap As Heap
Attribute p_Heap.VB_VarHelpID = -1
Private p_Ref As Long
Private p_Length As Long
Private p_ForceSerialize As Boolean
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
Friend Sub Constructor0(ByVal Heap As Heap, ByVal Reference As Long, ByVal Length As Long)
    Set p_Heap = Heap
    p_Ref = Reference
    p_Length = Length
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As HeapBlock)
    
    This_Constructed = True
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

Public Property Get Reference() As Long
THROW_IfNull
    Reference = p_Ref
End Property

Public Property Get Heap() As Heap
THROW_IfNull
    Set Heap = p_Heap
End Property

''<remarks>If heap block has been disposed this property returns zero.</remarks>
Public Property Get Size() As Long
    Size = p_Length
End Property

Public Property Get IsResizable() As Boolean
THROW_IfNull
    IsResizable = True
End Property

Public Property Get ActualSize() As Long
THROW_IfNull
    Dim Flags As HeapFlags
    If Not p_ForceSerialize Then Flags = hfNoSerialize
    ActualSize = API_HeapSize(p_Heap.Handle, Flags, p_Ref)
End Property

Public Property Get Synchronize() As Boolean
    Synchronize = p_ForceSerialize
End Property
Public Property Let Synchronize(ByVal Value As Boolean)
THROW_IfNull
    p_ForceSerialize = Value
End Property

''<summary>This function Re-Allocates current heap memory block to the given new length.</summary>
''<params><param name="NewLength">The new length of current heap memory block to grow of shrink to.</param></params>
''<remarks>
''  The functionality of this function is heavy and if
''</remarks>
Public Sub ReAlloc(ByVal NewLength As Long, Optional ByVal Flags As HeapFlags = hfNoFlags)
THROW_IfNull
    Dim Result As Long
    If Flags = NO_VALUE Then
        If Not p_ForceSerialize Then Flags = hfNoSerialize
    End If
    Result = API_HeapReAlloc(p_Heap.Handle, Flags, p_Ref, NewLength)
    If Result = vbNullPtr Then
        Dim HRESULT As Long
        HRESULT = API_GetLastError
        If HRESULT = STATUS_NO_MEMORY Then
            throw Exps.OutOfMemoryException
        ElseIf HRESULT = STATUS_ACCESS_VIOLATION Then
            throw Exps.AccessDeniedException
        Else
            throw Exps.InvalidOperationException
        End If
    Else
        Dim BufRef As Long
        BufRef = p_Ref
        p_Ref = Result
        If Result <> BufRef Then RaiseEvent BlockMoved(EventArgs(Me))
    End If
End Sub

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
    ToString = CLASSALIAS
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

Public Function Clone() As HeapBlock
    Set Clone = New HeapBlock
    Call Clone.ConstructorCopy(Me)
End Function

Public Sub Dispose()
    If p_Ref = vbNullPtr Then Exit Sub
    Dim Flags As HeapFlags
    If Not p_ForceSerialize Then Flags = hfNoSerialize
    If API_HeapFree(p_Heap.Handle, Flags, p_Ref) = NO_VALUE Then _
        throw Exps.IfError
    'Call p_Heap.BlockDestroyed(Me)
    p_Ref = vbNullPtr
    Set p_Heap = Nothing
    p_Length = NO_VALUE
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
Private Sub Class_Initialize()
    p_ForceSerialize = True
End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    Call Dispose
End Sub

Private Sub THROW_IfNull()
    If p_Ref = vbNullPtr Then _
        throw Exps.ObjectDisposedException
End Sub
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
'<section IMemoryBlock Implementations>
'
Private Property Get IMemoryBlock_Access() As MemoryAccessPermissions: IMemoryBlock_Access = Memory.GetAccessPermission(Me): End Property
Private Property Get IMemoryBlock_IsLimited() As Boolean: IMemoryBlock_IsLimited = False: End Property
Private Property Get IMemoryBlock_Limit() As Long: IMemoryBlock_Limit = Size: End Property
Private Property Get IMemoryBlock_Reference() As Long: IMemoryBlock_Reference = Reference: End Property
Private Property Get IMemoryBlock_IsResizable() As Boolean: IMemoryBlock_IsResizable = IsResizable: End Property
Private Sub IMemoryBlock_Resize(ByVal NewSize As Long): ReAlloc (NewSize): End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section HeapDisposed Callback>
'
Private Sub p_Heap_HeapDisposed(ByVal E As EventArgs)
    p_Ref = vbNullPtr
    p_Heap = vbNullPtr
    p_Length = NO_VALUE
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

