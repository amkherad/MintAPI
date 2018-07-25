VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Heap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="Heap" tags="">
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
'Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Heap" '#DO NOT REMOVE IT
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

Public Enum HeapFlags
    hfNoFlags = 0
    hfNoSerialize = &H1
    hfGrowable = &H2
    hfGenerateExceptions = &H4
    hfZeroMemory = &H8
    hfReallocInPlaceOnly = &H10
    hfTailCheckingEnabled = &H20
    hfFreeCheckingEnabled = &H40
    hfDisableCoalesceOnFree = &H80
    hfCreateAlign16 = &H10000
    hfCreateEnableTracing = &H20000
    hfMaximumTag = &HFFF
    hfPseudoTagFlag = &H8000
    hfTagShift = 18
End Enum
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables/Events>
'
Public Event HeapDisposed(ByVal E As EventArgs)

Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Private p_Heap As Long
Private p_IsGrowable As Boolean
Private p_AllocatedBlocks As New List
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
Friend Sub Constructor0(ByVal Heap As Long, ByVal IsGrowable As Boolean)
    If Heap = vbNullPtr Then throw Exps.ArgumentNullException
    
    p_Heap = Heap
    p_IsGrowable = IsGrowable
    
    This_Constructed = True
End Sub
'Friend Sub ConstructorCopy(ByVal That As Heap)
'    p_Heap = That.Handle
'
'    This_Constructed = True
'End Sub
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

Public Property Get Handle() As Long
THROW_IfNull
    Handle = p_Heap
End Property

Public Property Get IsGrowable() As Boolean
THROW_IfNull
    IsGrowable = p_IsGrowable
End Property

Public Function Allocate(ByVal Size As Long, Optional ByVal Flags As HeapFlags = hfNoFlags) As HeapBlock
THROW_IfNull
    Dim Ptr As Long
    Ptr = API_HeapAlloc(p_Heap, Flags, Size)
    If Ptr = vbNullPtr Then throw Exps.OutOfMemoryException
    Set Allocate = New HeapBlock
    Call Allocate.Constructor0(Me, Ptr, Size)
'
'    Dim WR As New WeakReference
'    Set WR.Object = Allocate
'    Call p_AllocatedBlocks.Append(WR)
End Function

'Friend Sub Release()
'THROW_IfNull
'    p_Heap = vbNullPtr
'End Sub

Public Sub Dispose()
THROW_IfNull
    If p_Heap <> API_GetProcessHeap Then
        If API_HeapDestroy(p_Heap) = NO_VALUE Then _
            throw Exps.IfError
        RaiseEvent HeapDisposed(EventArgs(Me))
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

'Public Function Clone() As Heap
'    Set Clone = New Memory
'    Call Clone.ConstructorCopy(Me)
'End Function
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
    If p_Heap = vbNullPtr Then throw Exps.ObjectDisposedException
End Sub

'Friend Sub BlockDestroyed(ByVal Block As HeapBlock)
'    Dim WR As WeakReference, RemoveNodes As New List
'
'    For Each WR In p_AllocatedBlocks
'        If WR.IsAlive Then
'            If WR.Object Is Block Then _
'                Call RemoveNodes.Append(WR)
'        Else
'            Call RemoveNodes.Append(WR)
'        End If
'    Next
'
'    Call p_AllocatedBlocks.RemoveAll(RemoveNodes)
'End Sub
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
'Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
