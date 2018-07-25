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
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="List" tags="">
''  <summary>
''    Represents a list of objects that can be accessed by index.
''    Provides methods to search, sort, and manipulate lists.
''    This list is optimized for both For and For Each and it doesn't matter
''    to which to use.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>LinkedList Vector ReadOnlyVector</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

''#######  NODE_PTR() IS NOT THREAD-SAFE , TODO: make it thread safe!.

Option Base 0
Option Explicit

Implements IObject
Implements ICloneable
Implements ICollection
Implements IEnumerable
Implements IList
'Implements IProvideEnumerableVersion

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "List" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"

Const QUERYINFORMATION_ALL              As String = "all"
Const QUERYINFORMATION_DEBUGINFO        As String = "dbginfo"
Const QUERYINFORMATION_ENUMVERSION      As String = "enumversion"

Const NOTIFICATION_CLEAR                As String = "clear"
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

'NODE STRUCTURE:
'   NODE*   Next
'   VARIANT Value
'   NODE*   Previous

Const STRUCTURE_LENGTH As Long = 24
Const STRUCTURE_NEXT As Long = 0
Const STRUCTURE_DATA As Long = 4
Const STRUCTURE_PREV As Long = 20
Const STRUCTURE_DATALENGTH As Long = VLEN_VARIANT

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

'=====================================
Dim pFirstTokenPtr As Long
Dim pMidTokenPtr As Long 'Not used yet ! I need to sleep dudes. :D
Dim pLastTokenPtr As Long
Dim pCounts As Long

Dim pCurrentIndex As Long
Dim pCurrentTokenM5Ptr As Long
Dim pCurrentTokenPtr As Long
'=====================================

Dim ChkTokenTypes As Boolean
Dim TokTypes As VbVarType
Dim EnumVersion As Long
'=====================================

Dim LQ As ListQueryable
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
Friend Sub Constructor0(Tokens As Variant)
    Call AppendRange(Tokens)

    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal List As List)
    Call NODE_REMOVEALL
    Call InsertRange(0, List)

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

Public Function GetEnumerator() As IEnumerator
    Dim LEnumerator As New ListEnumerator
    Call LEnumerator.Constructor0(Me)
    Set GetEnumerator = LEnumerator
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

Public Property Get IsEmpty() As Boolean
    IsEmpty = ((pFirstTokenPtr = vbNullPtr) Or (pCounts = 0))
End Property

Public Function Prepend(ByRef Value As Variant) As Long
    Call EnsureDataType(VarType(Value))
    EnumVersion = EnumVersion + 1
    Dim rPtr As Long
    rPtr = NODE_RESERVE
    'Prepend = 0
    If pCounts > 0 Then
        Call NODE_INSERT(vbNullPtr, rPtr, 0)
        pCounts = pCounts + 1
    Else
        pLastTokenPtr = rPtr
        pCurrentIndex = 0
        pCurrentTokenPtr = rPtr
        pCurrentTokenM5Ptr = rPtr
        pCounts = 1
    End If
    pFirstTokenPtr = rPtr
    Call NODE_SETVALUE(rPtr, Value)
End Function
Public Function Append(ByRef Value As Variant) As Long
Attribute Append.VB_UserMemId = -6
    Call EnsureDataType(VarType(Value))
    EnumVersion = EnumVersion + 1
    Dim rPtr As Long
    rPtr = NODE_RESERVE
    Append = pCounts
    If pCounts > 0 Then
        Call NODE_INSERT(pLastTokenPtr, rPtr, pCounts)
        pCounts = pCounts + 1
    Else
        pFirstTokenPtr = rPtr
        pCurrentIndex = 0
        pCurrentTokenPtr = rPtr
        pCurrentTokenM5Ptr = rPtr
        pCounts = 1
    End If
    pLastTokenPtr = rPtr
    Call NODE_SETVALUE(rPtr, Value)
End Function
Public Function AppendRange(ByRef Value As Variant) As Long
    Dim V As Variant, RetVal As Long
    If IsObject(Value) Then
        Dim Obj As Object
        Set Obj = Value
        If TypeOf Obj Is IEnumerable Then
            For Each V In Obj: Call Append(V): RetVal = RetVal + 1: Next
        Else
            throw Exps.InvalidArgumentException("Value")
        End If
    ElseIf IsArray(Value) Then
        For Each V In Obj: Call Append(V): RetVal = RetVal + 1: Next
    Else
        Call Append(Value)
        RetVal = 1
    End If
    AppendRange = RetVal
End Function

''<summary>Removes all elements from the List.</summary>
Public Sub Clear()
    EnumVersion = EnumVersion + 1
    Call NODE_REMOVEALL
End Sub


Public Function IndexOf(ByRef Value As Variant, Optional ByVal Index As Long = 0, Optional ByVal Count As Long = -1, Optional ByVal Comparer As IComparer) As Long
    If pCounts = 0 Then
        IndexOf = -1
        Exit Function
    End If
    If Comparer Is Nothing Then Set Comparer = MintAPI.Comparer.Default
    Dim cRef As Long, cIndex As Long
    cRef = pFirstTokenPtr
    Do While (cRef <> vbNullPtr) And (cIndex < pCounts)
        If Comparer.Compare(Value, NODE_VALUE(cRef)) = crEqual Then
            IndexOf = cIndex
            Exit Function
        End If
        cIndex = cIndex + 1
        Call memcpy(cRef, ByVal (cRef + STRUCTURE_NEXT), VLEN_PTR)
    Loop
    IndexOf = -1
End Function
Public Function LastIndexOf(ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    If pCounts = 0 Then
        LastIndexOf = -1
        Exit Function
    End If
    If Comparer Is Nothing Then Set Comparer = MintAPI.Comparer.Default
    Dim cRef As Long, cIndex As Long
    cRef = pLastTokenPtr
    cIndex = pCounts - 1
    Do While (cRef <> vbNullPtr) And (cIndex >= 0)
        If Comparer.Compare(Value, NODE_VALUE(cRef)) = crEqual Then
            LastIndexOf = cIndex
            Exit Function
        End If
        cIndex = cIndex - 1
        Call memcpy(cRef, ByVal (cRef + STRUCTURE_PREV), VLEN_PTR)
    Loop
    LastIndexOf = -1
End Function

Public Function Reverse(Optional ByVal StartAt As Long = 0, Optional ByVal Count As Long = -1) As List
    Dim mEndAt As Long
    mEndAt = (StartAt + Count)
    If (StartAt < 0) Or (mEndAt > pCounts) Then throw Exps.IndexOutOfRangeException
    Dim RetVal As New List
    Dim i As Long
    mEndAt = mEndAt - 1
    For i = mEndAt To StartAt Step -1
        Call RetVal.Append(Item(i))
    Next
    Set Reverse = RetVal
End Function
Public Sub ForEach(ByVal Predicate As Method, Optional ByVal StartAt As Long, Optional ByVal Count As Long)
    If Predicate Is Nothing Then throw Exps.ArgumentNullException("Predicate")
    
End Sub
Public Function Sort(ByVal Predicate As Method, Optional ByVal StartAt As Long, Optional ByVal Count As Long) As List
    If Predicate Is Nothing Then throw Exps.ArgumentNullException("Predicate")
    
End Function
Public Function Find(ByVal Predicate As Method) As Variant
    If Predicate Is Nothing Then throw Exps.ArgumentNullException("Predicate")
    
End Function
Public Function FindIndex(ByVal Predicate As Method) As Long
    If Predicate Is Nothing Then throw Exps.ArgumentNullException("Predicate")
    
End Function
Public Function Exists(ByVal Predicate As Method) As Boolean
    If Predicate Is Nothing Then throw Exps.ArgumentNullException("Predicate")
    
End Function
Public Function FindAll(ByVal Predicate As Method) As List
    If Predicate Is Nothing Then throw Exps.ArgumentNullException("Predicate")
    
End Function
Public Function TrueForAll(ByVal Predicate As Method) As Boolean
    If Predicate Is Nothing Then throw Exps.ArgumentNullException("Predicate")
    
End Function
Public Function GetRange(ByVal Index As Long, Optional ByVal Count As Long = -1) As List
    Dim mEndAt As Long
    If Count = -1 Then Count = pCounts - Index
    mEndAt = (Index + Count)
    If (Index < 0) Or (mEndAt > pCounts) Then throw Exps.IndexOutOfRangeException
    Dim RetVal As New List
    Dim i As Long
    mEndAt = mEndAt - 1
    
    For i = Index To mEndAt
        Call RetVal.Append(Item(i))
    Next
    
    Set GetRange = RetVal
End Function
Public Sub Insert(ByVal Index As Long, ByRef Value As Variant)
    If (Index < 0) Or (Index > pCounts) Then throw Exps.IndexOutOfRangeException
    Call EnsureDataType(VarType(Value))
    
    Dim NODEPTR As Long, PrevPtr As Long
    If Index > 0 Then PrevPtr = NODE_PTR(Index - 1, False)
    
    EnumVersion = EnumVersion + 1
    
    NODEPTR = NODE_RESERVE
    
    Call NODE_SETVALUE(NODEPTR, Value)
    
    Call NODE_INSERT(PrevPtr, NODEPTR, Index)
    pCounts = pCounts + 1
End Sub
Public Sub InsertRange(ByVal Index As Long, ByVal Values As IEnumerable)
    If (Index < 0) Or (Index > pCounts) Then throw Exps.IndexOutOfRangeException

    Dim NODEPTR As Long, PrevPtr As Long
    If Index > 0 Then PrevPtr = NODE_PTR(Index - 1, False)

    EnumVersion = EnumVersion + 1

    Dim V As Variant, i As Long
    i = 0
    For Each V In Values
        Call EnsureDataType(VarType(V))

        NODEPTR = NODE_RESERVE

        Call NODE_SETVALUE(NODEPTR, V)

        Call NODE_INSERT(PrevPtr, NODEPTR, Index + i)
        i = i + 1

        PrevPtr = NODEPTR
        pCounts = pCounts + 1
    Next
End Sub

''<summary>Swaps two specified elements of the List.</summary>
Public Sub Swap(ByVal Index1 As Long, ByVal Index2 As Long)
    EnumVersion = EnumVersion + 1
    Dim N1Ptr As Long, N2Ptr As Long
    Dim Buffer As Variant
    N1Ptr = NODE_PTR(Index1, True)
    N2Ptr = NODE_PTR(Index2, True)
    
    Call memcpy(Buffer, ByVal (N1Ptr + STRUCTURE_DATA), STRUCTURE_DATALENGTH)
    Call memcpy(ByVal (N1Ptr + STRUCTURE_DATA), ByVal (N2Ptr + STRUCTURE_DATA), STRUCTURE_DATALENGTH)
    Call memcpy(ByVal (N2Ptr + STRUCTURE_DATA), Buffer, STRUCTURE_DATALENGTH)
    
    Call memzero(Buffer, STRUCTURE_DATALENGTH)
End Sub
'***************

Public Sub RemoveAt(ByVal Index As Long)
    EnumVersion = EnumVersion + 1
    Call NODE_DELVALUE(NODE_PTR(Index, True), Index)
    If pCounts > 0 Then pCounts = pCounts - 1
End Sub
Public Sub Remove(ByRef Value As Variant, Optional ByVal Comparer As IComparer)
    Dim Index As Long
    Index = IndexOf(Value, Comparer)
    If Index = -1 Then throw Exps.ItemNotExistsException
    Call RemoveAt(Index)
End Sub
Public Sub RemoveOfType(ByVal vType As VariantTypes)
    EnumVersion = EnumVersion + 1
    
End Sub
Public Sub RemoveRange(ByVal Index As Long, ByVal Count As Long)
    If Index < 0 Then throw Exps.IndexOutOfRangeException
    If Count < 0 Then throw Exps.IndexOutOfRangeException
    If Count > pCounts Then throw Exps.IndexOutOfRangeException
    If (Index + Count) > pCounts Then throw Exps.IndexOutOfRangeException
    
    If (Index + Count) = pCounts Then
        If Index = 0 Then
            Call NODE_REMOVEALL
        Else
            
        End If
    Else
        
    End If
End Sub
Public Function RemoveAll(ByVal RemovingItems As IEnumerable, Optional ByVal Comparer As IComparer) As Long
    Dim Index As Long, RemovedItems As Long
    Dim V As Variant
    
    For Each V In RemovingItems
        Do
            Index = IndexOf(V, Comparer)
            If Index <> -1 Then
                Call NODE_DELVALUE(NODE_PTR(Index, True), Index)
                
                RemovedItems = RemovedItems + 1
            End If
        Loop While (Index <> -1)
    Next
    
    If RemovedItems > 0 Then EnumVersion = EnumVersion + 1
    RemoveAll = RemovedItems
End Function


Public Property Get First() As Variant: Call MoveVariant(First, Item(0)): End Property
Public Property Let First(ByRef Value As Variant): Call SetValue(0, Value): End Property
Public Property Set First(ByVal Value As Object): Call SetValue(0, Value): End Property

Public Property Get Last() As Variant: Call MoveVariant(Last, Item(pCounts - 1)): End Property
Public Property Let Last(ByRef Value As Variant): Call SetValue(pCounts - 1, Value): End Property
Public Property Set Last(ByVal Value As Object): Call SetValue(pCounts - 1, Value): End Property

Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Call MoveVariant(Item, NODE_VALUE(NODE_PTR(Index, True)))
End Property
Public Property Let Item(ByVal Index As Long, ByRef Value As Variant)
    Call EnsureDataType(VarType(Value))
    EnumVersion = EnumVersion + 1
    Call NODE_SETVALUE(NODE_PTR(Index, True), Value)
End Property
Public Property Set Item(ByVal Index As Long, ByVal Value As Object)
    Call EnsureDataType(VarType(Value))
    EnumVersion = EnumVersion + 1
    Call NODE_SETVALUE(NODE_PTR(Index, True), Value)
End Property

Public Function GetTypeOf(ByVal Index As Long) As VariantTypes
    GetTypeOf = NODE_VALUETYPE(NODE_PTR(Index, True))
End Function

Public Sub SetValue(ByVal Index As Long, ByRef Value As Variant)
    EnumVersion = EnumVersion + 1
    Call NODE_SETVALUE(NODE_PTR(Index, True), Value)
End Sub

Public Function GetValue(ByVal Index As Long) As Variant
    Call MoveVariantX(GetValue, NODE_VALUE(NODE_PTR(Index, True)))
End Function

Public Property Get Count() As Long
    Count = pCounts
End Property

Public Sub CopyTo(ByRef Arr As Variant, Optional ByVal Index As Long = 0, Optional ByVal Count As Long = -1)
    Dim Vector As New Vector
    Call Vector.Constructor1(Me, Index, Count)

    Call Vector.SwapBuffer(Arr)
End Sub
Public Function Contains(Value As Variant, Optional ByVal Comparer As IComparer = Nothing) As Boolean
    
End Function

Public Function AsQueryable() As IQueryable
    If LQ Is Nothing Then
        Set LQ = New ListQueryable
        Call LQ.Constructor0(Me)
    End If
    Set AsQueryable = LQ
End Function
'
Public Function ToVector() As Vector
    Set ToVector = New Vector
    Call ToVector.Constructor1(Me)
End Function
Public Function AsReadOnly() As ReadOnlyVector
    Set AsReadOnly = New ReadOnlyVector
    Call AsReadOnly.Constructor1(Me)
End Function
Public Function ToQueue() As Queue
    Set ToQueue = New Queue
    Call ToQueue.Constructor1(Me)
End Function
Public Function ToStack() As Stack
    Set ToStack = New Stack
    Call ToStack.Constructor1(Me)
End Function
Public Function ToArray() As Variant
    Call ToVector.SwapBuffer(ToArray)
End Function

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

Public Function Clone() As List
    Set Clone = New List
    Call Clone.ConstructorCopy(Me)
End Function
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
    pCurrentIndex = -1
End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    Call NODE_REMOVEALL
End Sub

Friend Function GetEnumerableVersion() As Long
    GetEnumerableVersion = EnumVersion
End Function

Friend Sub CopyItemToVariant(ByVal Index As Long, ByRef Var As Variant)
    Dim nPtr As Long
    nPtr = NODE_PTR(Index, True)
    Call memcpy(Var, nPtr + STRUCTURE_DATA, VLEN_VARIANT)
End Sub

Private Sub EnsureDataType(ByVal DataType As VbVarType)
    If ChkTokenTypes Then _
        If DataType <> TokTypes Then _
            throw Exps.InvalidOperationException
End Sub
Private Sub Throw_Empty()
    throw Exps.InvalidOperationException
End Sub

'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
''<summary>
''  Retrives the pointer to the list structure of the given index.
''</summary>
''<retval>Returns the list item determined by the index.</retval>
''<remarks>This function is highly optimized.</remarks>
Private Function NODE_PTR(ByVal pNODEINDEX As Long, ByVal ThrowIfNull As Boolean) As Long
    If pNODEINDEX < 0 Or pNODEINDEX >= pCounts Then throw Exps.IndexOutOfRangeException
    Dim i As Long, Count As Long
    Count = pCounts - 1
    
    If (pCurrentIndex <> -1) And (pNODEINDEX > pCurrentIndex) Then
        If pNODEINDEX = Count Then
            NODE_PTR = pLastTokenPtr
            pCurrentTokenM5Ptr = 0
        Else
            NODE_PTR = NODE_NEXT(pCurrentTokenPtr, True)
            For i = pCurrentIndex + 1 To Count
                If NODE_PTR = vbNullPtr Then throw Exps.InvalidOperationException
                If pNODEINDEX = i Then Exit For
                NODE_PTR = NODE_NEXT(NODE_PTR, True)
            Next
        End If
    ElseIf pNODEINDEX = pCurrentIndex Then
        NODE_PTR = pCurrentTokenPtr
'    ElseIf pNODEINDEX = pCurrentIndex - 1 Then
'        NODE_PTR = NODE_PREVIOUS(pCurrentTokenPtr, ThrowIfNull)
    ElseIf pNODEINDEX = 0 Then
        NODE_PTR = pFirstTokenPtr
    Else 'If NODEINDEX < pCurrentIndex Then
        If pCurrentTokenM5Ptr > 0 Then
            Dim cVRange As Long
            cVRange = pCurrentIndex - 5 '6
            If cVRange > 0 Then
                If pNODEINDEX > cVRange Then
                    NODE_PTR = pCurrentTokenM5Ptr
                    For i = cVRange To pCurrentIndex
                        If NODE_PTR = vbNullPtr Then throw Exps.InvalidOperationException
                        If pNODEINDEX = i Then Exit For
                        NODE_PTR = NODE_NEXT(NODE_PTR, True)
                    Next
                    GoTo JumpFirstToCurrent
                End If
            End If
        End If
        NODE_PTR = NODE_NEXT(pFirstTokenPtr, True)
        For i = 1 To pCurrentIndex - 1
            If NODE_PTR = vbNullPtr Then throw Exps.InvalidOperationException
            If pNODEINDEX = i Then Exit For
            NODE_PTR = NODE_NEXT(NODE_PTR, True)
        Next
JumpFirstToCurrent:
    End If
    
    pCurrentIndex = pNODEINDEX
    pCurrentTokenPtr = NODE_PTR
    If pCurrentIndex <= 5 Then
        pCurrentTokenM5Ptr = pFirstTokenPtr
    Else
        Dim TM5PtrBuff As Long
        TM5PtrBuff = pCurrentTokenPtr
        For i = 1 To 5
            TM5PtrBuff = NODE_PREVIOUS(TM5PtrBuff, False)
            If TM5PtrBuff = vbNullPtr Then
                Exit For
            Else
                pCurrentTokenM5Ptr = TM5PtrBuff
            End If
        Next
        'to be added: 6>------------|------------<Index
    End If

    If ThrowIfNull Then _
        If NODE_PTR = vbNullPtr Then throw Exps.InvalidOperationException
End Function
''<summary>Retrives the pointer to the next list structure of the given structure pointer.</summary>
''<retval>Returns the next structure pointer.</retval>
Private Function NODE_NEXT(ByVal NODEPTR As Long, ByVal ThrowIfNull As Boolean) As Long
    If NODEPTR <> vbNullPtr Then _
        NODE_NEXT = MemPtr(NODEPTR)
    
    If NODE_NEXT = vbNullPtr Then _
        If ThrowIfNull Then throw Exps.InvalidOperationException
End Function
''<summary>Retrives the pointer to the previous list structure of the given structure pointer.</summary>
''<retval>Returns the previous structure pointer.</retval>
Private Function NODE_PREVIOUS(ByVal NODEPTR As Long, ByVal ThrowIfNull As Boolean) As Long
    If NODEPTR <> vbNullPtr Then _
        NODE_PREVIOUS = MemPtr(NODEPTR + VLEN_PTR + STRUCTURE_DATALENGTH)
    
    If NODE_PREVIOUS = vbNullPtr Then _
        If ThrowIfNull Then throw Exps.InvalidOperationException
End Function
''<summary>Retrives the value of the given structure pointer.</summary>
''<retval>Returns the value of the item.</retval>
Private Function NODE_VALUE(ByVal NODEPTR As Long) As Variant
    If NODEPTR <> vbNullPtr Then
        Dim lngType As Long
        lngType = MemLong(NODEPTR + STRUCTURE_DATA + VARIANT_TYPE_OFFSET)
        If lngType = VT_DISPATCH Then
            Dim objRetVal As IUnknown
            ObjectPtr(objRetVal) = MemLong(NODEPTR + STRUCTURE_DATA + VARIANT_DATA_OFFSET)
            Set NODE_VALUE = objRetVal
            ObjectPtr(objRetVal) = vbNullPtr
        Else
            Dim vVal As Variant
            Call memcpy(vVal, ByVal (NODEPTR + STRUCTURE_DATA), STRUCTURE_DATALENGTH)
            Call Evaluate(NODE_VALUE, vVal)
            Call memzero(vVal, STRUCTURE_DATALENGTH)
        End If
    Else
        throw Exps.InvalidOperationException
    End If
End Function
''<summary>Retrives the value of the given structure pointer.</summary>
''<retval>Returns the value of the item.</retval>
Private Function NODE_VALUETYPE(ByVal NODEPTR As Long) As Variant
    If NODEPTR <> vbNullPtr Then
        Call memcpy(NODE_VALUETYPE, ByVal (NODEPTR + STRUCTURE_DATA + VARIANT_TYPE_OFFSET), VLEN_VARTYPE)
    Else
        throw Exps.InvalidOperationException
    End If
End Function
''<summary>Sets the value of the given structure pointer.</summary>
Private Sub NODE_SETVALUE(ByVal NODEPTR As Long, ByRef RIVALUE As Variant)
    If NODEPTR <> vbNullPtr Then
        Call NODE_RELEASEVALUE(NODEPTR, False)
        If IsObject(RIVALUE) Then
            Call IUnknown.AddRef(RIVALUE)
            Call memcpy(ByVal (NODEPTR + STRUCTURE_DATA + VARIANT_TYPE_OFFSET), VT_DISPATCH&, VLEN_LONG)
            Call memcpy(ByVal (NODEPTR + STRUCTURE_DATA + VARIANT_DATA_OFFSET), ObjPtr(RIVALUE), VLEN_PTR)
        Else
            Dim vVal As Variant
            Call Evaluate(vVal, RIVALUE)
            Call memcpy(ByVal (NODEPTR + STRUCTURE_DATA), ByVal VarPtr(vVal), STRUCTURE_DATALENGTH)
            Call memzero(ByVal VarPtr(vVal), VLEN_VARIANT)
        End If
    Else
        throw Exps.InvalidOperationException
    End If
End Sub
Private Sub NODE_RELEASEVALUE(ByVal NODEPTR As Long, ByVal ReleaseStruct As Boolean)
    If NODEPTR = vbNullPtr Then Exit Sub
    Dim lngType As Long
    Call memcpy(lngType, ByVal (NODEPTR + STRUCTURE_DATA + VARIANT_TYPE_OFFSET), VLEN_INTEGER)
    If lngType = VT_DISPATCH Then
        Dim objRetVal As IUnknown
        Call memcpy(objRetVal, MemLong(NODEPTR + STRUCTURE_DATA + VARIANT_DATA_OFFSET), VLEN_OBJECT)
        Call IUnknown.Release(objRetVal)
        Call memzero(objRetVal, VLEN_OBJECT)
    Else
        Dim vVal As Variant
        Call memcpy(vVal, ByVal (NODEPTR + STRUCTURE_DATA), STRUCTURE_DATALENGTH)
        vVal = vbEmpty
        Call API_VariantClear(vVal)
    End If
    
    If ReleaseStruct Then
        'Call memzero(ByVal MemPtrV, STRUCTURE_DATALENGTH)
        Call memzero(ByVal NODEPTR, STRUCTURE_LENGTH)
        Call API_CoTaskMemFree(NODEPTR)
    End If
End Sub
''<summary>Releases the value of the given structure pointer.</summary>
Private Sub NODE_DELVALUE(ByVal NODEPTR As Long, ByVal Index As Long)
    If NODEPTR <> vbNullPtr Then
        If (pCounts = 1) And (NODEPTR = pFirstTokenPtr) Then
            Call NODE_REMOVEALL
        Else
            Dim nPtr As Long, pPtr As Long
            nPtr = MemPtr(NODEPTR + STRUCTURE_NEXT)
            pPtr = MemPtr(NODEPTR + STRUCTURE_PREV)
            If nPtr <> vbNullPtr Then _
                MemPtr(nPtr + STRUCTURE_PREV) = pPtr
            If pPtr <> vbNullPtr Then _
                MemPtr(pPtr + STRUCTURE_NEXT) = nPtr
            
            If pCurrentIndex >= pCounts - 1 Then pCurrentIndex = pCurrentIndex - 1
            
            If NODEPTR = pFirstTokenPtr Then
                If pPtr <> 0 Then
                    pFirstTokenPtr = pPtr
                Else
                    pFirstTokenPtr = nPtr
                End If
            End If
            If NODEPTR = pLastTokenPtr Then
                If nPtr <> 0 Then
                    pLastTokenPtr = nPtr
                Else
                    pLastTokenPtr = pPtr
                End If
            End If
            
            If NODEPTR = pCurrentTokenPtr Then
                If nPtr <> 0 Then
                    pCurrentTokenPtr = nPtr
                Else
                    pCurrentTokenPtr = pPtr
                End If
            End If
            If NODEPTR = pCurrentTokenM5Ptr Then
                If pPtr <> 0 Then
                    pCurrentTokenM5Ptr = pPtr
                Else
                    pCurrentTokenM5Ptr = nPtr
                End If
            End If
        
            Call NODE_RELEASEVALUE(NODEPTR, True)
        End If
    Else
        throw Exps.InvalidOperationException
    End If
End Sub
''<summary>Insert a structure into list.</summary>
Private Sub NODE_INSERT(ByVal INSERTAFTER As Long, ByVal NODEPTR As Long, ByVal Index As Long)
    If NODEPTR <> vbNullPtr Then
        If INSERTAFTER = vbNullPtr Then
            MemPtr(NODEPTR + STRUCTURE_PREV) = vbNullPtr
            If pFirstTokenPtr <> 0 Then
                MemPtr(NODEPTR + STRUCTURE_NEXT) = pFirstTokenPtr
                MemPtr(pFirstTokenPtr + STRUCTURE_PREV) = NODEPTR
            End If
            If pCurrentTokenM5Ptr = pFirstTokenPtr Then pCurrentTokenM5Ptr = NODEPTR
            pFirstTokenPtr = NODEPTR
            If pLastTokenPtr = 0 Then pLastTokenPtr = NODEPTR
        Else
            Dim nPtr As Long
            nPtr = MemPtr(INSERTAFTER + STRUCTURE_NEXT)
            
            MemPtr(NODEPTR + STRUCTURE_PREV) = INSERTAFTER
            If nPtr <> vbNullPtr Then
                MemPtr(NODEPTR + STRUCTURE_NEXT) = nPtr
                MemPtr(nPtr + STRUCTURE_PREV) = NODEPTR
            End If
            MemPtr(INSERTAFTER + STRUCTURE_NEXT) = NODEPTR
            
            If (pLastTokenPtr = INSERTAFTER) Then pLastTokenPtr = NODEPTR
        End If
    Else
        throw Exps.InvalidOperationException
    End If
    
    Dim i As Long
    pCurrentIndex = Index
    pCurrentTokenPtr = NODEPTR
    If pCurrentIndex <= 5 Then
        pCurrentTokenM5Ptr = pFirstTokenPtr
    Else
        Dim TM5PtrBuff As Long
        TM5PtrBuff = pCurrentTokenPtr
        For i = 1 To 5
            TM5PtrBuff = NODE_PREVIOUS(TM5PtrBuff, False)
            If TM5PtrBuff = vbNullPtr Then
                Exit For
            Else
                pCurrentTokenM5Ptr = TM5PtrBuff
            End If
        Next
    End If
End Sub
''<summary>Removes all nodes of the list.</summary>
Private Sub NODE_REMOVEALL()
    If pCounts = 0 Then Exit Sub
    
    Dim cPtr As Long, sNext As Long
    
    cPtr = pFirstTokenPtr
    While cPtr <> vbNullPtr
        sNext = MemPtr(cPtr + STRUCTURE_NEXT)
        Call NODE_RELEASEVALUE(cPtr, True)
        cPtr = sNext
    Wend
    
    pFirstTokenPtr = vbNullPtr
    pLastTokenPtr = vbNullPtr
    pCurrentTokenPtr = vbNullPtr
    pCurrentTokenM5Ptr = vbNullPtr
    pCurrentIndex = -1
    pCounts = 0
End Sub
''<summary>Reserves a new structure and returns it's pointer.</summary>
Private Function NODE_RESERVE() As Long
    NODE_RESERVE = API_CoTaskMemAlloc(STRUCTURE_LENGTH)
    If NODE_RESERVE = vbNullPtr Then throw Exps.OutOfMemoryException
    Call memzero(ByVal NODE_RESERVE, STRUCTURE_LENGTH)
End Function
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
'*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*
'=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=*=
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
'<section IEnumerable Implementations>
'
Private Function IEnumerable_GetEnumerator() As IEnumerator: Set IEnumerable_GetEnumerator = GetEnumerator: End Function
Private Function IEnumerable_NewEnum() As IUnknown: Set IEnumerable_NewEnum = NewEnum: End Function
'Private Function IProvideEnumerableVersion_GetEnumerableVersion() As Long: IProvideEnumerableVersion_GetEnumerableVersion = EnumVersion: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICollection Implementations>
'
Private Sub ICollection_CopyTo(Arr As Variant, ByVal Index As Long): Call CopyTo(Arr, Index): End Sub
Private Property Get ICollection_Count() As Long: ICollection_Count = pCounts: End Property
Private Function ICollection_GetEnumerator() As IEnumerator: Set ICollection_GetEnumerator = GetEnumerator: End Function
Private Function ICollection_NewEnum() As stdole.IUnknown: Set ICollection_NewEnum = NewEnum: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IList Implementations>
'

Private Function IList_Add(Value As Variant) As Long: IList_Add = Append(Value): End Function
Private Sub IList_Clear(): Call Clear: End Sub
Private Function IList_Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean: IList_Contains = Contains(Value, Comparer): End Function
Private Sub IList_CopyTo(Arr As Variant, ByVal Index As Long): Call CopyTo(Arr, Index): End Sub
Private Property Get IList_Count() As Long: IList_Count = pCounts: End Property
Private Function IList_GetEnumerator() As IEnumerator: Set IList_GetEnumerator = GetEnumerator: End Function
Private Function IList_IndexOf(Value As Variant, Optional ByVal Comparer As IComparer) As Long: IList_IndexOf = IndexOf(Value): End Function
Private Sub IList_Insert(ByVal Index As Long, Value As Variant): Call Insert(Index, Value): End Sub
Private Property Get IList_IsFixedSize() As Boolean: IList_IsFixedSize = False: End Property
Private Property Get IList_IsReadOnly() As Boolean: IList_IsReadOnly = False: End Property
Private Property Get IList_Item(ByVal Index As Long) As Variant: Call MoveVariant(IList_Item, Item(Index)): End Property
Private Property Let IList_Item(ByVal Index As Long, RHS As Variant): Call SetValue(Index, RHS): End Property
Private Property Set IList_Item(ByVal Index As Long, ByVal RHS As Object): Set Item(Index) = RHS: End Property
Private Function IList_NewEnum() As IUnknown: Set IList_NewEnum = NewEnum: End Function
Private Sub IList_Remove(Value As Variant, Optional ByVal Comparer As IComparer): Call Remove(Value, Comparer): End Sub
Private Sub IList_RemoveAt(ByVal Index As Long): Call RemoveAt(Index): End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
