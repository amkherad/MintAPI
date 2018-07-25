VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LazyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Member1" ,"Provider"
'@PROJECT_LICENSE

''<class name="LazyList" tags="">
''  <summary>
''    Provides a lazy loaded list that loads data when its needed(accessed).
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
Implements ICollection
Implements IEnumerable
Implements IList
Implements List

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "LazyList" '#DO NOT REMOVE IT
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
Public Type LazyListItem
    IsLoad As Boolean
    Value As Variant
    Info As Variant
End Type
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

Dim LQ As ListQueryable

Dim InnerList As New List
Dim pProvider As IProvider, pEngine As String, pAlias As String
Dim pCallback As ICallBack
Dim pSuspendCallback As Boolean

Dim EnumVersion As Long
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
Friend Sub Constructor0(ByVal List As List)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As LazyList)
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
    Dim LEnumerator As New LazyListEnumerator
    Call LEnumerator.Constructor0(Me)
    Set GetEnumerator = LEnumerator
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

Public Property Get Provider() As IProvider
    Set Provider = pProvider
End Property
Public Property Let Provider(ByVal Value As IProvider)
    Set pProvider = Value
    
    EnumVersion = EnumVersion + 1
End Property
Public Property Set Provider(ByVal Value As IProvider)
    Set pProvider = Value
    
    EnumVersion = EnumVersion + 1
End Property

Public Property Get ProviderEngine() As String
    ProviderEngine = pEngine
End Property
Public Property Let ProviderEngine(ByVal Value As String)
    pEngine = Value
    
    EnumVersion = EnumVersion + 1
End Property

Public Property Get ProviderAlias() As String
    ProviderEngine = pAlias
End Property
Public Property Let ProviderAlias(ByVal Value As String)
    pAlias = Value
    
    EnumVersion = EnumVersion + 1
End Property


Public Property Get IsEmpty() As Boolean
    IsEmpty = InnerList.IsEmpty
End Property

Public Function Prepend(ByRef Value As Variant, Optional ByVal IsLoaded As Boolean = True, Optional ByRef Info As Variant) As Long
    Dim IB As LazyListItem
    Call Evaluate(IB.Value, Value)
    Call Evaluate(IB.Info, Info)
    IB.IsLoad = IsLoaded
    
    Dim AddToList As Boolean
    
    If Not pSuspendCallback Then
        AddToList = CBool(pCallback.Invoke(AL.Create(Value)))
    Else
        AddToList = True
    End If
    
    If Not AddToList Then throw Exps.InvalidOperationException
    
    Prepend = InnerList.Prepend(IB)
    
    EnumVersion = EnumVersion + 1
End Function
Public Function Append(ByRef Value As Variant, Optional ByVal IsLoaded As Boolean = True, Optional ByRef Info As Variant) As Long
    Dim IB As LazyListItem
    Call Evaluate(IB.Value, Value)
    Call Evaluate(IB.Info, Info)
    IB.IsLoad = IsLoaded
    
    Dim AddToList As Boolean
    
    If Not pSuspendCallback Then
        AddToList = CBool(pCallback.Invoke(AL.Create(Value)))
    Else
        AddToList = True
    End If
    
    If Not AddToList Then throw Exps.InvalidOperationException
    
    Append = InnerList.Append(IB)
    
    EnumVersion = EnumVersion + 1
End Function
Public Function AppendRange(Value As Variant) As Long
    throw Exps.NotImplementedException
End Function

Public Sub Clear()
    Call InnerList.Clear
    
    If Not pSuspendCallback Then _
        Call pCallback.Invoke(Nothing)
    
    EnumVersion = EnumVersion + 1
End Sub

Public Function IndexOf(ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim i As Long

    If Comparer Is Nothing Then Set Comparer = MintAPI.Comparer.Default

    For i = 0 To Count - 1
        If Comparer.Equals(Item(i), Value) Then
            IndexOf = i
            Exit Function
        End If
    Next
    IndexOf = -1
End Function
Public Function LastIndexOf(ByRef Value As Variant, Optional ByVal Comparer As IComparer) As Long
    Dim i As Long

    If Comparer Is Nothing Then Set Comparer = MintAPI.Comparer.Default

    For i = Count - 1 To 0 Step -1
        If Comparer.Equals(Item(i), Value) Then
            LastIndexOf = i
            Exit Function
        End If
    Next
    LastIndexOf = -1
End Function

Public Function Reverse(Optional ByVal StartAt As Long = 0, Optional ByVal Count As Long = -1) As LazyList
    Dim lst As New LazyList
    Call lst.SetMetaInfo(pProvider, pCallback, pSuspendCallback)

    Dim i As Long
    For i = 0 To Count - 1
        Call lst.AppendDirectlyToInnerList(InnerList(i))
    Next

    Set Reverse = lst
End Function
Public Sub ForEach(ByVal ForEachRoutine As Method, Optional ByVal StartAt As Long, Optional ByVal Count As Long)
    
End Sub
Public Function Sort(ByVal SortRoutine As Method, Optional ByVal StartAt As Long, Optional ByVal Count As Long) As List
    
End Function
Public Function Find(ByVal Predicate As Method) As Variant
    
End Function
Public Function Exists(ByVal Predicate As Method) As Boolean
    
End Function
Public Function FindAll(ByVal Predicate As Method) As List
    
End Function
Public Function TrueForAll(ByVal Predicate As Method) As Boolean
    
End Function

Public Function GetRange(ByVal Index As Long, Optional ByVal Count As Long = -1) As LazyList
'    Dim mEndAt As Long
'    If Count = -1 Then Count = pCounts - Index
'    mEndAt = (Index + Count)
'    If (Index < 0) Or (mEndAt > pCounts) Then throw Exps.OutOfRangeException
'    Dim RetVal As New List
'    Dim i As Long
'    mEndAt = mEndAt - 1
'
'    For i = Index To mEndAt
'        Call RetVal.Append(Item(i))
'    Next
'
'    Set GetRange = RetVal
End Function
Public Sub Insert(ByVal Index As Long, ByRef Value As Variant)
'    If (Index < 0) Or (Index > pCounts) Then throw Exps.OutOfRangeException
'    Call EnsureDataType(VarType(Value))
'
'    Dim NODEPTR As Long, PrevPtr As Long
'    If Index > 0 Then PrevPtr = NODE_PTR(Index - 1, False)
'
'    EnumVersion = EnumVersion + 1
'
'    NODEPTR = NODE_RESERVE
'
'    Call NODE_SETVALUE(NODEPTR, Value)
'
'    Call NODE_INSERT(PrevPtr, NODEPTR, Index)
'    pCounts = pCounts + 1
End Sub
Public Sub InsertRange(ByVal Index As Long, ByVal Values As IEnumerable)
'    If (Index < 0) Or (Index > pCounts) Then throw Exps.OutOfRangeException
'
'    Dim NODEPTR As Long, PrevPtr As Long
'    If Index > 0 Then PrevPtr = NODE_PTR(Index - 1, False)
'
'    EnumVersion = EnumVersion + 1
'
'    Dim V As Variant, i As Long
'    i = 0
'    For Each V In Values
'        Call EnsureDataType(VarType(V))
'
'        NODEPTR = NODE_RESERVE
'
'        Call NODE_SETVALUE(NODEPTR, V)
'
'        Call NODE_INSERT(PrevPtr, NODEPTR, Index + i)
'        i = i + 1
'
'        PrevPtr = NODEPTR
'        pCounts = pCounts + 1
'    Next
End Sub

''<summary>Swaps two specified elements of the List.</summary>
Public Sub Swap(ByVal Index1 As Long, ByVal Index2 As Long)
    EnumVersion = EnumVersion + 1
    Call InnerList.Swap(Index1, Index2)
End Sub
'***************

Public Sub RemoveAt(ByVal Index As Long)
    EnumVersion = EnumVersion + 1
    
End Sub
Public Sub Remove(ByRef Value As Variant, Optional ByVal Comparer As IComparer)
    EnumVersion = EnumVersion + 1
    
End Sub
Public Function RemoveAll(ByVal RemovingItems As IEnumerable, Optional ByVal Comparer As IComparer) As Long
    
End Function
Public Sub RemoveOfType(ByVal vType As VariantTypes)
    EnumVersion = EnumVersion + 1
    
End Sub
Public Sub RemoveRange(ByVal Index As Long, ByVal Count As Long)
    EnumVersion = EnumVersion + 1
    
End Sub


Public Property Get First() As Variant: Call MoveVariant(First, Item(0)): End Property
Public Property Let First(ByRef Value As Variant): Call SetValue(0, Value): End Property
Public Property Set First(ByVal Value As Object): Call SetValue(0, Value): End Property

Public Property Get Last() As Variant: Call MoveVariant(Last, Item(InnerList.Count - 1)): End Property
Public Property Let Last(ByRef Value As Variant): Call SetValue(InnerList.Count - 1, Value): End Property
Public Property Set Last(ByVal Value As Object): Call SetValue(InnerList.Count - 1, Value): End Property

Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Dim IB As LazyListItem
    IB = InnerList(Index)
    If Not IB.IsLoad Then
        Dim AL As ArgumentList
        If IsObject(IB.Info) Then
            If TypeOf IB.Info Is ArgumentList Then
                Set AL = IB.Info
            Else
                Set AL = MintAPI.AL.Create(IB.Info)
            End If
        Else
            Set AL = MintAPI.AL.Create(IB.Info)
        End If
        If Not pProvider.Provide(pEngine, pAlias, AL, IB.Value) Then _
            throw Exps.InvalidOperationException
        
        IB.IsLoad = True
        InnerList(Index) = IB
    End If
    
    Call EvaluateX(Item, IB.Value)
End Property
Public Property Let Item(ByVal Index As Long, ByRef Value As Variant)
    Dim IB As LazyListItem
    IB = InnerList(Index)
    IB.IsLoad = True
    Call Evaluate(IB.Value, Value)
    InnerList(Index) = IB
    
    EnumVersion = EnumVersion + 1
End Property
Public Property Set Item(ByVal Index As Long, ByVal Value As Object)
    Dim IB As LazyListItem
    IB = InnerList(Index)
    IB.IsLoad = True
    Call Evaluate(IB.Value, Value)
    InnerList(Index) = IB
    
    EnumVersion = EnumVersion + 1
End Property

Public Function GetTypeOf(ByVal Index As Long) As VariantTypes
    GetTypeOf = VariantType(Item(Index))
End Function

Public Sub SetValue(ByVal Index As Long, ByRef Value As Variant, Optional ByVal IsLoaded As Boolean = True)
    Dim IB As LazyListItem
    IB = InnerList(Index)
    IB.IsLoad = IsLoaded
    Call Evaluate(IB.Value, Value)
    InnerList(Index) = IB
    
    EnumVersion = EnumVersion + 1
End Sub

Public Function GetValue(ByVal Index As Long) As Variant
    Dim IB As LazyListItem
    IB = InnerList(Index)
    If Not IB.IsLoad Then
        Dim AL As ArgumentList
        If IsObject(IB.Info) Then
            If TypeOf IB.Info Is ArgumentList Then
                Set AL = IB.Info
            Else
                Set AL = MintAPI.AL.Create(IB.Info)
            End If
        Else
            Set AL = MintAPI.AL.Create(IB.Info)
        End If
        If Not pProvider.Provide(pEngine, pAlias, AL, IB.Value) Then _
            throw Exps.InvalidOperationException
        
        IB.IsLoad = True
        InnerList(Index) = IB
    End If
    
    Call EvaluateX(GetValue, IB.Value)
End Function

Public Property Get ItemInfo(ByVal Index As Long)
    Dim IB As LazyListItem
    IB = InnerList(Index)
    Call API_VariantCopy(ItemInfo, IB.Info)
End Property
Public Property Let ItemInfo(ByVal Index As Long, ByRef Value As Variant)
    Dim IB As LazyListItem
    IB = InnerList(Index)
    IB.IsLoad = True
    Call Evaluate(IB.Info, Value)
    InnerList(Index) = IB
    
    EnumVersion = EnumVersion + 1
End Property
Public Property Set ItemInfo(ByVal Index As Long, ByVal Value As Object)
    Dim IB As LazyListItem
    IB = InnerList(Index)
    IB.IsLoad = True
    Call Evaluate(IB.Info, Value)
    InnerList(Index) = IB
    
    EnumVersion = EnumVersion + 1
End Property

Public Property Get Count() As Long
    Count = InnerList.Count
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

Public Function ToList() As List
    Set ToList = New List
    Call ToList.ConstructorCopy(Me)
End Function
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

Public Function Clone() As LazyList
    Set Clone = New LazyList
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
    pSuspendCallback = True
End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Friend Sub SetMetaInfo(ByVal Provider As IProvider, ByVal Callback As ICallBack, ByVal SuspendCallback As Boolean)
    Set pProvider = Provider
    Set pCallback = Callback
    pSuspendCallback = SuspendCallback
End Sub
Friend Function AppendDirectlyToInnerList(Value) As Long
    AppendDirectlyToInnerList = InnerList.Append(Value)
End Function

Friend Function GetEnumerableVersion() As Long
    GetEnumerableVersion = EnumVersion
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
Private Property Get ICollection_Count() As Long: ICollection_Count = InnerList.Count: End Property
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
Private Property Get IList_Count() As Long: IList_Count = InnerList.Count: End Property
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

'=============================================
'=============================================
'=============================================
'<section List Implementations>
'
Private Function List_GetEnumerator() As IEnumerator: Set List_GetEnumerator = GetEnumerator: End Function
Private Function List_GetHashCode() As Long: List_GetHashCode = GetHashCode: End Function
Private Function List_GetRange(ByVal Index As Long, Optional ByVal Count As Long = -1&) As List: Set List_GetRange = GetRange(Index, Count): End Function
Private Function List_GetTypeOf(ByVal Index As Long) As VariantTypes: List_GetTypeOf = GetTypeOf(Index): End Function
Private Function List_IndexOf(Value As Variant, Optional ByVal Index As Long = 0&, Optional ByVal Count As Long = -1&, Optional ByVal Comparer As IComparer) As Long: List_IndexOf = IndexOf(Value, Comparer): End Function
Private Sub List_Insert(ByVal Index As Long, Value As Variant): Call Insert(Index, Value): End Sub
Private Sub List_InsertRange(ByVal Index As Long, ByVal Values As IEnumerable): Call InsertRange(Index, Values): End Sub
Private Property Get List_IsEmpty() As Boolean: List_IsEmpty = IsEmpty: End Property
Private Property Set List_Item(ByVal Index As Long, ByVal RHS As Object): Call SetValue(Index, RHS): End Property
Private Property Let List_Item(ByVal Index As Long, RHS As Variant): Call SetValue(Index, RHS): End Property
Private Property Get List_Item(ByVal Index As Long) As Variant: Call MoveVariant(List_Item, Item(Index)): End Property
Private Property Set List_First(ByVal RHS As Object): Call SetValue(0, RHS): End Property
Private Property Let List_First(RHS As Variant): Call SetValue(0, RHS): End Property
Private Property Get List_First() As Variant: Call MoveVariant(List_First, Item(0)): End Property
Private Property Set List_Last(ByVal RHS As Object): Call SetValue(Count - 1, RHS): End Property
Private Property Let List_Last(RHS As Variant): Call SetValue(Count - 1, RHS): End Property
Private Property Get List_Last() As Variant: Call MoveVariant(List_Last, Item(Count - 1)): End Property
Private Function List_LastIndexOf(Value As Variant, Optional ByVal Comparer As IComparer) As Long: List_LastIndexOf = LastIndexOf(Value, Comparer): End Function
Private Function List_MetaObject() As MetaObject: Set List_MetaObject = MetaObject: End Function
Private Function List_NewEnum() As stdole.IUnknown: Set List_NewEnum = NewEnum: End Function
Private Function List_Append(Value As Variant) As Long: List_Append = Append(Value): End Function
Private Function List_Prepend(Value As Variant) As Long: List_Prepend = Prepend(Value): End Function
Private Function List_AppendRange(Value As Variant) As Long: List_AppendRange = AppendRange(Value): End Function
Private Sub List_Remove(Value As Variant, Optional ByVal Comparer As IComparer): Call Remove(Value, Comparer): End Sub
Private Function List_RemoveAll(ByVal RemovingItems As IEnumerable, Optional ByVal Comparer As IComparer) As Long: List_RemoveAll = RemoveAll(RemovingItems, Comparer): End Function
Private Sub List_RemoveAt(ByVal Index As Long): Call RemoveAt(Index): End Sub
Private Sub List_RemoveOfType(ByVal vType As VariantTypes): Call RemoveOfType(vType): End Sub
Private Sub List_RemoveRange(ByVal Index As Long, ByVal Count As Long): Call RemoveRange(Index, Count): End Sub
Private Function List_Reverse(Optional ByVal StartAt As Long = 0&, Optional ByVal Count As Long = -1&) As List: Set List_Reverse = Reverse(StartAt, Count): End Function
Private Sub List_SetValue(ByVal Index As Long, Value As Variant): Call SetValue(Index, Value): End Sub
Private Function List_Sort(ByVal SortRoutine As Method, Optional ByVal StartAt As Long, Optional ByVal Count As Long) As List: Set List_Sort = Sort(SortRoutine, StartAt, Count): End Function
Private Sub List_Swap(ByVal Index1 As Long, ByVal Index2 As Long): Call Swap(Index1, Index2): End Sub
Private Function List_ToArray() As Variant: Call MoveVariant(List_ToArray, ToArray): End Function
Private Function List_ToQueue() As Queue: Set List_ToQueue = ToQueue: End Function
Private Function List_ToStack() As Stack: Set List_ToStack = ToStack: End Function
Private Function List_ToString() As String: List_ToString = ToString: End Function
Private Function List_ToVector() As Vector: Set List_ToVector = ToVector: End Function
Private Function List_TrueForAll(ByVal Predicate As Method) As Boolean: List_TrueForAll = TrueForAll(Predicate): End Function
Private Function List_AsQueryable() As IQueryable: Set List_AsQueryable = AsQueryable: End Function
Private Sub List_Clear(): Call Clear: End Sub
Private Function List_Clone() As List: Set List_Clone = Clone: End Function
Private Function List_Contains(Value As Variant, Optional ByVal Comparer As IComparer = Nothing) As Boolean: List_Contains = Contains(Value, Comparer): End Function
Private Sub List_CopyTo(Arr As Variant, Optional ByVal Index As Long = 0&, Optional ByVal Count As Long = -1&): Call CopyTo(Arr, Index, Count): End Sub
Private Property Get List_Count() As Long: List_Count = Count: End Property
Private Function List_Equals(Value As Variant) As Boolean: List_Equals = Equals(Value): End Function
Private Function List_Exists(ByVal Predicate As Method) As Boolean: List_Exists = Exists(Predicate): End Function
Private Function List_Find(ByVal Predicate As Method) As Variant: List_Find = Find(Predicate): End Function
Private Function List_FindAll(ByVal Predicate As Method) As List: Set List_FindAll = FindAll(Predicate): End Function
Private Sub List_ForEach(ByVal ForEachRoutine As Method, Optional ByVal StartAt As Long, Optional ByVal Count As Long): Call ForEach(ForEachRoutine, StartAt, Count): End Sub
Private Function List_FindIndex(ByVal Predicate As Method) As Long:: End Function
Private Function List_AsReadOnly() As ReadOnlyVector: Set List_AsReadOnly = AsReadOnly: End Function
Private Function List_GetValue(ByVal Index As Long) As Variant: Call MoveVariantX(List_GetValue, GetValue(Index)): End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
