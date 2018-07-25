VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgumentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Member1" ,"Argument"
'@PROJECT_LICENSE

''<class name="ArgumentList" tags="">
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
Implements ICollection
'Implements IList
Implements IEnumerable
Implements IDictionary

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ArgumentList" '#DO NOT REMOVE IT
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
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Dim eArgs_P As New List 'Extended Args
Dim pArgs As Variant, pCounts As Long, pRelease As Boolean, pShallowRelease As Boolean
Dim eArgs_N As New List 'Extended Args

Dim pNames As New List

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
Friend Sub Constructor0(ByVal ParamArrayVariantPtr As Long)
    If ParamArrayVariantPtr = vbNullPtr Then throw Exps.InvalidOperationException("Invalid Variant Pointer.")
    Dim SAPtr As Long
    SAPtr = MemPtr(MemPtr(ParamArrayVariantPtr))
    If API_SafeArrayGetDim(SAPtr) <> 1 Then throw Exps.MultiDimentionException
    pArgs = SAPtr
    VariantType(pArgs) = (VT_ARRAY Or VT_VARIANT)
    pRelease = False
    pShallowRelease = False
    pCounts = (API_SafeArrayGetUBound(SAPtr, 1) - API_SafeArrayGetLBound(SAPtr, 1) + 1)
    
    Dim i As Long
    For i = 0 To pCounts - 1
        Call pNames.Append("var" & i) 'Waste of memory. BUT is important
    Next
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal SafeArrayPtr As Long)
    If SafeArrayPtr = vbNullPtr Then throw Exps.InvalidOperationException("Invalid Safe Array Pointer.")
    If API_SafeArrayGetDim(SafeArrayPtr) <> 1 Then throw Exps.MultiDimentionException
    pArgs = SafeArrayPtr
    VariantType(pArgs) = (VT_ARRAY Or VT_VARIANT)
    pRelease = False
    pShallowRelease = False
    pCounts = (API_SafeArrayGetUBound(SafeArrayPtr, 1) - API_SafeArrayGetLBound(SafeArrayPtr, 1) + 1)
    
    Dim i As Long
    For i = 0 To pCounts - 1
        Call pNames.Append("var" & i) 'Waste of memory. BUT is important
    Next
    
    This_Constructed = True
End Sub
Friend Sub Constructor2(ArgsToRefer As Variant)
    Call memcpy(pArgs, ArgsToRefer, VLEN_VARIANT)
    pRelease = False
    pShallowRelease = False
    pCounts = ArraySize(pArgs)
    
    Dim i As Long
    For i = 0 To pCounts - 1
        Call pNames.Append("var" & i) 'Waste of memory. BUT is important
    Next
    
    This_Constructed = True
End Sub
Friend Sub Constructor3(ArgsToClone As Variant)
    Call API_VariantCopy(pArgs, ArgsToClone)
    pRelease = True
    pShallowRelease = False
    pCounts = ArraySize(pArgs)
    
    Dim i As Long
    For i = 0 To pCounts - 1
        Call pNames.Append("var" & i) 'Waste of memory. BUT is important
    Next
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal ArgumentList As ArgumentList)
    
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
    Dim LEnumerator As New ArgumentListEnumerator
    Call LEnumerator.Constructor0(Me)
    Set GetEnumerator = LEnumerator
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

Public Function Prepend(ByRef Arg As Variant, Optional ByRef Key As Variant) As Long
    Prepend = eArgs_P.Prepend(Arg)
    If pNames.IndexOf(Key) >= 0 Then throw Exps.InvalidOperationException("Element Names Exists.")
    Call pNames.Prepend(Key)
End Function
Public Function Append(ByRef Arg As Variant, Optional ByRef Key As Variant) As Long
    Append = pCounts + eArgs_P.Count + eArgs_N.Append(Arg)
    If pNames.IndexOf(Key) >= 0 Then throw Exps.InvalidOperationException("Element Names Exists.")
    Call pNames.Append(Key)
End Function

Public Sub Clear()
    Call Release_pArgs
    Call eArgs_N.Clear
    Call eArgs_P.Clear
    Call pNames.Clear
End Sub

Public Property Get Count() As Long
    Count = (eArgs_P.Count + pCounts + eArgs_N.Count)
End Property

Public Property Get Argument(ByRef Index As Variant) As Argument
Attribute Argument.VB_UserMemId = 0
    Set Argument = New Argument
    Dim myIndex As Long, outIndex As Long, lst As List
    If GetArrayIndex(Index, myIndex, outIndex, lst) Then
        Call Argument.Constructor1(pNames(myIndex), lst(outIndex))
    Else
        Call Argument.Constructor0(pNames(myIndex), pArgs(outIndex))
    End If
End Property

Public Property Get Item(ByRef Index As Variant) As Variant
    Dim myIndex As Long, outIndex As Long, lst As List
    If GetArrayIndex(Index, myIndex, outIndex, lst) Then
        Call API_VariantCopy(Item, lst(outIndex))
    Else
        Call API_VariantCopy(Item, pArgs(outIndex))
    End If
End Property
Public Property Let Item(ByRef Index As Variant, ByRef Value As Variant)
    Dim myIndex As Long, outIndex As Long, lst As List
    If GetArrayIndex(Index, myIndex, outIndex, lst) Then
        Call lst.SetValue(outIndex, Value)
    Else
        Call API_VariantCopy(pArgs(outIndex), Value)
    End If
End Property
Public Property Set Item(ByRef Index As Variant, ByVal Value As Object)
    Dim myIndex As Long, outIndex As Long, lst As List
    If GetArrayIndex(Index, myIndex, outIndex, lst) Then
        Call lst.SetValue(outIndex, Value)
    Else
        Call API_VariantCopy(pArgs(outIndex), Value)
    End If
End Property

Public Property Get ItemType(ByRef Index As Variant) As VariantTypes
    Dim myIndex As Long, outIndex As Long, lst As List
    If GetArrayIndex(Index, myIndex, outIndex, lst) Then
        ItemType = VariantType(lst.GetTypeOf(outIndex))
    Else
        ItemType = VariantType(pArgs(outIndex))
    End If
End Property

Public Function GetValue(ByRef Index As Variant, Optional ByRef DefaultValue As Variant) As Variant
    Dim myIndex As Long, outIndex As Long, lst As List
    If GetArrayIndex(Index, myIndex, outIndex, lst, False) Then
        Call API_VariantCopy(GetValue, lst(outIndex))
    Else
        Call API_VariantCopy(GetValue, pArgs(outIndex))
    End If
End Function
Public Sub SetValue(ByRef Index As Variant, ByRef Value As Variant)
    Dim myIndex As Long, outIndex As Long, lst As List
    If GetArrayIndex(Index, myIndex, outIndex, lst) Then
        Call lst.SetValue(outIndex, Value)
    Else
        Call API_VariantCopy(pArgs(outIndex), Value)
    End If
End Sub

Public Property Get ItemKey(ByVal Index As Long) As Variant
    Call API_VariantCopy(ItemKey, pNames(Index))
End Property
Public Property Let ItemKey(ByVal Index As Long, ByRef Value As Variant)
    Call pNames.SetValue(Index, Value)
End Property
Public Property Set ItemKey(ByVal Index As Long, ByVal Value As Object)
    Call pNames.SetValue(Index, Value)
End Property


Public Property Get Values() As Vector
    Dim Elements As Long
    Set Values = New Vector
    Elements = Count
    Call Values.Constructor0(Elements)
    
    Elements = Elements - 1
    Dim i As Long
    For i = 0 To Elements
        Call Values.SetValue(i, Item(i))
    Next
End Property
Public Property Get Keys() As Vector
    Set Keys = pNames.ToVector
End Property

Public Function IndexOf(ByRef Element As Variant) As Long
    Dim i As Long, mCount As Long
    Dim CMP As IComparer
    Set CMP = Comparer.Default
    mCount = Count - 1
    For i = 0 To mCount
        If CMP.Equals(pNames(i), Element) Then
            IndexOf = i
            Exit Function
        End If
    Next
    IndexOf = -1
End Function
Public Function LastIndexOf(ByRef Element As Variant) As Long
    Dim i As Long, mCount As Long
    Dim CMP As IComparer
    Set CMP = Comparer.Default
    mCount = Count - 1
    For i = mCount To 0 Step -1
        If CMP.Equals(pNames(i), Element) Then
            LastIndexOf = i
            Exit Function
        End If
    Next
    LastIndexOf = -1
End Function
Public Function Contains(ByRef Key As Variant) As Boolean
    Contains = pNames.Contains(Key)
End Function

Public Sub CopyTo(ByRef Arr As Variant, Optional ByVal Index As Long = 0, Optional ByVal Count As Long = -1)
    
End Sub
Public Sub MoveTo(ByRef Arr As Variant)
    
End Sub

Public Function ToArray() As Variant()
    Dim Elements As Long
    Elements = Count
    ToArray = Arrays.CreateSafeVector(atVariantArray, Elements)
    
    Dim i As Long
    For i = 0 To Elements - 1
        'Call API_VariantCopy(ToArray(i), Item(i))
    Next
End Function
Public Function ToDictionary() As Dictionary
    
End Function
Public Function ToList() As List
    Set ToList = New List
    
    Dim V As Variant
    For Each V In Me
        Call ToList.Append(V)
    Next
End Function
Public Function ToVector() As Vector
    Set ToVector = ToList.ToVector
End Function

Public Property Get IsReferring() As Boolean
    IsReferring = Info.IsByRef(pArgs)
End Property

Public Sub Localize()
    Dim PV As Variant
    Call Evaluate(PV, pArgs)
    Call Release_pArgs
    Call MoveVariant(pArgs, PV)
    Dim i As Long
    For i = 0 To pCounts - 1
        If Info.IsRecord(pArgs(i)) Then pArgs(i) = pArgs(i)
    Next
    
    'pArgs = pArgs
    pRelease = True
End Sub


Public Property Get Scheme() As ParametersScheme
    Dim i As Long, outIndex As Long, outZIndex As Long, lst As List
    Set Scheme = New ParametersScheme
    
    For i = 0 To pCounts - 1
        Dim cParam As New ParameterInfo, PType As VariantTypes, pAttrs As ParameterAttributes
        If GetArrayIndex(i, outZIndex, outIndex, lst) Then
            PType = lst.GetTypeOf(outIndex)
        Else
            PType = VariantType(pArgs(i))
        End If
        pAttrs = ((PType And VT_BYREF) = VT_BYREF)
        Call cParam.Constructor0(pNames(i), PType, pAttrs, Null, i)
        Call Scheme.Append(cParam)
    Next
End Property

Public Function ThrowNot(ByVal HowMax As Long, Optional ByVal HowMin As Long = 0)
    Dim mCount As Long
    mCount = Count
    If mCount > HowMax Then throw Exps.ArgumentCountException(HowMax)
    If mCount < HowMin Then throw Exps.ArgumentCountException(HowMin)
End Function
Public Function IsInRange(ByVal HowMax As Long, Optional ByVal HowMin As Long = 0) As Boolean
    Dim mCount As Long
    mCount = Count
    If mCount > HowMax Then Exit Function
    If mCount < HowMin Then Exit Function
    IsInRange = True
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

Public Function Clone() As ArgumentList
    Set Clone = New ArgumentList
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
'Private Sub Class_Initialize()
'   ...
'End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    Call Clear
    Set eArgs_N = Nothing
    Set eArgs_P = Nothing
    Set pNames = Nothing
End Sub

Friend Function GetEnumerableVersion() As Long
    GetEnumerableVersion = EnumVersion
End Function

Friend Function PackForPush(ByVal Scheme As ParametersScheme) As ArgumentList
    Dim V() As Variant, Buff As Variant, pCount As Long
    pCount = Count
    V = Arrays.CreateSafeVector(atVariantArray, pCount)
    
    Dim Attrs As ParameterAttributes, VT As VariantTypes
    Dim myIndex As Long, outIndex As Long, lst As List
    Dim i As Long, Index As Long
    For i = pCount - 1 To 0 Step -1
        If GetArrayIndex(Index, myIndex, outIndex, lst, False) Then
            Call lst.CopyItemToVariant(outIndex, Buff)
        Else
            Call memcpy(Buff, pArgs(outIndex), VLEN_VARIANT)
        End If
        
        'Attrs = Scheme.Item(Index).Attributes
        'VT = VariantType(Buff)
        
        'If Attrs <> VT Then
        '    If (Attrs And paByRef) = paByRef Then
        '        MsgBox "Hal na dashtam benevisam"
        '    Else
        '        MsgBox "Hal nadashtam benevisam"
        '    End If
        'End If
        
        Call memcpy(V(i), Buff, VLEN_VARIANT)
        Index = Index + 1
    Next
    Call memzero(Buff, VLEN_VARIANT)
    
    Dim X As New ArgumentList
    Call X.Set_pArgs(V, Count)
    Set PackForPush = X
    
    Call memzeroarr(V, VLEN_PTR)
End Function
Friend Sub Set_pArgs(That_pArgs() As Variant, ByVal Count As Long)
    pArgs = MemPtr(API_VarPtrArray(That_pArgs))
    VariantType(pArgs) = VT_ARRAY Or VT_VARIANT
    
    pCounts = Count
    pRelease = False
    pShallowRelease = True
    
    Dim i As Long
    For i = 0 To Count - 1
        Call pNames.Append(vbNullString)
    Next
    
    Call memzeroarr(That_pArgs, VLEN_PTR)
End Sub

Private Sub Release_pArgs()
    If pRelease Then
        Call Arrays.Release(pArgs, False)
        Call API_VariantClear(pArgs)
        Call memzero(pArgs, VLEN_VARIANT)
    ElseIf pShallowRelease Then
        Call Arrays.Release(pArgs, True)
    Else
        Call memzero(pArgs, VLEN_VARIANT)
    End If
End Sub

Private Function GetArrayIndex(ByRef Index As Variant, ByRef outZBIndex As Long, ByRef outIndex As Long, ByRef outList As List, Optional ByVal ThrowIfNull As Boolean = True) As Boolean
    Dim fCount As Long, sCount As Long, iIndex As Long
    fCount = eArgs_P.Count
    sCount = eArgs_N.Count
    
    Dim VT As VbVarType
    VT = VarType(Index)
    Select Case VT
        Case vbByte, vbInteger, vbLong, vbDecimal, vbDouble, vbSingle
            iIndex = CLng(Index)
        Case Else
            iIndex = IndexOf(Index)
            If iIndex = -1 Then throw Exps.IndexOutOfRangeException
    End Select
    
    If iIndex < 0 Then
        throw Exps.IndexOutOfRangeException
    ElseIf iIndex < fCount Then
        outZBIndex = iIndex
        outIndex = iIndex
        Set outList = eArgs_P
        GetArrayIndex = True
    ElseIf iIndex < (fCount + pCounts) Then
        outZBIndex = iIndex
        outIndex = iIndex - fCount
        GetArrayIndex = False
    ElseIf iIndex < (fCount + pCounts + sCount) Then
        outZBIndex = iIndex
        outIndex = iIndex - (fCount + pCounts)
        Set outList = eArgs_N
        GetArrayIndex = True
    Else
        If ThrowIfNull Then throw Exps.IndexOutOfRangeException
    End If
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
Private Property Get ICollection_Count() As Long: ICollection_Count = Count: End Property
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
'Private Function IList_Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean: IList_Contains = Contains(Value, Comparer): End Function
Private Sub IList_CopyTo(Arr As Variant, ByVal Index As Long): Call CopyTo(Arr, Index): End Sub
Private Property Get IList_Count() As Long: IList_Count = Count: End Property
Private Function IList_GetEnumerator() As IEnumerator: Set IList_GetEnumerator = GetEnumerator: End Function
Private Function IList_IndexOf(Value As Variant, Optional ByVal Comparer As IComparer) As Long: IList_IndexOf = IndexOf(Value): End Function
'Private Sub IList_Insert(ByVal Index As Long, Value As Variant): Call Insert(Index, Value): End Sub
Private Property Get IList_IsFixedSize() As Boolean: IList_IsFixedSize = False: End Property
Private Property Get IList_IsReadOnly() As Boolean: IList_IsReadOnly = False: End Property
Private Property Get IList_Item(ByVal Index As Long) As Variant: Call MoveVariant(IList_Item, Item(Index)): End Property
Private Property Let IList_Item(ByVal Index As Long, RHS As Variant): Call SetValue(Index, RHS): End Property
Private Property Set IList_Item(ByVal Index As Long, ByVal RHS As Object): Set Item(Index) = RHS: End Property
'Private Function IList_NewEnum() As IUnknown: Set IList_NewEnum = NewEnum: End Function
'Private Sub IList_Remove(Value As Variant, Optional ByVal Comparer As IComparer): Call Remove(Value, Comparer): End Sub
'Private Sub IList_RemoveAt(ByVal Index As Long): Call RemoveAt(Index): End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IDictionary Implementations>
'
Private Function IDictionary_GetEnumerator() As IEnumerator: Set IDictionary_GetEnumerator = GetEnumerator: End Function
Private Function IDictionary_NewEnum() As stdole.IUnknown: Set IDictionary_NewEnum = NewEnum: End Function
Private Sub IDictionary_Add(Key As Variant, Value As Variant): Call Append(Value, Key): End Sub
Private Sub IDictionary_Clear(): Call Clear: End Sub
Private Function IDictionary_Contains(Key As Variant) As Boolean: IDictionary_Contains = Contains(Key): End Function
Private Sub IDictionary_CopyTo(Arr As Variant, ByVal Index As Long): Call CopyTo(Arr, Index): End Sub
Private Property Get IDictionary_Count() As Long: IDictionary_Count = Count: End Property
Private Property Get IDictionary_Item(Key As Variant) As Variant: Call MoveVariant(IDictionary_Item, Item(Key)): End Property
Private Property Let IDictionary_Item(Key As Variant, RHS As Variant): Call SetValue(Key, RHS): End Property
Private Property Set IDictionary_Item(Key As Variant, RHS As Variant): Call SetValue(Key, RHS): End Property
Private Sub IDictionary_Remove(Key As Variant): throw Exps.NotSupportedException: End Sub
Private Property Get IDictionary_IsFixedSize() As Boolean: IDictionary_IsFixedSize = False: End Property
Private Property Get IDictionary_IsReadOnly() As Boolean: IDictionary_IsReadOnly = False: End Property
Private Property Get IDictionary_Keys() As ICollection: Set IDictionary_Keys = Keys: End Property
Private Property Get IDictionary_Values() As ICollection: Set IDictionary_Values = Values: End Property
'
'</section>
'---------------------------------------------
'---------------------------------------------
