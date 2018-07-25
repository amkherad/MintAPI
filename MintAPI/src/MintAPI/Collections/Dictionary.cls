VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="Dictionary" tags="">
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
Implements IEnumerable
Implements IDictionary

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Dictionary" '#DO NOT REMOVE IT
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
Public Type KeyValuePair
    Key As Variant
    Value As Variant
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

Dim InnerList As New List

Dim EnumVersion As Long
'Dim DQ As DictionaryQueryable

Dim p_AutoCreateOnSet As Boolean
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
Friend Sub Constructor0(ByVal ID As Long)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As Dictionary)

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
    Dim DEnumerator As New DictionaryEnumerator
    Call DEnumerator.Constructor0(Me)
    Set GetEnumerator = DEnumerator
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

Public Sub Add(ByRef Key As Variant, ByRef Value As Variant)
    Dim Index As Long, KV As KeyValuePair
    Index = IndexOf(Key)
    
    If Index <> -1 Then throw Exps.ItemExistsException
    
    EnumVersion = EnumVersion + 1
    Call Evaluate(KV.Key, Key)
    Call Evaluate(KV.Value, Value)
    
    Call InnerList.Append(KV)
End Sub

Public Property Get Count() As Long
    Count = InnerList.Count
End Property

Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim Index As Long, KV As KeyValuePair
    Index = IndexOf(Key)
    If Index = -1 Then throw Exps.ItemNotExistsException
    
    KV = InnerList(Index)
    
    Call EvaluateX(Item, KV.Value)
End Property
Public Property Let Item(ByRef Key As Variant, ByRef Value As Variant)
    Dim Index As Long, KV As KeyValuePair
    Index = IndexOf(Key)
    If Index = -1 Then
        If p_AutoCreateOnSet Then
            Call Add(Key, Value)
            Exit Property
        Else
            throw Exps.ItemNotExistsException
        End If
    End If
    
    KV = InnerList(Index)
    Call Evaluate(KV.Value, Value)
    EnumVersion = EnumVersion + 1
    InnerList(Index) = KV
End Property
Public Property Set Item(ByRef Key As Variant, ByVal Value As Object)
    Dim Index As Long, KV As KeyValuePair
    Index = IndexOf(Key)
    If Index = -1 Then
        If p_AutoCreateOnSet Then
            Call Add(Key, Value)
            Exit Property
        Else
            throw Exps.ItemNotExistsException
        End If
    End If
    
    KV = InnerList(Index)
    Call Evaluate(KV.Value, Value)
    EnumVersion = EnumVersion + 1
    InnerList(Index) = KV
End Property
Public Sub SetValue(ByRef Key As Variant, ByRef Value As Variant)
    Dim Index As Long, KV As KeyValuePair
    Index = IndexOf(Key)
    If Index = -1 Then
        If p_AutoCreateOnSet Then
            Call Add(Key, Value)
            Exit Sub
        Else
            throw Exps.ItemNotExistsException
        End If
    End If
    
    KV = InnerList(Index)
    Call Evaluate(KV.Value, Value)
    EnumVersion = EnumVersion + 1
    InnerList(Index) = KV
End Sub

Public Function TryGetValue(ByRef Key As Variant, ByRef Value As Variant) As Boolean
    Dim Index As Long, KV As KeyValuePair
    Index = IndexOf(Key)
    If Index = -1 Then TryGetValue = False: Exit Function
    
    KV = InnerList(Index)
    
    Call Evaluate(Value, KV.Value)
    TryGetValue = True
End Function

Friend Property Get AutoCreateOnSet() As Boolean
    AutoCreateOnSet = p_AutoCreateOnSet
End Property
Friend Property Let AutoCreateOnSet(ByVal Value As Boolean)
    p_AutoCreateOnSet = Value
End Property

Public Sub Remove(ByRef Key As Variant)
    Dim Index As Long
    Index = IndexOf(Key)
    If Index = -1 Then throw Exps.ItemNotExistsException
    EnumVersion = EnumVersion + 1
    Call InnerList.RemoveAt(Index)
End Sub

Public Sub Clear()
    Call InnerList.Clear
End Sub

Public Property Get Keys() As ICollection
    Dim lst As New List
    Dim KV As KeyValuePair
    
    Dim i As Long
    For i = 0 To InnerList.Count - 1
        KV = InnerList(i)
        Call lst.Append(KV.Key)
    Next
    
    Set Keys = lst.ToVector
End Property
Public Property Get Values() As ICollection
    Dim lst As New List
    Dim KV As KeyValuePair
    
    Dim i As Long
    For i = 0 To InnerList.Count - 1
        KV = InnerList(i)
        Call lst.Append(KV.Value)
    Next
    
    Set Values = lst.ToVector
End Property

Public Sub CopyTo(ByRef Arr As Variant, ByVal Index As Long)
    Call Values.CopyTo(Arr, Index)
End Sub
Public Function Contains(ByRef Key As Variant) As Boolean
    Contains = (IndexOf(Key) >= 0)
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

Public Function Clone() As Dictionary
    Set Clone = New Dictionary
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
'    ...
'End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Friend Function ItemAt(ByVal Index As Long) As KeyValuePair
    ItemAt = InnerList(Index)
End Function
Private Function IndexOf(ByRef Key As Variant) As Long
    Dim i As Long
    Dim Comp As IComparer
    Dim KV As KeyValuePair
    
    Set Comp = Comparer.Default
    For i = 0 To InnerList.Count - 1
        KV = InnerList(i)
        If Comp.Equals(Key, KV.Key) Then
            IndexOf = i
            Exit Function
        End If
    Next
    IndexOf = -1
End Function

Friend Function GetEnumerableVersion() As Long
    GetEnumerableVersion = EnumVersion + InnerList.GetEnumerableVersion
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
'Private Function IList_Add(Value As Variant) As Long: IList_Add = Append(Value): End Function
'Private Sub IList_Clear(): Call Clear: End Sub
'Private Function IList_Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean: IList_Contains = Contains(Value, Comparer): End Function
'Private Sub IList_CopyTo(Arr As Variant, ByVal Index As Long): Call CopyTo(Arr, Index): End Sub
'Private Property Get IList_Count() As Long: IList_Count = InnerList.Count: End Property
'Private Function IList_GetEnumerator() As IEnumerator: Set IList_GetEnumerator = GetEnumerator: End Function
'Private Function IList_IndexOf(Value As Variant, Optional ByVal Comparer As IComparer) As Long: IList_IndexOf = IndexOf(Value): End Function
'Private Sub IList_Insert(ByVal Index As Long, Value As Variant): Call Insert(Index, Value): End Sub
'Private Property Get IList_IsFixedSize() As Boolean: IList_IsFixedSize = False: End Property
'Private Property Get IList_IsReadOnly() As Boolean: IList_IsReadOnly = False: End Property
'Private Property Get IList_Item(ByVal Index As Long) As Variant: Call MoveVariant(IList_Item, Item(Index)): End Property
'Private Property Let IList_Item(ByVal Index As Long, RHS As Variant): Call SetValue(Index, RHS): End Property
'Private Property Set IList_Item(ByVal Index As Long, ByVal RHS As Object): Set Item(Index) = RHS: End Property
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
'<section List Implementations>
'
Private Sub IDictionary_Add(Key As Variant, Value As Variant): Call Add(Key, Value): End Sub
Private Sub IDictionary_Clear(): Call Clear: End Sub
Private Function IDictionary_Contains(Key As Variant) As Boolean: IDictionary_Contains = Contains(Key): End Function
Private Sub IDictionary_CopyTo(Arr As Variant, ByVal Index As Long): Call CopyTo(Arr, Index): End Sub
Private Property Get IDictionary_Count() As Long: IDictionary_Count = Count: End Property
Private Function IDictionary_GetEnumerator() As IEnumerator: Set IDictionary_GetEnumerator = GetEnumerator: End Function
Private Property Get IDictionary_IsFixedSize() As Boolean: IDictionary_IsFixedSize = False: End Property
Private Property Get IDictionary_IsReadOnly() As Boolean: IDictionary_IsReadOnly = False: End Property
Private Property Set IDictionary_Item(Key As Variant, RHS As Variant): Call SetValue(Key, RHS): End Property
Private Property Let IDictionary_Item(Key As Variant, RHS As Variant): Call SetValue(Key, RHS): End Property
Private Property Get IDictionary_Item(Key As Variant) As Variant: Call MoveVariant(IDictionary_Item, Item(Key)): End Property
Private Property Get IDictionary_Keys() As ICollection: Set IDictionary_Keys = Keys: End Property
Private Function IDictionary_NewEnum() As stdole.IUnknown: Set IDictionary_NewEnum = NewEnum: End Function
Private Sub IDictionary_Remove(Key As Variant): Call Remove(Key): End Sub
Private Property Get IDictionary_Values() As ICollection: Set IDictionary_Values = Values: End Property
'
'</section>
'---------------------------------------------
'---------------------------------------------
