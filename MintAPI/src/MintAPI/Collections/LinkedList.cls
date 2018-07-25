VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="LinkedList" tags="">
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
Implements IList
Implements IEnumerable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "LinkedList" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"
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

Dim p_First As LinkedListNode
Dim p_Last As LinkedListNode
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
Friend Sub Constructor0(ByRef InitialElements As Variant)

    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal IEnum As IEnumerable)
    If IEnum Is Nothing Then throw Exps.ArgumentNullException
    

    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal LinkedList As LinkedList)
    If LinkedList Is Nothing Then throw Exps.ArgumentNullException
    'Set myList = LinkedList.GetMyList

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

''<summary>Returns an enumerator that iterates through the LinkedList.</summary>
''<retval>An IEnumerator for the LinkedList.</retval>
Public Function GetEnumerator() As IEnumerator
    'Set GetEnumerator = myList.GetEnumerator
End Function
Public Function NewEnum() As IUnknown
    'Set NewEnum = myList.NewEnum 'CreateEnumerator(GetEnumerator)
End Function

''<summary>Gets the number of elements actually contained in the LinkedList.</summary>
Public Property Get Count() As Long
    
End Property

''<summary>Removes all elements from the LinkedList.</summary>
Public Sub Clear()
    
End Sub

''<summary>Determines the LinkedList is empty or not.</summary>
Public Property Get IsEmpty() As Boolean
    IsEmpty = Not ((p_First Is Nothing) Or (p_Last Is Nothing))
End Property

''<summary>Gets the first node of the LinkedList.</summary>
''<retval>The first element of the LinkedList.</retval>
Public Property Get First() As LinkedListNode: Set First = p_First: End Property
Public Property Let First(ByRef Value As LinkedListNode): Set p_First = Value: End Property
Public Property Set First(ByVal Value As LinkedListNode): Set p_First = Value: End Property

''<summary>Gets the last node of the LinkedList.</summary>
''<retval>The last element of the LinkedList.</retval>
Public Property Get Last() As LinkedListNode: Set Last = p_Last: End Property
Public Property Let Last(ByRef Value As LinkedListNode): Set p_Last = Value: End Property
Public Property Set Last(ByVal Value As LinkedListNode): Set p_Last = Value: End Property

''<summary>Gets the node at the Index of the LinkedList.</summary>
''<retval>The element at the Index of the LinkedList.</retval>
'Public Property Get Item(ByVal Index As Long) As Variant: Call API_VariantCopy(Item, myList(Index)): End Property
'Public Property Let Item(ByVal Index As Long, ByRef Value As Variant): Call myList.SetValue(Index, Value): End Property
'Public Property Set Item(ByVal Index As Long, ByVal Value As Object): Call myList.SetValue(Index, Value): End Property

''<summary>Adds the specified new element after the specified existing element in the LinkedList.</summary>
'Public Sub AddAfter(ByVal Index As Long, ByRef Value As Variant): Call myList.Insert(Index + 1, Value): End Sub
''<summary>Adds the specified new element before the specified existing element in the LinkedList.</summary>
'Public Sub AddBefore(ByVal Index As Long, ByRef Value As Variant): Call myList.Insert(Index, Value): End Sub
''<summary>Adds the specified new element at the start of the LinkedList.</summary>
'Public Sub AddFirst(ByRef Value As Variant): Call myList.Prepend(Value): End Sub
''<summary>Adds the specified new element at the end of the LinkedList.</summary>
'Public Sub AddLast(ByRef Value As Variant): Call myList.Append(Value): End Sub

'''<summary>Removes the first occurrence of the specified value from the LinkedList.</summary>
'Public Sub Remove(ByRef Value As Variant, Optional ByVal Comparer As IComparer): Call myList.Remove(Value, Comparer): End Sub
'''<summary>Removes the element at the Index from the LinkedList.</summary>
'Public Sub RemoveAt(ByVal Index As Long): Call myList.RemoveAt(Index): End Sub
'''<summary>Removes the element at the start of the LinkedList.</summary>
'Public Sub RemoveFirst(): Call myList.RemoveAt(0): End Sub
'''<summary>Removes the element at the end of the LinkedList.</summary>
'Public Sub RemoveLast(): Call myList.RemoveAt(myList.Count - 1): End Sub

''<summary>Copies the LinkedList elements to an existing one-dimensional array, starting at the specified array index.</summary>
''<params>
''  <param name="Arr">The one-dimensional array that is the destination of the elements copied from LinkedList. The array must have zero-based indexing.</param>
''  <param name="Index">The zero-based index in array at which copying begins.</param>
''</params>
''<exceptions>
''  <exception name="ArgumentNullException">Arr is null.</exception>
''  <exception name="ArgumentException">Arr is multidimensional.-or- index is equal to or greater than the length of array.-or- The number of elements in the source LinkedList is greater than the available space from index to the end of the destination array.</exception>
''  <exception name="OutOfRangeException">Index is less than zero. </exception>
''  <exception name="ArrayTypeMismatchException">The type of the source LinkedList cannot be cast automatically to the type of the destination array.</exception>
''</exceptions>
Public Sub CopyTo(ByRef Arr As Variant, ByVal Index As Long)
    
End Sub

''<summary>Determines whether an element is in the LinkedList.</summary>
''<params><param name="Value">The element to locate in the LinkedList. The value can be null.</param></params>
''<retval>true if Value is found in the LinkedList; otherwise, false.</retval>
Public Function Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean
    
End Function


'''<summary>Converts an IEnumerable to an IQueryable.</summary>
'''<retval>An IQueryable of this IEnumerable.</retval>
'Public Function AsQueryable() As IQueryable
'    Set AsQueryable = myList.AsQueryable
'End Function

Public Function ToList() As List
    Set ToList = New List
    Call ToList.Constructor0(Me)
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

Public Function Clone() As LinkedList
    Set Clone = New LinkedList
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
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
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
'<section IEnumerable Implementations>
'
Private Function IEnumerable_GetEnumerator() As IEnumerator: Set IEnumerable_GetEnumerator = GetEnumerator: End Function
Private Function IEnumerable_NewEnum() As IUnknown: Set IEnumerable_NewEnum = NewEnum: End Function
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
Private Function IList_Add(Value As Variant) As Long:: End Function
Private Sub IList_Clear(): Call Clear: End Sub
Private Function IList_Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean: IList_Contains = Contains(Value, Comparer): End Function
Private Sub IList_CopyTo(Arr As Variant, ByVal Index As Long): Call CopyTo(Arr, Index): End Sub
Private Property Get IList_Count() As Long: IList_Count = Count: End Property
Private Function IList_GetEnumerator() As IEnumerator: Set IList_GetEnumerator = GetEnumerator: End Function
Private Function IList_IndexOf(Value As Variant, Optional ByVal Comparer As IComparer) As Long:: End Function
Private Sub IList_Insert(ByVal Index As Long, Value As Variant):: End Sub
Private Property Get IList_IsFixedSize() As Boolean: IList_IsFixedSize = False: End Property
Private Property Get IList_IsReadOnly() As Boolean: IList_IsReadOnly = False: End Property
Private Property Get IList_Item(ByVal Index As Long) As Variant:: End Property
Private Property Set IList_Item(ByVal Index As Long, ByVal RHS As Object):: End Property
Private Property Let IList_Item(ByVal Index As Long, RHS As Variant):: End Property
Private Function IList_NewEnum() As stdole.IUnknown: Set IList_NewEnum = NewEnum: End Function
Private Sub IList_Remove(Value As Variant, Optional ByVal Comparer As IComparer):: End Sub
Private Sub IList_RemoveAt(ByVal Index As Long):: End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
