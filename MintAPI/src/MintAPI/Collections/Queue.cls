VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="Queue" tags="collection">
''  <summary>
''    Represents a first-in, last-out (FILO) collection of objects.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>List</using>
''  <idea>Me</idea>
''  <sample>
''      Dim Q As New Queue
''      Q.Enqueue("Ali")
''      stdout Q.Dequeue 'Ali
''</sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
Implements ICloneable
Implements IEnumerable
Implements ICollection

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Queue" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"

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

Dim myList As List
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
''  <carg type="variant" name="InitialElements"></carg>
''  <carg type="List" name="List"></carg>
''  <carg type="Queue" name="Queue"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByRef InitialElements As Variant)
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal IEnum As IEnumerable)
    If IEnum Is Nothing Then throw Exps.ArgumentNullException
    Set myList = Enumerable.ToList(IEnum)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal Queue As Queue)
    If Queue Is Nothing Then throw Exps.ArgumentNullException
    Set myList = Queue.GetMyList
    
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

''<summary>Returns an enumerator that iterates through the Queue.</summary>
''<retval>An IEnumerator for the Queue.</retval>
Public Function GetEnumerator() As IEnumerator
    Set GetEnumerator = myList.GetEnumerator
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = myList.NewEnum 'CreateEnumerator(GetEnumerator)
End Function

''<summary>Gets the number of elements contained in the Queue.</summary>
''<retval>The number of elements contained in the Queue.</retval>
Public Property Get Count() As Long
    Count = myList.Count
End Property

''<summary>Removes all elements from the Queue.</summary>
Public Sub Clear()
    Call myList.Clear
End Sub

''<summary>Adds an element to the end of the Queue.</summary>
''<params><param name="Value">The element to add to the Queue. The value can be null.</param></params>
Public Sub Enqueue(ByRef Value As Variant)
    Call myList.Append(Value) 'The only diffrence with Stack
End Sub
''<summary>Removes and returns the element at the beginning of the Queue.</summary>
''<retval>The element that is removed from the beginning of the Queue.</retval>
''<exceptions>
''  <exception name="InvalidOperationException">The Queue is empty.</exception>
''</exceptions>
Public Function Dequeue() As Variant
    Call API_VariantCopy(Dequeue, myList.Item(0))
    Call myList.RemoveAt(0)
End Function
''<summary>
''  Returns the element at the beginning of the Queue without removing it.
''</summary>
''<retval>The element at the beginning of the Queue.</retval>
''<exceptions>
''  <exception name="InvalidOperationException">The Queue is empty.</exception>
''</exceptions>
Public Function Peek() As Variant
    Call API_VariantCopy(Peek, myList.Item(0))
End Function

''<summary>Determines whether an element is in the Queue.</summary>
''<params><param name="Value">The element to locate in the Queue. The value can be null.</param></params>
''<retval>true if Value is found in the Queue; otherwise, false.</retval>
Public Function Contains(ByRef Value As Variant, Optional ByVal Comparer As IComparer = Nothing) As Boolean
    Contains = myList.Contains(Value, Comparer)
End Function

''<summary>Copies the Queue elements to an existing one-dimensional array, starting at the specified array index.</summary>
''<params>
''  <param name="Arr">The one-dimensional array that is the destination of the elements copied from Queue. The array must have zero-based indexing.</param>
''  <param name="Index">The zero-based index in array at which copying begins.</param>
''</params>
''<exceptions>
''  <exception name="ArgumentNullException">Arr is null.</exception>
''  <exception name="ArgumentException">Arr is multidimensional.-or- index is equal to or greater than the length of array.-or- The number of elements in the source Queue is greater than the available space from index to the end of the destination array.</exception>
''  <exception name="OutOfRangeException">Index is less than zero. </exception>
''  <exception name="ArrayTypeMismatchException">The type of the source Queue cannot be cast automatically to the type of the destination array.</exception>
''</exceptions>
Public Sub CopyTo(ByRef Arr As Variant, ByVal Index As Long)
    Call myList.CopyTo(Arr, Index)
End Sub

''<summary>Converts an IEnumerable to an IQueryable.</summary>
''<retval>An IQueryable of this IEnumerable.</retval>
Public Function AsQueryable() As IQueryable
    Set AsQueryable = myList.AsQueryable
End Function

''<summary>Provides access to Stack elements as list, which is editable.</summary>
''<retval>A list of Stack source elements.</retval>
Public Function AsList() As List
    Set AsList = myList
End Function

''<summary>Copies the Queue elements to a List.</summary>
''<retval>A List containing copies of the elements of the Queue.</retval>
Public Function ToList() As List
    Set ToList = myList.Clone
End Function

'<summary>Copies the Queue elements to a Vector.</summary>
'<retval>A Vector containing copies of the elements of the Queue.</retval>
Public Function ToVector() As Vector
    Set ToVector = myList.ToVector
End Function

'<summary>Copies the Queue elements to a new array.</summary>
'<retval>A new array containing copies of the elements of the Queue.</retval>
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

Public Function Clone() As Queue
    Set Clone = New Queue
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
    Set myList = New List
End Sub

Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    Set myList = Nothing
End Sub

Friend Function GetMyList() As List
    Set GetMyList = myList
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
Private Function IEnumerable_NewEnum() As stdole.IUnknown: Set IEnumerable_NewEnum = NewEnum: End Function
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
