VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParametersScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="ParametersScheme" tags="immutable">
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
Implements IList

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ParametersScheme" '#DO NOT REMOVE IT
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

Dim InnerList As New List
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
'Friend Sub Constructor0(ByVal ID As Long)
'
'    This_Constructed = True
'End Sub
Friend Sub ConstructorCopy(ByVal That As ParametersScheme)

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
    Set GetEnumerator = InnerList.GetEnumerator
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

Public Property Get Count() As Long
    Count = InnerList.Count
End Property


Public Property Get Size() As Long
    Dim pSize As Long, V, PI As ParameterInfo
    For Each V In Me
        Set PI = V
        pSize = pSize + PI.Size
    Next
    Size = pSize
End Property
Public Property Get OnStackSize() As Long
    Dim Size As Long, V, PI As ParameterInfo
    For Each V In Me
        Set PI = V
        Size = Size + PI.OnStackSize
    Next
    OnStackSize = Size
End Property

Friend Sub Append(ByVal Parameter As ParameterInfo)
    Call InnerList.Append(Parameter)
End Sub

Public Property Get Item(ByVal IndexKey As Variant) As ParameterInfo
Attribute Item.VB_UserMemId = 0
    Set Item = InnerList(IndexOf(IndexKey))
End Property

Public Function IndexOf(ByVal IndexKey As Variant) As Long
    Dim VT As VariantTypes
    VT = VarType(IndexKey)
    Select Case VT
        Case VT_I4, VT_I2, VT_INT, VT_I1, VT_I8, VT_Double, VT_Single, VT_Byte
            IndexOf = CLng(IndexKey)
        Case VT_BSTR, VT_LPSTR, VT_LPWSTR, VT_BSTR_BLOB
            Dim PI As ParameterInfo, i As Long, strKey As String
            strKey = CStr(IndexKey)
            For i = 0 To InnerList.Count - 1
                Set PI = InnerList(i)
                If PI.Name = strKey Then
                    IndexOf = i
                    Exit Function
                End If
            Next
        Case Else
            throw Exps.InvalidArgumentException("IndexKey")
    End Select
    IndexOf = -1
End Function
Public Function LastIndexOf(ByVal IndexKey As Variant) As Long
    Dim VT As VariantTypes
    VT = VarType(IndexKey)
    Select Case VT
        Case VT_I4, VT_I2, VT_INT, VT_I1, VT_I8, VT_Double, VT_Single, VT_Byte
            LastIndexOf = CLng(IndexKey)
        Case VT_BSTR, VT_LPSTR, VT_LPWSTR, VT_BSTR_BLOB
            Dim PI As ParameterInfo, i As Long, strKey As String
            strKey = CStr(IndexKey)
            For i = InnerList.Count - 1 To 0 Step -1
                Set PI = InnerList(i)
                If PI.Name = strKey Then
                    LastIndexOf = i
                    Exit Function
                End If
            Next
        Case Else
            throw Exps.InvalidArgumentException("IndexKey")
    End Select
    LastIndexOf = -1
End Function

Public Function Validate(ByVal Args As ArgumentList) As Boolean
    If Args Is Nothing Then
        If InnerList.Count = 0 Then Validate = True
        Exit Function
    End If
    If InnerList.Count <> Args.Count Then Exit Function 'Return False
    
    Dim i As Long, PI As ParameterInfo
    For i = 0 To Args.Count - 1
        Set PI = InnerList(i)
        If Not PI.IsMatch(Args(i).Value) Then Exit Function 'Return False
    Next
    
    Validate = True
End Function

Public Sub CopyTo(ByRef Arr As Variant, Optional ByVal Index As Long = 0, Optional ByVal Count As Long = -1)
    Call InnerList.CopyTo(Arr, Index, Count)
End Sub
Public Function Contains(Value As Variant, Optional ByVal Comparer As IComparer = Nothing) As Boolean
    Contains = InnerList.Contains(Value, Comparer)
End Function


Public Function AsQueryable() As IQueryable
    Set AsQueryable = InnerList.AsQueryable
End Function
'
Public Function ToList() As List
    Set ToList = InnerList.Clone
End Function
Public Function ToVector() As Vector
    Set ToVector = InnerList.ToVector
End Function
Public Function ToArray() As Variant
    Set ToArray = InnerList.ToArray
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

Public Function Clone() As ParametersScheme
    
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
Private Function IList_Add(Value As Variant) As Long: throw Exps.NotSupportedException: End Function
Private Sub IList_Clear(): throw Exps.NotSupportedException: End Sub
Private Function IList_Contains(Value As Variant, Optional ByVal Comparer As IComparer) As Boolean: IList_Contains = Contains(Value, Comparer): End Function
Private Sub IList_CopyTo(Arr As Variant, ByVal Index As Long): Call CopyTo(Arr, Index): End Sub
Private Property Get IList_Count() As Long: IList_Count = Count: End Property
Private Function IList_GetEnumerator() As IEnumerator: Set IList_GetEnumerator = GetEnumerator: End Function
Private Function IList_IndexOf(Value As Variant, Optional ByVal Comparer As IComparer) As Long: IList_IndexOf = IndexOf(Value): End Function
Private Sub IList_Insert(ByVal Index As Long, Value As Variant): throw Exps.NotSupportedException: End Sub
Private Property Get IList_IsFixedSize() As Boolean: IList_IsFixedSize = True: End Property
Private Property Get IList_IsReadOnly() As Boolean: IList_IsReadOnly = True: End Property
Private Property Get IList_Item(ByVal Index As Long) As Variant: Call MoveVariant(IList_Item, Item(Index)): End Property
Private Property Let IList_Item(ByVal Index As Long, RHS As Variant): throw Exps.NotSupportedException: End Property
Private Property Set IList_Item(ByVal Index As Long, ByVal RHS As Object): throw Exps.NotSupportedException: End Property
Private Function IList_NewEnum() As IUnknown: Set IList_NewEnum = NewEnum: End Function
Private Sub IList_Remove(Value As Variant, Optional ByVal Comparer As IComparer): throw Exps.NotSupportedException: End Sub
Private Sub IList_RemoveAt(ByVal Index As Long): throw Exps.NotSupportedException: End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

