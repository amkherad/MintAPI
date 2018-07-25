VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumerableDelegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="EnumerableDelegate" tags="">
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
Implements IEnumerable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "EnumerableDelegate" '#DO NOT REMOVE IT
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

Private p_Enumerable As List
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
    Set p_Enumerable = List
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal That As EnumerableDelegate)
    Set p_Enumerable = That.ToList
    
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
    Set GetEnumerator = p_Enumerable.GetEnumerator
End Function
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = 0
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = p_Enumerable.NewEnum
End Function

Public Function Count() As Long
    Count = p_Enumerable.Count
End Function
Public Function LongCount() As Int64
    LongCount = Convert.CInt64(p_Enumerable.Count)
End Function

Public Function IsAny(Optional ByVal Predicate As Method) As Boolean
'Bool Predicate(Variant)
    Dim V As Variant
    If Predicate Is Nothing Then
        For Each V In p_Enumerable
            If Not IsNullOrEmpty(V) Then
                IsAny = True
                Exit Function
            End If
        Next
    Else
        For Each V In p_Enumerable
            If CBool(Predicate(AL.Create(V))) Then
                IsAny = True
                Exit Function
            End If
        Next
    End If
    IsAny = False
End Function
Public Function IsAll(ByVal Predicate As Method) As Boolean
'Bool Predicate(Variant)
    Dim V As Variant
    IsAll = False
    For Each V In p_Enumerable
        If Not CBool(Predicate(AL.Create(V))) Then
            Exit Function
        Else
            IsAll = True
        End If
    Next
End Function
Public Function Where(ByVal Predicate As Method) As EnumerableDelegate
'Bool Predicate(Variant)
    Dim RetList As New List
    Dim V As Variant
    For Each V In p_Enumerable
        If Not CBool(Predicate(AL.Create(V))) Then
            Call RetList.Append(V)
        End If
    Next
    Set Where = CreateDelegate(RetList)
End Function
Public Function SelectAs(ByVal Predicate As Method) As EnumerableDelegate
'Variant Predicate(Variant)
    Dim RetList As New List
    Dim V As Variant
    For Each V In p_Enumerable
        Call RetList.Append(Predicate(AL.Create(V)))
    Next
    Set SelectAs = CreateDelegate(RetList)
End Function
Public Function Skip(ByVal Count As Long) As EnumerableDelegate
    
End Function
Public Function Take(ByVal Count As Long) As EnumerableDelegate
    
End Function
Public Sub ForEach(ByVal Action As Method)
'void Action(Variant)
    
End Sub
Public Function Distinct(Optional ByVal Predicate As Method) As EnumerableDelegate
'void Predicate(Variant)
    
End Function
'Public Function Sort(ByVal Predicate As Method) As IEnumerable
'
'End Function
Public Function Cast(ByVal Predicate As Method) As IEnumerable

End Function
Public Function OrderBy(Optional ByVal Predicate As Method) As EnumerableDelegate
    
End Function
Public Function OrderByDescending(Optional ByVal Predicate As Method) As EnumerableDelegate
    
End Function

Public Function ToList() As List
'    Set ToList = New List
'    Dim V As Variant
'    For Each V In p_Enumerable
'        Call ToList.Append(V)
'    Next
    Set ToList = p_Enumerable.Clone
End Function
Public Function ToVector() As Vector
    Set ToVector = ToList.ToVector
End Function
Public Function ToArray() As Variant
    Call EvaluateX(ToArray, ToList.ToArray)
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

Public Function Clone() As EnumerableDelegate
    Set Clone = New EnumerableDelegate
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

Private Function CreateDelegate(ByVal List As List) As EnumerableDelegate
    Set CreateDelegate = New EnumerableDelegate
    Call CreateDelegate.Constructor0(List)
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
'<section ICloneable Implementations>
'
Private Function IEnumerable_GetEnumerator() As IEnumerator: Set IEnumerable_GetEnumerator = GetEnumerator: End Function
Private Function IEnumerable_NewEnum() As stdole.IUnknown: Set IEnumerable_NewEnum = NewEnum: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
