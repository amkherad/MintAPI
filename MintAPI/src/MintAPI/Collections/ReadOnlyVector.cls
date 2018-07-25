VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadOnlyVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="ReadOnlyVector" tags="immutable">
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
'Implements Vector
Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ReadOnlyVector" '#DO NOT REMOVE IT
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

Dim pArray As Variant
Dim pSA As Long
Dim pCounts As Long
'Dim PType As VbVarType
'Dim pChkType As Boolean
'
'Dim EnumVersion As Long
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
Friend Sub Constructor0(ByVal Size As Long)
    If Size = 0 Then throw Exps.InvalidArgumentException("Size")
    If Size < 0 Then throw Exps.InvalidArgumentException("Size")
    
    Dim aPtr As Long
    aPtr = API_SafeArrayCreateVectorEx(vbVariant, 0, Size)
    If aPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    pArray = aPtr
    VariantType(pArray) = VT_ARRAY Or VT_VARIANT
    pSA = Arrays.GetSafeArrayPointer(pArray, True)
    pCounts = Size
End Sub
Friend Sub Constructor1(ByVal FromList As IList, Optional ByVal Index As Long = 0, Optional ByVal Count As Long = -1)
    Dim EndAt As Long, lstCount As Long
    lstCount = FromList.Count
    If Count = -1 Then
        EndAt = lstCount
    Else
        EndAt = Index + Count
    End If
    If (Index < 0) Or (EndAt > lstCount) Then throw Exps.IndexOutOfRangeException
    
    lstCount = EndAt - Index
    EndAt = EndAt - 1
    
    'If buffCount = 0 Then throw Exps.ListIsEmptyException
    
    Dim aPtr As Long
    aPtr = API_SafeArrayCreateVectorEx(vbVariant, 0, lstCount)
    If aPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    pArray = aPtr
    VariantType(pArray) = VT_ARRAY Or VT_VARIANT
    pSA = Arrays.GetSafeArrayPointer(pArray, True)
    Dim i As Long
    For i = Index To EndAt
        Call API_VariantCopy(pArray(i - Index), FromList(i))
    Next
    pCounts = lstCount
End Sub
Friend Sub Constructor2(ByRef ArrToRefer As Variant)
    
End Sub
Friend Sub Constructor3(ByRef ArrToClone As Variant)
    
End Sub
Friend Sub ConstructorCopy(ByVal ROV As ReadOnlyVector)
    
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
    Dim LEnumerator As New VectorEnumerator
    Call LEnumerator.Constructor0(Me)
    Set GetEnumerator = LEnumerator
End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = CreateEnumerator(GetEnumerator)
End Function

Public Property Get Length() As Long
    Length = pCounts
End Property

Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If (Index < 0) Or (Index >= pCounts) Then throw Exps.IndexOutOfRangeException
    Call API_VariantCopy(Item, pArray(Index))
End Property

Public Sub CopyTo(ByRef Arr As Variant, ByVal Index As Long)
    
End Sub

Public Property Get ElementSize() As Long
    ElementSize = API_SafeArrayGetElemsize(pSA)
End Property


Public Function AsQueryable() As IQueryable
    
End Function

'Public Function ToList() As List
'    Dim i As Long, Count As Long
'    Count = pCounts - 1
'    Dim lstRetVal As New List
'    For i = 0 To Count
'        Call lstRetVal.Append(Item(i))
'    Next
'    Set ToList = lstRetVal
'End Function
'Public Function ToArray() As Variant
'    ToArray = pArray
'End Function
'Public Function ToVector() As Vector
'    Set ToVector = New Vector
'    Call ToVector.Constructor3(pArray, atVariantArray, False)
'End Function
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
Public Property Get Alias() As String
    Alias = CLASSALIASSTRING
End Property
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

Public Function Clone() As ReadOnlyVector
    
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
'<section Vector Implementation>
'
Private Property Let Vector_CheckTypes(RHS As Boolean)

End Property
Private Property Get Vector_CheckTypes() As Boolean

End Property
Private Sub Vector_Clear()

End Sub
Private Function Vector_Clone() As Vector

End Function
Private Property Get Vector_DataHandle() As Long

End Property
Private Sub Vector_Dispose()

End Sub
Private Function Vector_GetUnlocked() As Vector

End Function
Private Property Get Vector_Handle() As Long

End Property
Private Property Set Vector_Item(Index As Long, RHS As Object)

End Property
Private Property Let Vector_Item(Index As Long, RHS As Variant)

End Property

Private Property Get Vector_Item(Index As Long) As Variant

End Property
Private Property Get Vector_Length() As Long

End Property
Private Sub Vector_MakeItReadOnly()

End Sub
Private Sub Vector_Swap(destinationIndex As Long, sourceIndex As Long)

End Sub
Private Function Vector_ToArray() As Variant

End Function
Private Function Vector_ToList() As List

End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
