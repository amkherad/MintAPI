VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="Vector" tags="">
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
'
Implements IObject
Implements ICollection
Implements ICloneable
Implements IEnumerable
'Implements IProvideEnumerableVersion

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Vector" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"
''
''</section>
''---------------------------------------------
''---------------------------------------------
'
''=============================================
''=============================================
''=============================================
''<section API/Type Declarations>
''
'
''
''</section>
''---------------------------------------------
''---------------------------------------------
'
''=============================================
''=============================================
''=============================================
''<section Variables/Events>
''
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Dim pArray As Variant
Dim pSA As Long
Dim pCounts As Long
Dim PType As VbVarType
Dim pChkType As Boolean

Dim EnumVersion As Long

Dim VQ As VectorQueryable
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
Friend Sub Constructor0(ByVal Size As Long, Optional ByVal ArrayType As ArrayTypes = atVariantArray, Optional ByVal CheckVarType As Boolean = False)
    If Size = 0 Then throw Exps.InvalidArgumentException("Size")
    If Size < 0 Then throw Exps.InvalidArgumentException("Size")

    Dim pPtr As Long
    pPtr = API_SafeArrayCreateVectorEx(vbVariant, 0, Size)
    If pPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    pArray = pPtr
    VariantType(pArray) = VT_ARRAY Or VT_VARIANT
    pSA = Arrays.GetSafeArrayPointer(pArray, True)
    pCounts = Size

    PType = ArrayType
    pChkType = CheckVarType
End Sub
Friend Sub Constructor1(ByVal FromList As IList, Optional ByVal Index As Long = 0, Optional ByVal Count As Long = -1, Optional ByVal ArrayType As ArrayTypes = atVariantArray, Optional ByVal CheckVarType As Boolean = False)
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

    Dim pPtr As Long
    pPtr = API_SafeArrayCreateVectorEx(vbVariant, 0, lstCount)
    If pPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    pArray = pPtr
    VariantType(pArray) = VT_ARRAY Or VT_VARIANT
    pSA = Arrays.GetSafeArrayPointer(pArray, True)
    Dim i As Long
    For i = Index To EndAt
        Call API_VariantCopy(pArray(i - Index), FromList(i))
    Next
    pCounts = lstCount

    PType = ArrayType
    pChkType = CheckVarType
End Sub
Friend Sub Constructor2(ByRef ArrToRefer As Variant, Optional ByVal ArrayType As ArrayTypes = vbVariant, Optional ByVal CheckVarType As Boolean = False)

End Sub
Friend Sub Constructor3(ByRef ArrToClone As Variant, Optional ByVal ArrayType As ArrayTypes = vbVariant, Optional ByVal CheckVarType As Boolean = False)

End Sub
Friend Sub ConstructorCopy(ByVal Vector As Vector)

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

Public Property Get CheckTypes() As Boolean
    CheckTypes = pChkType
End Property
Public Property Let CheckTypes(ByVal Value As Boolean)
    If pCounts > 0 Then throw Exps.InvalidOperationException("Vector must be empty to do this action.")
    pChkType = Value
End Property

Public Sub Clear()
    If IsArray(pArray) Then _
        Call Arrays.Release(pArray)

    pCounts = 0
    PType = vbNull
    pChkType = False

    EnumVersion = 0
End Sub

Friend Property Get Handle() As Long
    Handle = VariantAbsoluteDataPtr(pArray)
End Property
Friend Property Get DataHandle() As Long
    DataHandle = Arrays.GetSafeArrayPointer(pArray, False)
End Property

Public Property Get Length() As Long
    Length = pCounts
End Property

Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If (Index < 0) Or (Index >= pCounts) Then throw Exps.IndexOutOfRangeException
    Call API_VariantCopy(Item, pArray(Index))
End Property
Public Property Let Item(ByVal Index As Long, Value As Variant)
    If (Index < 0) Or (Index >= pCounts) Then throw Exps.IndexOutOfRangeException
    Call API_VariantCopy(pArray(Index), Value)
End Property
Public Property Set Item(ByVal Index As Long, ByVal Value As Object)
    If (Index < 0) Or (Index >= pCounts) Then throw Exps.IndexOutOfRangeException
    Call API_VariantCopy(pArray(Index), Value)
End Property

Public Sub SetValue(ByVal Index As Long, ByRef Value As Variant)
    Call API_VariantCopy(pArray(Index), Value)
End Sub

Public Sub CopyTo(ByRef Arr As Variant, ByVal Index As Long)

End Sub

Public Property Get ElementSize() As Long
    ElementSize = API_SafeArrayGetElemsize(pSA)
End Property

Public Sub Swap(ByVal Index1 As Long, ByVal Index2 As Long)
    If (Index1 < 0) Or (Index1 >= pCounts) Then throw Exps.IndexOutOfRangeException
    If (Index2 < 0) Or (Index2 >= pCounts) Then throw Exps.IndexOutOfRangeException

    Dim MemV(3) As Long, MemVPtr As Long '16 Bytes
    Dim p1Ptr As Long, p2Ptr As Long
    MemVPtr = VarPtr(MemV(0))
    p1Ptr = VarPtr(pArray(Index1))
    p2Ptr = VarPtr(pArray(Index2))
    Call memcpy(ByVal MemVPtr, ByVal p1Ptr, VLEN_VARIANT)
    Call memcpy(ByVal p1Ptr, ByVal p2Ptr, VLEN_VARIANT)
    Call memcpy(ByVal p2Ptr, ByVal MemVPtr, VLEN_VARIANT)
    Call memzero(ByVal MemVPtr, VLEN_VARIANT)
End Sub


Public Function AsQueryable() As IQueryable
    If VQ Is Nothing Then
        Set VQ = New VectorQueryable
        Call VQ.Constructor0(Me)
    End If
    Set AsQueryable = VQ
End Function
Public Function AsReadOnly() As ReadOnlyVector
    Set AsReadOnly = New ReadOnlyVector
    Call AsReadOnly.Constructor3(pArray)
End Function
Friend Function AsReadOnly_destroyer() As ReadOnlyVector
    Set AsReadOnly_destroyer = New ReadOnlyVector
    Call AsReadOnly_destroyer.Constructor2(pArray)
    Call memzero(pArray, VLEN_VARIANT)
End Function

Public Function ToList() As List
    Dim i As Long, Count As Long
    Count = pCounts - 1
    Dim lstRetVal As New List
    For i = 0 To Count
        Call lstRetVal.Append(Item(i))
    Next
    Set ToList = lstRetVal
End Function
Public Function ToArray() As Variant
    If IsArray(pArray) Then _
        ToArray = pArray
        'Call API_VariantCopyInd(ToArray, pArray)
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

Public Function Clone() As Vector
    Set Clone = New Vector
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

Friend Function GetEnumerableVersion() As Long
    GetEnumerableVersion = EnumVersion
End Function

''<summary>Swaps pArray to Buf and release the pArray (class array pointer) class become unusable unless calls to constructors.</summary>
Friend Sub SwapBuffer(ByRef Buf As Variant)
'* UNSAFE *
    Dim Buf_IsByRef As Boolean, BufVT As VariantTypes
    BufVT = VariantType(Buf)
    Buf_IsByRef = (BufVT And VT_BYREF) = VT_BYREF

    If BufVT = 0 Then
        Call memcpy(ByVal VarPtr(Buf), ByVal VarPtr(pArray), VLEN_VARIANT)
    Else
        If Buf_IsByRef Then BufVT = BufVT Xor VT_BYREF

        If BufVT <> VariantType(pArray) Then throw Exps.InvalidArgumentException
        Dim SAPtr As Long
        Call memcpy(SAPtr, ByVal (VarPtr(Buf) + VARIANT_DATA_OFFSET), VLEN_PTR)
        Call memcpy(SAPtr, ByVal SAPtr, VLEN_PTR)
        If SAPtr = 0 Then
            Call memcpy(ByVal VariantDataPtr(Buf), Arrays.GetSafeArrayPointer(pArray), VLEN_PTR)
        Else
            'Array is initialized.
            Dim aSize As Long
            If API_SafeArrayGetDim(SAPtr) <> 1 Then throw Exps.MultiDimentionException
            aSize = (API_SafeArrayGetUBound(SAPtr, 1) - API_SafeArrayGetLBound(SAPtr, 1) + 1)

            If aSize > pCounts Then aSize = pCounts

            aSize = aSize - 1
            Dim i As Long
            For i = 0 To aSize
                Call API_VariantCopy(Buf(i), pArray(i))
            Next
            Exit Sub 'Exit Is Very Important
        End If
    End If

    Call memzero(ByVal VarPtr(pArray), VLEN_VARIANT)
    pCounts = 0
    Call Clear
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
'<section IEnumerable Implementations>
'
Private Function IEnumerable_GetEnumerator() As IEnumerator: Set IEnumerable_GetEnumerator = GetEnumerator: End Function
Private Function IEnumerable_NewEnum() As stdole.IUnknown: Set IEnumerable_NewEnum = NewEnum: End Function

'Private Function IProvideEnumerableVersion_GetEnumerableVersion() As Long: IProvideEnumerableVersion_GetEnumerableVersion = EnumVersion: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
