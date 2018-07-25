VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Vector"
Attribute VB_Ext_KEY = "Member1" ,"ByteArray"
'@PROJECT_LICENSE

''<class name="ArrayStatic" static tags="">
''  <summary>
''    Provides static methods for arrays.
''  </summary>
''  <using>ArrayGeneric</using>
''  <idea>Me , :) Everywhere! , VBCorLib</idea>
''  <sample>
''      Dim c() As Byte
''      c = Arrays.CreateSafeVector(atByteArray, 100)
''      For i = 0 To 99
''          c(i) = i
''      Next
''  </sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ArrayStatic" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Public Enum ArrayTypes
    atLongArray = vbLong Or vbArray
    atIntegerArray = vbInteger Or vbArray
    atByteArray = vbByte Or vbArray
    atSingleArray = vbSingle Or vbArray
    atDoubleArray = vbDouble Or vbArray
    atStringArray = vbString Or vbArray
    atObjectArray = VBObject Or vbArray
    atBooleanArray = vbBoolean Or vbArray
    atDateArray = vbDate Or vbArray
    atCurrencyArray = vbCurrency Or vbArray
    atVariantArray = vbVariant Or vbArray
    atUserDefinedTypeArray = VariantTypes.VT_Record Or vbArray
    atUDT = atUserDefinedTypeArray
End Enum
Public Enum ArrayFlags
    afNone = 0
    afAuto = FADF_AUTO
    afStatic = FADF_STATIC
    afEmbedded = FADF_EMBEDDED
    afFixedSize = FADF_FIXEDSIZE
    afRecord = FADF_RECORD
    afHaveID = FADF_HAVEIID
    afHaveVarType = FADF_HAVEVARTYPE
    afBSTR = FADF_BSTR
    afUnknown = FADF_UNKNOWN
    afDispatch = FADF_DISPATCH
    afVariant = FADF_VARIANT
    afReserved = FADF_RESERVED
End Enum
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Public Function CreateVectorArray(ByVal ArrayType As ArrayTypes, ByVal Length As Long, Optional ByVal Flags As ArrayFlags = 0) As ArrayGeneric
    Dim AG As New ArrayGeneric
    Call AG.Constructor0(ArrayType, Length)
    Set CreateVectorArray = AG
End Function
Public Function Create(ByVal ArrayType As ArrayTypes, Dimensions() As Long, Optional ByVal LowerBounds As Long = 0, Optional ByVal Flags As ArrayFlags = 0) As ArrayGeneric
Attribute Create.VB_UserMemId = 0
    Dim AG As New ArrayGeneric
    Call AG.Constructor1(ArrayType, Dimensions, LowerBounds)
    Set Create = AG
End Function
'
'Public Function Vector(ByVal Size As Long, ByVal ArrayType As ArrayTypes) As Vector
'
'End Function

Public Function CreateSafeByteArray(ByVal Length As Long, Optional ByVal LowerBound As Long = 0, Optional ByVal Flags As ArrayFlags = 0) As Variant
    If Length = 0 Then throw Exps.InvalidArgumentException(Mtr("Length can't be declared as zero."))
    If Length < 0 Then throw Exps.InvalidArgumentException
    Dim RetPtr As Variant
    RetPtr = API_SafeArrayCreateVector(vbByte, LowerBound, Length)
    If RetPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    If Flags <> 0 Then _
        Call memcpy(ByVal (RetPtr + SAFEARRAY_FFEATURES_OFFSET), CLng(Flags), VLEN_LONG)
    CreateSafeByteArray = RetPtr
    VariantType(CreateSafeByteArray) = (vbByte Or vbArray)
End Function
Public Function CreateSafeVector(ByVal ArrayType As ArrayTypes, ByVal Length As Long, Optional ByVal LowerBound As Long = 0, Optional ByVal Flags As ArrayFlags = 0) As Variant
    If Length = 0 Then throw Exps.InvalidArgumentException(Mtr("Length can't be declared as zero."))
    If Length < 0 Then throw Exps.InvalidArgumentException
    Dim RetPtr As Long
    RetPtr = API_SafeArrayCreateVector((ArrayType Xor VT_ARRAY), LowerBound, Length)
    If RetPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    If Flags <> 0 Then _
        Call memcpy(ByVal (RetPtr + SAFEARRAY_FFEATURES_OFFSET), CLng(Flags), VLEN_LONG)
    CreateSafeVector = RetPtr
    VariantType(CreateSafeVector) = (ArrayType Or vbArray)
End Function
Public Function CreateSafeVectorPtr(ByVal ArrayType As ArrayTypes, ByVal Length As Long, Optional ByVal LowerBound As Long = 0, Optional ByVal Flags As ArrayFlags = 0) As Long
    If Length = 0 Then throw Exps.InvalidArgumentException(Mtr("Length can't be declared as zero."))
    If Length < 0 Then throw Exps.InvalidArgumentException
    CreateSafeVectorPtr = API_SafeArrayCreateVector((ArrayType Xor VT_ARRAY), LowerBound, Length)
    If CreateSafeVectorPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    If Flags <> 0 Then _
        Call memcpy(ByVal (CreateSafeVectorPtr + SAFEARRAY_FFEATURES_OFFSET), CLng(Flags), VLEN_LONG)
End Function
Public Function CreateSafeArray(ByVal ArrayType As ArrayTypes, Dimensions() As Long, Optional ByVal LowerBounds As Long = 0, Optional ByVal Flags As ArrayFlags = 0) As Variant
    Dim Dims As Long, i As Long, ln As Long, l_Bound As Long
    Dims = ArraySize(Dimensions)
    If Dims <= 0 Then throw Exps.InvalidArgumentException("Dementions must be positive non zero number.")

    ln = 1
    l_Bound = LBound(Dimensions)
    For i = l_Bound To UBound(Dimensions)
        ln = ln * Dimensions(i)
    Next
    If ln = 0 Then _
        throw Exps.InvalidArgumentException(Mtr("One (or more) of the dimentions declared as zero."))

    Dim DimenHndl As Long, Length As Long, Dimen As Long
    Length = Dims * SIZEOF_SAFEARRAYBOUND
    DimenHndl = API_CoTaskMemAlloc(Length)
    For i = 0 To Dims - 1
        Dimen = Dimensions(l_Bound + i)
        If Dimen < 0 Then
            MemLong(DimenHndl + (i * SIZEOF_SAFEARRAYBOUND)) = Dimen 'Dimensions(l_Bound + i)
            MemLong(DimenHndl + (i * SIZEOF_SAFEARRAYBOUND) + VLEN_LONG) = -Dimen
        Else
            MemLong(DimenHndl + (i * SIZEOF_SAFEARRAYBOUND)) = Dimen 'Dimensions(l_Bound + i)
            MemLong(DimenHndl + (i * SIZEOF_SAFEARRAYBOUND) + VLEN_LONG) = LowerBounds
        End If
    Next
    
    Dim RetPtr As Long
    RetPtr = API_SafeArrayCreateN((ArrayType Xor VT_ARRAY), Dims, ByVal DimenHndl)
    If RetPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    
    If Flags <> 0 Then _
        Call memcpy(ByVal (RetPtr + SAFEARRAY_FFEATURES_OFFSET), CLng(Flags), VLEN_LONG)
    
    Call API_CoTaskMemFree(DimenHndl)
    
    ' Elements of a size that are a multiple of 4 will not create a new array
    ' when being returned. All other datatype sizes will cause a new array
    ' to be created. The array must also be assigned at the end of the code
    ' so VB can optimize returning the Variant datatype and not copy it when possible.
    'CreateSafeArray = RetVal 'Kelly Ethridge
    CreateSafeArray = RetPtr
    VariantType(CreateSafeArray) = (ArrayType Or VT_ARRAY)
End Function

Public Function ByteArray(ByVal Length As Long, Optional ByVal Flags As ArrayFlags = 0) As ByteArray
    Dim BA As New ByteArray
    Call BA.Constructor0(Length)
    Set ByteArray = BA
End Function
Public Function ByteArrayOF(Optional ByRef Target As Variant, Optional ByVal ConvertToBinary As Boolean = True, Optional ByVal Length As Long = -1, Optional ByVal Flags As ArrayFlags = 0) As ByteArray
    Dim BA As New ByteArray
    Call BA.Constructor1(Target, ConvertToBinary, Length)
    Set ByteArrayOF = BA
End Function
Public Function BoundedByteArrayOF(ByRef Arr As Variant, Optional ByVal Flags As ArrayFlags = 0) As ByteArray
    Dim BA As New ByteArray
    Call BA.Constructor3(Arr)
    Set BoundedByteArrayOF = BA
End Function
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

Public Function GetSafeArrayPointer(ByRef NativeArray As Variant, Optional ByVal ThrowOnNull As Boolean = False) As Long 'by Kelly Ethridge added on 7/23/2013
    Const BYREF_ARRAY As Long = VT_BYREF Or VT_ARRAY
    Dim VT As Long
    VT = VariantType(NativeArray)
    Select Case VT And BYREF_ARRAY
        ' we have to double deref the original array pointer because
        ' the variant held a pointer to the original array variable.
        Case BYREF_ARRAY
            GetSafeArrayPointer = MemLong(MemLong(VarPtr(NativeArray) + VARIANT_DATA_OFFSET))
        ' we won't need to deref again if the original array was dimensioned
        ' as a variant ie:
        '    Dim Arr As Variant
        '    ReDim Arr(1 To 10) As Long
        '
        ' The passed in variant will be the array variable, not a ByRef
        ' pointer to the array variable.
        Case vbArray
            GetSafeArrayPointer = MemLong(VarPtr(NativeArray) + VARIANT_DATA_OFFSET)
        Case vbEmpty, vbNull
            throw Exps.ArgumentNullException("NativeArray")
        Case VBObject
            If TypeOf NativeArray Is ArrayGeneric Then
                Dim AG As ArrayGeneric
                Set AG = NativeArray
                GetSafeArrayPointer = AG.Handle
            'ElseIf TypeOf arr Is Vector Then
                'Dim Vec As Vector
                'Set Vec = arr
                'GetSafeArrayPointer = Vec.Handle
            Else
                throw Exps.InvalidArgumentException("This type of object is unknown.")
            End If
        Case Else
            throw Exps.ArrayExpectedException("NativeArray")
    End Select
    
    ' HACK HACK HACK 'This function from VBCorLib
    '
    ' When an uninitialized array of objects or UDTs is passed into a
    ' function as a ByRef Variant, the array is initialized with just the
    ' SafeArrayDescriptor, at which point, it is a valid array and can
    ' be used by UBound and LBound after the call. So, now we're just
    ' going to assume that any object or UDT array that has just the descriptor
    ' allocated was Null to begin with. That means whenever an Object or UDT
    ' array is passed to any cArray method, it will technically never
    ' be uninitialized, just zero-length.
    Select Case VT And VT_TypeMask
        Case VBObject, vbUserDefinedType
            If MemLong(GetSafeArrayPointer + SAFEARRAY_PVDATA_OFFSET) = vbNullPtr Then _
                GetSafeArrayPointer = vbNullPtr
    End Select
    
    If ThrowOnNull Then
        If GetSafeArrayPointer = vbNullPtr Then
            throw Exps.ArgumentNullException("NativeArray")
        End If
    End If
End Function
Public Function GetDataPointerOf(ByRef Arr As Variant) As Long
    Dim SAPtr As Long
    SAPtr = GetSafeArrayPointer(Arr)
    Call memcpy(GetDataPointerOf, ByVal (SAPtr + SAFEARRAY_PVDATA_OFFSET), VLEN_PTR)
End Function

'from VBCorLib.VerifyArrayRange()
'renamed to EnsureArrayRange because of library name compatibility.
'### ORIGINALLY NAMED: VerifyArrayRange ###
'This function and comments wrote by Kelly Ethridge - 2005
Public Function EnsureArrayRange1DSAPtr(ByVal pSafeArray As Long, ByVal Index As Long, ByRef Count As Long, ByRef outException As Exception) As Boolean
    ' This function is optimized by not refactoring
    ' common sections with other helper rountine in
    ' order to cut down on total function calls.
    Dim aLBound As Long, aUBound As Long
    
    ' Check if the array is a null array.
    If pSafeArray = vbNullPtr Then
        Set outException = Exps.ArgumentNullException
        EnsureArrayRange1DSAPtr = False
        Exit Function
    End If
    
    ' Ensure we only have a 1-Dimension array.
    If API_SafeArrayGetDim(pSafeArray) <> 1 Then
        Set outException = Exps.MultiDimentionException
        EnsureArrayRange1DSAPtr = False
        Exit Function
    End If
    
    aLBound = API_SafeArrayGetLBound(pSafeArray, 1)
    aUBound = API_SafeArrayGetUBound(pSafeArray, 1)
    
    ' Can't have an index before the beginning of the array.
    If Index < aLBound Then
        Set outException = Exps.IndexOutOfRangeException
        EnsureArrayRange1DSAPtr = False
        Exit Function
    End If
    
    ' Can't have a negative count.
    If Count < -1 Then
        Set outException = Exps.IndexOutOfRangeException
        EnsureArrayRange1DSAPtr = False
        Exit Function
    ElseIf Count = -1 Then
        Count = (aUBound - aLBound + 1)
    End If
    
    ' Can't have the range extend past the end of the array.
    If Index + Count - 1 > aUBound Then
        Set outException = Exps.IndexOutOfRangeException
        EnsureArrayRange1DSAPtr = False
        Exit Function
    End If
    
    Set outException = Exps.NoError
    EnsureArrayRange1DSAPtr = True
End Function

Public Function Size(ByRef Arr As Variant) As Long
    If Not IsArray(Arr) Then throw Exps.ArrayExpectedException("Arr")
    Dim ArrPtr As Long
    ArrPtr = Arrays.GetSafeArrayPointer(Arr)
    If API_SafeArrayGetDim(ArrPtr) > 1 Then _
        throw Exps.MultiDimentionException
    On Error GoTo ZeroLength
    Size = (API_SafeArrayGetUBound(ArrPtr, 1) - API_SafeArrayGetLBound(ArrPtr, 1) + 1)
ZeroLength:
End Function





Public Sub Release(ByRef Arr As Variant, Optional ByVal KeepData As Boolean = False)
    If Not IsArray(Arr) Then throw Exps.ArrayExpectedException
    Dim IntPtr As Long
    IntPtr = GetSafeArrayPointer(Arr, True)
    If Not KeepData Then Call API_SafeArrayDestroyData(IntPtr)
    Call API_CoTaskMemFree(IntPtr)
    If Info.IsByRef(Arr) Then
        Dim VPtr As Long
        Call memcpy(VPtr, ByVal (VarPtr(Arr) + VARIANT_DATA_OFFSET), VLEN_PTR)
        Call memzero(ByVal VPtr, VLEN_PTR)
    End If
    
    Call memzero(Arr, VLEN_VARIANT)
End Sub


Public Function GetRank(ByRef Arr As Variant) As Long
    GetRank = API_SafeArrayGetDim(GetSafeArrayPointer(Arr))
End Function
Public Function UpperBound(ByRef Arr As Variant, Optional ByVal Dimension As Long = 1) As Long
    UpperBound = API_SafeArrayGetUBound(GetSafeArrayPointer(Arr), Dimension)
End Function
Public Function LowerBound(ByRef Arr As Variant, Optional ByVal Dimension As Long = 1) As Long
    LowerBound = API_SafeArrayGetLBound(GetSafeArrayPointer(Arr), Dimension)
End Function

''<summary>Determines the sum of total dimension's length.</summary>
''<retval>The length of the array.</retval>
Public Function Length(ByRef Arr As Variant) As Long
    If IsObject(Arr) Then
        If TypeOf Arr Is ArrayGeneric Then
            Dim AG As ArrayGeneric
            Set AG = Arr
            Length = AG.Length
        'ElseIf TypeOf arr Is Vector Then
            'Dim Vec As Vector
            'Set Vec = arr
            'Length = Vec.Length
        ElseIf TypeOf Arr Is ByteArray Then
            Dim BA As ByteArray
            Set BA = Arr
            Length = BA.Length
        Else
            throw Exps.InvalidArgumentException("Arr")
        End If
    ElseIf IsArray(Arr) Then
        Dim pSafeArray  As Long
        Dim Rank        As Long
        
        pSafeArray = GetSafeArrayPointer(Arr, True)
        Rank = API_SafeArrayGetDim(pSafeArray)
        
        pSafeArray = pSafeArray + SIZEOF_SAFEARRAY
        ' optimize for 99% of calls
        If Rank = 1 Then
            Length = MemLong(pSafeArray)
            Exit Function
        End If
        
        ' calculate the product of all the dimension sizes.
        Length = 1
        Dim i As Long
        For i = pSafeArray To pSafeArray + (SIZEOF_SAFEARRAYBOUND * (Rank - 1)) Step SIZEOF_SAFEARRAYBOUND
            Length = Length * MemLong(i)
        Next i
    Else
        throw Exps.ArrayExpectedException
    End If
End Function

''<summary>Determines the length of the given (one based) dimension.</summary>
''<retval>The length of the array.</retval>
''<remarks>If the dimension set to zero function returns the sum of total dimension's length.</remarks>
Public Function GetLength(ByRef Arr As Variant, Optional ByVal Dimension As Long = 0) As Long 'by Kelly Ethridge added on 7/31/2013
    Dim pSafeArray  As Long
    Dim Rank        As Long
    
    pSafeArray = GetSafeArrayPointer(Arr, True)
    Rank = API_SafeArrayGetDim(pSafeArray)
    If Dimension > Rank Then _
        throw Exps.IndexOutOfRangeException
    
    pSafeArray = pSafeArray + SIZEOF_SAFEARRAY
    If Dimension < 1 Then
        ' optimize for 99% of calls
        If Rank = 1 Then
            GetLength = MemLong(pSafeArray)
            Exit Function
        End If
        
        ' calculate the product of all the dimension sizes.
        GetLength = 1
        Dim i As Long
        For i = pSafeArray To pSafeArray + (SIZEOF_SAFEARRAYBOUND * (Rank - 1)) Step SIZEOF_SAFEARRAYBOUND
            GetLength = GetLength * MemLong(i)
        Next i
    Else
        ' get the size of a specific dimension.
        pSafeArray = pSafeArray + (Rank - Dimension) * SIZEOF_SAFEARRAYBOUND
        GetLength = MemLong(pSafeArray)
    End If
End Function

Public Function ToOneDimentional(ByRef MultiDimArr As Variant) As Variant
    
End Function

Public Function IsEmpty(Arr) As Boolean
    If IsObject(Arr) Then
        If TypeOf Arr Is ArrayGeneric Then
            Dim AG As ArrayGeneric
            Set AG = Arr
            IsEmpty = AG.Length <= 0
        'ElseIf TypeOf arr Is Vector Then
            'Dim Vec As Vector
            'Set Vec = arr
            'IsEmpty = Vec.Length <= 0
        ElseIf TypeOf Arr Is ByteArray Then
            Dim BA As ByteArray
            Set BA = Arr
            IsEmpty = BA.IsEmpty
        Else
            throw Exps.InvalidArgumentException
        End If
    ElseIf IsArray(Arr) Then
        On Error GoTo ZeroLength
        IsEmpty = (UBound(Arr) - LBound(Arr) + 1) <= 0
        Exit Function
ZeroLength:
        IsEmpty = True
    Else
        throw Exps.ArrayExpectedException
    End If
End Function
''<summary>Indicates that the array is null or not.</summary>
''<retval>A boolean that determines the array is null or not.</retval>
Public Function IsNull(ByRef Arr As Variant) As Boolean
    IsNull = (GetSafeArrayPointer(Arr) = vbNullPtr)
End Function

'Public Sub CopyListToArray(ByVal List As List, ByRef arr As Variant, Optional ByVal Index As Long = 0, Optional ByVal Count As Long = -1)
'    Dim Vec As New Vector
'    Call Vec.Constructor1(List, Index, Count)
'End Sub
'Public Sub CopyVectorToArray(ByVal Vector As Vector, ByRef arr As Variant)
'    Call Vector.Clone.SwapBuffer(arr)
'End Sub
'Public Sub MoveVectorToArray(ByRef Vector As Vector, ByRef arr As Variant)
'    Call Vector.SwapBuffer(arr)
'    Set Vector = Nothing
'End Sub

Private Function IsWhiteSpace(ByRef Arr As Variant) As Boolean
    'space: ,tab: ,unbreakablespace: ,enter:
    
End Function
Public Function IsNullOrEmpty(ByRef Arr As Variant) As Boolean
    
End Function
Public Function IsNullOrWhiteSpace(ByRef Arr As Variant) As Boolean
    
End Function

Public Function ResizeArray(ByRef Arr As Variant, ByVal NewSize As Long, Optional ByVal PreserveData As Boolean = False) As Variant
    If IsObject(Arr) Then
        
    ElseIf IsArray(Arr) Then
        
    Else
        
    End If
End Function

Public Sub Copy()
    
End Sub
Public Function Find(ByRef Arr As Variant) As Long
    
End Function
Public Function Clone(ByRef Arr As Variant, Optional ByVal StartAt As Long = 0, Optional ByVal Length As Long = -1) As Variant
    
End Function

Public Function DupRef(ByRef Arr As Variant) As Variant
    If Not IsArray(Arr) Then throw Exps.ArrayExpectedException
    Dim RetVal As Variant, ptrArr As Long
    VariantType(RetVal) = VariantType(Arr)
    If Info.IsByRef(Arr) Then
        ptrArr = API_CoTaskMemAlloc(VLEN_PTR)
        MemPtr(ptrArr) = VariantAbsoluteDataPtr(Arr)
        VariantDataPtr(RetVal) = VarPtr(ptrArr)
    Else
        ptrArr = VariantDataPtr(Arr)
        VariantDataPtr(RetVal) = ptrArr
    End If
End Function

Public Function Compare(ByRef Arr1 As Variant, ByRef Arr2 As Variant) As CompareResults
    If IsNull(Arr1) Then throw Exps.ArgumentNullException("Arr1")
    If IsNull(Arr2) Then throw Exps.ArgumentNullException("Arr2")
    Dim A1Len As Long, A2Len As Long
    A1Len = Length(Arr1)
    A2Len = Length(Arr2)
    If A1Len > A2Len Then
        Compare = crGreater
    ElseIf A1Len < A2Len Then
        Compare = crLess
    Else
        Dim i As Long, A1LB As Long, A2LB As Long
        A1LB = LowerBound(Arr1)
        A2LB = LowerBound(Arr2)
        For i = 0 To A1Len
            'Compare = Cmp(
        Next
    End If
End Function

Public Function ToArray() As Variant
    
End Function
Public Function ToVector() As Vector
    
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
'<section Array Creators>
'
Public Function Variants(ParamArray Str() As Variant) As Variant()
    
End Function

Public Function Strings(ParamArray Str() As Variant) As String()
    
End Function
Public Function Objects(ParamArray EnumerableObjects() As Variant) As Object()
    
End Function
Public Function Doubles(ParamArray EnumerableDoubles() As Variant) As Double()
    
End Function
Public Function Singles(ParamArray EnumerableSingles() As Variant) As Single()
    
End Function
Public Function Longs(ParamArray EnumerableLongs() As Variant) As Long()
    
End Function
Public Function Integers(ParamArray EnumerableInts() As Variant) As Integer()
'    Dim Count As Long, RetVal() As Integer
'    On Error GoTo CatchErr
'    Count = UBound(EnumerableInts) - LBound(EnumerableInts) + 1
'CatchErr:
'    If Count = 0 Then Exit Function
'    Dim VPtr As Long
'    VPtr = CreateSafeVectorPtr(atIntegerArray, Count)
'    Call memcpy(ByVal API_VarPtrArray(RetVal), VPtr, VLEN_PTR)
'    Dim i As Long
'    For i = 0 To Count - 1
'        RetVal(i) = CInt(EnumerableInts(i))
'    Next
'    Call memcpyarr(Integers, RetVal, VLEN_PTR)
'    Call memzero(ByVal API_VarPtrArray(RetVal), VLEN_PTR)
    ReDim Integers(10) As Integer
    '= EnumerableInts
End Function
Public Function Bytes(ParamArray EnumerableBytes() As Variant) As Byte()
    
End Function
Public Function CloneStringArrayToVariantArray(EnumerableStrs() As String) As Variant()
    Dim StrsCount As Long, RetVars() As Variant
    StrsCount = ArraySize(EnumerableStrs)
    If StrsCount > 0 Then
        'ReDim retVars(StrsCount - 1)
        RetVars = Arrays.CreateSafeVector(atVariantArray, StrsCount)
        Dim i As Long, xCounter As Long
        For i = LBound(EnumerableStrs) To UBound(EnumerableStrs)
            RetVars(xCounter) = CStr(EnumerableStrs(i))
            xCounter = xCounter + 1
        Next
    End If
    CloneStringArrayToVariantArray = RetVars
End Function


Public Function CAnything(ByRef ArrayValue As Variant, ByVal ElemSize As Long, Optional ByVal ToOneDimensional As Boolean = False, Optional ByVal Flags As Long = 0) As Long
    Dim DataPtr As Long, TotalLen As Long, SA As SafeArray1d
    Dim VT As VariantTypes, IsByRef As Boolean
    VT = VariantType(ArrayValue)
    IsByRef = ((VT And VT_BYREF) = VT_BYREF)
    If IsByRef Then VT = VT Xor VT_BYREF
    
    If (VT And VT_ARRAY) = VT_ARRAY Then
        DataPtr = GetSafeArrayPointer(ArrayValue)
        If API_SafeArrayGetDim(DataPtr) <> 1 Then _
            If Not ToOneDimensional Then throw Exps.MultiDimentionException
        Call memcpy(DataPtr, ByVal (DataPtr + SAFEARRAY_PVDATA_OFFSET), VLEN_PTR)
        Call memcpy(TotalLen, ByVal (DataPtr + SAFEARRAY_CBELEMENTS_OFFSET), VLEN_LONG)
        TotalLen = TotalLen * GetLength(ArrayValue)
    Else
        Select Case VT
            Case VT_BSTR, VT_LPSTR, VT_LPWSTR, VT_BSTR_BLOB
                TotalLen = LenB(ArrayValue)
                DataPtr = StrPtr(ArrayValue)
            Case Else
                TotalLen = Info.GetArgumentLength(VT)
                Dim NDP As Long
                NDP = Memory.CoTaskMemAlloc(TotalLen)
                DataPtr = DerefVariantDataPtr(ArrayValue, IsByRef)
                Call memcpy(ByVal NDP, ByVal DataPtr, TotalLen)
                DataPtr = NDP
        End Select
    End If
    
    If ElemSize > TotalLen Then throw Exps.IndexOutOfRangeException
    
    SA.cbElements = ElemSize
    SA.cDims = 1
    SA.cElements = (TotalLen / ElemSize)
    SA.fFeatures = FADF_AUTO Or Flags
    SA.lLbound = 0
    SA.pvData = DataPtr
    
    CAnything = API_CoTaskMemAlloc(SIZEOF_GUIDSAFEARRAY1D)
    If CAnything = vbNullPtr Then throw Exps.OutOfMemoryException
    Call memcpy(ByVal CAnything, SA, SIZEOF_GUIDSAFEARRAY1D)
End Function

Public Function CLongArray(ByRef ArrayValue As Variant) As Long()
    Call memcpy(ByVal API_VarPtrArray(CLongArray), CAnything(ArrayValue, VLEN_LONG, True), VLEN_PTR)
End Function
Public Function CIntegerArray(ByRef ArrayValue As Variant) As Integer()
    Call memcpy(ByVal API_VarPtrArray(CIntegerArray), CAnything(ArrayValue, VLEN_INTEGER, True), VLEN_PTR)
End Function
Public Function CByteArray(ByRef ArrayValue As Variant) As Byte()
    Call memcpy(ByVal API_VarPtrArray(CByteArray), CAnything(ArrayValue, VLEN_BYTE, True), VLEN_PTR)
End Function
Public Function CDecimalArray(ByRef ArrayValue As Variant) As Variant()
    Call memcpy(ByVal API_VarPtrArray(CDecimalArray), CAnything(ArrayValue, VLEN_VARIANT, True), VLEN_PTR)
End Function
Public Function CCurrencyArray(ByRef ArrayValue As Variant) As Currency()
    Call memcpy(ByVal API_VarPtrArray(CCurrencyArray), CAnything(ArrayValue, VLEN_CURRENCY, True), VLEN_PTR)
End Function
Public Function CSingleArray(ByRef ArrayValue As Variant) As Single()
    Call memcpy(ByVal API_VarPtrArray(CSingleArray), CAnything(ArrayValue, VLEN_SINGLE, True), VLEN_PTR)
End Function
Public Function CDoubleArray(ByRef ArrayValue As Variant) As Double()
    Call memcpy(ByVal API_VarPtrArray(CDoubleArray), CAnything(ArrayValue, VLEN_DOUBLE, True), VLEN_PTR)
End Function
Public Function CDateArray(ByRef ArrayValue As Variant) As Date()
    Call memcpy(ByVal API_VarPtrArray(CDateArray), CAnything(ArrayValue, VLEN_DATE, True), VLEN_PTR)
End Function
Public Function CBooleanArray(ByRef ArrayValue As Variant) As Boolean()
    Call memcpy(ByVal API_VarPtrArray(CBooleanArray), CAnything(ArrayValue, VLEN_BOOLEAN, True), VLEN_PTR)
End Function
Public Function CStringArray(ByRef ArrayValue As Variant) As String()
    Call memcpy(ByVal API_VarPtrArray(CStringArray), CAnything(ArrayValue, VLEN_STRPTR, True), VLEN_PTR)
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private Helpers>
'
Public Sub MoveArray(V1, V2, Optional ByVal ChkArray As Boolean = True)
    Dim VT1 As VariantTypes, VT2 As VariantTypes
    Dim IsBR1 As Boolean, IsBR2 As Boolean
    VT1 = VariantType(V1): VT2 = VariantType(V2)
    IsBR1 = ((VT1 And VT_BYREF) = VT_BYREF)
    IsBR2 = ((VT2 And VT_BYREF) = VT_BYREF)
    
    If ChkArray Then
        If IsBR1 Then If (VT1 And VT_ARRAY) <> VT_ARRAY Then throw Exps.InvalidArgumentException
        If IsBR2 Then If (VT2 And VT_ARRAY) <> VT_ARRAY Then throw Exps.InvalidArgumentException
    End If
    
    Dim VPtr1 As Long, VPtr2 As Long
    If IsBR1 Then
        VPtr1 = VariantDataPtr(V1)
        If IsBR2 Then
            VPtr2 = VariantDataPtr(V2)
            Call memcpy(ByVal VPtr1, ByVal VPtr2, VLEN_PTR)
            Call memzero(ByVal VPtr2, VLEN_PTR)
        Else
            VPtr2 = VarPtr(V2)
            Call memcpy(ByVal VPtr1, ByVal (VPtr2 + VARIANT_DATA_OFFSET), VLEN_PTR)
        End If
    Else
        If IsBR2 Then
            VPtr2 = VariantDataPtr(V2)
            Call memcpy(ByVal (VarPtr(V1) + VARIANT_DATA_OFFSET), ByVal VPtr2, VLEN_PTR)
            Call memzero(ByVal VPtr2, VLEN_PTR)
            VT2 = VT2 Xor VT_BYREF
        Else
            Call memcpy(V1, V2, VLEN_VARIANT)
        End If
    End If
    Call memcpy(V1, VT2, VLEN_VARTYPE)
    Call memzero(V2, VLEN_VARIANT)
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
