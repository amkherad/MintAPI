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

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Public Function CreateVectorArray(ByVal ArrayType As ArrayTypes, ByVal Length As Long) As ArrayGeneric
    Dim AG As New ArrayGeneric
    Call AG.Constructor0(ArrayType, Length)
    Set CreateVectorArray = AG
End Function
Public Function Create(ByVal ArrayType As ArrayTypes, Dimensions() As Long, Optional ByVal LowerBounds As Long = 0) As ArrayGeneric
    Dim AG As New ArrayGeneric
    Call AG.Constructor1(ArrayType, Dimensions, LowerBounds)
    Set Create = AG
End Function
Public Sub Dispose(ByRef arr As Variant)
    If IsObject(arr) Then
        If TypeOf arr Is ArrayGeneric Then
            Dim AG As ArrayGeneric
            Set AG = arr
            Call AG.Dispose
        ElseIf TypeOf arr Is Vector Then
            Dim Vec As Vector
            Set Vec = arr
            Call Vec.Dispose
        ElseIf TypeOf arr Is ByteArray Then
            Dim BA As ByteArray
            Set BA = arr
            Call BA.Dispose
        Else
            throw Exps.UnknownValueException("This type of object is unknown.")
        End If
    ElseIf IsArray(arr) Then
        Erase arr
    Else
        throw Exps.OnlyArraysAcceptedException
    End If
End Sub
Public Sub Release(ByRef arr As Variant)
    Dim IntPtr As Long
    If IsObject(arr) Then
        If TypeOf arr Is ArrayGeneric Then
            Dim AG As ArrayGeneric
            Set AG = arr
            Call AG.Dispose
            Set arr = Nothing
        ElseIf TypeOf arr Is Vector Then
            Dim Vec As Vector
            Set Vec = arr
            Call Vec.Dispose
            Set arr = Nothing
        ElseIf TypeOf arr Is ByteArray Then
            Dim BA As ByteArray
            Set BA = arr
            Call BA.Dispose
            Set arr = Nothing
        Else
            throw Exps.UnknownValueException("This type of object is unknown.")
        End If
    ElseIf IsArray(arr) Then
        IntPtr = VariantAbsoluteDataPtr(arr)
        Call SetNullPtr(IntPtr)
        Call API_SafeArrayDestroyData(IntPtr)
    ElseIf IsInteger(arr) Then
        IntPtr = CLng(arr)
        Call memzero(ByVal IntPtr, Len(arr))
    Else
        throw Exps.OnlyArraysAcceptedException
    End If
End Sub

Public Function Vector(ByVal Size As Long, ByVal ArrayType As ArrayTypes) As Vector
    
End Function

Public Function CreateSafeByteArray(ByVal Length As Long, Optional ByVal LowerBound As Long = 0) As Variant
    If Length = 0 Then throw Exps.ZeroArgumentException(Mtr("Length can't be declared as zero."))
    Dim retVal As Variant
    VariantType(retVal) = (vbByte Or vbArray)
    VariantDataPtr(retVal) = API_SafeArrayCreateVector(vbByte, LowerBound, Length)
    CreateSafeByteArray = retVal
End Function
Public Function CreateSafeVector(ByVal ArrayType As ArrayTypes, ByVal Length As Long, Optional ByVal LowerBound As Long = 0) As Variant
    If Length = 0 Then throw Exps.ZeroArgumentException(Mtr("Length can't be declared as zero."))
    Dim retVal As Variant
    VariantType(retVal) = (ArrayType Or vbArray)
    VariantDataPtr(retVal) = API_SafeArrayCreateVector((ArrayType Xor VT_ARRAY), LowerBound, Length)
    CreateSafeVector = retVal
End Function
Public Function CreateSafeArray(ByVal ArrayType As ArrayTypes, Dimensions() As Long, Optional ByVal LowerBounds As Long = 0) As Variant
    Dim Dims As Long, i As Long, ln As Long, l_Bound As Long
    Dims = ArraySize(Dimensions)
    If Dims <= 0 Then throw Exps.NegativeArgumentException("Dementions must be positive non zero number.")

    ln = 1
    l_Bound = LBound(Dimensions)
    For i = l_Bound To UBound(Dimensions)
        ln = ln * Dimensions(i)
    Next
    If ln = 0 Then _
        throw Exps.InvalidArgumentValueException(Mtr("One (or more) of the dimentions declared as zero."))

    Dim DimenHndl As Long, Length As Long
    Length = Dims * SIZEOF_SAFEARRAYBOUND
    DimenHndl = API_CoTaskMemAlloc(Length)
    For i = 0 To Dims - 1
        MemLong(DimenHndl + (i * SIZEOF_SAFEARRAYBOUND)) = Dimensions(l_Bound + i)
        MemLong(DimenHndl + (i * SIZEOF_SAFEARRAYBOUND) + VLEN_LONG) = LowerBounds
    Next
    
    Dim retVal As Variant
    VariantType(retVal) = (ArrayType Or vbArray)
    VariantDataPtr(retVal) = API_SafeArrayCreateN((ArrayType Xor VT_ARRAY), Dims, ByVal DimenHndl)
    
    Call API_CoTaskMemFree(DimenHndl)
    
    ' Elements of a size that are a multiple of 4 will not create a new array
    ' when being returned. All other datatype sizes will cause a new array
    ' to be created. The array must also be assigned at the end of the code
    ' so VB can optimize returning the Variant datatype and not copy it when possible.
    CreateSafeArray = retVal 'Kelly Ethridge
End Function

Public Function ByteArray(ByVal Length As Long) As ByteArray
    Dim BA As New ByteArray
    Call BA.Constructor0(Length)
    Set ByteArray = BA
End Function
Public Function ByteArrayOF(Optional ByRef target As Variant, Optional ByVal Length As Long = -1) As ByteArray
    Dim BA As New ByteArray
    Call BA.Constructor1(target, Length)
    Set ByteArrayOF = BA
End Function
Public Function BoundedByteArrayOF(ByRef arr As Variant) As ByteArray
    Dim BA As New ByteArray
    Call BA.Constructor3(arr)
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

Public Function GetSafeArrayPointer(ByRef arr As Variant) As Long
    If IsObject(arr) Then
        If TypeOf arr Is ArrayGeneric Then
            Dim AG As ArrayGeneric
            Set AG = arr
            GetSafeArrayPointer = AG.Handle
        ElseIf TypeOf arr Is Vector Then
            Dim Vec As Vector
            Set Vec = arr
            GetSafeArrayPointer = Vec.Handle
        ElseIf TypeOf arr Is ByteArray Then
            Dim BA As ByteArray
            Set BA = arr
            GetSafeArrayPointer = BA.Handle
        Else
            throw Exps.UnknownValueException("This type of object is unknown.")
        End If
    ElseIf IsArray(arr) Then
        GetSafeArrayPointer = GetArrayPointer(arr, True)
    Else
        throw Exps.OnlyArraysAcceptedException
    End If
End Function
Private Function GetSafeArrayOf(ByRef arr As Variant) As SafeArray
    MemPtr(VarPtr(GetSafeArrayOf)) = GetSafeArrayPointer(arr)
End Function
Public Function GetDataPointerOf(ByRef arr As Variant) As Long
    If IsObject(arr) Then
        If TypeOf arr Is ArrayGeneric Then
            Dim AG As ArrayGeneric
            Set AG = arr
            GetDataPointerOf = AG.DataHandle
        ElseIf TypeOf arr Is Vector Then
            Dim Vec As Vector
            Set Vec = arr
            GetDataPointerOf = Vec.DataHandle
        ElseIf TypeOf arr Is ByteArray Then
            Dim BA As ByteArray
            Set BA = arr
            GetDataPointerOf = BA.DataHandle
        Else
            throw Exps.UnknownValueException("This type of object is unknown.")
        End If
    ElseIf IsArray(arr) Then
        GetDataPointerOf = GetSafeArrayOf(arr).pvData
    Else
        throw Exps.OnlyArraysAcceptedException
    End If
End Function

Public Function GetRank(ByRef arr As Variant) As Long
    If IsObject(arr) Then
        If TypeOf arr Is ArrayGeneric Then
            Dim AG As ArrayGeneric
            Set AG = arr
            GetRank = AG.Length
        ElseIf TypeOf arr Is Vector Then
            Dim Vec As Vector
            Set Vec = arr
            GetRank = Vec.Length
        ElseIf TypeOf arr Is ByteArray Then
            Dim BA As ByteArray
            Set BA = arr
            GetRank = BA.Length
        Else
            throw Exps.InvalidArgumentTypeException
        End If
    ElseIf IsArray(arr) Then
        GetRank = API_SafeArrayGetDim(GetArrayPointer(arr, True))
    Else
        throw Exps.OnlyArraysAcceptedException
    End If
End Function
Public Function UpperBound(ByRef arr As Variant, Optional ByVal Dimension As Long = 0) As Long
    
End Function
Public Function LowerBound(ByRef arr As Variant, Optional ByVal Dimension As Long = 0) As Long
    
End Function
Public Function Length(ByRef arr As Variant) As Long
    If IsObject(arr) Then
        If TypeOf arr Is ArrayGeneric Then
            Dim AG As ArrayGeneric
            Set AG = arr
            Length = AG.Length
        ElseIf TypeOf arr Is Vector Then
            Dim Vec As Vector
            Set Vec = arr
            Length = Vec.Length
        ElseIf TypeOf arr Is ByteArray Then
            Dim BA As ByteArray
            Set BA = arr
            Length = BA.Length
        Else
            throw Exps.InvalidArgumentTypeException
        End If
    ElseIf IsArray(arr) Then
        Dim pSafeArray  As Long
        Dim Rank        As Long
        
        pSafeArray = GetArrayPointer(arr, True)
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
        throw Exps.OnlyArraysAcceptedException
    End If
End Function
Public Function GetLength(ByRef arr As Variant, Optional ByVal Dimension As Long = 0) As Long 'by Kelly Ethridge added on 7/31/2013
    Dim pSafeArray  As Long
    Dim Rank        As Long
    
    pSafeArray = GetArrayPointer(arr, True)
    Rank = API_SafeArrayGetDim(pSafeArray)
    If Dimension > Rank Then _
        throw Exps.OutOfRangeException
    
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

Public Function IsEmpty(arr) As Boolean
    If IsObject(arr) Then
        If TypeOf arr Is ArrayGeneric Then
            Dim AG As ArrayGeneric
            Set AG = arr
            IsEmpty = AG.Length <= 0
        ElseIf TypeOf arr Is Vector Then
            Dim Vec As Vector
            Set Vec = arr
            IsEmpty = Vec.Length <= 0
        ElseIf TypeOf arr Is ByteArray Then
            Dim BA As ByteArray
            Set BA = arr
            IsEmpty = BA.IsEmpty
        Else
            throw Exps.InvalidArgumentTypeException
        End If
    ElseIf IsArray(arr) Then
        On Error GoTo zeroLength
        IsEmpty = (UBound(arr) - LBound(arr) + 1) <= 0
        Exit Function
zeroLength:
        IsEmpty = True
    Else
        throw Exps.OnlyArraysAcceptedException
    End If
End Function
Public Function IsNull(ByRef arr As Variant) As Boolean
    IsNull = (GetArrayPointer(arr) = vbNullPtr)
End Function

Private Function IsWhiteSpace(ByRef arr As Variant) As Boolean
    'space: ,tab: ,unbreakablespace: ,enter:
    
End Function
Public Function IsNullOrEmpty(ByRef arr As Variant) As Boolean
    
End Function
Public Function IsNullOrWhiteSpace(ByRef arr As Variant) As Boolean
    
End Function

Public Function ResizeArray(arr, ByVal NewSize As Long, Optional ByVal PreserveData As Boolean = False) As Variant
    If IsObject(arr) Then
        
    ElseIf IsArray(arr) Then
        
    Else
        
    End If
End Function

Public Sub Copy()
    
End Sub
Public Function Find(ByRef arr As Variant) As Long
    
End Function
Public Function Clone(ByRef arr As Variant, Optional ByVal StartAt As Long = 0, Optional ByVal Length As Long = -1) As Variant

End Function
Public Function DuplicateRef(ByRef arr As Variant) As Variant
    If Not IsArray(arr) Then throw Exps.InvalidArgumentTypeException("Only arrays accepted.")
    Dim retVal As Variant, ptrArr As Long
    VariantType(retVal) = VariantType(arr)
    If IsByRef(arr) Then
        ptrArr = API_CoTaskMemAlloc(VLEN_PTR)
        MemPtr(ptrArr) = VariantAbsoluteDataPtr(arr)
        VariantDataPtr(retVal) = VarPtr(ptrArr)
    Else
        ptrArr = VariantDataPtr(arr)
        VariantDataPtr(retVal) = ptrArr
    End If
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
Public Function Strings(ParamArray str() As Variant) As String()
    Dim i As Long, strSize As Long, retVal() As String
    On Error GoTo zeroLength
    strSize = UBound(str) - LBound(str) + 1
zeroLength:
    If strSize > 0 Then
        ReDim retVal(strSize - 1)
        For i = 0 To strSize - 1
            retVal(i) = CStr(str(i))
        Next
    End If
    Strings = retVal
End Function
Public Function Objects(ParamArray EnumerableObjects() As Variant) As Object()
    Dim i As Long, objSize As Long, retVal() As Object
    On Error GoTo zeroLength
    objSize = UBound(EnumerableObjects) - LBound(EnumerableObjects) + 1
zeroLength:
    If objSize > 0 Then
        ReDim retVal(objSize - 1)
        For i = 0 To objSize - 1
            Set retVal(i) = EnumerableObjects(i)
        Next
    End If
    Objects = retVal
End Function
Public Function Doubles(ParamArray EnumerableDoubles() As Variant) As Double()
    Dim i As Long, dblSize As Long, retVal() As Double
    On Error GoTo zeroLength
    dblSize = UBound(EnumerableDoubles) - LBound(EnumerableDoubles) + 1
zeroLength:
    If dblSize > 0 Then
        ReDim retVal(dblSize - 1)
        For i = 0 To dblSize - 1
            retVal(i) = CDbl(EnumerableDoubles(i))
        Next
    End If
    Doubles = retVal
End Function
Public Function Singles(ParamArray EnumerableSingles() As Variant) As Single()
    Dim i As Long, sngSize As Long, retVal() As Single
    On Error GoTo zeroLength
    sngSize = UBound(EnumerableSingles) - LBound(EnumerableSingles) + 1
zeroLength:
    If sngSize > 0 Then
        ReDim retVal(sngSize - 1)
        For i = 0 To sngSize - 1
            retVal(i) = CSng(EnumerableSingles(i))
        Next
    End If
    Singles = retVal
End Function
Public Function Longs(ParamArray EnumerableLongs() As Variant) As Long()
    Dim i As Long, lngSize As Long, retVal() As Long
    On Error GoTo zeroLength
    lngSize = UBound(EnumerableLongs) - LBound(EnumerableLongs) + 1
zeroLength:
    If lngSize > 0 Then
        ReDim retVal(lngSize - 1)
        For i = 0 To lngSize - 1
            retVal(i) = CLng(EnumerableLongs(i))
        Next
    End If
    Longs = retVal
End Function
Public Function Integers(ParamArray Ints() As Variant) As Integer()
    Dim i As Long, intSize As Long, retVal() As Integer
    On Error GoTo zeroLength
    intSize = UBound(Ints) - LBound(Ints) + 1
zeroLength:
    If intSize > 0 Then
        ReDim retVal(intSize - 1)
        For i = 0 To intSize - 1
            retVal(i) = CLng(Ints(i))
        Next
    End If
    Integers = retVal
End Function
Public Function Bytes(ParamArray EnumerableBytes() As Variant) As Byte()
    Dim i As Long, btSize As Long, retVal() As Byte
    On Error GoTo zeroLength
    btSize = UBound(EnumerableBytes) - LBound(EnumerableBytes) + 1
zeroLength:
    If btSize > 0 Then
        ReDim retVal(btSize - 1)
        For i = 0 To btSize - 1
            retVal(i) = CByte(EnumerableBytes(i))
        Next
    End If
    Bytes = retVal
End Function
Public Function ByteArrays(ParamArray BAs() As Variant) As ByteArray()
    Dim i As Long, baSize As Long, retVal() As ByteArray
    On Error GoTo zeroLength
    baSize = UBound(BAs) - LBound(BAs) + 1
zeroLength:
    On Error GoTo 0
    If baSize > 0 Then
        ReDim retVal(baSize - 1)
        For i = 0 To baSize - 1
            If VarType(BAs(i)) <> VBObject Then throw Exps.InvalidArgumentTypeException("Only ByteArrays Accepted.")
            If Not TypeOf BAs(i) Is ByteArray Then throw Exps.InvalidArgumentTypeException("Only ByteArrays Accepted.")
            Set retVal(i) = BAs(i)
        Next
    End If
    ByteArrays = retVal
End Function
Public Function StringArrayToVariantArray(strs() As String) As Variant()
    Dim strsCount As Long, retVars() As Variant
    strsCount = ArraySize(strs)
    If strsCount > 0 Then
        ReDim retVars(strsCount - 1)
        Dim i As Long, xCounter As Long
        For i = LBound(strs) To UBound(strs)
            retVars(xCounter) = CStr(strs(i))
            xCounter = xCounter + 1
        Next
    End If
    StringArrayToVariantArray = retVars
End Function

Public Function CLongArray(ByRef ArrayValue As Variant) As Long()
    Dim VT As VariantTypes
    VT = VariantType(ArrayValue)
    If Not CanCastToXArray(VT, vbLong) Then throw Exps.InvalidCastException
    Dim IntPtr As Long
    IntPtr = API_VarPtrArray(CLongArray)
    MemPtr(IntPtr) = VariantAbsoluteDataPtr(ArrayValue)
End Function
Public Function CIntegerArray(ByRef ArrayValue As Variant) As Integer()
    Dim VT As VariantTypes
    VT = VariantType(ArrayValue)
    If Not CanCastToXArray(VT, vbInteger) Then throw Exps.InvalidCastException
    Dim IntPtr As Long
    IntPtr = API_VarPtrArray(CIntegerArray)
    MemPtr(IntPtr) = VariantAbsoluteDataPtr(ArrayValue)
End Function
Public Function CByteArray(ByRef ArrayValue As Variant) As Byte()
    Dim VT As VariantTypes
    VT = VariantType(ArrayValue)
    If Not CanCastToXArray(VT, vbByte) Then throw Exps.InvalidCastException
    Dim IntPtr As Long
    IntPtr = API_VarPtrArray(CByteArray)
    MemPtr(IntPtr) = VariantAbsoluteDataPtr(ArrayValue)
End Function
Public Function CDecimalArray(ByRef ArrayValue As Variant) As Variant()
    Dim VT As VariantTypes
    VT = VariantType(ArrayValue)
    If Not CanCastToXArray(VT, vbDecimal) Then throw Exps.InvalidCastException
    Dim IntPtr As Long
    IntPtr = API_VarPtrArray(CDecimalArray)
    MemPtr(IntPtr) = VariantAbsoluteDataPtr(ArrayValue)
End Function
Public Function CCurrencyArray(ByRef ArrayValue As Variant) As Currency()
    Dim VT As VariantTypes
    VT = VariantType(ArrayValue)
    If Not CanCastToXArray(VT, vbCurrency) Then throw Exps.InvalidCastException
    Dim IntPtr As Long
    IntPtr = API_VarPtrArray(CCurrencyArray)
    MemPtr(IntPtr) = VariantAbsoluteDataPtr(ArrayValue)
End Function
Public Function CSingleArray(ByRef ArrayValue As Variant) As Single()
    Dim VT As VariantTypes
    VT = VariantType(ArrayValue)
    If Not CanCastToXArray(VT, vbSingle) Then throw Exps.InvalidCastException
    Dim IntPtr As Long
    IntPtr = API_VarPtrArray(CSingleArray)
    MemPtr(IntPtr) = VariantAbsoluteDataPtr(ArrayValue)
End Function
Public Function CDoubleArray(ByRef ArrayValue As Variant) As Double()
    Dim VT As VariantTypes
    VT = VariantType(ArrayValue)
    If Not CanCastToXArray(VT, vbDouble) Then throw Exps.InvalidCastException
    Dim IntPtr As Long
    IntPtr = API_VarPtrArray(CDoubleArray)
    MemPtr(IntPtr) = VariantAbsoluteDataPtr(ArrayValue)
End Function
Public Function CDateArray(ByRef ArrayValue As Variant) As Date()
    Dim VT As VariantTypes
    VT = VariantType(ArrayValue)
    If Not CanCastToXArray(VT, vbDate) Then throw Exps.InvalidCastException
    Dim IntPtr As Long
    IntPtr = API_VarPtrArray(CDateArray)
    MemPtr(IntPtr) = VariantAbsoluteDataPtr(ArrayValue)
End Function
Public Function CBooleanArray(ByRef ArrayValue As Variant) As Boolean()
    Dim VT As VariantTypes
    VT = VariantType(ArrayValue)
    If Not CanCastToXArray(VT, vbBoolean) Then throw Exps.InvalidCastException
    Dim IntPtr As Long
    IntPtr = API_VarPtrArray(CBooleanArray)
    MemPtr(IntPtr) = VariantAbsoluteDataPtr(ArrayValue)
End Function
Public Function CStringArray(ByRef ArrayValue As Variant) As String()
    Dim VT As VariantTypes
    VT = VariantType(ArrayValue)
    If Not CanCastToXArray(VT, vbString) Then throw Exps.InvalidCastException
    Dim IntPtr As Long
    IntPtr = API_VarPtrArray(CStringArray)
    MemPtr(IntPtr) = VariantAbsoluteDataPtr(ArrayValue)
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
Private Function CanCastToXArray(ByVal VT As VariantTypes, ByVal ArrayType As VbVarType) As Boolean
    If (VT And VT_BYREF) = VT_BYREF Then _
        VT = VT Xor VT_BYREF
    CanCastToXArray = (VT = (ArrayType Or VT_ARRAY))
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
