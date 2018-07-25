Attribute VB_Name = "mint_api"
Option Explicit

Public Declare Function API_VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
'Public Declare Function API_VarPtr Lib "msvbvm60" Alias "VarPtr" (Ptr As Any) As Long
Public Declare Sub memcpyarr Lib "Kernel32" Alias "RtlMoveMemory" (Destination() As Any, Source() As Any, ByVal Length As Long)
Public Declare Sub memzeroarr Lib "Kernel32" Alias "RtlZeroMemory" (Destination() As Any, ByVal Length As Long)

' SafeArray Constants
Public Const SIZEOF_SAFEARRAY               As Long = 16
Public Const SIZEOF_SAFEARRAYBOUND          As Long = 8
Public Const SIZEOF_SAFEARRAY1D             As Long = SIZEOF_SAFEARRAY + SIZEOF_SAFEARRAYBOUND
Public Const SIZEOF_GUID                    As Long = 16
Public Const SIZEOF_GUIDSAFEARRAY1D         As Long = SIZEOF_SAFEARRAY1D + SIZEOF_GUID

' Byte offsets into the SafeArray structure.
Public Const SAFEARRAY_IID_OFFSET           As Long = -16
Public Const SAFEARRAY_IRECORDINFO_OFFSET   As Long = -4
Public Const SAFEARRAY_VARTYPE_OFFSET       As Long = -4
Public Const SAFEARRAY_DIMS_OFFSET          As Long = 0
Public Const SAFEARRAY_FFEATURES_OFFSET     As Long = 2
Public Const SAFEARRAY_CBELEMENTS_OFFSET    As Long = 4
Public Const SAFEARRAY_CLOCKS_OFFSET        As Long = 8
Public Const SAFEARRAY_PVDATA_OFFSET        As Long = 12
Public Const SAFEARRAY_CELEMENTS_OFFSET     As Long = 16
Public Const SAFEARRAY_LBOUND_OFFSET        As Long = 20

' Variant descriptions and offsets into the layout.
Public Const VARIANT_TYPE_OFFSET            As Long = 0
Public Const VARIANT_DATA_OFFSET            As Long = 8
Public Const VARIANT_DATAEXT_OFFSET         As Long = 8
Public Const VARIANT_DECIMAL_OFFSET         As Long = &HC


Public Const OBJECTBUFFER_STREAMINGVALUE As String = "streamingvalue"
Public Const OBJECTBUFFER_HANDLE As String = "handle"
Public Const OBJECTBUFFER_DISPLAY_CONTEXT As String = "hdc"
Public Const OBJECTBUFFER_EVENTNAMES_LIST As String = "eventnameslist"
Public Const OBJECTBUFFER_TEXEDOBJECT As String = "itexedobject"
Public Const OBJECTBUFFER_MINTLOCALTYPE As String = "mintlocaltype"
Public Const OBJECTBUFFER_INHERIT As String = "inherit"
Public Const OBJECTBUFFER_RENDERTARGET As String = "rendertarget"
Public Const OBJECTBUFFER_SIGNAL As String = "signal"
Public Const OBJECTBUFFER_SLOT As String = "slot"
Public Const OBJECTBUFFER_NULL As String = "null"
Public Const OBJECTBUFFER_CONTROLLERINSTANCE As String = "controllerinstance"

Public Const OBJECTBUFFER_VALUES_STREAMING_CLEAR As String = "clear"
Public Const OBJECTBUFFER_VALUES_STREAMING_NEWLINE As String = "newline"
Public Const OBJECTBUFFER_VALUES_STREAMING_FLUSH As String = "flush"
Public Const OBJECTBUFFER_VALUES_STREAMING_SEEK As String = "seek"

Public Declare Function DllGetClassObject Lib "mintapi0.dll" (ByVal rclsid As Any, ByVal riid As Any, out_ppv As Any) As Long
Public Declare Function DllCanUnloadNow Lib "mintapi0.dll" () As Long

'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Pointer And Variant Methods>
'
Public Function Using(ByVal Object As IDisposable) As UsingBlock
    Set Using = New UsingBlock
    Call Using.Constructor0(Object)
End Function
Public Function UsingH(ByVal Object As IHandle) As UsingHandleBlock
    Set UsingH = New UsingHandleBlock
    Call UsingH.Constructor0(Object)
End Function


'*:object
Public Property Get ObjectPtr(ByRef Object As Object) As Long
    ObjectPtr = ObjPtr(Object)
End Property
Public Property Let ObjectPtr(ByRef Object As Object, ByVal Value As Long)
    Call memcpy(ByVal VarPtr(Object), Value, VLEN_PTR)
End Property

Public Property Get StringPtr(Str As String) As Long
    StringPtr = StrPtr(Str)
End Property
Public Property Let StringPtr(Str As String, ByVal Value As Long)
    Call memcpy(ByVal VarPtr(Str), Value, VLEN_PTR)
End Property
Public Property Get VTablePtr(pObj As Object) As Long
    VTablePtr = ObjPtr(pObj)
End Property
Public Property Let VTablePtr(pObj As Object, ByVal Value As Long)
    Call memcpy(ByVal VarPtr(pObj), ByVal VarPtr(Value), VLEN_PTR)
End Property

Public Property Get CallerObject() As Object
    Dim EBP As Long
    EBP = mHelper.ReadCallerEBP
    If Not Memory.CanAccessTo(EBP, VLEN_PTR) Then throw Exps.AccessDeniedException
    Call memcpy(EBP, ByVal EBP, VLEN_PTR)
    If Not Memory.CanAccessTo(EBP + 8, VLEN_PTR) Then throw Exps.AccessDeniedException
    Call memcpy(CallerObject, ByVal (EBP + 8), VLEN_PTR)
    Call IUnknown.AddRef(CallerObject)
End Property
'Public Function IUnknown_AddRef(ByVal Target As Object) As Long 'IUnknown::AddRef()
'    'If ObjPtr(Target) = vbNullPtr Then Exit Function
'
'    Dim TIVBUnknown As IVBUnknown
'    Call memcpy(TIVBUnknown, Target, VLEN_OBJECT)
'    IUnknown_AddRef = TIVBUnknown.AddRef
'    Call memzero(TIVBUnknown, VLEN_OBJECT)
'End Function
'Public Function IUnknown_Release(ByVal Target As Object) As Long 'IUnknown::Release()
'    'If ObjPtr(Target) = vbNullPtr Then Exit Function
'
'    Dim TIVBUnknown As IVBUnknown
'    Call memcpy(TIVBUnknown, Target, VLEN_OBJECT)
'    IUnknown_Release = TIVBUnknown.Release
'    Call memzero(TIVBUnknown, VLEN_OBJECT)
'End Function
'Public Function IUnknown_QueryInterface(ByVal Target As Object, ByVal Guid As Guid, ByRef RetPtr As Long, ByRef RetObj As Object, Optional ByVal AutoAddRef As Boolean = False, Optional ByVal ThrowOnError As Boolean = True) As Long 'IUnknown::QueryInterface()
'    'If ObjPtr(Target) = vbNullPtr Then Exit Function
'    If Guid Is Nothing Then throw Exps.ArgumentNullException
'    Dim TIVBUnknown As IVBUnknown, RetRef As Long
'    Call memcpy(TIVBUnknown, Target, VLEN_OBJECT)
'    IUnknown_QueryInterface = TIVBUnknown.QueryInterface(Guid.ToStdGuid, RetRef)
'    If ThrowOnError Then _
'        If IUnknown_QueryInterface <> S_OK Then throw Exps.InvalidCastException("Unable To Cast To Specified Interface.")
'    If AutoAddRef Then Call TIVBUnknown.AddRef
'    Call memzero(TIVBUnknown, VLEN_OBJECT)
'    'ObjectPtr(RetObj) = RetRef
'    RetPtr = RetRef
'    If Not RetObj Is Nothing Then _
'        Call memcpy(ByVal VarPtr(RetObj), RetRef, VLEN_PTR)
'End Function

Public Sub MoveVariant(ByRef Dst As Variant, ByRef Src As Variant)
    Call API_VariantClear(Dst)
    
    If IsObject(Src) Then
        If Info.IsByRef(Src) Then
            Call IUnknown.AddRef(Src)
        End If
    End If
    
    Call memcpy(ByVal VarPtr(Dst), ByVal VarPtr(Src), VLEN_VARIANT)
    Call memzero(ByVal VarPtr(Src), VLEN_VARIANT)
End Sub
Public Sub MoveVariantX(ByRef Dst As Variant, ByRef Src As Variant)
    Call memcpy(ByVal VarPtr(Dst), ByVal VarPtr(Src), VLEN_VARIANT)
    Call memzero(ByVal VarPtr(Src), VLEN_VARIANT)
End Sub
Public Sub Evaluate(ByRef Dst As Variant, ByRef Src As Variant)
    If IsObject(Src) Then
        Set Dst = Src
    Else
            Dst = Src
    End If
End Sub
Public Sub EvaluateX(ByRef Dst As Variant, ByRef Src As Variant)
    If IsObject(Src) Then
        Set Dst = Src
    Else
            Dst = Src
    End If
End Sub

Public Property Let VariantType(ByRef Var As Variant, ByVal VariantType As VariantTypes)
    Dim VT As Integer
    VT = VariantType
    Call memcpy(ByVal VarPtr(Var), VT, VLEN_INTEGER)
End Property
Public Property Get VariantType(ByRef Var As Variant) As VariantTypes
    Dim VT As Integer
    Call memcpy(VT, ByVal VarPtr(Var), VLEN_INTEGER)
    VariantType = VT
End Property

Public Function DerefVariantDataPtr(ByRef Var As Variant, ByVal DeRef As Long) As Long
    Dim VPtr As Long
    VPtr = VarPtr(Var)
    'Call memcpy(VT, ByVal (VPtr + VARIANT_TYPE_OFFSET), VLEN_VARTYPE)
    Call memcpy(DerefVariantDataPtr, ByVal (VPtr + VARIANT_DATA_OFFSET), VLEN_PTR)
    If DeRef Then _
        Call memcpy(DerefVariantDataPtr, ByVal DerefVariantDataPtr, VLEN_PTR)
End Function

Public Property Let VariantDataPtr(ByRef Var As Variant, ByVal DataPtr As Long)
    Call memcpy(ByVal (VarPtr(Var) + VARIANT_DATA_OFFSET), DataPtr, VLEN_PTR)
End Property
Public Property Get VariantDataPtr(ByRef Var As Variant) As Long
    Dim VT As Long, VPtr As Long
    VPtr = VarPtr(Var)
    Call memcpy(VT, ByVal (VPtr + VARIANT_TYPE_OFFSET), VLEN_VARTYPE)
    VariantDataPtr = VPtr + VARIANT_DATA_OFFSET
    If (VT And VT_BYREF) = VT_BYREF Then _
        Call memcpy(VariantDataPtr, ByVal VariantDataPtr, VLEN_PTR)
End Property
'Public Property Let VariantAbsoluteDataPtr(Var, ByVal DataPtr As Long)
'    Call memcpy(ByVal (VarPtr(Var) + VARIANT_DATA_OFFSET), ByVal VarPtr(DataPtr), VLEN_PTR)
'End Property
Public Property Get VariantAbsoluteDataPtr(ByRef Var As Variant, Optional ByVal GetAbsoluteDataPtr As Boolean = True) As Long
    Dim VT As Long, VPtr As Long
    VPtr = VarPtr(Var)
    Call memcpy(VT, ByVal (VPtr + VARIANT_TYPE_OFFSET), VLEN_VARTYPE)
    VariantAbsoluteDataPtr = VPtr + VARIANT_DATA_OFFSET
    If (VT And VT_BYREF) = VT_BYREF Then _
        Call memcpy(VariantAbsoluteDataPtr, ByVal VariantAbsoluteDataPtr, VLEN_PTR)
    If GetAbsoluteDataPtr Then _
        If (VT And VT_ARRAY) = VT_ARRAY Then _
            Call memcpy(VariantAbsoluteDataPtr, ByVal VariantAbsoluteDataPtr, VLEN_PTR) 'Address of begin of SafeArray structure itself (not array pointer).
            
        Select Case VT
            Case VT_BSTR, VT_LPSTR, VT_LPWSTR, VT_UDT
                Call memcpy(VariantAbsoluteDataPtr, ByVal VariantAbsoluteDataPtr, VLEN_PTR) 'Address of begin of string itself (not string pointer).
            'Case VT_UDT
            '    Call memcpy(VariantAbsoluteDataPtr, ByVal VariantAbsoluteDataPtr, VLEN_PTR) 'Address of begin of structure itself (not structure pointer).
        End Select
End Property

Public Property Get ArrayDataPtr(ByVal Address As Long) As Long
    Call memcpy(ArrayDataPtr, ByVal (Address + SAFEARRAY_PVDATA_OFFSET), VLEN_PTR)
End Property
Public Property Let ArrayDataPtr(ByVal Address As Long, ByVal DataPtr As Long)
    Call memcpy(ByVal (Address + SAFEARRAY_PVDATA_OFFSET), ByVal VarPtr(DataPtr), VLEN_PTR)
End Property
'Public Property Get ArrayLock(ByVal Address As Long) As Long
'    Call memcpy(ByVal VarPtr(ArrayLock), ByVal (Address + SAFEARRAY_CLOCKS_OFFSET), VLEN_PTR)
'End Property
'Public Property Let ArrayLock(ByVal Address As Long, ByVal DataPtr As Long)
'    Call memcpy(ByVal (Address + SAFEARRAY_CLOCKS_OFFSET), ByVal VarPtr(DataPtr), VLEN_PTR)
'End Property

'Public Sub memcpy(ByVal pDest As Long, ByVal pSrc As Long, ByVal lngLength As Long)
'
'End Sub
'Public Sub memzero(ByVal pDest As Long, ByVal lngLength As Long)
'    Dim Dummy1 As Long, Dummy2 As Long, Dummy3 As Long
'    Dim Dummy4 As Long, Dummy5 As Long, Dummy6 As Long
'End Sub '18 Bytes Dummy

'these must replace with assembly.
Public Property Get MemPtr(ByVal Address As Long) As Long
    Call memcpy(ByVal VarPtr(MemPtr), ByVal Address, VLEN_PTR)
End Property
Public Property Let MemPtr(ByVal Address As Long, ByVal Value As Long)
    Call memcpy(ByVal Address, ByVal VarPtr(Value), VLEN_PTR)
End Property
Public Property Get MemLong(ByVal Address As Long) As Long
    Call memcpy(ByVal VarPtr(MemLong), ByVal Address, VLEN_PTR)
End Property
Public Property Let MemLong(ByVal Address As Long, ByVal Value As Long)
    Call memcpy(ByVal Address, ByVal VarPtr(Value), VLEN_PTR)
End Property
'------

Public Property Get MemPtrV(V As Variant) As Long
    MemPtrV = VariantDataPtr(V)
End Property
Public Property Let MemPtrV(V As Variant, ByVal Value As Long)
    Call memcpy(ByVal VariantDataPtr(V), ByVal VarPtr(Value), VLEN_PTR)
End Property
Public Property Get MemLongV(V As Variant) As Long
    MemLongV = VariantDataPtr(V)
End Property
Public Property Let MemLongV(V As Variant, ByVal Value As Long)
    Call memcpy(ByVal VariantDataPtr(V), ByVal VarPtr(Value), VLEN_PTR)
End Property

Public Function GetLengthOf(ByRef Var As Variant) As Long
    GetLengthOf = Len(Var)
End Function

'Public Function GetSafeArrayPtr(ByRef Arr As Variant) As Long
'    If IsNull(Arr) Or IsEmpty(Arr) Then throw Exps.ArgumentNullException
'    If Not IsArray(Arr) Then throw Exps.OnlyArraysAcceptedException
'
'
'End Function



'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Castings>
'

''<summary>Cast objects into IUnknown.</summary>
Public Function CUnk(ByVal Obj As Object) As IUnknown
    Set CUnk = Obj
End Function
Public Function CVBUnk(ByVal Obj As Object) As IVBUnknown
    Call memcpy(CVBUnk, Obj, VLEN_PTR)
End Function
Public Function CVBDisp(ByVal Obj As Object) As IVBDispatch
    Call memcpy(CVBDisp, Obj, VLEN_PTR)
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Generics>
'

'API_SysAllocString
'API_SysAllocStringLen
'Public Function InternString(ByRef tString As String) As String
'    Call memcpy(ByVal VarPtr(InternString), ByVal VarPtr(tString), VLEN_PTR)
'End Function




Public Function IsNullOrMissing(ByRef Target As Variant) As Boolean
    If IsMissing(Target) Then IsNullOrMissing = True: Exit Function
    If IsNull(Target) Then IsNullOrMissing = True: Exit Function
End Function
Public Function IsNullOrMissingOrEmpty(ByRef Target As Variant) As Boolean
    If IsMissing(Target) Then IsNullOrMissingOrEmpty = True: Exit Function
    If IsNull(Target) Then IsNullOrMissingOrEmpty = True: Exit Function
    If IsEmpty(Target) Then IsNullOrMissingOrEmpty = True: Exit Function
End Function
Public Function IsNullOrEmpty(ByRef Target As Variant) As Boolean
    If IsNull(Target) Then IsNullOrEmpty = True: Exit Function
    If IsEmpty(Target) Then IsNullOrEmpty = True: Exit Function
End Function

Public Function ArraySize(ByRef TargetArray As Variant) As Long
If Not IsArray(TargetArray) Then throw Exps.ArrayExpectedException("TargetArray")
    Dim ArrPtr As Long
    ArrPtr = Arrays.GetSafeArrayPointer(TargetArray)
    If API_SafeArrayGetDim(ArrPtr) > 1 Then _
        throw Exps.MultiDimentionException
    On Error GoTo zeroLength
    ArraySize = (API_SafeArrayGetUBound(ArrPtr, 1) - API_SafeArrayGetLBound(ArrPtr, 1) + 1)
zeroLength:
End Function
'Public Sub EmptyArray(ByRef TargetArray As Variant)
'If Not (VarType(TargetArray) And vbArray = vbArray) Then throw Exps.ArrayExpectedException("TargetArray")
'    Erase TargetArray
'End Sub
Public Function IsEmptyArray(ByRef TargetArray As Variant) As Boolean
If Not (VarType(TargetArray) And vbArray) = vbArray Then throw Exps.ArrayExpectedException("TargetArray")
    On Error GoTo zeroLength
    IsEmptyArray = (UBound(TargetArray) - LBound(TargetArray) + 1) <= 0
    Exit Function
zeroLength:
    IsEmptyArray = True
End Function
'Public Sub EmptyVar(ByRef Variable As Variant)
'    If IsObject(Variable) Then
'        Set Variable = Nothing
'    Else
'        If IsArray(Variable) Then
'            Erase Variable
'        Else
'            Variable = Empty
'        End If
'    End If
'End Sub
Public Function IsEmptyVariable(ByRef TargetVariable As Variant) As Boolean
    Select Case VarType(TargetVariable)
        Case vbArray
            IsEmptyVariable = (IsEmptyArray(TargetVariable))
        Case VBObject
            IsEmptyVariable = (TargetVariable Is Nothing)
        Case vbVariant
            IsEmptyVariable = (TargetVariable = Null)
        Case Else
            IsEmptyVariable = (TargetVariable = Empty)
    End Select
End Function


Public Function RecordCompare(ByRef R1 As Variant, ByRef R2 As Variant) As CompareResults
    
End Function

Public Function ArrayCompare(ByRef A1 As Variant, ByRef A2 As Variant, Optional LengthToCompare As Long = -1) As CompareResults
    Dim a1Type As VbVarType, a2Type As VbVarType
    a1Type = VarType(A1): a2Type = VarType(A2)
    If ((a1Type And vbArray) <> vbArray) Or ((a2Type And vbArray) <> vbArray) Then throw Exps.ArrayExpectedException
    If (a1Type <> a2Type) Then GoTo some_returnFalse 'throw Exps.InvalidArgumentTypeException("a1 Array Type Must Equal To a2 Array Type.")
    Dim A1Len As Long, A2Len As Long
    Dim l1Bound As Long, l2Bound As Long
    On Error GoTo a1_zeroLength
    l1Bound = LBound(A1)
    A1Len = UBound(A1) - l1Bound + 1
a1_zeroLength:
    On Error GoTo a2_zeroLength
    l2Bound = LBound(A2)
    A2Len = UBound(A2) - l2Bound + 1
a2_zeroLength:
    If A1Len = 0 And A2Len = 0 Then
        ArrayCompare = crEqual
        Exit Function
    End If
    If (l1Bound <> l2Bound) Then GoTo some_returnFalse
    Dim ln As Long
    If LengthToCompare = -1 Then
        If A1Len <> A2Len Then GoTo some_returnFalse
        LengthToCompare = A1Len
    End If
    If A1Len < LengthToCompare Then GoTo some_returnFalse
    If A2Len < LengthToCompare Then GoTo some_returnFalse
    ln = LengthToCompare - 1
    On Error GoTo 0
    Dim i As Long
    For i = LBound(A1) To ln
        If A1(i) <> A2(i) Then GoTo some_returnFalse
    Next
    ArrayCompare = crEqual
    Exit Function
some_returnFalse:
    ArrayCompare = crNotEqual
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ByteArray Castings>
'
'Public Sub CopyByteArrayToByteArray(DestinationBA() As Byte, SourceBA() As Byte)
'    Call memcpy(ByVal API_VarPtrArray(DestinationBA), ByVal API_VarPtrArray(SourceBA), VLEN_PTR)
'End Sub
'MemoryToByteArray : Copies specified memory address value to a bytearray.
'targetAddress : The address of memory to copy to byte array.
'SourceSize : Source memory address content size to copy to byte array.
'Times : Determines the times that source memory value with length equals to SourceSize copies to byte array.
'IT'S AN UNSAFE METHOD!!
Public Function MemoryToByteArray(ByVal TargetAddress As Long, ByVal SourceSize As Long) As Byte()
    Dim outLen As Long
    outLen = SourceSize
    If outLen <= 0 Then throw Exps.InvalidArgumentException("SourceSize")
    
    Dim RetVal() As Byte
    RetVal = Arrays.CreateSafeByteArray(outLen)
    
    Dim IntPtr As Long
    IntPtr = Arrays.GetDataPointerOf(RetVal)
    
    Call memcpy(ByVal IntPtr, ByVal TargetAddress, outLen)
    MemoryToByteArray = RetVal
End Function
'IT'S AN UNSAFE METHOD!!
Public Sub CopyMemoryToByteArray(ByVal TargetAddress As Long, ByVal SourceSize As Long, targetByteArray() As Byte, Optional ByVal Times As Long = 1)
    Dim outLen As Long
    outLen = (SourceSize * Times) - 1
    If outLen <= 0 Then throw Exps.InvalidArgumentException("SourceSize or Times")
    'Dim targetByteArray() As Byte
    ReDim targetByteArray(outLen)

    Dim c_byte_value As Byte

    Dim i As Long
    For i = 0 To outLen
        Call memcpy(ByVal VarPtr(c_byte_value), ByVal (TargetAddress + i), 1)
        targetByteArray(i) = c_byte_value
    Next

    'targetByteArray = retVal
End Sub
'ByteArrayToMemory : copies specified byte array to memory target's address.
'targetArray : source array to copy to memory.
'BytesToCopy : number of bytes used to copy byte array.
'IT'S AN UNSAFE METHOD!!
Public Sub ByteArrayToMemory(ByVal TargetAddress As Long, targetByteArray() As Byte, BytesToCopy As Long, Optional FillWithNull As Boolean = False)
    If BytesToCopy <= 0 Then Exit Sub
    Dim arrSize As Long
    arrSize = ArraySize(targetByteArray)
    If arrSize < BytesToCopy Then
        If Not FillWithNull Then throw Exps.InvalidOperationException("targetByteArray Length Is Less Than BytesToCopy And FillWithNull Is Not Allowed.")
    End If
    Dim i As Long, c_byte_value As Byte
    For i = 0 To BytesToCopy - 1
        If (i < arrSize) Then
            c_byte_value = targetByteArray(i)
        Else
            c_byte_value = 0
        End If
        Call memcpy(ByVal (TargetAddress + i), ByVal VarPtr(c_byte_value), 1)
    Next
End Sub
'IT'S AN UNSAFE METHOD!!

Public Function StringToIntegerArray(Str As String, Optional Length As Long = -1) As Integer()
    If Length = 0 Then Exit Function
    Dim StrLen As Long
    Dim RetVal() As Integer
    StrLen = Len(Str)
    If StrLen = 0 Then Exit Function
    If Length < 0 Then Length = StrLen
    ReDim RetVal(Length - 1)
    Dim i As Long
    For i = 1 To Length
        RetVal(i - 1) = AscW(Mid(Str, i, 1))
    Next
    StringToIntegerArray = RetVal()
End Function
Public Function IntegerArrayToString(intArr() As Integer, Optional Length As Long = -1) As String
    If Length = 0 Then Exit Function
    Dim Str As String, arrSize As Long
    arrSize = ArraySize(intArr)
    If arrSize = 0 Then Exit Function
    If Length < 0 Then Length = arrSize
    Dim i As Long, arrlBound As Long
    arrlBound = LBound(intArr)
    For i = 0 To Length - 1
        Str = Str & ChrW(intArr(i + arrlBound))
    Next
    IntegerArrayToString = Str
End Function
Public Function StringToByteArraySpeed(ByVal Str As String, Optional Length As Long = -1) As Byte()
    Dim lLength As Long
    If Length = -1 Then
        lLength = Len(Str)
    Else
        lLength = Length
    End If
    StringToByteArraySpeed = MemoryToByteArray(VarPtr(Str), lLength)
End Function
Public Function StringToByteArray(Str As String, Optional Length As Long = -1) As Byte()
    If Length = 0 Then Exit Function
    Dim StrLen As Long
    Dim RetVal() As Byte
    StrLen = Len(Str)
    If StrLen = 0 Then Exit Function
    If Length < 0 Then Length = StrLen
    ReDim RetVal(Length - 1)
    Dim i As Long
    For i = 1 To Length
        RetVal(i - 1) = Asc(Mid(Str, i, 1))
    Next
    StringToByteArray = RetVal()
End Function
Public Function ByteArrayToString(B() As Byte, Optional Length As Long = -1) As String
    If Length = 0 Then Exit Function
    Dim Str As String, arrSize As Long
    On Error GoTo zeroLength
    arrSize = UBound(B) - LBound(B) + 1
zeroLength:
    If arrSize = 0 Then Exit Function
    If Length < 0 Then Length = arrSize
    Dim i As Long, arrlBound As Long
    arrlBound = LBound(B)
    For i = 0 To Length - 1
        Str = Str & Chr(B(i + arrlBound))
    Next
    ByteArrayToString = Str
End Function
Public Function ByteArrayToSafeString(B() As Byte, Optional Length As Long = -1) As String
    If Length = 0 Then Exit Function
    Dim Str As String, arrSize As Long
    On Error GoTo zeroLength
    arrSize = UBound(B) - LBound(B) + 1
zeroLength:
    If arrSize = 0 Then Exit Function
    If Length < 0 Then Length = arrSize
    Dim i As Long, arrlBound As Long, cIndex As Long
    arrlBound = LBound(B)
    For i = 0 To Length - 1
        cIndex = i + arrlBound
        If B(cIndex) = 0 Then Exit For
        Str = Str & Chr(B(cIndex))
    Next
    ByteArrayToSafeString = Str
End Function
'Public Function LongToByteArray(lngNum As Long) As Byte()
'    LongToByteArray = MemoryToByteArray(VarPtr(lngNum), 4)
'End Function
'Public Function ByteArrayToLong(B() As Byte) As Long
'    Call ByteArrayToMemory(VarPtr(ByteArrayToLong), B, 4)
'End Function
'Public Function IntegerToByteArray(intNum As Integer) As Byte()
'    IntegerToByteArray = MemoryToByteArray(VarPtr(intNum), 2)
'End Function
'Public Function ByteArrayToInteger(B() As Byte) As Integer
'    Call ByteArrayToMemory(API_VarPtr(ByteArrayToInteger), B, 2)
'End Function
'Public Function ByteToByteArray(btByte As Byte) As Byte()
'    ByteToByteArray = MemoryToByteArray(VarPtr(btByte), 1)
'End Function
'Public Function ByteArrayToByte(B() As Byte) As Byte
'    Call ByteArrayToMemory(API_VarPtr(ByteArrayToByte), B, 1)
'End Function
'Public Function DateToByteArray(dtDate As Date) As Byte()
'    DateToByteArray = MemoryToByteArray(VarPtr(dtDate), 8)
'End Function
'Public Function ByteArrayToDate(B() As Byte) As Date
'    Call ByteArrayToMemory(API_VarPtr(ByteArrayToDate), B, 8)
'End Function
'Public Function CurrencyToByteArray(cyCurrency As Currency) As Byte()
'    CurrencyToByteArray = MemoryToByteArray(VarPtr(cyCurrency), 8)
'End Function
'Public Function ByteArrayToCurrency(B() As Byte) As Currency
'    Call ByteArrayToMemory(API_VarPtr(ByteArrayToCurrency), B, 8)
'End Function
'Public Function DoubleToByteArray(dblNum As Double) As Byte()
'    DoubleToByteArray = MemoryToByteArray(VarPtr(dblNum), 8)
'End Function
'Public Function ByteArrayToDouble(B() As Byte) As Double
'    Call ByteArrayToMemory(API_VarPtr(ByteArrayToDouble), B, 8)
'End Function
'Public Function SingleToByteArray(sngNum As Single) As Byte()
'    SingleToByteArray = MemoryToByteArray(VarPtr(sngNum), 4)
'End Function
'Public Function ByteArrayToSingle(B() As Byte) As Single
'    Call ByteArrayToMemory(API_VarPtr(ByteArrayToSingle), B, 4)
'End Function
'Public Function BooleanToByteArray(boolValue As Boolean) As Byte()
'    BooleanToByteArray = MemoryToByteArray(VarPtr(boolValue), 2)
'End Function
'Public Function ByteArrayToBoolean(B() As Byte) As Boolean
'    Call ByteArrayToMemory(API_VarPtr(ByteArrayToBoolean), B, 2)
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Var Castings>
'


'function Write Arguments to Format.
Public Function funcwArgs_(ByVal Format As String, ParamArray Args()) As String
    Dim cArgs() As Variant
    cArgs = Args
    funcwArgs_ = funcwArgs(Format, cArgs)
End Function
'function Read Arguments from Stream using Format.
Public Function funcrArgs_(ByVal Format As String, tStream As IClassStream, ParamArray Args()) As String
    Dim cArgs() As Variant
    cArgs = Args
    funcrArgs_ = funcrArgs(Format, tStream, cArgs)
End Function
'---------------------------------------------
'function Write Arguments to Format.
Public Function funcwArgs(ByVal Format As String, Args()) As String
    '%d,%s,%i,%u,%l,%f,%c,%b,%o  ,\\,\n,\r,\a,\t,\c,\xFF,\255,\0377
    '%o:object
End Function
'function Read Arguments from Stream using Format.
Public Function funcrArgs(ByVal Format As String, tStream As IClassStream, Args()) As String
    '{n},%d,%s,%i,%u,%l,%f,%c,%b,%o  ,\\,\n,\r,\a,\t,\c,\xFF,\255,\0377
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------


Public Sub mint_setstream_state(State As Boolean, inoutState As Boolean, State_LOCK As String, Optional Reserved)
    Dim strReserved As String
    strReserved = IIf(IsMissing(Reserved), "", CStr(Reserved))
    If State Then
        If inoutState Then
            If State_LOCK <> strReserved Then GoTo errThrow
        Else
            If State_LOCK <> "" Then
                If strReserved <> State_LOCK Then GoTo errThrow
            End If
            inoutState = True
            State_LOCK = strReserved
        End If
    Else
        If inoutState Then
            If State_LOCK <> strReserved Then GoTo errThrow
        Else
            If State_LOCK <> "" Then
                If strReserved <> State_LOCK Then GoTo errThrow
            End If
            inoutState = False
            State_LOCK = strReserved
        End If
    End If
    Exit Sub
errThrow:
    throw Exps.InvalidArgumentException("Invalid reserved value.")
End Sub


