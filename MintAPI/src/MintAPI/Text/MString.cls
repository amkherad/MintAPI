VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="mString" static tags="immutable">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

'reminder: minttext.dll (RegEx{})

'NOTE: DOES NOT SUPPORT CONCURRENCY.

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "mString" '#DO NOT REMOVE IT
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
Public Enum CharacterType
    ctNone = 0
    ctNull = &H100 'Includes null character.
    ctAlphabetic = 1 'Includes alphabetical characters.
    ctNumeric = 2 'Includes numeric characters.
    ctWhiteSpace = 4 'Includes white-spaces.
    ctNewLine = 8 'Includes new-line character.
    ctOthers = &H10 'Includes not alphabetic, numeric and white-space characters.
    ctSpecialCharacters = &H20 'Includes all special characters except NULL.
    ctAllHumanReadables = ctAlphabetic Or ctNumeric Or ctWhiteSpace Or ctOthers 'Includes characters those are printable.
    ctBinary = ctAllHumanReadables Or ctOthers Or ctSpecialCharacters 'Includes not printable characters. i.e. NULL
    ctAll = ctBinary Or ctAllHumanReadables
End Enum

Private WS() As Integer
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'

Public Property Get SystemDefaultCharSize() As Long: SystemDefaultCharSize = 2: End Property

Public Property Get sEmpty() As String
    sEmpty = vbNullString
End Property

''<default/>
Public Function Create(ByVal Length As Long, Optional ByRef Ptr As Long) As String
Attribute Create.VB_UserMemId = 0
    If Length = 0 Then Exit Function
    If Length < 0 Then throw Exps.InvalidArgumentException("Length")
    Dim SPtr As Long, StrLen As Long
    StrLen = ((Length + 1) * 2)
    SPtr = Memory.FastAllocate(StrLen + 4)
    If SPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    Call memcpy(ByVal SPtr, (Length * 2), VLEN_LONG)
    SPtr = SPtr + VLEN_LONG
    Call memzero(ByVal SPtr, StrLen)
    Call memcpy(Create, SPtr, VLEN_PTR)
    Ptr = SPtr
End Function 'OK
Friend Sub InitStr(ByRef Str As Variant, ByRef IntAStr() As Integer, ByVal Length As Long)
    If Length = 0 Then Exit Sub
    If Length < 0 Then throw Exps.InvalidArgumentException("Length")
    Dim SPtr As Long, StrLen As Long
    StrLen = ((Length + 1) * 2)
    SPtr = API_CoTaskMemAlloc(StrLen + VLEN_PTR) '+VLEN_PTR to store the length of string before the real string data.
    If SPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    Call memcpy(ByVal SPtr, (Length * 2), VLEN_LONG)
    SPtr = SPtr + VLEN_PTR
    Call memzero(ByVal SPtr, StrLen)
    Call memcpy(ByVal VariantDataPtr(Str), SPtr, VLEN_PTR)
    VariantType(Str) = VT_BSTR
    
    Dim S As String
    Call memcpy(S, SPtr, VLEN_PTR)
    Call memcpyarr(IntAStr, InternStrToIntArray(S), VLEN_PTR)
End Sub


Public Function GetStringArrayItem(ByRef Str As Variant, ByRef outIntArr() As Integer, ByVal cIndex As Long) As Boolean
    Dim RetPtr As Long
    Dim VT As VariantTypes ', IsByRef As Boolean
    VT = VarType(Str)
    'IsByRef = ((VT And VT_BYREF) = VT_BYREF)
    'If IsByRef Then VT = VT Xor VT_BYREF
    RetPtr = API_VarPtrArray(outIntArr)
    Select Case VT
        Case (VT_ARRAY Or VT_VARIANT), (VT_ARRAY Or VT_BSTR)
            Dim ALen As Long
            ALen = ArraySize(Str)
            If cIndex >= ALen Then throw Exps.IndexOutOfRangeException
            Call memcpyarr(outIntArr, CastToIntArray(Str(cIndex)), VLEN_PTR)
            If cIndex = ALen - 1 Then GetStringArrayItem = False Else GetStringArrayItem = True
        Case Else
            Call memcpyarr(outIntArr, CastToIntArray(Str), VLEN_PTR)
            GetStringArrayItem = False
    End Select
End Function
Public Function CastToIntArray(ByRef Str As Variant) As Integer()
    Dim RetPtr As Long
    Dim VT As VariantTypes, IsByRef As Boolean
    VT = VariantType(Str)
    IsByRef = ((VT And VT_BYREF) = VT_BYREF)
    If IsByRef Then VT = VT Xor VT_BYREF
    RetPtr = API_VarPtrArray(CastToIntArray)
    Select Case VT
        Case VT_BSTR, VT_LPSTR, VT_LPWSTR, VT_BSTR_BLOB
            Dim mStr As String
            Call memcpy(mStr, StrPtr(Str), VLEN_PTR)
            Call memcpy(ByVal RetPtr, ByVal API_VarPtrArray(InternStrToIntArray(mStr)), VLEN_PTR)
            Call memzero(mStr, VLEN_PTR)
        Case VT_ARRAY Or VT_I2
            Dim VDPtr As Long
            VDPtr = DerefVariantDataPtr(Str, IsByRef)
            If API_SafeArrayGetDim(VDPtr) <> 1 Then throw Exps.MultiDimentionException
            Call memcpy(ByVal RetPtr, VDPtr, VLEN_PTR)
        Case VT_I2, VT_UI2
            If Not IsByRef Then throw Exps.NotSupportedException
            Dim SAPtr As Long, SA As SafeArray1d, sLen As Long
            SA.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
            SA.cDims = 1
            SA.cbElements = 2
            SA.lLbound = 0
            SA.cElements = 1
            SA.pvData = VariantDataPtr(Str)
            
            SAPtr = API_CoTaskMemAlloc(SIZEOF_SAFEARRAY1D)
            If SAPtr = vbNullPtr Then throw Exps.OutOfMemoryException
            Call memcpy(ByVal SAPtr, SA, SIZEOF_SAFEARRAY1D)
            Call memcpy(ByVal RetPtr, SAPtr, VLEN_PTR)
            SA.fFeatures = 0
        Case Else
            throw Exps.InvalidArgumentException
    End Select
End Function
Public Function Intern(ByRef Str As String) As String 'UnSafe care for interned string to be removed
    Call memcpy(Intern, Str, VLEN_PTR)
End Function
Public Function InternStrToIntArray(ByRef Str As String) As Integer() 'UnSafe must remove the structure from the memory.
    Dim SAPtr As Long, SA As SafeArray1d, sLen As Long
    sLen = Len(Str)
    
    If sLen = 0 Then Exit Function
    
    SA.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
    SA.cDims = 1
    SA.cbElements = 2
    SA.lLbound = 0
    SA.cElements = sLen
    SA.pvData = StrPtr(Str)
    
    SAPtr = API_CoTaskMemAlloc(SIZEOF_SAFEARRAY1D)
    If SAPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    Call memcpy(ByVal SAPtr, SA, SIZEOF_SAFEARRAY1D)
    
    SA.fFeatures = 0
    
    Call memcpy(ByVal API_VarPtrArray(InternStrToIntArray), SAPtr, VLEN_PTR)
End Function
Public Function InternIntArrayToStr(ByRef Str() As Integer) As String 'Safe
    Dim SAPtr As Long
    SAPtr = Arrays.GetSafeArrayPointer(Str, False)
    If SAPtr = vbNullPtr Then Exit Function
    If API_SafeArrayGetDim(SAPtr) <> 1 Then throw Exps.MultiDimentionException
    Call memcpy(InternIntArrayToStr, ArrayDataPtr(SAPtr), VLEN_PTR)
End Function
Public Function ToIntArray(ByRef Str As Variant) As Integer()
    Dim SAPtr As Long, NStrPtr As Long, sLen As Long
    Dim iStr() As Integer, SA As SafeArray1d
    iStr = CastToIntArray(Str)
    sLen = Arrays.Size(iStr)
    
    If sLen = 0 Then Exit Function
    
    NStrPtr = API_SysAllocStringLen(Arrays.GetDataPointerOf(iStr), sLen)
    
    SA.cDims = 1
    SA.cbElements = 2
    SA.lLbound = 0
    SA.cElements = sLen
    SA.pvData = NStrPtr
    
    SAPtr = API_CoTaskMemAlloc(SIZEOF_SAFEARRAY1D)
    If SAPtr = vbNullPtr Then throw Exps.OutOfMemoryException
    
    Call memcpy(ByVal SAPtr, SA, SIZEOF_SAFEARRAY1D)
    Call memcpy(ByVal API_VarPtrArray(ToIntArray), SAPtr, VLEN_PTR)
End Function
Public Sub DestoryInternedIntArray(ByRef IntAStr() As Integer, Optional ByVal KeepData As Boolean = False)
    Call Arrays.Release(IntAStr, KeepData)
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

Public Sub Clear(ByRef Str As Variant)
    Dim VT As VariantTypes
    VT = VBA.VarType(Str)
    Select Case VT
        Case (VT_ARRAY Or VT_I2)
            Call Arrays.Release(Str)
        Case VT_BSTR, VT_LPSTR, VT_LPWSTR
            Call API_SysFreeString(StrPtr(Str))
        Case VT_I2
            Str = 0
        Case Else
            throw Exps.InvalidArgumentException
    End Select
    Call memzero(Str, VLEN_VARIANT)
End Sub

Public Function Length(ByRef Str As Variant) As Long
    Dim VT As VariantTypes
    VT = VBA.VarType(Str)
    Select Case VT
        Case (VT_ARRAY Or VT_I2)
            Length = ArraySize(Str)
        Case VT_BSTR, VT_LPSTR, VT_LPWSTR
            Length = VBA.Len(Str)
        Case VT_I2
            Length = 2
        Case Else
            throw Exps.InvalidArgumentException
    End Select
End Function
Public Function LengthStr(ByRef Str As String) As Long
    LengthStr = Len(Str)
End Function

Public Function ToLower(ByRef Str As Variant) As String
    Dim iStr() As Integer, retStr() As Integer, iLen As Long
    iStr = CastToIntArray(Str)
    iLen = ArraySize(iStr)
    
    If iLen = 0 Then GoTo End_Cleaner
    
    Call InitStr(ToLower, retStr, iLen)
    
    Dim i As Long, cChar As Integer
    For i = 0 To iLen - 1
        retStr(i) = ToLower_(iStr(i))
    Next
    
End_Cleaner:
'    Call Arrays.Release(iStr, True)
'    Call Arrays.Release(retStr, True)
End Function
Public Function ToUpper(ByRef Str As Variant) As String
    Dim iStr() As Integer, retStr() As Integer, iLen As Long
    iStr = CastToIntArray(Str)
    iLen = ArraySize(iStr)
    
    If iLen = 0 Then GoTo End_Cleaner
    
    Call InitStr(ToUpper, retStr, iLen)
    
    Dim i As Long, cChar As Integer
    For i = 0 To iLen - 1
        retStr(i) = ToUpper_(iStr(i))
    Next
    
End_Cleaner:
'    Call Arrays.Release(iStr, True)
'    Call Arrays.Release(retStr, True)
End Function
Public Function ToProper(ByRef Str As Variant, Optional ByVal KeepOtherCase As Boolean = False) As String
    Dim iStr() As Integer, retStr() As Integer, iLen As Long
    iStr = CastToIntArray(Str)
    iLen = ArraySize(iStr)
    
    If iLen = 0 Then GoTo End_Cleaner
    
    Call InitStr(ToProper, retStr, iLen)
    
    Dim i As Long, IsPrevNotAlphabetic As Boolean, cChar As Integer
    IsPrevNotAlphabetic = True
    If KeepOtherCase Then
        For i = 0 To iLen - 1
            cChar = iStr(i)
            If Not IsAlphabetic_(cChar) Then
                IsPrevNotAlphabetic = True
            ElseIf IsPrevNotAlphabetic Then
                cChar = ToUpper_(cChar)
                IsPrevNotAlphabetic = False
            End If
            retStr(i) = cChar
        Next
    Else
        For i = 0 To iLen - 1
            cChar = iStr(i)
            If Not IsAlphabetic_(cChar) Then
                IsPrevNotAlphabetic = True
            ElseIf IsPrevNotAlphabetic Then
                cChar = ToUpper_(cChar)
                IsPrevNotAlphabetic = False
            Else
                cChar = ToLower_(cChar)
            End If
            retStr(i) = cChar
        Next
    End If
    
End_Cleaner:
'    Call Arrays.Release(iStr, True)
'    Call Arrays.Release(retStr, True)
End Function
Public Function ToLowerInvariant(ByRef Str As Variant) As String
    
End Function
Public Function ToUpperInvariant(ByRef Str As Variant) As String
    
End Function

Public Function StartsWith(ByRef Str As Variant, ByRef What As Variant, Optional ByVal IgnoreCase As Boolean) As Boolean
    Dim iStr() As Integer, wStr() As Integer, iLen As Long, wLen As Long
    iStr = CastToIntArray(Str)
    iLen = ArraySize(iStr)
    wStr = CastToIntArray(What)
    wLen = ArraySize(wStr)
    
    If iLen = 0 Then GoTo End_Cleaner
    If wLen = 0 Then GoTo End_Cleaner
    If wLen > iLen Then GoTo End_Cleaner
    
    Dim i As Long ', cChar As Integer
    If IgnoreCase Then
        For i = 0 To wLen - 1
            If ToLower_(iStr(i)) <> ToLower_(wStr(i)) Then GoTo End_Cleaner
        Next
    Else
        For i = 0 To wLen - 1
            If iStr(i) <> wStr(i) Then GoTo End_Cleaner
        Next
    End If
    StartsWith = True
    
End_Cleaner:
'    Call Arrays.Release(iStr, True)
'    Call Arrays.Release(wStr, True)
End Function
Public Function EndsWith(ByRef Str As Variant, ByRef What As Variant, Optional ByVal IgnoreCase As Boolean) As Boolean
    Dim iStr() As Integer, wStr() As Integer, iLen As Long, wLen As Long
    iStr = CastToIntArray(Str)
    iLen = ArraySize(iStr)
    wStr = CastToIntArray(What)
    wLen = ArraySize(wStr)
    
    If iLen = 0 Then GoTo End_Cleaner
    If wLen = 0 Then GoTo End_Cleaner
    If wLen > iLen Then GoTo End_Cleaner
    
    Dim i As Long, ii As Long
    ii = iLen - 1
    If IgnoreCase Then
        For i = wLen - 1 To 0 Step -1
            If ToLower_(iStr(ii)) <> ToLower_(wStr(i)) Then GoTo End_Cleaner
            ii = ii - 1
        Next
    Else
        For i = wLen - 1 To 0 Step -1
            If iStr(ii) <> wStr(i) Then GoTo End_Cleaner
            ii = ii - 1
        Next
    End If
    EndsWith = True
    
End_Cleaner:
'    Call Arrays.Release(iStr, True)
'    Call Arrays.Release(wStr, True)
End Function

Public Function Replace(ByRef Str As Variant, ByRef FindWhat As Variant, ByRef ReplaceWith As Variant, Optional ByVal StartIndex As Long = 0, Optional ByVal Count As Long = -1) As String
    Dim iStr() As Integer, wStr() As Integer, rStr() As Integer
    Dim iLen As Long, wLen As Long, rLen As Long
    iStr = CastToIntArray(Str)
    iLen = ArraySize(iStr)
    wStr = CastToIntArray(FindWhat)
    wLen = ArraySize(wStr)
    rStr = CastToIntArray(ReplaceWith)
    rLen = ArraySize(rStr)
    
    
End Function

Public Function Remove(ByRef Str As Variant, ByVal StartIndex As Long, ByVal Count As Long) As String
    
End Function

Public Function SubSring(ByRef Str As Variant, ByVal StartIndex As Long, Optional ByVal Length As Long = -1) As String
    If StartIndex < 0 Then throw Exps.IndexOutOfRangeException
    If Length < 0 Then _
        If Length <> -1 Then throw Exps.InvalidArgumentException("Length")
    Dim Err As Exception
    Dim iStr() As Integer, sLen As Long
    iStr = CastToIntArray(Str)
    sLen = ArraySize(iStr)
    
    If sLen = 0 Then
        If StartIndex = 0 Then
            GoTo End_Cleaner
        Else
            Set Err = Exps.IndexOutOfRangeException: GoTo End_ErrorManager
        End If
    End If
    
    If Length = -1 Then
        Length = sLen - StartIndex
    Else
        If (StartIndex + Length) > sLen Then Set Err = Exps.IndexOutOfRangeException: GoTo End_ErrorManager
    End If
    
    SubSring = Create(Length)
    
    Call memcpy(ByVal StrPtr(SubSring), iStr(StartIndex), Length * 2)
    
End_ErrorManager:
    If Not Err Is Nothing Then throw Err
End_Cleaner:
'    Call Arrays.Release(iStr, True)
End Function

Public Function Insert(ByRef Str As Variant, ByVal StartIndex As Long, ByRef Value As Variant) As String
    If StartIndex < 0 Then throw Exps.IndexOutOfRangeException
    Dim Err As Exception
    Dim iStr() As Integer, sLen As Long, iVal() As Integer, VLen As Long
    iVal = CastToIntArray(Value) 'must be first to optimize if any errors happens.
    VLen = ArraySize(iVal)
    iStr = CastToIntArray(Str)
    sLen = ArraySize(iStr)
    
    If (sLen = 0) Or (VLen = 0) Then Set Err = Exps.ArgumentNullException: GoTo End_ErrorManager
    If StartIndex >= sLen Then Set Err = Exps.IndexOutOfRangeException: GoTo End_ErrorManager
    
    Dim NewLen As Long, NewPtr As Long
    NewLen = sLen + VLen
    
    Insert = Create(NewLen)
    NewPtr = StrPtr(Insert)
    
    Dim StartIndexBuffer As Long
    StartIndexBuffer = StartIndex
    StartIndex = StartIndex * 2
    VLen = VLen * 2
    sLen = sLen * 2
    'First Part
    If StartIndex > 0 Then _
        Call memcpy(ByVal NewPtr, iStr(0), StartIndex)
    
    'Middle Part
    NewPtr = NewPtr + StartIndex
    Call memcpy(ByVal NewPtr, iVal(0), VLen)
    
    'Last Part
    NewPtr = NewPtr + VLen
    Dim LastLen As Long
    LastLen = sLen - StartIndex
    If LastLen > 0 Then _
        Call memcpy(ByVal NewPtr, iStr(StartIndexBuffer), LastLen)
        
End_ErrorManager:
    If Not Err Is Nothing Then throw Err
End_Cleaner:
'    Call Arrays.Release(iStr, True)
'    Call Arrays.Release(iVal, True)
End Function

'TrimStrs: String, Integer(), Integer, Array<Variant>(String, Integer(), Integer)
Public Function Trim(ByRef Str As Variant, Optional ByRef TrimStrs As Variant) As String
    Dim Err As Exception
    Dim TSIndex As Long, TEIndex As Long, iLen As Long
    Dim TLSIndex As Long, TLEIndex As Long
    Dim tStr() As Integer, tLen As Long, cIndex As Long
    Dim TS As Variant, TSHaveNext As Boolean, Continue As Boolean
    Dim iStr() As Integer
    iStr = CastToIntArray(Str)
    iLen = ArraySize(iStr)
    
    MsgBox "iLen:" & iLen
    
    If IsMissing(TrimStrs) Then
        TS = Arrays.GetSafeArrayPointer(WS, True)
        VariantType(TS) = VT_ARRAY Or VT_I2
    Else
        Call memcpy(TS, TrimStrs, VLEN_VARIANT)
    End If
    
    Dim Iteration As Long
    
    TSIndex = 0: TEIndex = -1
    While ((TLSIndex <> TSIndex) Or (TLEIndex <> TEIndex))
        TLSIndex = TSIndex
        TLEIndex = TEIndex
        MsgBox "Iteration:" & Iteration
        Iteration = Iteration + 1
        MsgBox "TLSIndex:" & TLSIndex
        MsgBox "TLEIndex:" & TEIndex
        cIndex = 0
        Do
            TSHaveNext = GetStringArrayItem(TS, tStr, cIndex)
            tLen = ArraySize(tStr)
            TSIndex = TrimStartIndex(iStr, iLen, tStr, tLen, TSIndex)
            TEIndex = TrimEndIndex(iStr, iLen, tStr, tLen, TEIndex)
        Loop While TSHaveNext
    Wend
    
    Call memzero(TS, VLEN_VARIANT)
    Call memzeroarr(tStr, VLEN_VARIANT)
    
    If TEIndex = -1 Then TEIndex = (iLen - TSIndex) - 1
    MsgBox "TEIndex:" & TEIndex
    
    If TEIndex = 0 Then GoTo End_Cleaner
    If TEIndex < TSIndex Then GoTo End_Cleaner
    
    Call memcpy(Trim, SubSring(iStr, TSIndex, (TEIndex - TSIndex) + 1), VLEN_PTR)

End_ErrorManager:
    If Not Err Is Nothing Then throw Err
End_Cleaner:
'    Call Arrays.Release(iStr, True)
End Function
Public Function TrimStart(ByRef Str As Variant, Optional ByRef TrimStrs As Variant) As String
    Dim TSIndex As Long, iLen As Long
    Dim tStr() As Integer, tLen As Long, cIndex As Long
    Dim TS As Variant, IsVMissing As Boolean
    Dim iStr() As Integer
    iStr = CastToIntArray(Str)
    iLen = ArraySize(iStr)
    
    IsVMissing = IsMissing(TrimStrs)
    If IsVMissing Then
        TS = Arrays.GetSafeArrayPointer(WS, True)
        VariantType(TS) = VT_ARRAY Or VT_I2
    Else
        Call memcpy(TS, TrimStrs, VLEN_VARIANT)
    End If
    
    TSIndex = 0
    While GetStringArrayItem(TS, tStr, cIndex)
        tLen = ArraySize(tStr)
        TSIndex = TrimStartIndex(iStr, iLen, tStr, tLen, TSIndex)
    Wend
    
    If IsVMissing Then _
        Call memzero(TS, VLEN_VARIANT)
    
    Call memcpy(TrimStart, SubSring(iStr, TSIndex), VLEN_PTR)
End Function
Public Function TrimEnd(ByRef Str As Variant, Optional ByRef TrimStrs As Variant) As String
    Dim TEIndex As Long, iLen As Long
    Dim tStr() As Integer, tLen As Long, cIndex As Long
    Dim TS As Variant, IsVMissing As Boolean
    Dim iStr() As Integer
    iStr = CastToIntArray(Str)
    iLen = ArraySize(iStr)
    
    IsVMissing = IsMissing(TrimStrs)
    If IsVMissing Then
        TS = Arrays.GetSafeArrayPointer(WS, True)
        VariantType(TS) = VT_ARRAY Or VT_I2
    Else
        Call memcpy(TS, TrimStrs, VLEN_VARIANT)
    End If
    
    TEIndex = -1
    While GetStringArrayItem(TS, tStr, cIndex)
        tLen = ArraySize(tStr)
        TEIndex = TrimEndIndex(iStr, iLen, tStr, tLen, TEIndex)
    Wend
    
    If IsVMissing Then _
        Call memzero(TS, VLEN_VARIANT)
    
    Call memcpy(TrimEnd, SubSring(iStr, 0, TEIndex + 1), VLEN_PTR)
End Function

Public Function PadLeft(ByRef Str As String, ByVal TotalLength As Long, Optional ByRef PaddingStr As Variant) As String 'VBCorLib
    Dim mPadStr As String
    If IsMissing(PaddingStr) Then PaddingStr = 32 Else PaddingStr = PaddingStr
    Dim lStr As Long
    lStr = Len(Str)
    If TotalLength > lStr Then
        PadLeft = String$(TotalLength - lStr, PaddingStr) & Str
    Else
        PadLeft = Str
    End If
End Function
Public Function PadRight(ByRef Str As String, ByVal TotalLength As Long, Optional ByRef PaddingStr As Variant) As String 'VBCorLib
    Dim mPadStr As String
    If IsMissing(PaddingStr) Then PaddingStr = 32 Else PaddingStr = PaddingStr
    Dim lStr As Long
    lStr = Len(Str)
    If TotalLength > lStr Then
        PadRight = Str & String$(TotalLength - lStr, mPadStr)
    Else
        PadRight = Str
    End If
End Function

Public Function Escape(ByRef Str As Variant, ByRef Elements As Variant) As String
    
End Function

Public Function IsNumber(ByRef Str As String) As Boolean
    IsNumber = VBA.IsNumeric(Str)
End Function
Public Function IsNumeric(ByRef Str As Variant) As Boolean
    Dim iStr() As Integer, iLen As Long
    iStr = CastToIntArray(Str)
    iLen = Arrays.Size(iStr)
    
    If iLen = 0 Then Exit Function
    
    Dim i As Long
    For i = 0 To iLen - 1
        If Not IsNumeric_(iStr(i)) Then
            IsNumeric = False
            Exit Function
        End If
    Next
    IsNumeric = True
End Function
Public Function IsAlphabetic(ByRef Str As Variant) As Boolean
    Dim iStr() As Integer, iLen As Long
    iStr = CastToIntArray(Str)
    iLen = Arrays.Size(iStr)
    
    If iLen = 0 Then Exit Function
    
    Dim i As Long
    For i = 0 To iLen - 1
        If Not IsAlphabetic_(iStr(i)) Then
            IsAlphabetic = False
            Exit Function
        End If
    Next
    IsAlphabetic = True
End Function
Public Function IsEmptyOrNull(ByRef Str As Variant) As Boolean
    Dim VT As VariantTypes
    VT = VarType(Str)
    Select Case VT
        Case vbString
            If StrPtr(Str) = vbNullPtr Then
                IsEmptyOrNull = True
                Exit Function
            End If
    End Select
    Dim sLen As Long
    sLen = Length(Str)
    If sLen <= 0 Then
        IsEmptyOrNull = True
        Exit Function
    End If
End Function
Public Function IsWhiteSpace(ByRef Str As Variant) As Boolean
    Dim iStr() As Integer, iLen As Long
    iStr = CastToIntArray(Str)
    iLen = ArraySize(iStr)
    
    If iLen = 0 Then Exit Function
    
    Dim i As Long
    For i = 0 To iLen - 1
        If Not IsWhiteSpace_(iStr(i)) Then
            IsWhiteSpace = False
            Exit Function
        End If
    Next
    IsWhiteSpace = True
End Function
Public Function IsEmptyOrWhiteSpace(ByRef Str As Variant) As Boolean
    IsEmptyOrWhiteSpace = IsEmptyOrNull(Str)
    If IsEmptyOrWhiteSpace Then Exit Function
    IsEmptyOrWhiteSpace = IsWhiteSpace(Str)
End Function
Public Function IsString(ByRef Str As Variant) As Boolean
    IsString = (VarType(Str) = vbString)
End Function

Public Property Get WhiteSpaces() As Integer()
    WhiteSpaces = WS
End Property

Public Function GetRandomString(ByVal CharType As CharacterType, ByVal MinLength As Long, Optional ByVal MaxLength As Long) As String
    Call Randomize(Timer)
    Dim Length As Long, Diff As Long
    If MaxLength = 0 Then MaxLength = MinLength * 2
    Diff = MaxLength - MinLength
    Length = MinLength + (Diff * Rnd)
    Dim RetVal As String, i As Long
    
    Select Case CharType
        Case ctAlphabetic
            Dim c As Long
            For i = 1 To Length
                c = CLng(Rnd * 50)
                If c < 25 Then c = c + 65 Else c = ((c - 25) + 97)
                RetVal = RetVal & Chr(c)
            Next
        Case ctNumeric
            For i = 1 To Length: RetVal = RetVal & CStr(Fix(Rnd * 10)): Next
        Case ctWhiteSpace '{TAB,SPACE,NBSP}
            
        Case ctAllHumanReadables
            
        Case ctBinary
            
        Case ctAll
            For i = 1 To Length: RetVal = RetVal & Chr$(Rnd * 256): Next
        Case Else
            throw Exps.InvalidArgumentException
    End Select
    
    GetRandomString = RetVal
End Function


Public Function Format(ByVal StrFormat As Variant, ParamArray Args() As Variant) As String
    
End Function

Public Function HashString(ByRef Str As Variant, Range As RangedLength) As Long
    
End Function

Public Function GetHashCode(ByRef Str As String) As Long
    GetHashCode = StrPtr(Str)
End Function

Public Function Compare(ByRef Str1 As String, ByRef Str2 As String, Optional ByVal IgnoreCase As Boolean = False) As Boolean
    Compare = StrComp(Str1, Str2, IIf(IgnoreCase, vbTextCompare, vbBinaryCompare))
End Function

Public Function Stringify(ByRef Expression) As String
    
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
'<section privates>
'
Friend Function ToUpper_(ByVal InInt As Integer) As Integer
    'Temporary
    If (InInt >= 97) And (InInt <= 122) Then ToUpper_ = InInt - 32 Else ToUpper_ = InInt
End Function
Friend Function ToLower_(ByVal InInt As Integer) As Integer
    'Temporary
    If (InInt >= 65) And (InInt <= 90) Then ToLower_ = InInt + 32 Else ToLower_ = InInt
End Function
Friend Function IsWhiteSpace_(ByVal InInt As Integer) As Boolean
    Dim i As Long
    For i = 0 To ArraySize(WS) - 1
        If InInt = WS(i) Then
            IsWhiteSpace_ = True
            Exit Function
        End If
    Next
End Function
Friend Function IsNumeric_(ByVal InInt As Integer) As Boolean
    If (InInt >= 48) And (InInt <= 57) Then IsNumeric_ = True
End Function
Friend Function IsAlphabetic_(ByVal InInt As Integer) As Boolean
    If InInt < 65 Then Exit Function
    If InInt > 122 Then Exit Function
    If ((InInt <= 90) Or (InInt >= 97)) Then Exit Function
    IsAlphabetic_ = True
End Function

Public Function TrimStartIndex(ByRef iStr() As Integer, ByVal iStrLen As Long, ByRef tStr() As Integer, ByVal tStrLen As Long, ByVal StartIndex As Long)
    If tStrLen > iStrLen Then Exit Function
    Dim i As Long
    For i = StartIndex To tStrLen - 1
        If iStr(i) <> tStr(i) Then
            TrimStartIndex = i
            Exit Function
        End If
    Next
    TrimStartIndex = i
End Function
Public Function TrimEndIndex(ByRef iStr() As Integer, ByVal iStrLen As Long, ByRef tStr() As Integer, ByVal tStrLen As Long, ByVal StartIndex As Long)
    If tStrLen > iStrLen Then
        If iStrLen > 0 Then TrimEndIndex = iStrLen - 1
        Exit Function
    End If
    If StartIndex = -1 Then StartIndex = iStrLen - 1
    Dim i As Long, ii As Long
    ii = StartIndex
    For i = tStrLen - 1 To 0 Step -1
        If iStr(ii) <> tStr(i) Then
            TrimEndIndex = ii
            Exit Function
        End If
        ii = ii - 1
    Next
    TrimEndIndex = ii
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
Private Sub Class_Initialize()
    WS = Arrays.CreateSafeVector(atIntegerArray, 21)
    WS(0) = &H20
    WS(1) = &HD
    WS(2) = &H9
    WS(3) = &HA
    WS(4) = &HB
    WS(5) = &HC
    WS(6) = &HA0
    WS(7) = &H2000
    WS(8) = &H2001
    WS(9) = &H2002
    WS(10) = &H2003
    WS(11) = &H2004
    WS(12) = &H2005
    WS(13) = &H2006
    WS(14) = &H2007
    WS(15) = &H2008
    WS(16) = &H2009
    WS(17) = &H200A
    WS(18) = &H200B
    WS(19) = &H3000
    WS(20) = &HFEFF
    
    'WS = Arrays.Integers(&H20, &HD, &H9, &HA, &HB, &HC, &HA0, &H2000, &H2001, &H2002, &H2003, &H2004, &H2005, &H2006, &H2007, &H2008, &H2009, &H200A, &H200B, &H3000, &HFEFF)
End Sub
Private Sub Class_Terminate()
    'Call memzeroarr(WS, VLEN_PTR)
End Sub
