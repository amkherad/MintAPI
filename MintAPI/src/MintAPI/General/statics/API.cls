VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "API"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Global MintAPI methods provided for user applications."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Application"
Attribute VB_Ext_KEY = "Member1" ,"ApplicationMessageEventArgs"
Attribute VB_Ext_KEY = "Member2" ,"ArgumentList"
Attribute VB_Ext_KEY = "Member3" ,"BigNumber"
Attribute VB_Ext_KEY = "Member4" ,"ByteArray"
Attribute VB_Ext_KEY = "Member5" ,"Configuration"
Attribute VB_Ext_KEY = "Member6" ,"Console"
Attribute VB_Ext_KEY = "Member7" ,"Directory"
Attribute VB_Ext_KEY = "Member8" ,"DiskDrive"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member9" ,"Environment"
Attribute VB_Ext_KEY = "Member10" ,"EventArgs"
Attribute VB_Ext_KEY = "Member11" ,"EventArgs"
Attribute VB_Ext_KEY = "Member12" ,"EventArgs"
Attribute VB_Ext_KEY = "Member13" ,"EventArgs"
Attribute VB_Ext_KEY = "Member14" ,"ExceptionOccuredEventArgs"
Attribute VB_Ext_KEY = "Member15" ,"File"
Attribute VB_Ext_KEY = "Member16" ,"Generics"
Attribute VB_Ext_KEY = "Member17" ,"Language"
Attribute VB_Ext_KEY = "Member18" ,"List"
Attribute VB_Ext_KEY = "Member19" ,"Memory"
Attribute VB_Ext_KEY = "Member20" ,"Method"
Attribute VB_Ext_KEY = "Member21" ,"Process"
Attribute VB_Ext_KEY = "Member22" ,"Registry"
Attribute VB_Ext_KEY = "Member23" ,"StringParser"
Attribute VB_Ext_KEY = "Member24" ,"StringParser"
Attribute VB_Ext_KEY = "Member25" ,"Thread"
Attribute VB_Ext_KEY = "Member26" ,"Thread"
Attribute VB_Ext_KEY = "Member27" ,"Timer"
Attribute VB_Ext_KEY = "Member28" ,"Timer"
Attribute VB_Ext_KEY = "Member29" ,"Vector"
Attribute VB_Ext_KEY = "Member30" ,"Vector"
Attribute VB_Ext_KEY = "Member31" ,"Vector"
Attribute VB_Ext_KEY = "Member32" ,"Vector"
Attribute VB_Ext_KEY = "Member33" ,"EventArgs"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "API"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING
''<class>
''
''</class>

''
Private p_VarResult As IResult
Private p_VarValue As Variant
Private p_StateValue As Variant
Private p_VarCollection As New Dictionary


Private Sub Class_Initialize()
    p_VarCollection.AutoCreateOnSet = True
End Sub
'=============================================
'=============================================
'=============================================
'<section Generic Methods>
'

Public Function tr(ByVal Key As String) As String
    If Not mint_application.CurrentTranslation.TryTranslate(Key, tr) Then _
        tr = Key
End Function
Public Function GetAddressOf(ByVal Address As Long) As Long
    GetAddressOf = Address
End Function

Public Function Using(ByVal Object As IDisposable) As UsingBlock
    Set Using = New UsingBlock
    Call Using.Constructor0(Object)
End Function
Public Function UsingH(ByVal Object As IHandle) As UsingHandleBlock
    Set UsingH = New UsingHandleBlock
    Call UsingH.Constructor0(Object)
End Function

Public Function VariantType(ByRef Var As Variant) As VariantTypes
    VariantType = mint_api.VariantType(Var)
End Function
Public Function VariantTypeNoByRef(ByRef Var As Variant) As VariantTypes
    VariantTypeNoByRef = mint_api.VariantType(Var)
    If (VariantTypeNoByRef And VT_BYREF) = VT_BYREF Then _
        VariantTypeNoByRef = VariantTypeNoByRef Xor VT_BYREF
End Function
Public Sub VariantCopy(ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
    Call API_VariantCopy(pvargDest, pvargSrc)
End Sub
Public Sub VariantCopyInd(ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
    Call API_VariantCopyInd(pvargDest, pvargSrc)
End Sub
Public Sub VariantMove(ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
    Call memcpy(pvargDest, pvargSrc, VLEN_VARIANT)
    Call memzero(pvargSrc, VLEN_VARIANT)
End Sub
Public Sub Eval(ByRef Dst As Variant, ByRef Src As Variant)
    If IsObject(Src) Then
        Set Dst = Src
    Else
            Dst = Src
    End If
End Sub
Public Sub EvalX(ByRef Dst As Variant, ByRef Src As Variant)
    If IsObject(Src) Then
        Set Dst = Src
    Else
            Dst = Src
    End If
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
Public Function Succeeded(ByRef Value As Variant, Optional ByRef RetValue As Variant) As Boolean
    Dim VT As VariantTypes
    VT = VariantType(Value)
    Select Case VT
'        Case VT_BOOL, VT_INT, VT_INT_PTR, VT_UI1, _
'             VT_I1, VT_I2, VT_I4, VT_I8, VT_R4, _
'             VT_R8, VT_UI2, VT_UI4, VT_UI8, VT_UR4, _
'             VT_UR8, VT_CY, VT_DATE, VT_DECIMAL, _
'             VT_HRESULT, VT_VARIANT
'            Succeeded = CBool(Value)
        Case VT_BSTR, VT_BSTR_BLOB, VT_LPSTR, VT_LPWSTR
            Succeeded = (Len(Value) > 0)
        Case VT_DISPATCH, VT_Unknown
            Succeeded = (ObjPtr(Value) <> 0)
        Case VT_HRESULT
            Succeeded = (CLng(Value) > 0)
        Case Else
On Error GoTo CatchErr
            If IsArray(Value) Then
                Succeeded = (Arrays.GetLength(Value) > 0)
            Else
                Succeeded = CBool(Value)
CatchErr:
            End If
    End Select
    If Not IsMissing(RetValue) Then Call Evaluate(RetValue, Value)
End Function

Public Function IsReferenceType(ByVal Var As Variant) As Boolean
    If IsObject(Var) Then IsReferenceType = True: Exit Function
    If IsArray(Var) Then IsReferenceType = True: Exit Function
End Function

Public Function IsNullOrMissing(ByRef Target As Variant) As Boolean
    If IsMissing(Target) Then IsNullOrMissing = True: Exit Function
    If IsNull(Target) Then IsNullOrMissing = True: Exit Function
End Function
Public Function IsNullOrMissingOrEmpty(ByRef Target As Variant) As Boolean
    If IsMissing(Target) Then IsNullOrMissingOrEmpty = True: Exit Function
    If IsNull(Target) Then IsNullOrMissingOrEmpty = True: Exit Function
    If IsEmpty(Target) Then IsNullOrMissingOrEmpty = True: Exit Function
End Function

Public Function IsIUnknown(ByRef Target As Variant) As Boolean
    IsIUnknown = IsObject(Target)
End Function
''<summary>Cast anything into IUnknown.</summary>
Public Function CUnk(ByVal Unknown As Object) As IUnknown ', Optional ByVal AddRef As Boolean = False
    Set CUnk = Unknown
End Function

''<summary>Cast a variant to specified data type, but it remains variant.</summary>
''<retval type="variant">A variant refering to casted value.</retval>
''<see>CUnk</see>
Public Function Cast(ByVal CastType As VbVarType, ByRef From As Variant, ByRef Default As Variant) As Variant
    On Error GoTo errDefault
    Select Case CastType
        Case vbInteger:     Cast = CInt(From)
        Case vbLong:        Cast = CLng(From)
        Case vbSingle:      Cast = CSng(From)
        Case vbDouble:      Cast = CDbl(From)
        Case vbCurrency:    Cast = CCur(From)
        Case vbDate:        Cast = CDate(From)
        Case vbBoolean:     Cast = CBool(From)
        Case vbDecimal:     Cast = CDec(From)
        Case vbByte:        Cast = CByte(From)
        Case VBObject:      Set Cast = CUnk(From)
    End Select
    Exit Function
    'a VBColLib method
errDefault:
    Call API_VariantCopy(Cast, Default)
End Function

Public Function AsByVal(ByRef Val As Variant) As Variant
    If Info.IsByRef(Val) Then
        If Info.IsRecord(Val) Then
            AsByVal = Val
        Else
            Call API_VariantCopyInd(AsByVal, Val)
        End If
    Else
        Call EvaluateX(AsByVal, Val)
    End If
End Function

Public Function CInt64(ByVal Number As Currency) As Int64
    Dim CY As Currency
    CY = (CCur(Number) * 0.0001@)
    Call memcpy(ByVal VarPtr(CInt64), ByVal VarPtr(CY), VLEN_CURRENCY)
End Function
Public Function CInt64ToCur(IntValue As Int64) As Currency
    Call memcpy(ByVal VarPtr(CInt64ToCur), ByVal VarPtr(IntValue), VLEN_CURRENCY)
    CInt64ToCur = (CInt64ToCur * 10000@)
End Function

Public Function VarPtrArr(ByRef Arr As Variant) As Long
    If Not IsArray(Arr) Then throw Exps.ArrayExpectedException("Arr")
    VarPtrArr = VarPtr(Arr) + VARIANT_DATA_OFFSET
    If Info.IsByRef(Arr) Then _
        Call memcpy(VarPtrArr, ByVal VarPtrArr, VLEN_PTR)
End Function
Public Sub EmptyVar(ByRef Variable As Variant)
    If IsObject(Variable) Then
        Set Variable = Nothing
    Else
        If IsArray(Variable) Then
            Erase Variable
        Else
            Variable = Empty
        End If
    End If
End Sub
Public Function IsEmptyVariable(ByRef TargetVariable As Variant) As Boolean
    Select Case VarType(TargetVariable)
        Case vbArray
            IsEmptyVariable = (IsEmptyArray(TargetVariable))
        Case VBObject
            IsEmptyVariable = (TargetVariable Is Nothing)
        Case vbString
            IsEmptyVariable = (TargetVariable = "")
        Case Else
            IsEmptyVariable = (TargetVariable = Empty)
    End Select
End Function

Public Function Query(ByVal strQuery As String, ParamArray Params() As Variant) As IResult
    Dim Args As ArgumentList
    Set Args = AL.FromParamArray(vsString)
    
    
    
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Streaming API>
'
Public Sub out(ByVal Stream As IClassStream, ByRef Data As Variant, Optional ByVal ConvertToBinary As Boolean = True, Optional Length As Long = -1)
    If Stream Is Nothing Then Set Stream = AppOutput
    If Not Stream.GetState(sdOutStream) Then throw Exps.InvalidOperationException
    Dim cBuffer As ClassStreamBuffer
    If ConvertToBinary Then
        Set cBuffer = Buffer.ClassStreamBuffer(csbtBinary Or csbtOutput, Data, Length)
    Else
        Set cBuffer = Buffer.ClassStreamBufferV(Data, Length)
    End If
    Call Stream.OutStream(cBuffer, Length)
End Sub
Public Sub inp(ByVal Stream As IClassStream, ByRef Data As Variant, Optional Length As Long = -1)
    If Stream Is Nothing Then Set Stream = AppInput
    If Not Stream.GetState(sdInStream) Then throw Exps.InvalidOperationException
    Dim cBuffer As ClassStreamBuffer
    Set cBuffer = Buffer.ClassStreamBuffer(csbtInput, Data, Length)
    Call Stream.InStream(cBuffer, Length)
    Call cBuffer.MoveToInput
End Sub
Public Sub lOut(ByVal Stream As IClassStream, ParamArray Data() As Variant)
    If Stream Is Nothing Then Set Stream = AppOutput
    If Not Stream.GetState(sdOutStream) Then throw Exps.InvalidOperationException
    Dim V As Variant
    For Each V In Data
        Call Stream.OutStream(Buffer.ClassStreamBufferV(V))
    Next
zeroLength:
End Sub
Public Sub lInp(ByVal Stream As IClassStream, ParamArray Data() As Variant)
    If Stream Is Nothing Then Set Stream = AppInput
    If Not Stream.GetState(sdInStream) Then throw Exps.InvalidOperationException
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Dim cBuffer As ClassStreamBuffer
        Set cBuffer = Buffer.ClassStreamBufferV(Data(i))
        Call Stream.InStream(cBuffer)
        Call cBuffer.MoveToInput
    Next
zeroLength:
End Sub
Public Function Endl() As Object
    Set Endl = Buffer.ClassStreamBuffer(csbtNewLine)
End Function
Public Function SClr() As Object
    Set SClr = Buffer.ClassStreamBuffer(csbtClear)
End Function
Public Function SFlush() As Object
    Set SFlush = Buffer.ClassStreamBuffer(csbtFlush)
End Function
Public Function SSeek(ByVal Index As Long) As Object
    Set SSeek = Buffer.ClassStreamBuffer(csbtSeek, Index, True)
End Function

Public Sub stdin(ParamArray Data() As Variant)
    Dim i As Long, AppIn As IClassStream
    Dim Count As Long
    On Error GoTo zeroLength
    Count = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Count = 0 Then throw Exps.InvalidOperationException
    Set AppIn = AppInput
    If Not AppIn.GetState(sdInStream) Then throw Exps.InvalidOperationException
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Dim cBuffer As ClassStreamBuffer
        Set cBuffer = Buffer.ClassStreamBuffer(csbtInput, Data(i))
        Call AppIn.InStream(cBuffer)
        Call cBuffer.MoveToInput
    Next
End Sub
Public Sub stdinb(ParamArray Data() As Variant)
    Dim i As Long, AppIn As IClassStream
    Dim Count As Long
    On Error GoTo zeroLength
    Count = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Count = 0 Then throw Exps.InvalidOperationException
    Set AppIn = AppInput
    If Not AppIn.GetState(sdInStream) Then throw Exps.InvalidOperationException
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Dim cBuffer As ClassStreamBuffer
        Set cBuffer = Buffer.ClassStreamBuffer(csbtBinary Or csbtInput, Data(i))
        Call AppIn.InStream(cBuffer)
        Call cBuffer.MoveToInput
    Next
End Sub
Public Sub stdout(ParamArray Data() As Variant)
    Dim V As Variant, AppOut As IClassStream
    Set AppOut = AppOutput
    If Not AppOut.GetState(sdOutStream) Then throw Exps.InvalidOperationException
    On Error GoTo 0
    For Each V In Data
        Call AppOut.OutStream(Buffer.ClassStreamBufferV(V))
    Next
End Sub
Public Sub stdoutb(ParamArray Data() As Variant)
    Dim V As Variant, AppOut As IClassStream
    Set AppOut = AppOutput
    If Not AppOut.GetState(sdOutStream) Then throw Exps.InvalidOperationException
    On Error GoTo 0
    For Each V In Data
        Call AppOut.OutStream(Buffer.ClassStreamBuffer(csbtBinary Or csbtOutput, V))
    Next
End Sub
Public Sub stderr(ParamArray Data() As Variant)
    Dim V As Variant, AppErr As IClassStream
    Set AppErr = AppError
    If Not AppErr.GetState(sdOutStream) Then throw Exps.InvalidOperationException
    On Error GoTo 0
    For Each V In Data
        Call AppErr.OutStream(Buffer.ClassStreamBufferV(V))
    Next
End Sub
Public Sub stdWarning(ParamArray Data() As Variant)
'    Dim Size As Long
'    On Error GoTo zeroLength
'    Size = UBound(Data) - LBound(Data) + 1
'zeroLength:
'    If Size = 0 Then Exit Sub
'    Dim i As Long
'    On Error GoTo 0
'    For i = LBound(Data) To UBound(Data)
'        Call mint_application.stdWarning(Data(i))
'    Next
End Sub
Public Sub stdDebug(ParamArray Data() As Variant)
'    Dim Size As Long
'    On Error GoTo zeroLength
'    Size = UBound(Data) - LBound(Data) + 1
'zeroLength:
'    If Size = 0 Then Exit Sub
'    Dim i As Long
'    On Error GoTo 0
'    For i = LBound(Data) To UBound(Data)
'        Call mint_application.stdDebug(Data(i))
'    Next
End Sub
Public Sub stdError(ParamArray Data() As Variant)
'    Dim Size As Long
'    On Error GoTo zeroLength
'    Size = UBound(Data) - LBound(Data) + 1
'zeroLength:
'    If Size = 0 Then Exit Sub
'    Dim i As Long
'    On Error GoTo 0
'    For i = LBound(Data) To UBound(Data)
'        Call mint_application.stdError(Data(i))
'    Next
End Sub
Public Sub stdCritical(ParamArray Data() As Variant)
'    Dim Size As Long
'    On Error GoTo zeroLength
'    Size = UBound(Data) - LBound(Data) + 1
'zeroLength:
'    If Size = 0 Then Exit Sub
'    Dim i As Long
'    On Error GoTo 0
'    For i = LBound(Data) To UBound(Data)
'        Call mint_application.stdCritical(Data(i))
'    Next
End Sub

Public Sub ConIn(ParamArray Data() As Variant)
    Dim i As Long, Count As Long
    On Error GoTo zeroLength
    Count = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Count = 0 Then throw Exps.InvalidOperationException
    For i = LBound(Data) To UBound(Data)
        inp Console, Data(i)
    Next
End Sub
Public Sub ConOut(ParamArray Data() As Variant)
    Dim V As Variant
    For Each V In Data
        out Console, V
    Next
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Object Buffer API>
'
'!UNSAFE METHOD.
'Public Function MintAPIType(ByVal TypeRefrence As Long, ByVal TypeLength As Long) As Object
'    Set MintAPIType = Objects.CreateObjectBuffer(OBJECTBUFFER_MINTLOCALTYPE, TypeRefrence, TypeLength)
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................

'==========================================================================================
'== Signal & Slot Connection Section ==================================================================
'==========================================================================================


'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................

'==========================================================================================
'== Bonus Section ===============================================================
'==========================================================================================


Public Property Get This() As Object
    Call memcpy(This, mHelper.CallerThis, VLEN_PTR)
    Call IUnknown.AddRef(This)
End Property
Public Property Get That() As Object
    Dim EBP As Long
    EBP = mHelper.ReadCallerEBP
    If Not Memory.CanAccessTo(EBP, VLEN_PTR) Then throw Exps.AccessDeniedException
    Call memcpy(EBP, ByVal EBP, VLEN_PTR)
    If Not Memory.CanAccessTo(EBP + 8, VLEN_PTR) Then throw Exps.AccessDeniedException
    Call memcpy(That, ByVal (EBP + 8), VLEN_PTR)
    Call IUnknown.AddRef(That)
End Property

''<summary>Free global safe-thread variable to be used by end-user.</summary>
''<remarks>The value of this property is thread static and holds a specified value on each thread.</remarks>
''<samples>Var = Query()</samples>
Public Property Get Var() As Variant
    Call EvaluateX(Var, p_VarValue)
End Property
Public Property Let Var(ByRef Value As Variant)
    Call Evaluate(p_VarValue, Value)
End Property
Public Property Set Var(ByVal Value As Object)
    Call Evaluate(p_VarValue, Value)
End Property
Public Property Get V() As Dictionary
    Set V = p_VarCollection
End Property
Public Property Get State() As Variant
    Call EvaluateX(State, p_StateValue)
End Property
Public Property Let State(ByRef Value As Variant)
    Call Evaluate(p_StateValue, Value)
End Property
Public Property Set State(ByVal Value As Object)
    Call Evaluate(p_StateValue, Value)
End Property

Public Property Get Result() As IResult
    Set Result = p_VarResult
End Property
Public Property Let Result(ByRef Value As IResult)
    Set p_VarResult = Value
End Property
Public Property Set Result(ByVal Value As IResult)
    Set p_VarResult = Value
End Property

''<summary>IncVar(x) is the same as the C operator ++x</summary>
Public Function IncVar(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
    Value = Value + Amount
    IncVar = Value
End Function
''<summary>DecVar(x) is the same as the C operator --x</summary>
Public Function DecVar(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
    Value = Value - Amount
    DecVar = Value
End Function
''<summary>VarInc(x) is the same as the C operator x++</summary>
Public Function VarInc(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
    VarInc = Value
    Value = Value + Amount
End Function
''<summary>VarDec(x) is the same as the C operator x--</summary>
Public Function VarDec(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
    VarDec = Value
    Value = Value - Amount
End Function
''<summary>assigns a value and returns it.</summary>
Public Function LetVar(ByRef Variable As Variant, Value As Variant) As Variant
    If IsObject(Value) Then
        Set Variable = Value
        Set LetVar = Variable
    Else
        Variable = Value
        LetVar = Variable
    End If
End Function
''<summary>Returns the variable value and, after that, assigns new value.</summary>
Public Function VarLet(ByRef Variable As Variant, Value As Variant) As Variant
    If IsObject(Value) Then
        Set VarLet = Variable
        Set Variable = Value
    Else
        VarLet = Variable
        Variable = Value
    End If
End Function
'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................
