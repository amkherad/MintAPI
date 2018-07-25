Attribute VB_Name = "mint_assemblies"
Option Explicit

Private Const KERNEL32_GLOBAL_ALLOC As String = "GlobalAlloc"
Private Const KERNEL32_GLOBAL_FREE As String = "GlobalFree"

Public Const FUNC_ORDER_SKIP_IUNKNOWN       As Long = 3 'Base 0
Public Const FUNC_ORDER_SKIP_IUNKNOWN_B1    As Long = 2 'Base 1

Public Const FUNC_ORDER_SKIP_IDISPATCH      As Long = FUNC_ORDER_SKIP_IUNKNOWN + 4    'Base 0
Public Const FUNC_ORDER_SKIP_IDISPATCH_B1   As Long = FUNC_ORDER_SKIP_IUNKNOWN_B1 + 4 'Base 1

Private Const FUNC_ORDER_FIRSTMETHOD As Long = FUNC_ORDER_SKIP_IUNKNOWN

'Private Const FUNC_ORDER_MEMCPY As Long = 0
Public Const FUNC_ORDER_MEMZERO As Long = FUNC_ORDER_FIRSTMETHOD + 0

Public Const FUNC_ORDER_READEAX As Long = FUNC_ORDER_FIRSTMETHOD + 1

Public Const FUNC_ORDER_READFS As Long = FUNC_ORDER_FIRSTMETHOD + 2
Public Const FUNC_ORDER_WRITEFS As Long = FUNC_ORDER_FIRSTMETHOD + 3

Public Const FUNC_ORDER_READESP As Long = FUNC_ORDER_FIRSTMETHOD + 4
Public Const FUNC_ORDER_WRITEESP As Long = FUNC_ORDER_FIRSTMETHOD + 5

Public Const FUNC_ORDER_READCALLEEEBP As Long = FUNC_ORDER_FIRSTMETHOD + 6
Public Const FUNC_ORDER_WRITECALLEEEBP As Long = FUNC_ORDER_FIRSTMETHOD + 7

Public Const FUNC_ORDER_READCALLEREBP As Long = FUNC_ORDER_FIRSTMETHOD + 8
Public Const FUNC_ORDER_WRITECALLEREBP As Long = FUNC_ORDER_FIRSTMETHOD + 9

Public Const FUNC_ORDER_SHIFTLEFT As Long = FUNC_ORDER_FIRSTMETHOD + 10
Public Const FUNC_ORDER_SHIFTRIGHT As Long = FUNC_ORDER_FIRSTMETHOD + 11

Public Const FUNC_ORDER_CALLEETHIS As Long = FUNC_ORDER_FIRSTMETHOD + 12
Public Const FUNC_ORDER_CALLERTHIS As Long = FUNC_ORDER_FIRSTMETHOD + 13

Public Const FUNC_ORDER_GETIP As Long = FUNC_ORDER_FIRSTMETHOD + 14

Public Const FUNC_ORDER_RESERVE As Long = FUNC_ORDER_FIRSTMETHOD + 15
Public Const FUNC_ORDER_RETURN As Long = FUNC_ORDER_FIRSTMETHOD + 16

Public Const FUNC_ORDER_CALLINT32 As Long = FUNC_ORDER_FIRSTMETHOD + 17
Public Const FUNC_ORDER_CALLDBL As Long = FUNC_ORDER_FIRSTMETHOD + 18
Public Const FUNC_ORDER_CALLINT64 As Long = FUNC_ORDER_FIRSTMETHOD + 19
Public Const FUNC_ORDER_CALL As Long = FUNC_ORDER_FIRSTMETHOD + 20

Public Const FUNC_ORDER_INCVAR32 As Long = FUNC_ORDER_FIRSTMETHOD + 21
Public Const FUNC_ORDER_VARINC32 As Long = FUNC_ORDER_FIRSTMETHOD + 22

Public Const FUNC_ORDER_PopInt32 As Long = FUNC_ORDER_FIRSTMETHOD + 25
Public Const FUNC_ORDER_PopInt64 As Long = FUNC_ORDER_FIRSTMETHOD + 26
Public Const FUNC_ORDER_PopInt128 As Long = FUNC_ORDER_FIRSTMETHOD + 27

'----- Ext

Public Const FUNC_ORDER_ROTATELEFT As Long = FUNC_ORDER_FIRSTMETHOD + 30
Public Const FUNC_ORDER_ROTATERIGHT As Long = FUNC_ORDER_FIRSTMETHOD + 31
'
'Public Const FUNC_ORDER_ARGUMENTLISTSTATIC_FROMPARAMARRAY As Long = FUNC_ORDER_FIRSTMETHOD + 32

'================================================================
'
'  mHelper interface structure:
'   '01. memcpy()
'   02. memzero()
'
'   03. ReadEAX()
'   04. WriteEAX()
'
'   05. ReadFS()   'for TES structure.
'   06. WriteFS()  'for TES structure.
'
'   07. ReadESP()
'   08. WriteESP()
'
'   09. ReadCalleeEBP()
'   10. WriteCalleeEBP()
'
'   11. ReadCallerEBP()
'   12. WriteCallerEBP()
'
'   13. ShiftLeft()
'   14. ShiftRight()
'
'   15. PushB()
'   16. PopB()
'
'   17. PushI()
'   18. PopI()
'
'   19. PushL()
'   20. PopL()
'
'   21. PushF()
'   22. PopF()
'
'   23. PushD()
'   24. PopD()
'
'   25. CalleeThis()
'   26. CallerThis()
'
'   27. GetIP()
'
'   28. Return()
'
'   29. Reserve()
'
'   30. Call() as void
'
'================================================================

Private Type HelperStructure
    pVTable As Long
    cRefs As Long
    tCount As Long
    
    CallProlog As New ByteArray
    CallEpilog As New ByteArray
    
    Funcs As Dictionary
    FuncTable(50) As Long
End Type

Public mHelper As IMintHelper

Private pHelperS As HelperStructure


Public Sub Construct()
    Set pHelperS.Funcs = New Dictionary
    
    Call Init_Headers(Library.Kernel32)
    
    Call Init_ClassMethods32
    
    Call Init_InternalMethods32
    
    pHelperS.cRefs = 1
    Dim VTable As Long
    pHelperS.pVTable = VarPtr(pHelperS.FuncTable(0))
    Call memcpy(mHelper, VarPtr(pHelperS.pVTable), VLEN_PTR)
End Sub

Private Sub Init_Headers(ByVal Lib As Library)
    Set pHelperS.CallProlog = Create_CallProlog(Lib)
    Set pHelperS.CallEpilog = Create_CallEpilog(Lib)
End Sub

Private Sub Init_ClassMethods32()
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_MEMZERO, Create_Helper_MemZero)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_READFS, Create_Helper_ReadFS)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_WRITEFS, Create_Helper_WriteFS)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_READESP, Create_Helper_ReadESP)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_WRITEESP, Create_Helper_WriteESP)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_READCALLEEEBP, Create_Helper_ReadCalleeEBP)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_WRITECALLEEEBP, Create_Helper_WriteCalleeEBP)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_READCALLEREBP, Create_Helper_ReadCallerEBP)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_WRITECALLEREBP, Create_Helper_WriteCallerEBP)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_SHIFTLEFT, Create_Helper_ShiftLeft)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_SHIFTRIGHT, Create_Helper_ShiftRight)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_RETURN, Create_Helper_Return)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_CALLEETHIS, Create_Helper_CalleeThis)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_CALLERTHIS, Create_Helper_CallerThis)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_CALL, Create_Helper_Call)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_CALLDBL, Create_Helper_CallDouble)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_CALLINT32, Create_Helper_CallInt32)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_CALLINT64, Create_Helper_CallInt64)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_GETIP, Create_Helper_GetIP)
'---------------------------------------------------------------------
    Call InsertMethodInHelperStructure(pHelperS, FUNC_ORDER_RESERVE, Create_Helper_Reserve)
'---------------------------------------------------------------------
End Sub


Private Sub Init_InternalMethods32()
    pHelperS.FuncTable(FUNC_ORDER_QueryInterface) = GetAddressOf(AddressOf IMintHelper_IUnknown_QueryInterface)
    pHelperS.FuncTable(FUNC_ORDER_AddRef) = GetAddressOf(AddressOf IMintHelper_IUnknown_AddRef)
    pHelperS.FuncTable(FUNC_ORDER_Release) = GetAddressOf(AddressOf IMintHelper_IUnknown_Release)
    
    pHelperS.FuncTable(FUNC_ORDER_ROTATELEFT) = GetAddressOf(AddressOf bitOperations.RotateLeft)
    pHelperS.FuncTable(FUNC_ORDER_ROTATERIGHT) = GetAddressOf(AddressOf bitOperations.RotateRight)
End Sub

'Can cast to IUnknown.
Private Function IMintHelper_IUnknown_QueryInterface(ByRef This As HelperStructure, ByVal riid As Long, ByRef pvObj As Long) As Long
    If riid = vbNullPtr Then GoTo CatchError
    'MsgBox "IMintHelper_IUnknown_QueryInterface"
    Dim Gu As Guid
    On Error GoTo CatchError
    Set Gu = Guid.FromReference(riid)
    If Gu.Equals(Guid.IUnknown) Then
        IMintHelper_IUnknown_QueryInterface = S_OK
    ElseIf Gu.Equals(Guid.IDispatch) Then
        IMintHelper_IUnknown_QueryInterface = S_OK
    ElseIf Gu.Equals(Guid.FromParts(&H7F87FFFF, &H7710, &H45, Arrays.Bytes(&HFF, &H7F, &H34, &H12, &H66, &HC6, &HCD, &H7A))) Then
        '{7F87FFFF-7710-0045-7FFF-7ACDC6661234} 'IMintHelper
        IMintHelper_IUnknown_QueryInterface = S_OK
    Else
        'Dim IMH As IMintHelper
        'If Gu.Equals(Objects.GetDispatch(IMH).Guid) Then
        '    pvObj = VarPtr(This.pVTable)
        '    IMintHelper_IUnknown_QueryInterface = S_OK
        'Else
            GoTo CatchError
        'End If
    End If
    pvObj = VarPtr(This.pVTable)
    This.cRefs = This.cRefs + 1
    Exit Function
CatchError:
    IMintHelper_IUnknown_QueryInterface = ERROR_E_NOINTERFACE
End Function
Private Function IMintHelper_IUnknown_AddRef(ByRef This As HelperStructure) As Long
    Dim cRefs As Long
    cRefs = This.cRefs + 1
    This.cRefs = cRefs
    IMintHelper_IUnknown_AddRef = cRefs
End Function
Private Function IMintHelper_IUnknown_Release(ByRef This As HelperStructure) As Long
    Dim cRefs As Long
    cRefs = This.cRefs - 1

    This.cRefs = cRefs
    IMintHelper_IUnknown_Release = cRefs

    Set This.Funcs = Nothing
    Set This.CallEpilog = Nothing
    Set This.CallProlog = Nothing

    If cRefs <= 0 Then _
        Call memzero(mHelper, VLEN_PTR)
End Function

Private Function IMintHelper_IDispatch_Invoke(ByVal dispIdMember As Long, ByVal riid As Long, ByVal LCID As Long, ByVal wFlags As Long, pDispParams As API_DISPPARAMS, pVarResult, pExcepInfo As API_EXCEPINFO, ByRef puArgErr As Long) As Long 'HRESULT
    
    MsgBox "IMintHelper_IDispatch_Invoke"
End Function





'*************************************************************************************
'**=================================================================================**
'*************************************************************************************
'**|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||**
'*************************************************************************************
'**=================================================================================**
'*************************************************************************************




Public Function GetMintAssembliesMethodGlobalInfo(ByVal MethodIndex As Long) As Method
    Set GetMintAssembliesMethodGlobalInfo = pHelperS.Funcs(MethodIndex)
End Function

Private Sub InsertMethodInHelperStructure(ByRef pHelper As HelperStructure, _
        ByVal Index As Long, ByVal Method As Method)
    
    Call pHelper.Funcs.Add(Index, Method)
    pHelper.FuncTable(Index) = Method.Reference
    
    pHelper.tCount = pHelper.tCount + 1
End Sub


''<summary>Reads the value of EAX register.</summary>
''<retval type="long">long value over EAX</retval>
Private Function Create_Helper_ReadEAX() As Method
    Set Create_Helper_ReadEAX = Runtime.CreateMethod("ReadEAX", _
        Prototype.Factory.SetRetVal(Prototype.AsEAXValue Or VT_Long).Clone, _
        ByteArray(CLng(&HC3)))          ' C3 'RET
End Function
''<summary>Returns FS segment register.</summary>
''<retval type="long">long value over EAX</retval>
Private Function Create_Helper_ReadFS() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H18A164))      ' 64 A1 18 00 'MOV EAX, FS:[18]
    Call BA.Append(CLng(&HC30000))      ' 00 00 C3 00 'RET 0
    
    Set Create_Helper_ReadFS = Runtime.CreateMethod("ReadFS", _
        Prototype.Factory.SetRetVal(Prototype.AsEAXValue Or VT_Long).Clone, BA)
End Function
''<summary>Writes fsValue into FS segment register.</summary>
''<params><param type="long" name="fsValue"></param></params>
Private Function Create_Helper_WriteFS() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H424448B))     ' 8B 44 24 08 'Mov EAX, SS:[ESP+8]
    Call BA.Append(CLng(&H18A364))      ' 64 A3 18 00 'Mov FS:[18], EAX
    Call BA.Append(CLng(&H4C20000))     ' 00 00 C2 04 'Ret 4
    
    Set Create_Helper_WriteFS = Runtime.CreateMethod("WriteFS", _
        Prototype.Factory.NewArg("fsValue", VT_In Or VT_Long).Clone, _
        BA)
End Function
''<summary>Shifts the number to left.</summary>
''<params>
''  <param type="long" name="Value"></param>
''  <param type="long" name="Count"></param>
''</params>
Private Function Create_Helper_ShiftLeft() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H824448B))     ' 8B 44 24 08 'Mov EAX, SS:[ESP+8]
    Call BA.Append(CLng(&HC244C8B))     ' 8B 4C 24 0C '
    Call BA.Append(CLng(&HCC2E0D3))     ' D3 E0 C2 0C '
    Call BA.Append(CLng(&HCCCCCC00))    ' 00 CC CC CC '
    
    Set Create_Helper_ShiftLeft = Runtime.CreateMethod("ShiftLeft", _
        Prototype.Factory.NewArg("Value", VT_In Or VT_Long) _
        .NewArg("Count", VT_In Or VT_Long) _
        .SetRetVal(Prototype.AsEAXValue Or VT_Long).Clone, _
        BA)
End Function
''<summary>Shifts the number to right.</summary>
''<params>
''  <param type="long" name="Value"></param>
''  <param type="long" name="Count"></param>
''</params>
Private Function Create_Helper_ShiftRight() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H824448B))     ' 8B 44 24 08 'Mov EAX, SS:[ESP+8]
    Call BA.Append(CLng(&HC244C8B))     ' 8B 4C 24 0C '
    Call BA.Append(CLng(&HCC2E8D3))     ' D3 E8 C2 0C '
    Call BA.Append(CLng(&HCCCCCC00))    ' 00 CC CC CC 'Ret
    
    Set Create_Helper_ShiftRight = Runtime.CreateMethod("ShiftRight", _
        Prototype.Factory.NewArg("Value", VT_In Or VT_Long) _
        .NewArg("Count", VT_In Or VT_Long) _
        .SetRetVal(Prototype.AsEAXValue Or VT_Long).Clone, _
        BA)
End Function
''<summary>Fill the given memory block with zeros.</summary>
''<params>
''  <param type="long" name="pDest"></param>
''  <param type="long" name="lngSize"></param>
''</params>
Private Function Create_Helper_MemZero() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H26C03257))    ' 57 32 C0 26 'PUSH EDI  'XOR Al, AL
    Call BA.Append(CLng(&H8247C8B))     ' 8B 7C 24 08 'MOV EDI, ES:[ESP+8]
    Call BA.Append(CLng(&HC244C8B))     ' 8B 4C 24 0C 'MOV ECX, SS:[ESP+C]
    Call BA.Append(CLng(&H5FAAF3FC))    ' FC F3 AA 5F 'CLD  'REP STOS BYTE PTR ES:[EDI]
    Call BA.Append(CLng(&H8C2))         ' C2 08 00 00 'POP EDI 'RET 8

    Set Create_Helper_MemZero = Runtime.CreateMethod("memzero", _
        Prototype.Factory.NewArg("pDest", VT_In Or VT_Long) _
        .NewArg("lngSize", VT_In Or VT_Long) _
        .SetRetVal(Prototype.AsEAXValue Or VT_Long).Clone, _
        BA)
End Function
''<summary>Reads the value of EBP register.</summary>
Private Function Create_Helper_ReadCalleeEBP() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&HC3C58B))    ' 8B C5 C3 00 'MOV EAX, EBP 'RET 0
    
    Set Create_Helper_ReadCalleeEBP = Runtime.CreateMethod("ReadCalleeEBP", _
        Prototype.Factory.SetRetVal(Prototype.AsEAXValue Or VT_Long).Clone, _
        BA)
End Function
''<summary>Writes the EBP register.</summary>
''<params>
''  <param type="long" name="vEBP"></param>
''</params>
Private Function Create_Helper_WriteCalleeEBP() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H8B08458B))  ' 8B 45 08 8B 'MOV EAX, [EBP+8]
    Call BA.Append(CLng(&H4C2E8))     ' E0 C2 04 00 'MOV EBP, EAX 'RET 4
    
    Set Create_Helper_WriteCalleeEBP = Runtime.CreateMethod("WriteCalleeEBP", _
        Prototype.Factory.NewArg("vEBP", VT_In Or VT_Long).Clone, _
        BA)
End Function
''<summary>Reads the value of EBP register of the caller function.</summary>
Private Function Create_Helper_ReadCallerEBP() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&HC300458B))  ' 8B 45 00 C3 'MOV EAX, dword ptr [EBP] 'RET
    
    Set Create_Helper_ReadCallerEBP = Runtime.CreateMethod("ReadCallerEBP", _
        Prototype.Factory.SetRetVal(Prototype.AsEAXValue Or VT_Long).Clone, _
        BA)
End Function
''<summary>Writes the EBP register of the caller function.</summary>
''<params>
''  <param type="long" name="vEBP"></param>
''</params>
Private Function Create_Helper_WriteCallerEBP() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H824448B))   ' 8B 44 24 08 'MOV EAX, dword ptr [ESP + 0x8]
    Call BA.Append(CLng(&HC2004589))  ' 89 45 00 C2 'MOV dword ptr [EBP], EAX
    Call BA.Append(CLng(&H4))         ' 04 00 00 00 'RET 4
    
    Set Create_Helper_WriteCallerEBP = Runtime.CreateMethod("WriteCallerEBP", _
        Prototype.Factory.NewArg("vEBP", VT_In Or VT_Long).Clone, _
        BA)
End Function
''<summary>Reads the value of ESP register.</summary>
Private Function Create_Helper_ReadESP() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&HC3C48B))    ' 8B C4 C3 00 'MOV EAX, ESP 'RET 0
    
    Set Create_Helper_ReadESP = Runtime.CreateMethod("ReadESP", _
        Prototype.Factory.SetRetVal(Prototype.AsEAXValue Or VT_Long).Clone, _
        BA)
End Function
''<summary>Writes the ESP register.</summary>
''<params>
''  <param type="long" name="vESP"></param>
''</params>
Private Function Create_Helper_WriteESP() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H8B08458B))  ' 8B 45 08 8B 'MOV EAX, [ESP + 8]
    Call BA.Append(CLng(&H4C2E0))     ' E0 C2 04 00 'MOV ESP, EAX 'RET 4
    
    Set Create_Helper_WriteESP = Runtime.CreateMethod("WriteESP", _
        Prototype.Factory.NewArg("vESP", VT_In Or VT_Long).Clone, _
        BA)
End Function
''<summary>Returns the current object pointer (same as Me).</summary>
Private Function Create_Helper_CalleeThis() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&HC308458B))  '8B 45 08 C3 'MOV EAX, dword ptr [EBP + 0x8] 'RET
    
    Set Create_Helper_CalleeThis = Runtime.CreateMethod("CalleeThis", _
        Prototype.Factory.SetRetVal(Prototype.AsEAXValue Or VT_Long).Clone, _
        BA)
End Function
''<summary>Returns the caller method's object pointer (same as Me in caller).</summary>
Private Function Create_Helper_CallerThis() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H8B00458B))  ' 8B 45 00 8B 'MOV EAX, dword ptr [EBP]
    Call BA.Append(CLng(&HC30840))    ' 40 08 C3 00 'MOV EAX, dword ptr [EAX + 0x8] 'RET
    
    Set Create_Helper_CallerThis = Runtime.CreateMethod("CallerThis", _
        Prototype.Factory.SetRetVal(Prototype.AsEAXValue Or VT_Long).Clone, _
        BA)
End Function
''<summary>Reserves the specified amount of stack.</summary>
''<params>
''  <param type="long" name="Length"></param>
''</params>
Private Function Create_Helper_Reserve() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H2B585859))  ' 59 58 58 2B 'POP ECX 'POP EAX 'POP EAX
    Call BA.Append(CLng(&HFFC033E0))  ' E0 33 C0 FF 'SUB ESP, EAX 'XOR EAX, EAX
    Call BA.Append(CLng(&HE1))        ' E1 00 00 00 'JMP ECX
    
    Set Create_Helper_Reserve = Runtime.CreateMethod("Reserve", _
        Prototype.Factory.NewArg("Length", VT_In Or VT_Long).Clone, _
        BA)
End Function
''<summary>Increaments a int32 value and then returns it.</summary>
''<params>
''  <param type="long" name="int32Var"></param>
''</params>
Private Function Create_Helper_IncVar32() As Method
    Dim BA As New ByteArray
    
    Set Create_Helper_IncVar32 = Runtime.CreateMethod("IncVar32", _
        Prototype.Factory.NewArg("int32Var", VT_In Or VT_Long) _
        .SetRetVal(Prototype.AsEAXValue Or VT_Long).Clone, _
        BA)
End Function
Private Function Create_Helper_GetIP() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H4C2))       ' C2 04 00 00 'RET 4
    
    Set Create_Helper_GetIP = Runtime.CreateMethod("GetIP", _
        Prototype.VoidMethod, _
        BA)
End Function

'=====================================================================================

''<summary>Simply return and all given parameters will remain onto stack (but it reclaim object pointer).</summary>
Private Function Create_Helper_Return() As Method
    Dim BA As New ByteArray
    Call BA.Append(CLng(&H4C2))       ' C2 04 00 00 'RET 4
    
    Set Create_Helper_Return = Runtime.CreateMethod("Return", _
        Prototype.VoidMethod, _
        BA)
End Function
Private Function Create_Helper_Call() As Method
    Dim BA As New ByteArray
'   pop         eax // remove return address from stack
'   pop         ecx // remove pointer to object from stack
'   pop         ecx // remove given forwarding address from stack
'   push        eax // push return address onto stack
'   jmp         ecx // jump to forwarding address
    Call BA.Append(CLng(&H50595958))  ' 58 59 59 50 'POP EAX 'POP ECX 'POP ECX 'PUSH EAX
    Call BA.Append(CLng(&HE1FF))      ' FF E1 00 00 'JMP ECX
    
    Set Create_Helper_Call = Runtime.CreateMethod("Call", _
        Prototype.Factory.NewArg("Reference", VT_Long).Clone, _
        BA)
End Function
Private Function Create_Helper_CallDouble() As Method
    Dim BA As New ByteArray
'   fstp        qword ptr [ecx] // return value gets double result
    Call BA.Append(pHelperS.CallProlog)
    Call BA.Append(CLng(&H19DD))      ' DD 19 00 00 'fstp qword ptr [ecx]
    Call BA.Append(pHelperS.CallEpilog)
    
    Set Create_Helper_CallDouble = Runtime.CreateMethod("CallDouble", _
        Prototype.Factory.NewArg("Reference", VT_Long).SetRetVal(VT_Double).Clone, _
        BA)
End Function
Private Function Create_Helper_CallInt32() As Method
    Dim BA As New ByteArray
'   mov         dword ptr [ecx],eax // copy first 32-bit of the return value
    Call BA.Append(pHelperS.CallProlog)
    Call BA.Append(CLng(&H189))       ' 89 01 00 00 'mov dword ptr [ecx],eax
    Call BA.Append(pHelperS.CallEpilog)
    
    Set Create_Helper_CallInt32 = Runtime.CreateMethod("CallInt32", _
        Prototype.Factory.NewArg("Reference", VT_Long).SetRetVal(VT_Int32).Clone, _
        BA)
End Function
Private Function Create_Helper_CallInt64() As Method
    Dim BA As New ByteArray
'   mov         dword ptr [ecx],eax // copy first 32-bit of the return value
'   mov         dword ptr [ecx+4],edx // copy second 32-bit of the return value
    Call BA.Append(pHelperS.CallProlog)
    Call BA.Append(CLng(&H51890189))  ' 89 01 89 51
    Call BA.Append(CByte(&H4))        ' 04 -- -- --
    Call BA.Append(pHelperS.CallEpilog)
    
    Set Create_Helper_CallInt64 = Runtime.CreateMethod("CallInt64", _
        Prototype.Factory.NewArg("Reference", VT_Long).SetRetVal(VT_Int64).Clone, _
        BA)
End Function

Private Function Create_CallProlog(ByVal Lib As Library) As ByteArray
    Dim BA As New ByteArray
'   push        0Ch // we need 12 bytes allocated
'   push        0 // we need fixed memory (GMEM_FIXED)
'   mov         eax,770001h // replace here 00770001 with the address of GlobalAlloc
'   call        eax // allocate new list node
'   pop         ecx // remove return address from stack
'   mov         dword ptr [eax],ecx // store return address in list node
'   pop         ecx // remove pointer to object from stack
'   pop         ebx // remove given forwarding address from stack
'   pop         ecx // get the location where return value is expected by VB caller
'   mov         dword ptr [eax+4],ecx // store location of return value in list node
'   mov         ecx,dword ptr fs:[18h] // get pointer to TIB structure
'   mov         edx,dword ptr [ecx+14h] // get pointer to previous list node from pvArbitrary
'   mov         dword ptr [eax+8],edx // link list nodes
'   mov         dword ptr [ecx+14h],eax // store new head of list at pvArbitrary
'   call        ebx // call the forwarding address
'   mov         ecx,dword ptr fs:[18h] // get pointer to TIB structure
'   push        ecx // save pointer to TIB structure
'   mov         ebx,dword ptr [ecx+14h] // get pointer to head of list from pvArbitrary
'   mov         ecx,dword ptr [ebx+4] // get location of return value from list node
    BA.Capacity = 54
    Call BA.Append(CLng(&H6A0C6A))    ' 6A 0C 6A 00 '
    Call BA.Append(CByte(&HB8))       ' B8 -- -- -- '
    Call BA.Append(CLng(Lib.LoadSymbol(KERNEL32_GLOBAL_ALLOC)))
    Call BA.Append(CByte(&HFF))
    Call BA.Append(CInt(&H59D0))      ' -- FF D0 59 '
    Call BA.Append(CLng(&H5B590889))  ' 89 08 59 5B '
    Call BA.Append(CLng(&H4488959))   ' 59 89 48 04 '
    Call BA.Append(CLng(&H180D8B64))  ' 64 8B 0D 18 '
    Call BA.Append(CLng(&H8B000000))  ' 00 00 00 8B '
    Call BA.Append(CLng(&H50891451))  ' 51 14 89 50 '
    Call BA.Append(CLng(&H14418908))  ' 08 89 41 14 '
    Call BA.Append(CLng(&H8B64D3FF))  ' FF D3 64 8B '
    Call BA.Append(CLng(&H180D))      ' 0D 18 00 00 '
    Call BA.Append(CLng(&H598B5100))  ' 00 51 8B 59 '
    Call BA.Append(CLng(&H44B8B14))   ' 14 8B 4B 04 '
    
    Set Create_CallProlog = BA
End Function
Private Function Create_CallEpilog(ByVal Lib As Library) As ByteArray
    Dim BA As New ByteArray
'   pop         ecx // get back the pointer to TIB structure
'   mov         edx,dword ptr [ebx] // get return address from list node
'   push        edx // restore return address onto stack
'   mov         edx,dword ptr [ebx+8] // get pointer to previous list node
'   mov         dword ptr [ecx+14h],edx // store pointer to previous node at pvArbitrary
'   push        ebx // we need to free the list node
'   mov         eax,770002h // replace here 00770002 with the address of GlobalFree
'   call        eax // free list node
'   ret // return to VB caller
    Call BA.Append(CLng(&H52138B59))  ' 59 8B 13 52 '
    Call BA.Append(CLng(&H8908538B))  ' 8B 53 08 89 '
    Call BA.Append(CLng(&HB8531451))  ' 51 14 53 B8 '
    Call BA.Append(CLng(Lib.LoadSymbol(KERNEL32_GLOBAL_FREE))) ' -- -- -- -- '
    Call BA.Append(CLng(&HC3D0FF))    ' FF D0 C3 00 '
    
    Set Create_CallEpilog = BA
End Function

