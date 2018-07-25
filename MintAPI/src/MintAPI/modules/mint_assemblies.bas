Attribute VB_Name = "mint_assemblies"
Option Explicit

Private Const NUMBEROF_METHODS As Long = 40
Private Const ASSEMBLYLENGTH_TOTALLONGS As Long = 70
Private Const ASSEMBLYLENGTH_TOTALBYTES As Long = ASSEMBLYLENGTH_TOTALLONGS * VLEN_LONG
Private Const METHOD_GAP As Long = 3

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

Public Const FUNC_ORDER_PushL As Long = FUNC_ORDER_FIRSTMETHOD + 12
Public Const FUNC_ORDER_PopL As Long = FUNC_ORDER_FIRSTMETHOD + 13

Public Const FUNC_ORDER_PushD As Long = FUNC_ORDER_FIRSTMETHOD + 14
Public Const FUNC_ORDER_PopD As Long = FUNC_ORDER_FIRSTMETHOD + 15

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

Public Type HelperFunctionMetaData
    FuncLength As Long
    Name As String
    'ArgumentLength As Long
    'ReturnLength As Long
    FuncPtr As Long
    Method As Method
End Type
Private Type HelperStructure
    pVTable As Long
    cRefs As Long
    bTop As Long
    
    Funcs(NUMBEROF_METHODS - 1) As Method
    FuncTable(NUMBEROF_METHODS - 1) As Long
End Type

Public mHelper As IMintHelper

Private pHelperS As HelperStructure
Private pASM As Long
Private bASM() As Byte, SAB1 As SafeArray1d
Private lASM() As Long, SAL1 As SafeArray1d
Private fCall_Prolog() As Long


Public Sub Construct()
    Call Init_LoadedMethods32

'===================================================================
' Initializing memory and arrays.
'===================================================================
    
    pASM = API_CoTaskMemAlloc(ASSEMBLYLENGTH_TOTALBYTES)
    If pASM = vbNullPtr Then throw Exps.OutOfMemoryException
    
    Call Init_Arrays(pASM) '** Initializes the arrays.
    
'===================================================================
'===================================================================
    Call Init_ClassMethods32
    Call Init_InternalMethods32
    
    Call Destroy_Arrays '** Destroys the arrays.
    
    Dim outOldProtect As Long
    If API_VirtualProtect(pASM, ASSEMBLYLENGTH_TOTALBYTES, PAGE_EXECUTE_READWRITE, outOldProtect) = NO_VALUE Then _
        throw Exps.IfError
    
    pHelperS.cRefs = 1
    pHelperS.pVTable = VarPtr(pHelperS.FuncTable(0))
    
    Call memcpy(mHelper, VarPtr(pHelperS.pVTable), VLEN_PTR)
End Sub
Private Sub Init_Arrays(ByVal pASM As Long)
    SAL1.cDims = 1: SAB1.cDims = 1
    SAL1.cLocks = 0: SAB1.cLocks = 0
    SAL1.pvData = pASM: SAB1.pvData = pASM
    SAL1.cbElements = 4: SAB1.cbElements = 1
    SAL1.lLbound = 0: SAB1.lLbound = 0
    SAL1.cElements = ASSEMBLYLENGTH_TOTALLONGS
    SAB1.cElements = ASSEMBLYLENGTH_TOTALBYTES
    SAL1.fFeatures = FADF_STATIC Or FADF_FIXEDSIZE
    SAB1.fFeatures = SAL1.fFeatures
    
    Call memcpy(ByVal API_VarPtrArray(lASM), VarPtr(SAL1.cDims), VLEN_PTR)
    Call memcpy(ByVal API_VarPtrArray(bASM), VarPtr(SAB1.cDims), VLEN_PTR)
    'Call memzero(VarPtr(SAL1), VLEN_PTR)
    'Call memzero(VarPtr(SAB1), VLEN_PTR)
End Sub
Private Sub Destroy_Arrays()
    Call API_CoTaskMemFree(VarPtr(SAL1))
    Call API_CoTaskMemFree(VarPtr(SAB1))
    
    Call memzero(ByVal API_VarPtrArray(lASM), VLEN_PTR)
    Call memzero(ByVal API_VarPtrArray(bASM), VLEN_PTR)
End Sub


Private Sub Init_LoadedMethods32()
    'Dim bASM(120) As Byte
    'Load optimized functions into dummy functions.
End Sub

Private Sub Init_ClassMethods32()
    Dim M As Method
    Call memzero(ByVal pASM, ASSEMBLYLENGTH_TOTALBYTES)
'=====================================================================
' 00 - 0000: long ReadFS()
'   - Returns FS segment register.
' Argument length   : 0
' Return length     : 4B
' Function length   : 8B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_READFS) = VarPtr(lASM(0))
    pHelperS.FuncMetaData(FUNC_ORDER_READFS).FuncLength = 8 'B
    pHelperS.FuncMetaData(FUNC_ORDER_READFS).Name = "ReadFS"
    lASM(0) = &H18A164      ' 64 A1 18 00 'MOV EAX, FS:[18]
    lASM(1) = &HC30000      ' 00 00 C3 00 'RET 0
    'lASM(2) = vbNullPtr     ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 01 - 0008: void WriteFS([in] long fsValue)
'   - Writes fsValue into FS segment register.
' Argument length   : 4B
' Return length     : 0
' Function length   : 12B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_WRITEFS) = VarPtr(lASM(3))
    pHelperS.FuncMetaData(FUNC_ORDER_WRITEFS).FuncLength = 12 'B
    pHelperS.FuncMetaData(FUNC_ORDER_WRITEFS).Name = "WriteFS"
    lASM(3) = &H424448B     ' 8B 44 24 08 'Mov EAX, SS:[ESP+8]
    lASM(4) = &H18A364      ' 64 A3 18 00 'Mov FS:[18], EAX
    lASM(5) = &H4C20000     ' 00 00 C2 04 'Ret 4
    'lASM(6) = vbNullPtr     ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 02 - 0028: long ShiftLeft([in] long Value, [in] long Count)
'   - Shift the number to left.
' Argument length   : 8B
' Return length     : 4B
' Function length   : 16B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_SHIFTLEFT) = VarPtr(lASM(7))
    pHelperS.FuncMetaData(FUNC_ORDER_SHIFTLEFT).FuncLength = 16 'B
    pHelperS.FuncMetaData(FUNC_ORDER_SHIFTLEFT).Name = "ShiftLeft"
    lASM(7) = &H824448B     ' 8B 44 24 08 'Mov EAX, SS:[ESP+8]
    lASM(8) = &HC244C8B     ' 8B 4C 24 0C '
    lASM(9) = &HCC2E0D3     ' D3 E0 C2 0C '
    lASM(10) = &HCCCCCC00   ' 00 CC CC CC '
    'lASM(11) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 03 - 0048: long ShiftRight([in] long Value, [in] long Count)
'   - Shift the number to left.
' Argument length   : 8B
' Return length     : 4B
' Function length   : 16B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_SHIFTRIGHT) = VarPtr(lASM(12))
    pHelperS.FuncMetaData(FUNC_ORDER_SHIFTRIGHT).FuncLength = 16 'B
    pHelperS.FuncMetaData(FUNC_ORDER_SHIFTRIGHT).Name = "ShiftRight"
    lASM(12) = &H824448B    ' 8B 44 24 08 'Mov EAX, SS:[ESP+8]
    lASM(13) = &HC244C8B    ' 8B 4C 24 0C '
    lASM(14) = &HCC2E8D3    ' D3 E8 C2 0C '
    lASM(15) = &HCCCCCC00   ' 00 CC CC CC 'Ret
    'lASM(16) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 04 - 0068: long memzero([in] long pDest, [in] long lngSize)
'   - Fill the given memory block with zeros.
' Argument length   : 8B
' Return length     : 4B
' Function length   : 20B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_MEMZERO) = VarPtr(lASM(17))
    pHelperS.FuncMetaData(FUNC_ORDER_MEMZERO).FuncLength = 20 'B
    pHelperS.FuncMetaData(FUNC_ORDER_MEMZERO).Name = "memzero"
    lASM(17) = &H26C03257   ' 57 32 C0 26 'PUSH EDI  'XOR AL, AL
    lASM(18) = &H8247C8B    ' 8B 7C 24 08 'MOV EDI, ES:[ESP+8]
    lASM(19) = &HC244C8B    ' 8B 4C 24 0C 'MOV ECX, SS:[ESP+C]
    lASM(20) = &H5FAAF3FC   ' FC F3 AA 5F 'CLD  'REP STOS BYTE PTR ES:[EDI]
    lASM(20) = &H8C2        ' C2 08 00 00 'POP EDI 'RET 8
    'lASM(21) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 05 - 0088: long ReadCalleeEBP()
'   - Reads the value of EBP register.
' Argument length   : 0B
' Return length     : 4B
' Function length   : 4B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_READCALLEEEBP) = VarPtr(lASM(22))
    pHelperS.FuncMetaData(FUNC_ORDER_READCALLEEEBP).FuncLength = 4 'B
    pHelperS.FuncMetaData(FUNC_ORDER_READCALLEEEBP).Name = "ReadCalleeEBP"
    lASM(22) = &HC3C58B       ' 8B C5 C3 00 'MOV EAX, EBP 'RET 0
    'lASM(23) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 06 - 0096: void WriteCalleeEBP([in] long vEBP)
'   - Writes the EBP register.
' Argument length   : 4B
' Return length     : 0B
' Function length   : 8B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_WRITECALLEEEBP) = VarPtr(lASM(24))
    pHelperS.FuncMetaData(FUNC_ORDER_WRITECALLEEEBP).FuncLength = 8 'B
    pHelperS.FuncMetaData(FUNC_ORDER_WRITECALLEEEBP).Name = "WriteCalleeEBP"
    lASM(24) = &H8B08458B     ' 8B 45 08 8B 'MOV EAX, [EBP+8]
    lASM(25) = &H4C2E8        ' E0 C2 04 00 'MOV EBP, EAX 'RET 4
    'lASM(26) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 05 - 0088: long ReadCallerEBP()
'   - Reads the value of EBP register of the caller function.
' Argument length   : 0B
' Return length     : 4B
' Function length   : 4B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_READCALLEREBP) = VarPtr(lASM(27))
    pHelperS.FuncMetaData(FUNC_ORDER_READCALLEREBP).FuncLength = 4 'B
    pHelperS.FuncMetaData(FUNC_ORDER_READCALLEREBP).Name = "ReadCallerEBP"
    lASM(27) = &HC300458B     ' 8B 45 00 C3 'MOV EAX, dword ptr [EBP] 'RET
    'lASM(28) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 06 - 0096: void WriteCallerEBP([in] long vEBP)
'   - Writes the EBP register of the caller function.
' Argument length   : 4B
' Return length     : 0B
' Function length   : 8B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_WRITECALLEREBP) = VarPtr(lASM(29))
    pHelperS.FuncMetaData(FUNC_ORDER_WRITECALLEREBP).FuncLength = 12 'B
    pHelperS.FuncMetaData(FUNC_ORDER_WRITECALLEREBP).Name = "WriteCallerEBP"
    lASM(29) = &H824448B      ' 8B 44 24 08 'MOV EAX, dword ptr [ESP + 0x8]
    lASM(30) = &HC2004589     ' 89 45 00 C2 'MOV dword ptr [EBP], EAX
    lASM(31) = &H4            ' 04 00 00 00 'RET 4
    'lASM(32) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 07 - 0108: long ReadESP()
'   - Reads the value of ESP register.
' Argument length   : 0B
' Return length     : 4B
' Function length   : 4B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_READESP) = VarPtr(lASM(33))
    pHelperS.FuncMetaData(FUNC_ORDER_READESP).FuncLength = 4 'B
    pHelperS.FuncMetaData(FUNC_ORDER_READESP).Name = "ReadESP"
    lASM(33) = &HC3C48B       ' 8B C4 C3 00 'MOV EAX, ESP 'RET 0
    'lASM(34) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 08 - 0116: void WriteESP([in] long vESP)
'   - Writes the ESP register.
' Argument length   : 4B
' Return length     : 0B
' Function length   : 8B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_WRITEESP) = VarPtr(lASM(35))
    pHelperS.FuncMetaData(FUNC_ORDER_WRITEESP).FuncLength = 8 'B
    pHelperS.FuncMetaData(FUNC_ORDER_WRITEESP).Name = "WriteESP"
    lASM(35) = &H8B08458B     ' 8B 45 08 8B 'MOV EAX, [ESP + 8]
    lASM(36) = &H4C2E0        ' E0 C2 04 00 'MOV ESP, EAX 'RET 4
    'lASM(37) = vbNullPtr    ' Ext
'---------------------------------------------------------------------

'---------------------------------------------------------------------
    Set M = Create_Helper_ReadEAX
    pHelperS.FuncTable(FUNC_ORDER_READEAX) = M.Reference
    Set pHelperS.Funcs(FUNC_ORDER_READEAX) = M
'---------------------------------------------------------------------

'=====================================================================
' 10 - 0128: object CalleeThis()
'   - Returns the current object pointer (same as Me).
' Argument length   : 0B
' Return length     : 4B
' Function length   : 4B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_CALLEETHIS) = VarPtr(lASM(40))
    pHelperS.FuncMetaData(FUNC_ORDER_CALLEETHIS).FuncLength = 4  'B
    pHelperS.FuncMetaData(FUNC_ORDER_CALLEETHIS).Name = "CalleeThis"
    lASM(40) = &HC308458B     '8B 45 08 C3 'MOV EAX, dword ptr [EBP + 0x8] 'RET
    'lASM(41) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 10 - 0128: object CallerThis()
'   - Returns the caller method's object pointer (same as Me in caller).
' Argument length   : 0B
' Return length     : 4B
' Function length   : 8B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_CALLERTHIS) = VarPtr(lASM(42))
    pHelperS.FuncMetaData(FUNC_ORDER_CALLERTHIS).FuncLength = 8  'B
    pHelperS.FuncMetaData(FUNC_ORDER_CALLERTHIS).Name = "CallerThis"
    lASM(42) = &H8B00458B     ' 8B 45 00 8B 'MOV EAX, dword ptr [EBP]
    lASM(43) = &HC30840       ' 40 08 C3 00 'MOV EAX, dword ptr [EAX + 0x8] 'RET
    'lASM(44) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 11 - 0128: void Return()
'   - Simply return and all given parameters will remain onto stack (but it reclaim object pointer).
' Argument length   : 0B
' Return length     : 4B
' Function length   : 4B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_RETURN) = VarPtr(lASM(45))
    pHelperS.FuncMetaData(FUNC_ORDER_RETURN).FuncLength = 8  'B
    pHelperS.FuncMetaData(FUNC_ORDER_RETURN).Name = "Return"
    lASM(45) = &H4C2         ' C2 04 00 00 'RET 4
    'lASM(46) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 12 - 0128: void Reserve(ByVal Length As Long)
'   - Reserves the specified amount of stack.
' Argument length   : 4B
' Return length     : 4B
' Function length   : 12B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_RESERVE) = VarPtr(lASM(47))
    pHelperS.FuncMetaData(FUNC_ORDER_RESERVE).FuncLength = 12  'B
    pHelperS.FuncMetaData(FUNC_ORDER_RESERVE).Name = "Reserve"
    lASM(47) = &H2B585859    ' 59 58 58 2B 'POP ECX 'POP EAX 'POP EAX
    lASM(48) = &HFFC033E0    ' E0 33 C0 FF 'SUB ESP, EAX 'XOR EAX, EAX
    lASM(49) = &HE1          ' E1 00 00 00 'JMP ECX
    'lASM(51) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 12 - 0128: void Call(ByVal FuncPtr As Long)
'   - Calls a forwarding function.
' Argument length   : 4B
' Return length     : 0B
' Function length   : 8B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_CALL) = VarPtr(lASM(52))
    pHelperS.FuncMetaData(FUNC_ORDER_CALL).FuncLength = 8  'B
    pHelperS.FuncMetaData(FUNC_ORDER_CALL).Name = "Call"
    lASM(52) = &H50595958    ' 58 59 59 50 'POP EAX 'POP ECX 'POP ECX 'PUSH EAX
    lASM(53) = &HE1FF        ' FF E1 00 00 'JMP ECX
    'lASM(54) = vbNullPtr    ' Ext
'---------------------------------------------------------------------
'=====================================================================
' 12 - 0128: long IncVar32(ByVal int32Var As Long)
'   - Increaments a int32 value and then returns it.
' Argument length   : 4B
' Return length     : 4B
' Function length   : 8B
'=====================================================================
    pHelperS.FuncTable(FUNC_ORDER_INCVAR32) = VarPtr(lASM(52))
    pHelperS.FuncMetaData(FUNC_ORDER_INCVAR32).FuncLength = 8  'B
    pHelperS.FuncMetaData(FUNC_ORDER_INCVAR32).Name = "Call"
    lASM(52) = &H50595958    ' 58 59 59 50 'POP EAX 'POP ECX 'POP ECX 'PUSH EAX
    lASM(53) = &HE1FF        ' FF E1 00 00 'JMP ECX
    'lASM(54) = vbNullPtr    ' Ext
'---------------------------------------------------------------------

    Dim i As Long, Ptr As Long
    For i = 0 To NUMBEROF_METHODS - 1
        Ptr = pHelperS.FuncTable(i)
        pHelperS.FuncMetaData(i).FuncPtr = Ptr
        If Method.IsExecutable(Ptr) Then
            Set pHelperS.FuncMetaData(i).Method = _
                Method.FromReference(pHelperS.FuncMetaData(i).Name, Ptr)
        End If
    Next
End Sub '

''<summary>Reads the value of EAX register.</summary>
Private Function Create_Helper_ReadEAX() As Method
    Set Create_Helper_ReadEAX = Runtime.CreateMethod("ReadEAX", Nothing, _
        ByteArray(CLng(&HC3)))      'RET
End Function


Public Function GetMintAssembliesMethodGlobalInfo(ByVal MethodIndex As Long) As HelperFunctionMetaData
    GetMintAssembliesMethodGlobalInfo = pHelperS.FuncMetaData(MethodIndex)
End Function

Public Sub LoadFunctionIntoMemory(ByVal FuncPtr As Long, ByVal MemPtr As Long, ByVal FuncLen As Long)
    '
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
    Dim Gu As Guid
    On Error GoTo CatchError
    Set Gu = Guid.FromMemory(Memory.FromReference(riid, SIZEOF_GUID))
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
    Exit Function
CatchError:
    IMintHelper_IUnknown_QueryInterface = ERROR_E_NOINTERFACE
End Function
Private Function IMintHelper_IUnknown_AddRef(ByRef This As HelperStructure) As Long
    This.cRefs = This.cRefs + 1
    IMintHelper_IUnknown_AddRef = This.cRefs
End Function
Private Function IMintHelper_IUnknown_Release(ByRef This As HelperStructure) As Long
    This.cRefs = This.cRefs - 1
    IMintHelper_IUnknown_Release = This.cRefs
    
    If This.cRefs <= 0 Then
        Call API_CoTaskMemFree(VarPtr(This))
        
        Call Destroy_Arrays
        Call API_CoTaskMemFree(pASM)
        
        ObjectPtr(mHelper) = vbNullPtr
    End If
End Function

Private Function IMintHelper_IDispatch_Invoke(ByVal dispIdMember As Long, ByVal riid As Long, ByVal LCID As Long, ByVal wFlags As Long, pDispParams As API_DISPPARAMS, pVarResult, pExcepInfo As API_EXCEPINFO, ByRef puArgErr As Long) As Long 'HRESULT
    
End Function
