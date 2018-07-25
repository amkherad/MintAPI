VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemoryStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="MemoryStatic" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

'##################@@@@@@@@@@@@@@@@@@ Function replacement - becarefull!

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "MemoryStatic" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const pFUNC_ORDER_MemoryMarshal = FUNC_ORDER_SKIP_IDISPATCH_B1 + 1

Const pFUNC_ORDER_CopyMemory As Long = pFUNC_ORDER_MemoryMarshal + 1
Const pFUNC_ORDER_ZeroMemory As Long = pFUNC_ORDER_MemoryMarshal + 2
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Public Enum MemoryAllocationFlags
    mafDefault = 0
    mafResizable
    mafMovable
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

Public Sub Copy(ByVal DestinationAddres As Long, ByVal SourceAddress As Long, ByVal Size As Long)
    If SourceAddress = 0 Then throw Exps.AccessDeniedException
    If DestinationAddres = 0 Then throw Exps.AccessDeniedException
    If Size = 0 Then Exit Sub
    
    Call memcpy(ByVal DestinationAddres, ByVal SourceAddress, Size)
End Sub
Public Sub ZeroMemory(ByVal SourceAddress As Long, ByVal Size As Long)
    If SourceAddress = 0 Then throw Exps.AccessDeniedException
    If Size = 0 Then Exit Sub
    
    Call memzero(ByVal SourceAddress, Size)
End Sub

Public Function Delegate(ByVal Reference As Long, ByVal Limit As Long, _
        Optional ByVal Access As MemoryAccessPermissions = mapDefault) As IMemoryBlock
    Dim Dlg As New MemoryDelegate
    Call Dlg.Constructor0(Reference, Limit, Access)
    Set Delegate = Dlg
End Function

'Public Function FromReference(ByVal Reference As Long, ByVal Limit As Long, _
'        Optional ByVal Access As MemoryAccessPermissions = mapDefault) As IMemoryBlock
'    Dim Mem As New MemoryDelegate
'    Call Mem.Constructor0(Reference, Limit, Access)
'    Set FromReference = Mem
'End Function

''<summary>Allocates some memory in current proccess.</summary>
''<see>Memory</see>
''<params><param name="Size">The size of the allocated memory.</param></params>
''<default/>
Public Function Allocate(ByVal Size As Long, Optional ByVal Flags As MemoryAllocationFlags = mafDefault) As IMemoryBlock
Attribute Allocate.VB_UserMemId = 0
    If Flags = mafDefault Then
        Set Allocate = ProcessHeap.Allocate(Size)
    Else
        '### TEMPERORY!!!
        Set Allocate = ProcessHeap.Allocate(Size)
    End If
End Function

'Public Function LightWeightAllocate(ByVal Size As Long) As Long
'    LightWeightAllocate = API_CoTaskMemAlloc(Size)
'    If LightWeightAllocate = vbNullPtr Then _
'        throw Exps.OutOfMemoryException
'End Function

Public Function FastAllocate(ByVal Size As Long) As Long
    FastAllocate = API_CoTaskMemAlloc(Size)
    If FastAllocate = vbNullPtr Then _
        throw Exps.OutOfMemoryException
End Function
Public Sub FastFree(ByVal Ptr As Long)
    'If API_HeapFree(ProcessHeap, 0, Ptr) <> NO_ERROR Then _
        throw Exps.IfError
End Sub

Public Function CreateMemory(ByRef Data As Variant, Optional ByVal ConvertToBinary As Boolean = True) As IMemoryBlock
    
End Function

Public Function MemoryOf(ByRef Data As Variant) As IMemoryBlock
    
End Function

Public Function CreateSharedMemory(ByVal SharedMemoryName As String, ByVal Size As Long) As IMemoryBlock
    
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

Public Property Get ProcessHeap() As Heap
    Dim pHeap As Long
    pHeap = API_GetProcessHeap
    If pHeap = NO_VALUE Then throw Exps.IfError
    Set ProcessHeap = New Heap
    Call ProcessHeap.Constructor0(pHeap, True)
End Property
Public Function CreateHeap( _
            ByVal CreationFlags As HeapFlags, _
            ByVal InitialSize As Long, _
            ByVal MaximumSize As Long) As Heap
    Dim Flags As Long
    If (CreationFlags And hfNoSerialize) = hfNoSerialize Then _
        Flags = Flags Or HeapFlags.hfNoSerialize
    Dim Result As Long
    Result = API_HeapCreate(Flags, InitialSize, MaximumSize)
    If Result = vbNullPtr Then _
        throw Exps.IfError("CreateHeap")
    Set CreateHeap = New Heap
    Call CreateHeap.Constructor0(Result, (MaximumSize = 0))
End Function

'Public Sub Deallocate(ByRef Mem As IMemoryBlock)
'    If Mem Is Nothing Then throw Exps.ArgumentNullException
'    Call Mem.Mem
'    Set Mem = Nothing
'End Sub
Public Sub Free(ByVal Ptr As Long)
    If API_HeapFree(ProcessHeap, 0, Ptr) <> NO_VALUE Then _
        throw Exps.IfError
End Sub

Public Function SizeOf(ByRef MemoryBlock As Variant) As Long
    If IsObject(MemoryBlock) Then
        If Not TypeOf MemoryBlock Is IMemoryBlock Then throw Exps.InvalidOperationException
        Dim Mem As IMemoryBlock
        Set Mem = MemoryBlock
        SizeOf = Mem.Limit
    Else
        SizeOf = API_HeapSize(ProcessHeap, 0, CLng(MemoryBlock))
    End If
End Function

Public Function CoTaskMemAlloc(ByVal Size As Long, Optional ByVal ThrowIfOutOfMemory As Boolean = True) As Long
    CoTaskMemAlloc = API_CoTaskMemAlloc(Size)
    If ThrowIfOutOfMemory Then _
        If CoTaskMemAlloc = vbNullPtr Then throw Exps.OutOfMemoryException
End Function
Public Sub CoTaskMemFree(ByVal PV As Long)
    Call API_CoTaskMemFree(PV)
End Sub

Public Function CanAccessTo(ByVal Ptr As Long, Optional ByVal Length As Long = 0, Optional ByVal Permission As MemoryAccessPermissions = mapReadWrite) As Boolean
    CanAccessTo = True
End Function
Public Function GetAccessPermission(ByVal MemoryBlock As IMemoryBlock, Optional ByVal TrustMemoryBlockInternalInfo As Boolean = True) As MemoryAccessPermissions
    If MemoryBlock Is Nothing Then throw Exps.ArgumentNullException("MemoryBlock")
    
    If TrustMemoryBlockInternalInfo Then
        GetAccessPermission = MemoryBlock.Access
    Else
        'if VirtualMemory.Query(nothing).Protect =
    End If
End Function

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

Public Sub CopyBlock(ByVal Destination As IMemoryBlock, ByVal Source As IMemoryBlock, ByVal Size As Long)
    If Destination Is Nothing Then throw Exps.ArgumentNullException("Destination")
    If Source Is Nothing Then throw Exps.ArgumentNullException("Source")
    If Size = 0 Then Exit Sub
    
    Dim Src As Long, Dest As Long
    Src = Source.Reference
    Dest = Destination.Reference
    If Src = 0 Then throw Exps.AccessDeniedException
    If Dest = 0 Then throw Exps.AccessDeniedException
    
    Call memcpy(ByVal Destination.Reference, ByVal Source.Reference, Size)
End Sub

Public Function DeRef(ByVal Target As Long) As Long
    Call memcpy(DeRef, ByVal Target, VLEN_PTR)
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------


Private Sub Class_Initialize()
    Dim VTbl As Long
    Call memcpy(VTbl, ByVal ObjPtr(Me), VLEN_PTR)
    Call memcpy(ByVal (VTbl + (pFUNC_ORDER_CopyMemory * VLEN_PTR)), Library.Kernel32.LoadSymbol("CopyMemory"), VLEN_PTR)
    Call memcpy(ByVal (VTbl + (pFUNC_ORDER_ZeroMemory * VLEN_PTR)), Library.Kernel32.LoadSymbol("ZeroMemory"), VLEN_PTR)
End Sub
