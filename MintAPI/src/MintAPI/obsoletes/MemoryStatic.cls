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
Option Explicit

Public Function FromHandle(ByVal Handle As Long, Optional ByVal Limit As Long = 4) As Memory
    Dim Mem As New Memory
    Call Mem.Constructor2(Handle, Limit)
    Set FromHandle = Mem
End Function

Public Function CreateMemory(ByRef Data As Variant, ByVal Access As MemoryAccess, Optional ByVal ConvertToBinary As Boolean = True) As Memory
    Dim Mem As New Memory
    If ConvertToBinary Then
        Call Mem.Constructor5(Data)
    Else
        Call Mem.Constructor6(Data)
    End If
    Set CreateMemory = Mem
End Function

Public Function MemoryOf(ByRef Data As Variant, ByVal Access As MemoryAccess) As Memory
    Dim Mem As New Memory
    Call Mem.Constructor3(Data)
    Set MemoryOf = Mem
End Function

Public Function CreateSharedMemory(ByVal SharedMemoryName As String) As Memory
    
End Function

''<summary>Allocates some memory in current proccess.</summary>
''<see>Memory</see>
''<params><param name="Size">The size of the allocated memory.</param></params>
Public Function Allocate(ByVal Size As Long) As Memory
    Dim Mem As New Memory
    Call Mem.Constructor0(Size)
    Set Allocate = Mem
End Function

Public Sub Deallocate(ByRef Mem As Memory) 'Must Be ByRef !
    If Mem Is Nothing Then throw Exps.ArgumentNullException
    Call Mem.Dispose
    Set Mem = Nothing
End Sub
Public Function CoTaskMemAlloc(ByVal Size As Long, Optional ByVal ThrowIfOutOfMemory As Boolean = True) As Long
    CoTaskMemAlloc = API_CoTaskMemAlloc(Size)
    If ThrowIfOutOfMemory Then throw Exps.OutOfMemoryException
End Function
Public Sub CoTaskMemFree(ByVal PV As Long)
    Call API_CoTaskMemFree(PV)
End Sub

Public Function SetDaisyChainNext(ByVal Mem As Memory, ByVal DCNext As Memory, Optional ByVal RetSource As Boolean = False) As Memory
    Set Mem.DCNext = DCNext
    Set SetDaisyChainNext = IIf(RetSource, Mem, DCNext)
End Function
Public Function SetDaisyChainPrevious(ByVal Mem As Memory, ByVal DCPrev As Memory, Optional ByVal RetSource As Boolean = False) As Memory
    Set Mem.DCPrevious = DCPrev
    Set SetDaisyChainPrevious = IIf(RetSource, Mem, DCPrev)
End Function
Public Function LinkDaisyChainNext(ByVal Mem As Memory, ByVal DCNext As Memory, Optional ByVal RetSource As Boolean = False) As Memory
    Set Mem.DCNext = DCNext
    Set DCNext.DCPrevious = Mem
    Set LinkDaisyChainNext = IIf(RetSource, Mem, DCNext)
End Function
Public Function LinkDaisyChainPrevious(ByVal Mem As Memory, ByVal DCPrev As Memory, Optional ByVal RetSource As Boolean = False) As Memory
    Set Mem.DCPrevious = DCPrev
    Set DCPrev.DCNext = Mem
    Set LinkDaisyChainPrevious = IIf(RetSource, Mem, DCPrev)
End Function



Public Function FromString(ByVal Value As String) As Memory
    
End Function
Public Function FromLong(ByVal Value As Long) As Memory
    
End Function
Public Function FromInteger(ByVal Value As Long) As Memory
    
End Function
Public Function FromPointer(ByVal Value As Long) As Memory
    
End Function
Public Function FromMethod(ByVal Value As Method) As Memory
    
End Function
Public Function FromArrayGeneric(ByVal Value As ArrayGeneric) As Memory
    
End Function
Public Function FromObject(ByVal Value As Object) As Memory
    
End Function
Public Function FromIObject(ByVal Value As IObject) As Memory
    
End Function
Public Function FromCurrency(ByVal Value As Currency) As Memory
    
End Function
Public Function FromDate(ByVal Value As Date) As Memory
    
End Function
Public Function FromByte(ByVal Value As Byte) As Memory
    
End Function
Public Function FromSingle(ByVal Value As Single) As Memory
    
End Function
Public Function FromDouble(ByVal Value As Double) As Memory
    
End Function


Public Property Get MemoryPtr(ByVal Address As Long) As Long
    Call memcpy(ByVal VarPtr(MemoryPtr), ByVal Address, VLEN_PTR)
End Property
Public Property Let MemoryPtr(ByVal Address As Long, ByVal Value As Long)
    Call memcpy(ByVal Address, ByVal VarPtr(Value), VLEN_PTR)
End Property
Public Property Get MemoryLong(ByVal Address As Long) As Long
    Call memcpy(ByVal VarPtr(MemoryLong), ByVal Address, VLEN_PTR)
End Property
Public Property Let MemoryLong(ByVal Address As Long, ByVal Value As Long)
    Call memcpy(ByVal Address, ByVal VarPtr(Value), VLEN_PTR)
End Property
Public Sub CopyMemory(ByVal DestinationAddres As Long, ByVal SourceAddress As Long, ByVal Length As Long)
    Call memcpy(ByVal DestinationAddres, ByVal SourceAddress, Length)
End Sub
Public Sub ZeroMemory(ByVal SourceAddress As Long, ByVal Length As Long)
    Call memzero(ByVal SourceAddress, Length)
End Sub

Public Function GetAddressOf(ByRef Target As Variant) As Long
    If IsByRef(Target) Then
        Call memcpy(GetAddressOf, ByVal (VarPtr(Target) + VARIANT_DATA_OFFSET), VLEN_PTR)
    Else
        GetAddressOf = VarPtr(Target)
    End If
End Function
