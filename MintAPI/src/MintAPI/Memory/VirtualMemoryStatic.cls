VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VirtualMemoryStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="VirtualMemoryStatic" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "VirtualMemoryStatic" '#DO NOT REMOVE IT
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
   
Private Const MEM_COMMIT           As Long = &H1000
Private Const MEM_RESERVE          As Long = &H2000
Private Const MEM_DECOMMIT         As Long = &H4000
Private Const MEM_RELEASE          As Long = &H8000
Private Const MEM_FREE             As Long = &H10000
Private Const MEM_PRIVATE          As Long = &H20000
Private Const MEM_MAPPED           As Long = &H40000
Private Const MEM_RESET            As Long = &H80000
Private Const MEM_TOP_DOWN         As Long = &H100000
Private Const MEM_4MB_PAGES        As Long = &H80000000
Private Const SEC_FILE             As Long = &H800000
Private Const SEC_IMAGE            As Long = &H1000000
Private Const SEC_VLM              As Long = &H2000000
Private Const SEC_RESERVE          As Long = &H4000000
Private Const SEC_COMMIT           As Long = &H8000000
Private Const SEC_NOCACHE          As Long = &H10000000
Private Const MEM_IMAGE            As Long = SEC_IMAGE

Public Enum VirtualMemoryAllocationTypes
    vmatNormal = 0
    vmatCommit = MEM_COMMIT
    vmatReserve = MEM_RESERVE
    vmatReset = MEM_RESET
    vmatTopDown = MEM_TOP_DOWN
End Enum
Public Enum VirtualMemoryFreeTypes
    vmftNormal = 0
    vmftDecommit = MEM_DECOMMIT
    vmftRelease = MEM_RELEASE
End Enum
Public Enum VirtualMemoryStates
    vmsCommit = MEM_COMMIT
    vmsFree = MEM_FREE
    vmsReserve = MEM_RESERVE
End Enum
Public Enum VirtualMemoryTypes
    vmtImage = MEM_IMAGE
    vmtMapped = MEM_MAPPED
    vmtPrivate = MEM_PRIVATE
End Enum

Public Type MemoryBasicInformation
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As MemoryAccessPermissions
    RegionSize As Long
    State As VirtualMemoryStates
    Protect As MemoryAccessPermissions
    Type As VirtualMemoryTypes
End Type
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Friend Function Create(ByVal Reference As Long, _
            ByVal Limit As Long, _
            ByVal Access As MemoryAccessPermissions, _
            ByVal AllocateType As VirtualMemoryAllocationTypes, _
            ByVal Process As Process) As VirtualMemory
    Set Create = New VirtualMemory
    Call Create.Constructor0(Reference, Limit, Access, AllocateType, Process)
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

Public Function Allocate(ByVal Size As Long, _
        ByVal AllocationType As VirtualMemoryAllocationTypes, _
        ByVal Access As MemoryAccessPermissions, _
        Optional ByVal PreferedAddress As Long = 0, _
        Optional ByVal Process As Process = Nothing) As IMemoryBlock
    
    Dim Result As Long
    If Process Is Nothing Then
        Result = API_VirtualAlloc(PreferedAddress, Size, AllocationType, Access)
    Else
        With UsingH(MintAPI.Process.CreateHandle(Process, PROCESS_VM_OPERATION, False))
            Result = API_VirtualAllocEx(.Handle, PreferedAddress, Size, AllocationType, Access)
        End With
    End If
    
    If Result = vbNullPtr Then throw Exps.IfError
    
    Set Allocate = Create(Result, Size, Access, AllocationType, Process)
End Function

'Public Sub Copy(ByVal Destination As IMemoryBlock, ByVal Source As IMemoryBlock, _
'        ByVal Size As Long, ByVal Access As MemoryAccessPermissions)
'
'    If Destination Is Nothing Then throw Exps.ArgumentNullException("Destination")
'    If Source Is Nothing Then throw Exps.ArgumentNullException("Source")
'    If Not TypeOf Destination Is VirtualMemory Then throw Exps.InvalidArgumentException("Destination")
'    If Not TypeOf Source Is VirtualMemory Then throw Exps.InvalidArgumentException("Source")
'
''    If API_VirtualCopy(Destination.Reference, Source.Reference, Size, Access) = NO_VALUE Then _
''        throw Exps.IfError
'End Sub

Public Sub Release(ByVal VirtualMemory As IMemoryBlock, _
        ByVal FreeType As VirtualMemoryFreeTypes, _
        Optional ByVal Process As Process = Nothing)
    
    If VirtualMemory Is Nothing Then throw Exps.ArgumentNullException("VirtualMemory")
    If Not TypeOf VirtualMemory Is VirtualMemory Then throw Exps.InvalidArgumentException("VirtualMemory")
    
    Dim Result As Long
    If Process Is Nothing Then
        Result = API_VirtualFree(VirtualMemory.Reference, VirtualMemory.Limit, FreeType)
    Else
        With UsingH(MintAPI.Process.CreateHandle(Process, PROCESS_VM_OPERATION, False))
            Result = API_VirtualFreeEx(.Handle, VirtualMemory.Reference, VirtualMemory.Limit, FreeType)
        End With
    End If
    
    If Result = vbNullPtr Then throw Exps.IfError
End Sub

Public Function Protect(ByVal VirtualMemory As IMemoryBlock, _
        ByVal Size As Long, _
        ByVal NewAccess As MemoryAccessPermissions, _
        Optional ByVal Process As Process = Nothing) As MemoryAccessPermissions
    
    If VirtualMemory Is Nothing Then throw Exps.ArgumentNullException("VirtualMemory")
    If Not TypeOf VirtualMemory Is VirtualMemory Then throw Exps.InvalidArgumentException("VirtualMemory")
    
    Dim Result As Long
    If Process Is Nothing Then
        Result = API_VirtualProtect(ByVal VirtualMemory.Reference, Size, NewAccess, Protect)
    Else
        With UsingH(MintAPI.Process.CreateHandle(Process, PROCESS_VM_OPERATION, False))
            Result = API_VirtualProtectEx(.Handle, ByVal VirtualMemory.Reference, Size, NewAccess, Protect)
        End With
    End If
    
    If Result = vbNullPtr Then throw Exps.IfError
End Function

Public Function Query(ByVal VirtualMemory As IMemoryBlock, _
        Optional ByVal Process As Process = Nothing) As MemoryBasicInformation
    
    If VirtualMemory Is Nothing Then throw Exps.ArgumentNullException("VirtualMemory")
    If Not TypeOf VirtualMemory Is VirtualMemory Then throw Exps.InvalidArgumentException("VirtualMemory")
    
    Dim MemoryBasicInformationLen As Long
    MemoryBasicInformationLen = LenB(Query)
    
    Dim Result As Long
    If Process Is Nothing Then
        Result = API_VirtualQuery(VirtualMemory.Reference, Query, MemoryBasicInformationLen)
    Else
        With UsingH(MintAPI.Process.CreateHandle(Process, PROCESS_VM_OPERATION, False))
            Result = API_VirtualQueryEx(.Handle, VirtualMemory.Reference, Query, MemoryBasicInformationLen)
        End With
    End If
    
    If Result = vbNullPtr Then throw Exps.IfError
End Function

Public Sub LockMemory(ByVal VirtualMemory As IMemoryBlock, ByVal Size As Long)
    If VirtualMemory Is Nothing Then throw Exps.ArgumentNullException("VirtualMemory")
    If Not TypeOf VirtualMemory Is VirtualMemory Then throw Exps.InvalidArgumentException("VirtualMemory")
    
    If API_VirtualLock(VirtualMemory.Reference, Size) = NO_VALUE Then _
        throw Exps.IfError
End Sub

Public Sub UnLockMemory(ByVal VirtualMemory As IMemoryBlock, ByVal Size As Long)
    If VirtualMemory Is Nothing Then throw Exps.ArgumentNullException("VirtualMemory")
    If Not TypeOf VirtualMemory Is VirtualMemory Then throw Exps.InvalidArgumentException("VirtualMemory")
    
    If API_VirtualUnlock(VirtualMemory.Reference, Size) = NO_VALUE Then _
        throw Exps.IfError
End Sub



'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
