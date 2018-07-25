VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="Process" tags="immutable">
''  <summary>
''    Represents an application process.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Thread</using>
''  <idea>Everywhere...</idea>
''  <sample>
''      Process.Current.EndProcess Successfull
''      <br>
''      Process.Current.GetOutput.Write "Test data to stdout."
''      <br>
''      Process.OpenProcess "test.exe"
''  </sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
Implements ICloneable
'Implements ITargetStream

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Process" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
'Private Declare Function API_GetProcessAffinityMask Lib "kernel32" Alias "GetProcessAffinityMask" (ByVal hProcess As Long, lpProcessAffinityMask As Long, SystemAffinityMask As Long) As Long
'Private Declare Function API_GetProcessDefaultLayout Lib "user32" Alias "GetProcessDefaultLayout" (ByRef pdwDefaultLayout As Long) As Long
'Private Declare Function API_GetProcessHeap Lib "kernel32" Alias "GetProcessHeap" () As Long
'Private Declare Function API_GetProcessHeaps Lib "kernel32" Alias "GetProcessHeaps" (ByVal NumberOfHeaps As Long, ProcessHeaps As Long) As Long
''Private Declare Function API_GetProcessIoCounters Lib "kernel32" Alias "GetProcessIoCounters" (ByVal hProcess As Long, ByRef lpIoCounters As PIO_COUNTERS) As Long
''Private Declare Function API_GetProcessMemoryInfo Lib "psapi" Alias "GetProcessMemoryInfo" (ByVal Process As Long, ByRef ppsmemCounters As PPROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
'Private Declare Function API_GetProcessPriorityBoost Lib "kernel32" Alias "GetProcessPriorityBoost" (ByVal hProcess As Long, ByRef pDisablePriorityBoost As Long) As Long
'Private Declare Function API_GetProcessShutdownParameters Lib "kernel32" Alias "GetProcessShutdownParameters" (lpdwLevel As Long, lpdwFlags As Long) As Long
'Private Declare Function API_GetProcessTimes Lib "kernel32" Alias "GetProcessTimes" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
'Private Declare Function API_GetProcessVersion Lib "kernel32" Alias "GetProcessVersion" (ByVal ProcessID As Long) As Long
'Private Declare Function API_GetProcessWindowStation Lib "user32" Alias "GetProcessWindowStation" () As Long
'Private Declare Function API_GetProcessWorkingSetSize Lib "kernel32" Alias "GetProcessWorkingSetSize" (ByVal hProcess As Long, lpMinimumWorkingSetSize As Long, lpMaximumWorkingSetSize As Long) As Long
'
'Private Declare Function API_SetProcessAffinityMask Lib "kernel32" Alias "SetProcessAffinityMask" (ByVal hProcess As Long, ByRef dwProcessAffinityMask As Long) As Long
'Private Declare Function API_SetProcessDefaultLayout Lib "user32" Alias "SetProcessDefaultLayout" (ByVal dwDefaultLayout As Long) As Long
'Private Declare Function API_SetProcessPriorityBoost Lib "kernel32" Alias "SetProcessPriorityBoost" (ByVal hProcess As Long, ByVal bDisablePriorityBoost As Long) As Long
'Private Declare Function API_SetProcessShutdownParameters Lib "kernel32" Alias "SetProcessShutdownParameters" (ByVal dwLevel As Long, ByVal dwFlags As Long) As Long
'Private Declare Function API_SetProcessWindowStation Lib "user32" Alias "SetProcessWindowStation" (ByVal hWinSta As Long) As Long
'Private Declare Function API_SetProcessWorkingSetSize Lib "kernel32" Alias "SetProcessWorkingSetSize" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
'
'Private Declare Function API_GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
'
''Private Declare Function API_TerminateProcess Lib "kernel32" Alias "TerminateProcess" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
'
'Private Declare Function API_IsWow64Process Lib "kernel32" Alias "IsWow64Process" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long

'Public Declare Function EnumProcesses Lib "psapi" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
'Public Declare Function EnumProcessModules Lib "psapi" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long

'Private Const STD_ERROR_HANDLE = -12&
'Private Const STD_INPUT_HANDLE = -10&
'Private Const STD_OUTPUT_HANDLE = -11&

'Private Const PROCESS_ALL_ACCESS = &H1F0FFF
'Private Const PROCESS_CREATE_THREAD = &H2
'Private Const PROCESS_CREATE_PROCESS = &H80
'Private Const PROCESS_DUP_HANDLE = &H40
'Private Const PROCESS_HEAP_ENTRY_BUSY = &H4
'Private Const PROCESS_HEAP_ENTRY_DDESHARE = &H20
'Private Const PROCESS_HEAP_ENTRY_MOVEABLE = &H10
'Private Const PROCESS_HEAP_REGION = &H1
'Private Const PROCESS_HEAP_UNCOMMITTED_RANGE = &H2
'Private Const PROCESS_QUERY_INFORMATION = &H400
'Private Const PROCESS_SET_INFORMATION = &H200
'Private Const PROCESS_SET_QUOTA = &H100
'Private Const PROCESS_TERMINATE = &H1
'Private Const PROCESS_VM_OPERATION = &H8
'Private Const PROCESS_VM_READ = &H10
'Private Const PROCESS_VM_WRITE = &H20


Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables/Events>
'
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Dim PID As Long 'process id
'Dim pHndl As Long

'Dim h_stdInput As Long
'Dim h_stdOutput As Long
'Dim h_stdError As Long

'------------------------------------
'Dim target_pID As Long
'Dim target_tID As Long

'Dim h_target_stdInput As Long
'Dim h_target_stdOutput As Long
'Dim h_target_stdError As Long

Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Constructors>
'
''<summary>
''  Basic constructor to use by MintAPI class provider,
''  When creating new instance of this class.
''</summary>
''<params>
''  <param name="Arguments">
''      A list of construction arguments.
''  </param>
''</params>
''<cargs order="0" count="1">
''  <carg type="long" name="ID"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal hProcess As Long)
    PID = API_GetProcessId(hProcess)
    If Not Me.IsOpen Then
        PID = 0
        Throw_InvalidPID
    End If
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal ProcessID As Long) 'Unsafe if process being closed.
    PID = ProcessID
    
    This_Constructed = True
End Sub
Friend Sub Constructor2(ByVal ProcessID As Long) 'Safe anyway.
    PID = ProcessID
    If Not Me.IsOpen Then
        PID = 0
        Throw_InvalidPID
    End If
    
    This_Constructed = True
End Sub
Friend Sub Constructor3(ProcessEntry As API_PROCESSENTRY32)
    PID = ProcessEntry.th32ProcessID
    If Not Me.IsOpen Then
        PID = 0
        Throw_InvalidPID
    End If
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal Process As Process)
    PID = Process.ProcessID
    If Not Me.IsOpen Then
        PID = 0
        Throw_InvalidPID
    End If
    
    This_Constructed = True
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

''<summary>Returns the name of the process shown in system process list.</summary>
''<retval>Process name.</retval>
Public Property Get Name() As String
    Dim PEntry As API_PROCESSENTRY32
    If Not GetProcessEntryFor(Me.ProcessID, PEntry) Then Throw_InvalidPID
    Name = GetLPSTR(StrConv(PEntry.szExeFile, vbUnicode))
End Property

''<summary>Represents the main class handle of process.</summary>
''<retval>A handle to the system process class with full access.</retval>
'Public Property Get Handle() As Long
'   Handle = OpenNewProcess(PROCESS_ALL_ACCESS, True)
'End Property
''<summary>Represents the id of the process to control it by Win32 api.</summary>
''<retval>A 32 bit integer id of the process.</retval>
Public Property Get ProcessID() As Long
    ProcessID = PID
End Property
''<summary>Represents the module id of the process executing file.</summary>
''<retval>Process specified module id.</retval>
Public Property Get ModuleID() As Long
    Dim PEntry As API_PROCESSENTRY32
    If Not GetProcessEntryFor(Me.ProcessID, PEntry) Then Throw_InvalidPID
    ModuleID = PEntry.th32ModuleID
End Property
''<summary>The address of the process execution file.</summary>
''<retval>Process specified executing file path.</retval>
Public Property Get Executable() As String
    Dim Str As String, retLength As Long
    Str = String$(MAX_PATH_A + 1, 0)
    'retLength = API_GetModuleFileNameUnicode(Me.ModuleID, str, MAX_PATH_A)
    With UsingH(Process.CreateHandle(Me, PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False))
        retLength = API_GetModuleFileNameExUnicode(.Handle, vbNullPtr, Str, MAX_PATH_A)
        Executable = StrConv(Str, vbFromUnicode)
        Executable = Left$(Str, retLength)
    End With
End Property

''<summary>Indicates the process specified by it's ID is still open or not.</summary>
''<retval>A boolean that determines the process is open or not.</retval>
Public Property Get IsOpen() As Boolean
    Dim PEntry As API_PROCESSENTRY32
    IsOpen = GetProcessEntryFor(Me.ProcessID, PEntry)
End Property

'Public Function ReadProcessMemory(ByVal BaseAddress As Long, ByVal Length As Long) As Memory
'    If BaseAddress < 0 Then throw Exps.NegativeArgumentException
'    If Length <= 0 Then throw Exps.OutOfRangeException("Only unsigned numbers accepted.")
'    Dim Buffer() As Byte, BytesReaded As Long
'    MemLong(API_VarPtrArray(Buffer)) = API_SafeArrayCreateVector(vbByte, 0, Length)
'    If API_Toolhelp32ReadProcessMemory(Me.ProcessID, ByVal BaseAddress, Buffer(0), Length, BytesReaded) = False Then _
'        throw Exps.IfError
'
'    Set ReadProcessMemory = New Memory
'    Call ReadProcessMemory.Constructor4(Buffer, 0, BytesReaded)
'    MemLong(API_VarPtrArray(Buffer)) = 0
'End Function

Public Property Get MainThread() As Thread
    Set MainThread = New Thread
    Call MainThread.Constructor2(0)
End Property
Public Property Get Threads() As Vector
    Dim lSnapShot As Long
    Dim tTE As API_THREADENTRY32
    Dim lNextThread As Long
    Dim tList As New List
    Dim PID As Long

    PID = ProcessID

    lSnapShot = INVALID_HANDLE_VALUE
    ' Take a snapshot of all running threads on current process.
    lSnapShot = API_CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, PID)
    If lSnapShot <> INVALID_HANDLE_VALUE Then
        tTE.dwSize = LenB(tTE)
        ' Retrieve information about the first thread,
        ' and exit if unsuccessful
        '@@@@@@@@@@@@@@@@lNextThread = API_Thread32First(lSnapShot, tTE)
        ' Now walk the thread list of the system,
        ' and display information about each thread
        ' associated with the specified process
        Do While lNextThread
            If tTE.th32OwnerProcessID = PID Then
                'If tList.FindIndex() < 0 Then
                    Dim cThread As New Thread
                    Call cThread.Constructor2(tTE.th32ThreadID) ' From ThreadID (unsafe)
                    Call tList.Append(cThread)
                'End If
            End If
            ' Get next thread
            Call memzero(tTE, LenB(tTE))
            tTE.dwSize = LenB(tTE)
            '@@@@@@@@@@@@@@@@@@@lNextThread = API_Thread32Next(lSnapShot, tTE)
        Loop
        ' Close the handle, Don't forget to clean up the snapshot object.
        Call API_CloseHandle(lSnapShot)
    Else
        throw Exps.IfError
    End If
    Set Threads = tList.ToVector
End Property

Public Property Get MainModule() As Module
    Dim tPE As API_PROCESSENTRY32
    If Not GetProcessEntryFor(ProcessID, tPE) Then _
        Throw_InvalidPID
        
    Set MainModule = New Module
    Call MainModule.Constructor0(tPE.th32ModuleID)
End Property
Public Property Get Modules() As Vector
    Dim lSnapShot As Long
    Dim tME As API_MODULEENTRY32
    Dim lNextModule As Long
    Dim mList As New List
    Dim PID As Long

    PID = ProcessID

    lSnapShot = INVALID_HANDLE_VALUE
    ' Take a snapshot of all running threads on current process.
    lSnapShot = API_CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, PID)
    If lSnapShot <> INVALID_HANDLE_VALUE Then
        tME.dwSize = LenB(tME)
        ' Retrieve information about the first thread,
        ' and exit if unsuccessful
        lNextModule = API_Module32First(lSnapShot, tME)
        ' Now walk the thread list of the system,
        ' and display information about each thread
        ' associated with the specified process
        Do While lNextModule
            'If tME.th32ProcessID = PID Then
                'If tList.FindIndex() < 0 Then
                    Dim cModule As New Module
                    Call cModule.Constructor0(tME.hModule) ' From ThreadID (unsafe)
                    Call mList.Append(cModule)
                'End If
            'End If
            ' Get next thread
            lNextModule = API_Module32Next(lSnapShot, tME)
        Loop
        ' Close the handle, Don't forget to clean up the snapshot object.
        Call API_CloseHandle(lSnapShot)
    Else
        throw Exps.IfError
    End If
    Set Modules = mList.ToVector
End Property

Public Property Get MainHeap() As Heap
    Dim PEntry As API_PROCESSENTRY32
    Dim HeapID As Long, HeapLength As Long
    If GetProcessEntryFor(Me.ProcessID, PEntry) Then Throw_InvalidPID
    HeapID = PEntry.th32DefaultHeapID
    'HeapLength = MintCore.KernelHeap.API_HeapSize(HeapID
    
    'Set MainHeap = New Memory
    'call MainHeap.Constructor1(
End Property
Public Property Get Heaps() As Vector
    Dim lSnapShot As Long
    Dim tHL As API_HEAPLIST32
    Dim lNextHeap As Long
    Dim hList As New List
    Dim PID As Long

    PID = ProcessID

    lSnapShot = INVALID_HANDLE_VALUE
    ' Take a snapshot of all running threads on current process.
    lSnapShot = API_CreateToolhelp32Snapshot(TH32CS_SNAPHEAPLIST, PID)
    If lSnapShot <> INVALID_HANDLE_VALUE Then
        tHL.dwSize = LenB(tHL)
        ' Retrieve information about the first thread,
        ' and exit if unsuccessful
        lNextHeap = API_Heap32ListFirst(lSnapShot, tHL)
        ' Now walk the thread list of the system,
        ' and display information about each thread
        ' associated with the specified process
        Do While lNextHeap
            If tHL.th32ProcessID = PID Then

            End If
            ' Get next thread
            lNextHeap = API_Heap32ListNext(lSnapShot, tHL)
        Loop
        ' Close the handle, Don't forget to clean up the snapshot object.
        Call API_CloseHandle(lSnapShot)
    Else
        throw Exps.IfError
    End If
    Set Heaps = hList.ToVector
End Property

Public Property Get HaveParent() As Boolean
    Dim PPID As Long, PEntry As API_PROCESSENTRY32
    If Not GetProcessEntryFor(Me.ProcessID, PEntry) Then _
        Throw_InvalidPID
    PPID = PEntry.th32ParentProcessID
    
    HaveParent = (PPID <> vbNullPtr)
End Property
Public Property Get Parent() As Process
    Dim PPID As Long, PEntry As API_PROCESSENTRY32
    If Not GetProcessEntryFor(Me.ProcessID, PEntry) Then _
        Throw_InvalidPID
    PPID = PEntry.th32ParentProcessID
    If PPID = vbNullPtr Then Exit Function
    
    Set Parent = New Process
    Call Parent.Constructor1(PPID)
End Property

''<summary>
''  Returns the master process which started this process or its parents.
''  Note: If no processes started this process current process itself will return.
''</summary>
''<retval>A process which representing the root executor process.</retval>
Public Property Get RootProcess() As Process
    Dim PEntry As API_PROCESSENTRY32
    If Not GetProcessEntryFor(Me.ProcessID, PEntry) Then _
        Throw_InvalidPID
    
    While GetParent(PEntry): Wend
    
    Set RootProcess = New Process
    Call RootProcess.Constructor3(PEntry)
End Property

Public Property Get ChildProcesses() As Vector
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As API_PROCESSENTRY32
    Dim Proc As Process
    Dim pList As New List
    Dim CPID As Long

    CPID = API_GetCurrentProcessId

    lSnapShot = INVALID_HANDLE_VALUE
    ' Create snapshot
    lSnapShot = API_CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, vbNullPtr)
    If lSnapShot <> INVALID_HANDLE_VALUE Then
        ' Length of the structure
        tPE.dwSize = Len(tPE)
        ' Find first process
        lNextProcess = API_Process32First(lSnapShot, tPE)
        Do While lNextProcess
            If tPE.th32ParentProcessID = CPID Then
                Set Proc = New Process
                Call Proc.Constructor3(tPE)
                Call pList.Append(Proc)
            End If
            ' Get next process
            lNextProcess = API_Process32Next(lSnapShot, tPE)
        Loop
        ' Close handle
        Call API_CloseHandle(lSnapShot)
    Else
        throw Exps.IfError
    End If

    Set ChildProcesses = pList.ToVector
End Property

''<summary>Indicates that the executing process if a 64bit process.</summary>
Public Function Is64BitProcess() As Boolean
    Dim RetVal As Long
    With UsingH(Process.CreateHandle(Me, PROCESS_QUERY_INFORMATION, False))
        If API_IsWow64Process(.Handle, RetVal) = NO_VALUE Then throw Exps.IfError
        Is64BitProcess = False '(RetVal = NO_VALUE)
    End With
End Function


''<summary>
''  EndProcess is the preferred method of ending a process.
''  This function provides a clean process shutdown.
''  This includes calling the entry-point function of all attached
''  dynamic-link libraries (DLLs) with a value indicating that the
''  process is detaching from the DLL.
''  If a process terminates by calling Terminate,
''  the DLLs that the process is attached to are not notified
''  of the process termination.
''</summary>
''<params><param name="ExitCode">Determines the exit code for the process.</param></params>
Public Sub EndProcess(Optional ByVal ExitCode As Long = SUCCESSFULL)
    If IsDesignMode And (ProcessID = API_GetCurrentProcessId) Then
        Call EndDebugging
    Else
        Call API_ExitProcess(ExitCode)
    End If
End Sub
'Public Sub EndProcessTree(Optional ByVal ExitCode As Long = SUCCESSFULL)
'    Call PrivateEndProcessTree(Me.RootProcess, ExitCode)
'End Sub
Public Function Terminate(Optional ByVal ExitCode As Long = SUCCESSFULL) As Long
    If IsDesignMode And (ProcessID = API_GetCurrentProcessId) Then
        Call EndDebugging
    Else
        With UsingH(Process.CreateHandle(Me, PROCESS_TERMINATE, False))
            Terminate = API_GetExitCodeProcess(.Handle, ExitCode)
            Call API_TerminateProcess(.Handle, Terminate)
        End With
    End If
End Function

Public Function ExitCode(Optional ByVal lExitCode As Long = SUCCESSFULL) As Long
    With UsingH(Process.CreateHandle(Me, PROCESS_TERMINATE, False))
        ExitCode = API_GetExitCodeProcess(.Handle, lExitCode)
    End With
End Function

Public Property Get TaskVisible() As Boolean
    
End Property
Public Property Let TaskVisible(Value As Boolean)
    
End Property


'Public Sub SetPrivilegeState()
'
'End Sub
'Public Function GetPrivilegeState() As Boolean
'
'End Function


'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Basic Members>
'
Public Function Equals(Value As Variant) As Boolean
    Equals = Objects.Equals(Me, Value)
End Function
Public Function ToString() As String
    ToString = CLASSALIAS
End Function
Public Function GetHashCode() As Long
    GetHashCode = Objects.GetHashCode(Me)
End Function
Public Function MetaObject() As MetaObject
    If myMetaObject Is Nothing Then
        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    End If
    Set MetaObject = myMetaObject
End Function

Public Function Clone() As Process
    Set Clone = New Process
    Call Clone.ConstructorCopy(Me)
End Function

'Public Function Fork() As Process
'   ...moved to mintposix module
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

'Private Function GetModulePath(ByVal hModule As Long) As String
'    Dim retVal As String,
'    retVal = String$(MAX_PATH_A, 0)
'    Call API_GetModuleFileName(hModule, retVal, Len(retVal))
'    GetModulePath = GetLPSTR(retVal)
'End Function

Private Function GetProcessEntryFor(ByVal PID As Long, ByRef PEntry As API_PROCESSENTRY32) As Boolean
    If PID = vbNullPtr Then Exit Function
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As API_PROCESSENTRY32
    
    GetProcessEntryFor = False
    
    lSnapShot = INVALID_HANDLE_VALUE
    ' Create snapshot
    lSnapShot = API_CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, vbNullPtr)
    If lSnapShot <> INVALID_HANDLE_VALUE Then
        ' Length of the structure
        tPE.dwSize = Len(tPE)
        ' Find first process
        lNextProcess = API_Process32First(lSnapShot, tPE)
        Do While lNextProcess
            If tPE.th32ProcessID = PID Then
                PEntry = tPE
                GetProcessEntryFor = True
                Exit Do
            End If
            ' Get next process
            lNextProcess = API_Process32Next(lSnapShot, tPE)
        Loop
        ' Close handle
        Call API_CloseHandle(lSnapShot)
    Else
        throw Exps.IfError
    End If
End Function

Private Sub PrivateEndProcessTree(ByVal Proc As Process, ByVal ExitCode As Long)
    If Proc Is Nothing Then throw Exps.ArgumentNullException
    If Not Proc.IsOpen Then Throw_InvalidPID
    Dim Children As Vector, CP As Variant
    Set Children = Proc.ChildProcesses
    If Not Children Is Nothing Then
        For Each CP In Children
            Call PrivateEndProcessTree(CP, ExitCode)
        Next
    End If

    Call Proc.EndProcess(ExitCode)
End Sub
Private Function GetParent(ByRef outPEntry As API_PROCESSENTRY32) As Boolean
    Dim PPID As Long, PEntry As API_PROCESSENTRY32
    If GetProcessEntryFor(outPEntry.th32ParentProcessID, PEntry) Then
        outPEntry = PEntry
        GetParent = True
    Else
        GetParent = False
    End If
End Function

Private Sub Throw_InvalidPID()
    throw Exps.InvalidOperationException("Specified process handle is not open.")
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo): Call Constructor(CInfo): End Sub
Private Function IObject_Equals(Value As Variant) As Boolean: IObject_Equals = Equals(Value): End Function
Private Function IObject_GetHashCode() As Long: IObject_GetHashCode = GetHashCode: End Function
Private Function IObject_MetaObject() As MetaObject: Set IObject_MetaObject = MetaObject: End Function
Private Function IObject_ToString() As String: IObject_ToString = ToString: End Function
Private Function IObject_Stub(ByVal ProxyInfo As ObjectProxyInfo) As Boolean
    
    IObject_Stub = False
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'HANDLE              hToken;
'LUID                SeDebugNameValue;
'TOKEN_PRIVILEGES    TokenPrivileges;
'
'if(OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken))
'{
'    if(LookupPrivilegeValue(NULL, SE_DEBUG_NAME, &SeDebugNameValue))
'    {
'        TokenPrivileges.PrivilegeCount              = 1;
'        TokenPrivileges.Privileges[0].Luid          = SeDebugNameValue;
'        TokenPrivileges.Privileges[0].Attributes    = SE_PRIVILEGE_ENABLED;
'
'        if(AdjustTokenPrivileges(hToken, FALSE, &TokenPrivileges, sizeof(TOKEN_PRIVILEGES), NULL, NULL))
'        {
'            CloseHandle(hToken);
'        }
'        Else
'        {
'            CloseHandle(hToken);
'            throw std::exception("Couldn't adjust token privileges!");
'        }
'    }
'    Else
'    {
'        CloseHandle(hToken);
'        throw std::exception("Couldn't look up privilege value!");
'    }
'}
'Else
'{
'    throw std::exception("Couldn't open process token!");
'}
