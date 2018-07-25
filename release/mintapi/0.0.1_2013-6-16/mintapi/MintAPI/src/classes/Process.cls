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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Application"
Attribute VB_Ext_KEY = "Member1" ,"Thread"
Attribute VB_Ext_KEY = "Member2" ,"Process"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Process"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements ITargetStream
Implements IAliasable


Private Declare Sub API_CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hPass As Long)

Private Declare Function API_CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function API_CreateProcessAsUser Lib "kernel32" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function API_GetCurrentProcess Lib "kernel32" Alias "GetCurrentProcess" () As Long
Private Declare Function API_GetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Function API_GetExitCodeProcess Lib "kernel32" Alias "GetExitCodeProcess" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function API_GetProcAddress Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function API_GetProcessAffinityMask Lib "kernel32" Alias "GetProcessAffinityMask" (ByVal hProcess As Long, lpProcessAffinityMask As Long, SystemAffinityMask As Long) As Long
Private Declare Function API_GetProcessDefaultLayout Lib "user32" Alias "GetProcessDefaultLayout" (ByRef pdwDefaultLayout As Long) As Long
Private Declare Function API_GetProcessHeap Lib "kernel32" Alias "GetProcessHeap" () As Long
Private Declare Function API_GetProcessHeaps Lib "kernel32" Alias "GetProcessHeaps" (ByVal NumberOfHeaps As Long, ProcessHeaps As Long) As Long
'Private Declare Function API_GetProcessIoCounters Lib "kernel32" Alias "GetProcessIoCounters" (ByVal hProcess As Long, ByRef lpIoCounters As PIO_COUNTERS) As Long
'Private Declare Function API_GetProcessMemoryInfo Lib "psapi" Alias "GetProcessMemoryInfo" (ByVal Process As Long, ByRef ppsmemCounters As PPROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function API_GetProcessPriorityBoost Lib "kernel32" Alias "GetProcessPriorityBoost" (ByVal hProcess As Long, ByRef pDisablePriorityBoost As Long) As Long
Private Declare Function API_GetProcessShutdownParameters Lib "kernel32" Alias "GetProcessShutdownParameters" (lpdwLevel As Long, lpdwFlags As Long) As Long
Private Declare Function API_GetProcessTimes Lib "kernel32" Alias "GetProcessTimes" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Private Declare Function API_GetProcessVersion Lib "kernel32" Alias "GetProcessVersion" (ByVal ProcessId As Long) As Long
Private Declare Function API_GetProcessWindowStation Lib "user32" Alias "GetProcessWindowStation" () As Long
Private Declare Function API_GetProcessWorkingSetSize Lib "kernel32" Alias "GetProcessWorkingSetSize" (ByVal hProcess As Long, lpMinimumWorkingSetSize As Long, lpMaximumWorkingSetSize As Long) As Long

Private Declare Function API_SetProcessAffinityMask Lib "kernel32" Alias "SetProcessAffinityMask" (ByVal hProcess As Long, ByRef dwProcessAffinityMask As Long) As Long
Private Declare Function API_SetProcessDefaultLayout Lib "user32" Alias "SetProcessDefaultLayout" (ByVal dwDefaultLayout As Long) As Long
Private Declare Function API_SetProcessPriorityBoost Lib "kernel32" Alias "SetProcessPriorityBoost" (ByVal hProcess As Long, ByVal bDisablePriorityBoost As Long) As Long
Private Declare Function API_SetProcessShutdownParameters Lib "kernel32" Alias "SetProcessShutdownParameters" (ByVal dwLevel As Long, ByVal dwFlags As Long) As Long
Private Declare Function API_SetProcessWindowStation Lib "user32" Alias "SetProcessWindowStation" (ByVal hWinSta As Long) As Long
Private Declare Function API_SetProcessWorkingSetSize Lib "kernel32" Alias "SetProcessWorkingSetSize" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long

Private Declare Function API_ReadProcessMemory Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function API_RegisterServiceProcess Lib "kernel32" Alias "RegisterServiceProcess" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long

Private Declare Function API_GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private Declare Function API_CreateToolhelp32Snapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function API_Process32First Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function API_Process32Next Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function API_OpenProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function API_TerminateProcess Lib "kernel32" Alias "TerminateProcess" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function API_IsWow64Process Lib "kernel32" Alias "IsWow64Process" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
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
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
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

Private Const TH32CS_SNAPPROCESS = &H2&

Dim ppID As Long 'parent process id
Dim pHndl As Long 'process handle
Dim pID As Long 'process id
Dim mID As Long 'module id
Dim tID As Long 'main thread id
Dim tCount As Long 'number of threads

Dim executable_file_path As String

Dim hID As Long 'heap id

Dim h_stdInput As Long
Dim h_stdOutput As Long
Dim h_stdError As Long

'------------------------------------
Dim target_pID As Long
Dim target_tID As Long

Dim h_target_stdInput As Long
Dim h_target_stdOutput As Long
Dim h_target_stdError As Long

Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String


Friend Sub InitializeW(ProcessEntry As PROCESSENTRY32)
    pID = ProcessEntry.th32ProcessID
    ppID = ProcessEntry.th32ParentProcessID
    mID = ProcessEntry.th32ModuleID
    hID = ProcessEntry.th32DefaultHeapID
    executable_file_path = GetLPSTR(ProcessEntry.szExeFile)
End Sub
Friend Sub Initialize(Optional ByVal Process As Long = -1, Optional ByVal ProcessId As Long = -1)
    If Process = -1 Then Process = 0
    If ProcessId = -1 Then ProcessId = 0
    pHndl = Process
    pID = ProcessId
End Sub

Public Property Get Name() As String

End Property
Public Property Get ExecutableFile() As File

End Property

Public Property Get Handle() As Long
    If pID = 0 Then
        Handle = CurrentProcess.Handle
    Else
        Handle = pID
    End If
End Property
Public Property Get ClassHandle() As Long
    If pHndl = 0 Then
        pHndl = CurrentProcess.ClassHandle
    End If
    ClassHandle = pHndl
End Property

Public Function CurrentProcess() As Process
    Dim p As New Process
    Call p.Initialize(API_GetCurrentProcessId)
    Set CurrentProcess = p
End Function
Public Function OpenProcess(Path As String, Arguments As String, Optional Environment As String) As Process

End Function
Public Function OpenProcessAs(Path As String, User As String, Arguments As String, Optional Environment As String) As Process

End Function
Public Function Process(ByVal ProcessId As Long) As Process
    Dim p As New Process
    Call p.Initialize(API_GetCurrentProcessId)
    Set CurrentProcess = p
End Function


Public Function Thread() As Thread
    
End Function
Public Function Threads() As Thread()

End Function
Public Sub DataOutput()

End Sub
Public Sub DataInput()

End Sub

Public Function EndProcess(Optional ByVal ExitCode As Long = 0, Optional ByVal Force As Boolean = False) As Long
    
End Function
Public Function EndProcessTree(Optional ByVal ExitCode As Long = 0, Optional ByVal Force As Boolean = False) As Long
    
End Function
Public Function Terminate(Optional ByVal ExitCode As Long = 0, Optional ByVal Force As Boolean = False) As Long
    Terminate = API_GetExitCodeProcess(ClassHandle, ExitCode)
    Call API_CloseHandle(ClassHandle)
End Function
Public Function Application() As IApplication

End Function
Public Property Get TaskVisible() As Boolean

End Property
Public Property Let TaskVisible(Value As Boolean)

End Property

Public Function CountProcesses() As Long
    CountProcesses = ArraySize(EnumProcesses)
End Function
Public Function EnumProcesses() As Process()
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As PROCESSENTRY32
    Dim p() As Process
    Dim pCount As Long

    ' Create snapshot
    lSnapShot = API_CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
        ' Length of the structure
        tPE.dwSize = Len(tPE)
        ' Find first process
        lNextProcess = API_Process32First(lSnapShot, tPE)
        Do While lNextProcess
'            Found specified process
'            If sExeName = Left$(tPE.szExeFile, Len(sExeName)) And Len(sExeName) > 0 Then
'                Dim lProcess As Long
'                Dim lExitCode As Long
'                ' Open process
'                lProcess = OpenProcess(0, False, tPE.th32ProcessID)
'                ' Terminate process
'                TerminateProcess lProcess, lExitCode
'                ' Close handle
'                CloseHandle lProcess
                ' Add exe to list
                ReDim Preserve p(pCount)
                Set p(pCount) = New Process
                Call p(pCount).InitializeW(tPE)
                pCount = pCount + 1
            'End If
            ' Get next process
            lNextProcess = API_Process32Next(lSnapShot, tPE)
        Loop
        ' Close handle
        Call API_CloseHandle(lSnapShot)
    Else
        throw SystemCallFailureException
    End If
End Function

Public Function ChildProcess(Name As String) As Process

End Function
Public Function RegisterChildProcesses(Processes) As Process

    Set RegisterChildProcesses = Me
End Function
Public Function ParentProcess() As Process

End Function

Public Function IsWow64Process() As Boolean
    Dim retVal As Long
    Call API_IsWow64Process(pHndl, retVal)
    IsWow64Process = (retVal <> 0)
End Function

Public Sub SetPrivilegeState()

End Sub
Public Function GetPrivilegeState() As Boolean

End Function

Private Function ITargetStream_Available(Dir As StreamDirection) As Long

End Function

Private Sub ITargetStream_CloseStream(Optional direction As StreamDirection = 3&)

End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean

End Function

Private Sub ITargetStream_inStream(Data, Optional Length As Long = -1)

End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    
End Function

Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = StreamDirection.sdBoth)

End Sub

Private Sub ITargetStream_outStream(Data, Optional Length As Long = -1)

End Sub

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved)
    Call mint_setstream_state(State, inState, inState_LOCK, Reserved)
End Sub

Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved)
    Call mint_setstream_state(State, outState, outState_LOCK, Reserved)
End Sub


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
