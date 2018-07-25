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
'@PROJECT_LICENSE

''<class name="Process" tags="">
''  <summary>
''    Represents an application process.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Thread</using>
''  <idea>Me , Qt , .net Framework , Java , Andriod activities.</idea>
''  <sample>
''      Call Process.Current.EndProcess(Successfull)
''
''      If Process.Current.Fork.IsParent Then
''          Process.OpenProcess "test.exe"
''      End If
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
Private Declare Function API_CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function API_CreateProcessAsUser Lib "kernel32" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function API_GetExitCodeProcess Lib "kernel32" Alias "GetExitCodeProcess" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Sub API_ExitProcess Lib "kernel32" Alias "ExitProcess" (uExitCode As Long)
Private Declare Function API_TerminateProcess Lib "kernel32" Alias "TerminateProcess" (hProcess As Long, uExitCode As Long) As Boolean
Private Declare Sub API_FatalExit Lib "kernel32" Alias "FatalExit" (ExitCode As Long)
    
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
Private Declare Function API_GetProcessVersion Lib "kernel32" Alias "GetProcessVersion" (ByVal ProcessID As Long) As Long
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

Private Declare Function API_OpenProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
'Private Declare Function API_TerminateProcess Lib "kernel32" Alias "TerminateProcess" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

'Private Declare Function API_GetProcessHeap Lib "kernel32" Alias "GetProcessHeap" () As Long
'Private Declare Function API_GetProcessHeaps Lib "kernel32" Alias "GetProcessHeaps" (ByVal NumberOfHeaps As Long, ProcessHeaps As Long) As Long

Private Declare Function API_IsWow64Process Lib "kernel32" Alias "IsWow64Process" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long

Private Declare Function API_GetProcessId Lib "kernel32" Alias "GetProcessId" (ByVal Process As Long) As Long



Private Const STD_ERROR_HANDLE = -12&
Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&



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
Private Type API_PROCESSENTRY32
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
'Private Const TH32CS_SNAPPROCESS = &H2&
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

Dim ppID As Long 'parent process id
Dim pHndl As Long 'process handle
Dim pID As Long 'process id
Dim pmID As Long 'module id
Dim pPath As String
Dim ptID As Long 'main thread id
Dim ptCount As Long 'number of threads
Dim phID As Long 'heap id

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
Friend Sub ConstructorArg(ByVal Arguments As ArgumentList)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal hProcess As Long)
    pHndl = hProcess
    pID = API_GetProcessId(pHndl)
    pmID = API_GetModuleHandleA(vbNullString)
    pPath = GetModulePath(pmID)
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal ProcessID As Long)
    'pHndl = hProcess
    pID = ProcessID
    pmID = API_GetModuleHandleA(vbNullString)
    pPath = GetModulePath(pmID)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal Process As Process)
    
End Sub

'Friend Sub InitializeW(ProcessEntry As PROCESSENTRY32)
'    pID = ProcessEntry.th32ProcessID
'    ppID = ProcessEntry.th32ParentProcessID
'    pmID = ProcessEntry.th32ModuleID
'    hID = ProcessEntry.th32DefaultHeapID
'    pPath = GetLPSTR(ProcessEntry.szExeFile)
'End Sub
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
    
End Property

''<summary>Represents the main class handle of process.</summary>
''<retval>A handle to the system process class.</retval>
Public Property Get Handle() As Long
    Handle = pHndl
End Property
''<summary>Represents the id of the process to control it by Win32 api.</summary>
''<retval>A 32 bit integer id of the process.</retval>
Public Property Get ProcessID() As Long
    ProcessID = pID
End Property
''<summary>Represents the module id of the process executing file.</summary>
''<retval>Process specified module id.</retval>
Public Property Get ModuleID() As Long
    ModuleID = pmID
End Property
''<summary>The address of the process executing file.</summary>
''<retval>Process specified executing file path.</retval>
Public Property Get Executable() As String
    Executable = pPath
End Property


Public Property Get StandardOutput() As Stream
    Set StandardOutput = OpenProcessStream(API_GetStdHandle(STD_OUTPUT_HANDLE), sdOutStream, "Standard Output")
End Property
Public Property Get StandardInput() As Stream
    Set StandardInput = OpenProcessStream(API_GetStdHandle(STD_INPUT_HANDLE), sdInStream, "Standard Input")
End Property
Public Property Get StandardError() As Stream
    Set StandardError = OpenProcessStream(API_GetStdHandle(STD_ERROR_HANDLE), sdOutStream, "Standard Error Output")
End Property

Public Function Memory(ByVal BaseAddress As Long) As Memory
    
End Function

Public Sub VirtualProtect(ByVal Memory As Memory)
    
End Sub

Public Function Thread() As Thread
    
End Function
Public Function Threads() As Thread()
    
End Function

Public Function EndProcess(Optional ByVal ExitCode As Long = SUCCESSFULL) As Long
    If IsDesignMode Then
        Call EndDebugging
    Else
        Call API_ExitProcess(ExitCode)
    End If
End Function
Public Function EndProcessTree(Optional ByVal ExitCode As Long = SUCCESSFULL) As Long
    
End Function
Public Function Terminate(Optional ByVal ExitCode As Long = SUCCESSFULL) As Long
    If IsDesignMode Then
        Call EndDebugging
    Else
        Terminate = API_GetExitCodeProcess(Handle, ExitCode)
        Call API_TerminateProcess(Handle, Terminate)
        Call API_CloseHandle(Handle) 'unreachable code !!!
    End If
End Function

Public Function ExitCode() As Long
    
End Function

Public Property Get TaskVisible() As Boolean

End Property
Public Property Let TaskVisible(Value As Boolean)

End Property

Public Property Get MainHeap() As Memory
    
End Property
Public Property Get Heaps() As Memory()
    
End Property

Public Property Get Parent() As Process
    
End Property

Public Function Fork() As Process
    
End Function

Public Function ChildProcesses(Name As String) As Process()
    
End Function

''<summary>Indicates that the executing process if a Wow64(64bit) process.</summary>
Public Function IsWow64Process() As Boolean
    Dim retVal As Long
    If API_IsWow64Process(Handle, retVal) = NO_VALUE Then throw Exps.IfError
    IsWow64Process = (retVal <> vbZero)
End Function

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
Public Property Get Alias() As String
    Alias = CLASSALIASSTRING
End Property
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
    Set MetaObject = InitMetaObject
End Function

Public Function Clone() As Object
    
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
Private Sub Class_Initialize()
    
End Sub
Private Sub Class_Terminate()
    Signal.Emit Me, SIGNAL_TERMINATING
    If Not myMetaObject Is Nothing Then
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Private Function GetModulePath(ByVal hModule As Long) As String
    Dim retVal As String
    retVal = String$(XLARGELPSTR, 0)
    Call API_GetModuleFileNameA(hModule, retVal, Len(retVal))
    GetModulePath = GetLPSTR(retVal)
End Function
Private Function OpenProcessStream(ByVal Handle As Long, ByVal Direction As StreamDirection, ByVal DirName As String) As ProcessStream
    Select Case Direction
        Case sdInStream, sdOutStream, sdBoth
        Case Else
            throw Exps.InvalidArgumentValueException
    End Select
    
    Set OpenProcessStream = New ProcessStream
    Call OpenProcessStream.Constructor0(Handle, Direction, DirName)
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Initialize(ByVal Arguments As ArgumentList)
    Call ConstructorArg(Arguments)
End Sub
Private Property Get IObject_Alias() As String
    IObject_Alias = CLASSALIASSTRING
End Property
Private Function IObject_Equals(Value As Variant) As Boolean
    IObject_Equals = Equals(Value)
End Function
Private Function IObject_GetHashCode() As Long
    IObject_GetHashCode = GetHashCode
End Function
Private Function IObject_MetaObject() As MetaObject
    Set IObject_MetaObject = MetaObject
End Function
Private Function IObject_ToString() As String
    IObject_ToString = ToString
End Function
Private Function IObject_Notify(ByVal Name As String, ByVal Arguments As ArgumentList) As Variant
    Select Case LCase$(Name)
        'Case ""
            
        Case Else
            throw Exps.UnknownNotificationException, Me
    End Select
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ITargetStream Implementation>
'
Private Function ITargetStream_Available(Dir As StreamDirection) As Long

End Function

Private Sub ITargetStream_CloseStream(Optional Direction As StreamDirection = 3&)

End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean

End Function

Private Sub ITargetStream_inStream(Data, Optional Length As Long = -1)

End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    
End Function

Private Sub ITargetStream_OpenStream(Optional Direction As StreamDirection = StreamDirection.sdBoth)

End Sub

Private Sub ITargetStream_outStream(Data, Optional Length As Long = -1)

End Sub

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved)
    Call mint_setstream_state(State, inState, inState_LOCK, Reserved)
End Sub

Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved)
    Call mint_setstream_state(State, outState, outState_LOCK, Reserved)
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------


'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------


'=============================================
'=============================================
'=============================================
'<section MetaObject constructor>
'
Private Function InitMetaObject() As MetaObject
    If Not myMetaObject Is Nothing Then
        Set InitMetaObject = myMetaObject
    End If
    Call Objects.MetaObject.Setup(MetaObject, Me, True, False, CLASSVERSIONSTRING)
    '...
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
