VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ProcessStatic" static tags="">
''  <summary>
''    Provides process static methods.
''  </summary>
''  <using>Process</using>
''  <idea>Me , Qt , .net Framework , Java , Andriod activities.</idea>
''  <sample>
''      Process.Current. ... 'Returns current process.
''      stdout Process.OpenProcess("C:\testproc.exe", "").ProcessID
''  </sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ProcessStatic" '#DO NOT REMOVE IT
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
Private Const THREAD_BASE_PRIORITY_LOWRT = 15    'value that gets a thread to LowRealtime-1
Private Const THREAD_BASE_PRIORITY_MAX = 2       'maximum thread base priority boost
Private Const THREAD_BASE_PRIORITY_MIN = (-2)    'minimum thread base priority boost
Private Const THREAD_BASE_PRIORITY_IDLE = (-15)  'value that gets a thread to idle

Public Enum ProcessCreationFlags
    pcfDefault = 0
    
    pcfDebugProcess = &H1
    pcfDebugOnlyThisProcess = &H2
    
    pcfCreateSuspended = &H4
    pcfDetachedProcess = &H8
    pcfCreateNewConsole = &H10
    
    pcfNormalPriorityClass = &H20
    pcfIdlePriorityClass = &H40
    pcfHighPriorityClass = &H80
    pcfRealtimePriorityClass = &H100
    
    pcfCreateNewProcessGroup = &H200
    pcfCreateUnicodeEnvironment = &H400
    pcfCreateSeparateWOWVDM = &H800
    pcfCreateSharedWOWVDM = &H1000
    pcfCreateForceDos = &H2000
    
    pcfBelowNormalPriorityClass = &H4000
    pcfAboveNormalPriorityClass = &H8000
    
    pcfStackSizeParamIsAReservation = &H10000
    
    pcfCreateBreakawayFromJob = &H1000000
    pcfCreatePreserveCodeAuthzLevel = &H2000000
    pcfCreate_DefaultErrorMode = &H4000000
    pcfCreateNoWindow = &H8000000
    
    pcfProfileUser = &H10000000
    pcfProfileKernel = &H20000000
    pcfProfileServer = &H40000000
    pcfCreateIgnoreSystemDefault = &H80000000
    
    pcfThreadPriorityLowest = THREAD_BASE_PRIORITY_MIN
    pcfThreadPriorityBelowNormal = (pcfThreadPriorityLowest + 1)
    pcfThreadPriorityNormal = 0
    pcfThreadPriorityHighest = THREAD_BASE_PRIORITY_MAX
    pcfThreadPriorityAboveNormal = (pcfThreadPriorityHighest - 1)
    pcfThreadPriorityErrorReturn = MAXLONG
    pcfThreadPriorityTimeCritical = THREAD_BASE_PRIORITY_LOWRT
    pcfThreadPriorityIdle = THREAD_BASE_PRIORITY_IDLE
End Enum
Public Enum WaitForInputIdleState
    wfiSuccessfull = 0
    wfiTimeout = 258
    wfiFailed = -1
End Enum
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables>
'
Dim CurrentProcess_Static As Process
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Public Function FromHandle(ByVal Handle As Long) As Process
    Dim cProc As Process
    Set cProc = New Process
    Call cProc.Constructor0(Handle)
    Set FromHandle = cProc
End Function
Public Function FromProcessID(ByVal ProcessID As Long) As Process
    Dim cProc As Process
    Set cProc = New Process
    Call cProc.Constructor2(ProcessID)
    Set FromProcessID = cProc
End Function

Public Function OpenProcess(ByVal Path As String, ByVal Arguments As String, Optional ByVal Environment As String, Optional ByVal ProcessCreationFlags As ProcessCreationFlags = pcfDefault) As Process
Attribute OpenProcess.VB_UserMemId = 0
    Dim Sec As API_SECURITY_ATTRIBUTES, Startup As API_STARTUPINFO
    Dim ProcInfo As API_PROCESS_INFORMATION, EnvPtr As Long
    Sec.nLength = LenB(Sec)
    Startup.cb = LenB(Startup)
    If IsMissing(Environment) Then
        EnvPtr = 0
    Else
        EnvPtr = StrPtr(Environment)
    End If
    If API_CreateProcess(Path, Arguments, _
        Sec, Sec, False, ProcessCreationFlags, ByVal EnvPtr, _
        vbNullString, Startup, ProcInfo) = NO_VALUE Then
        throw Exps.IfError
    End If
    
    Set OpenProcess = FromProcessID(ProcInfo.dwProcessId)
End Function
Public Function OpenProcessAs(ByVal Path As String, ByVal User As String, ByVal Arguments As String, Optional ByVal Environment As String, Optional ByVal ProcessCreationFlags As ProcessCreationFlags = pcfDefault) As Process
    Dim Sec As API_SECURITY_ATTRIBUTES, Startup As API_STARTUPINFO
    Dim ProcInfo As API_PROCESS_INFORMATION, EnvPtr As Long
    Dim UserToken As Long
    Sec.nLength = LenB(Sec)
    Startup.cb = LenB(Startup)
    If IsMissing(Environment) Then
        EnvPtr = 0
    Else
        EnvPtr = StrPtr(Environment)
    End If
'    If API_CreateProcessAsUser(UserToken, Path, _
'        Arguments, Sec, Sec, False, ProcessCreationFlags, _
'        ByVal EnvPtr, vbNullString, Startup, ProcInfo) = NO_VALUE Then
'        throw Exps.IfError
'    End If
    
    Set OpenProcessAs = FromProcessID(ProcInfo.dwProcessId)
End Function
Public Function OpenPipedProcess(ByVal Path As String, ByVal Arguments As String, Optional ByVal Environment As String) As Process
    
End Function
Public Function OpenPipedProcessAs(ByVal Path As String, ByVal User As String, ByVal Arguments As String, Optional ByVal Environment As String) As Process
    
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

''<summary>Returns the parent process of the current process.</summary>
''<retval>A process which representing the current process.</retval>
Public Property Get Current() As Process: Set Current = FromProcessID(API_GetCurrentProcessId): End Property
Public Property Get CurrentProcessID() As Long: CurrentProcessID = API_GetCurrentProcessId: End Property

Public Property Get CurrentHandle() As Win32Handle
    Set CurrentHandle = New Win32Handle
    Call CurrentHandle.Constructor0(API_GetCurrentProcess)
End Property

Public Function CreateHandle(ByVal Process As Process, ByVal dwDesiredAccess As Long, Optional ByVal bInheritHandle As Long = False) As Win32Handle
    If Process Is Nothing Then throw Exps.ArgumentNullException("Process")
    Dim pHndl As Long
    pHndl = API_OpenProcess(dwDesiredAccess, bInheritHandle, Process.ProcessID)
    If pHndl = NO_VALUE Then throw Exps.IfError
    Set CreateHandle = New Win32Handle
    Call CreateHandle.Constructor0(pHndl)
    Call GC.AddToFamilyChain(CreateHandle, Process)
End Function

''<summary>
''  Calling this method is equal to calling Process.Current.EndProcess(ExitCode)
''</summary>
''<params><param name="ExitCode">Determines the exit code for the process.</param></params>
Public Sub EndProcess(Optional ByVal ExitCode As Long = SUCCESSFULL)
    If mint_application.IsDesignMode Then
        Call EndDebugging
    Else
        Call Current.EndProcess(ExitCode)
    End If
End Sub


Public Function CountProcesses() As Long
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As API_PROCESSENTRY32
    Dim RetVal As Long
    
    ' Create snapshot
    lSnapShot = API_CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, vbNullPtr)
    If lSnapShot <> INVALID_HANDLE_VALUE Then
        ' Length of the structure
        tPE.dwSize = Len(tPE)
        ' Find first process
        lNextProcess = API_Process32First(lSnapShot, tPE)
        While lNextProcess
            RetVal = RetVal + 1
            ' Get next process
            lNextProcess = API_Process32Next(lSnapShot, tPE)
        Wend
        ' Close handle
        Call API_CloseHandle(lSnapShot)
    Else
        throw Exps.SystemCallFailureException
    End If
    CountProcesses = RetVal
End Function
''<summary>Enumerates the system processes.</summary>
''<retval>A vector of system processes.</retval>
''<remarks>
''  This also returns system processes. ex:winlogon.exe
''</remarks>
Public Function EnumProcesses() As Vector
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As API_PROCESSENTRY32
    Dim pList As New List
    Dim cProc As Process

    ' Create snapshot
    lSnapShot = API_CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, vbNullPtr)
    If lSnapShot <> INVALID_HANDLE_VALUE Then
        ' Length of the structure
        tPE.dwSize = Len(tPE)
        ' Find first process
        lNextProcess = API_Process32First(lSnapShot, tPE)
        While lNextProcess
            Set cProc = New Process
            Call cProc.Constructor3(tPE)
            Call pList.Append(cProc)
            ' Get next process
            lNextProcess = API_Process32Next(lSnapShot, tPE)
        Wend
        ' Close handle
        Call API_CloseHandle(lSnapShot)
    Else
        throw Exps.SystemCallFailureException
    End If

    Set EnumProcesses = pList.ToVector
End Function

Public Function WaitForInputIdle(ByVal Process As Process, ByVal Timeout As Long) As WaitForInputIdleState
    
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
