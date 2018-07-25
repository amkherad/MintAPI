VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Thread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Thread class provides threading methods and actions."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Thread"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Thread"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Public Enum ThreadState
    tsExecuting
    tsStop
    tsWaiting
    tsRunning
    tsRun
    tsInitializing
    tsMute
    tsTerminating
    tsTerminated
    tsDisposed
    tsLocked
    tsIdle
    tsFree
    tsReadyToExecute
End Enum

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type CONTEXT
    FltF0 As Double
    FltF1 As Double
    FltF2 As Double
    FltF3 As Double
    FltF4 As Double
    FltF5 As Double
    FltF6 As Double
    FltF7 As Double
    FltF8 As Double
    FltF9 As Double
    FltF10 As Double
    FltF11 As Double
    FltF12 As Double
    FltF13 As Double
    FltF14 As Double
    FltF15 As Double
    FltF16 As Double
    FltF17 As Double
    FltF18 As Double
    FltF19 As Double
    FltF20 As Double
    FltF21 As Double
    FltF22 As Double
    FltF23 As Double
    FltF24 As Double
    FltF25 As Double
    FltF26 As Double
    FltF27 As Double
    FltF28 As Double
    FltF29 As Double
    FltF30 As Double
    FltF31 As Double

    IntV0 As Double
    IntT0 As Double
    IntT1 As Double
    IntT2 As Double
    IntT3 As Double
    IntT4 As Double
    IntT5 As Double
    IntT6 As Double
    IntT7 As Double
    IntS0 As Double
    IntS1 As Double
    IntS2 As Double
    IntS3 As Double
    IntS4 As Double
    IntS5 As Double
    IntFp As Double
    IntA0 As Double
    IntA1 As Double
    IntA2 As Double
    IntA3 As Double
    IntA4 As Double
    IntA5 As Double
    IntT8 As Double
    IntT9 As Double
    IntT10 As Double
    IntT11 As Double
    IntRa As Double
    IntT12 As Double
    IntAt As Double
    IntGp As Double
    IntSp As Double
    IntZero As Double

    Fpcr As Double
    SoftFpcr As Double

    Fir As Double
    Psr As Long

    ContextFlags As Long
    Fill(4) As Long
End Type
Private Type LDT_ENTRY
    LimitLow As Integer
    BaseLow As Integer
    HighWord As Long    ' Can use LDT_BYTES Type
End Type
Private Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type



Private Declare Function API_GetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Long
Private Declare Function API_GetCurrentThread Lib "kernel32" Alias "GetCurrentThread" () As Long

Private Declare Function API_SuspendThread Lib "kernel32" Alias "SuspendThread" (ByVal hThread As Long) As Long
Private Declare Function API_ResumeThread Lib "kernel32" Alias "ResumeThread" (ByVal hThread As Long) As Long
Private Declare Function API_CreateThread Lib "kernel32" Alias "CreateThread" (lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long 'any is SECURITY_ATTRIBUTES
Private Declare Sub API_ExitThread Lib "kernel32" Alias "ExitThread" (ByVal dwExitCode As Long)
Private Declare Function API_GetExitCodeThread Lib "kernel32" Alias "GetExitCodeThread" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function API_OpenThread Lib "kernel32" Alias "OpenThread" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As Long
Private Declare Function API_SetThreadIdealProcessor Lib "kernel32" Alias "SetThreadIdealProcessor" (ByVal hThread As Long, ByVal dwIdealProcessor As Long) As Long
Private Declare Function API_SwitchToThread Lib "kernel32" Alias "SwitchToThread" () As Long
Private Declare Function API_TerminateThread Lib "kernel32" Alias "TerminateThread" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

Private Declare Function API_GetThreadContext Lib "kernel32" Alias "GetThreadContext" (ByVal hThread As Long, lpContext As CONTEXT) As Long
Private Declare Function API_GetThreadDesktop Lib "user32" Alias "GetThreadDesktop" (ByVal dwThread As Long) As Long
Private Declare Function API_GetThreadLocale Lib "kernel32" Alias "GetThreadLocale" () As Long
Private Declare Function API_GetThreadPriority Lib "kernel32" Alias "GetThreadPriority" (ByVal hThread As Long) As Long
Private Declare Function API_GetThreadSelectorEntry Lib "kernel32" Alias "GetThreadSelectorEntry" (ByVal hThread As Long, ByVal dwSelector As Long, lpSelectorEntry As LDT_ENTRY) As Long
Private Declare Function API_GetThreadTimes Lib "kernel32" Alias "GetThreadTimes" (ByVal hThread As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Private Declare Function API_SetThreadAffinityMask Lib "kernel32" Alias "SetThreadAffinityMask" (ByVal hThread As Long, ByVal dwThreadAffinityMask As Long) As Long
Private Declare Function API_SetThreadDesktop Lib "user32" Alias "SetThreadDesktop" (ByVal hDesktop As Long) As Long
Private Declare Function API_SetThreadContext Lib "kernel32" Alias "SetThreadContext" (ByVal hThread As Long, lpContext As CONTEXT) As Long
Private Declare Function API_SetThreadLocale Lib "kernel32" Alias "SetThreadLocale" (ByVal Locale As Long) As Long
Private Declare Function API_SetThreadPriority Lib "kernel32" Alias "SetThreadPriority" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function API_SetThreadToken Lib "advapi32" Alias "SetThreadToken" (Thread As Long, ByVal Token As Long) As Long
Private Declare Function API_GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private Declare Function API_SetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function API_KillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub API_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Private Declare Function API_WaitCommEvent Lib "kernel32" Alias "WaitCommEvent" (ByVal hFile As Long, lpEvtMask As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function API_WaitForDebugEvent Lib "kernel32.dll" Alias "WaitForDebugEvent" (ByRef lpDebugEvent As Any, ByVal dwMilliseconds As Long) As Long 'any is DEBUG_EVENT
Private Declare Function API_WaitForInputIdle Lib "user32" Alias "WaitForInputIdle" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function API_WaitForMultipleObjects Lib "kernel32" Alias "WaitForMultipleObjects" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function API_WaitForMultipleObjectsEx Lib "kernel32" Alias "WaitForMultipleObjectsEx" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare Function API_WaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function API_WaitForSingleObjectEx Lib "kernel32" Alias "WaitForSingleObjectEx" (ByVal hHandle As Long, ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

Private Declare Sub API_EnterCriticalSection Lib "kernel32" Alias "EnterCriticalSection" (ByVal lpCriticalSection_dummy As Long) 'CRITICAL_SECTION { dummy as long }
Private Declare Sub API_LeaveCriticalSection Lib "kernel32" Alias "LeaveCriticalSection" (ByVal lpCriticalSection_dummy As Long) 'CRITICAL_SECTION { dummy as long }


Private Declare Function API_DisableThreadLibraryCalls Lib "kernel32" Alias "DisableThreadLibraryCalls" (ByVal hLibModule As Long) As Long

Public Event ThreadInvoked(ByVal E As EventArgs)
Public Event ThreadConcurrentInvoked(ByVal E As EventArgs)
Public Event ThreadExit(ByVal E As EventArgs)

Dim lastThreadYelpedID As Long
Dim tID As Long
Dim t_cH As Long
Dim targetFunction As Method

Public Sub Initialize(Optional ByVal ThreadID As Long = -1, Optional ByVal ThreadClassID As Long = -1, Optional ByVal targetFuncHandle As Method = Nothing)
    If ThreadID = -1 Then ThreadID = 0
    If ThreadClassID = -1 Then ThreadClassID = 0
    tID = ThreadID
    t_cH = ThreadClassID
    Set targetFunction = targetFuncHandle
End Sub

Public Property Get Handle() As Long
    If tID = 0 Then
        Handle = CurrentThread.Handle
    Else
        Handle = tID
    End If
End Property
Friend Property Get ClassHandle() As Long
    If t_cH = 0 Then
        ClassHandle = CurrentThread.ClassHandle
    Else
        ClassHandle = t_cH
    End If
End Property

Public Function CurrentThread() As Thread
    Dim t As New Thread
    Call t.Initialize(API_GetCurrentThreadId, API_GetCurrentThread)
    Set CurrentThread = t
End Function
Public Function CallingThread() As Thread: Set CallingThread = CurrentThread: End Function
Public Function Thread(ByVal ThreadID As Long) As Thread
    Dim t As New Thread
    Call t.Initialize(ThreadID)
    Set Thread = t
End Function
Public Function MintAPIThread() As Thread
    Dim t As New Thread
    Call t.Initialize(App.ThreadID)
    Set MintAPIThread = t
End Function

Public Function Equals(ByVal t1 As Thread, ByVal t2 As Thread) As Boolean
    Equals = (t1.Handle = t2.Handle)
End Function

Public Function Invoke(Optional Args, Optional ByVal targetSite As Method = Nothing) As Long
    Dim tS As Method
    If targetSite Is Nothing Then Set tS = targetFunction Else Set tS = targetSite
    If tS Is Nothing Then throw InvalidCallException("Target CallBack Method Is Not Set.")
    If Not tS.Executable Then throw InvalidCallException("Target CallBack Method Is Not Executable.")
    Dim Addr As Long
    Addr = tS.Handle
    t_cH = API_CreateThread(ByVal 0, 0, ByVal Addr, ByVal 0, 0, tID)
    If t_cH = 0 Then
        throw Exception("Failed To Invoke Thread With Error [" & API_GetLastError & "]")
    Else
        RaiseEvent ThreadInvoked(EventArgs(Me))
    End If
    Call API_WaitForSingleObject(tID, INFINITE)
End Function
Public Function ConcurrentInvoke(Optional Args, Optional ByVal targetSite As Method = Nothing) As Thread
    Dim tS As Method
    If targetSite Is Nothing Then Set tS = targetFunction Else Set tS = targetSite
    If tS Is Nothing Then throw InvalidCallException("Target site is not set.")
    If Not tS.Executable Then throw InvalidStatusException("Target CallBack Is Invalid.")
    If t_cH = 0 Then
        If Not tS.Executable Then throw InvalidCallException("Target site is not executable.")
        Dim Addr As Long
        Addr = tS.Handle
        t_cH = API_CreateThread(ByVal 0, 0, ByVal Addr, ByVal 0, 0, tID)
        If t_cH = 0 Then
            throw InvalidCallException("Unable to create thread ,module kernel32.dll.")
        Else
            RaiseEvent ThreadConcurrentInvoked(EventArgs(Me))
        End If
        Set ConcurrentInvoke = Me
    Else
        Dim t As Thread
        Set t = New Thread
        Call t.Initialize(targetFuncHandle:=tS)
        Call t.ConcurrentInvoke
        Set ConcurrentInvoke = t
    End If
End Function

Public Function Yield() As Thread
    lastThreadYelpedID = tID
    'tID = API_SwitchToThread()
    Set Yield = Me
End Function

Public Sub Sleep(ByVal Milliseconds As Long)
    Call API_Sleep(Milliseconds)
End Sub

Public Sub SuspendThread()
    Call API_SuspendThread(Handle)
End Sub
Public Sub ResumeThread()
    Call API_ResumeThread(Handle)
End Sub

Public Property Get Priority() As Long
    
End Property
Public Property Let Priority(Value As Long)
    
End Property

Public Property Get State() As ThreadState
    
End Property

Public Function CreateTimer(ByVal Interval As Long, Optional ByVal targetSite As Method = Nothing) As Long
    
End Function
Public Sub DistroyTimer(ByVal TimerHandle As Long)
    
End Sub

Public Sub Loud()
    
End Sub
Public Sub Quiet()
    
End Sub

Public Sub Mute()
    
End Sub
Public Sub UnMute()
    
End Sub
Public Function IsMuted() As Boolean
    
End Function

Public Sub SetIdealProcessor(ByVal Processor As Object)
    
End Sub

Public Sub LockThread()
    
End Sub
Public Sub UnlockThread()
    
End Sub
Public Function IsLocked() As Boolean
    
End Function

Public Sub EnterCriticalSection()
    Call API_EnterCriticalSection(0)
End Sub
Public Sub LeaveCriticalSection()
    Call API_LeaveCriticalSection(0)
End Sub

Public Sub WaitForSingleThread(ByVal targetThread As Thread, Optional Milliseconds As Long = INFINITE)
    If targetThread Is Nothing Then throw ArgumentNullException
    Dim Handle As Long
    Handle = targetThread.Handle
    If Handle = 0 Then throw InvalidArgumentValueException
    Call API_WaitForSingleObject(Handle, Milliseconds)
End Sub
Public Sub WaitForSingleObject(ByVal ObjectHandle As Long, Optional Milliseconds As Long = INFINITE)
    If ObjectHandle = 0 Then throw InvalidArgumentValueException
    Call API_WaitForSingleObject(ObjectHandle, Milliseconds)
End Sub
Public Sub WaitForMultipleThread(targetThread() As Thread)
    'API_WaitForMultipleObjects
End Sub
Public Sub WaitForMultipleObject(ObjectHandle() As Long)
    
End Sub
Public Sub ContinueExecuting()
    
End Sub

Public Property Get ExecutionEndMethod() As Method
    
End Property
Public Property Let ExecutionEndMethod(Value As Method)
    
End Property
Public Property Set ExecutionEndMethod(Value As Method)
    
End Property

Public Function LastExecute() As Long
    
End Function
Public Function LastException() As Exception
    
End Function

Public Function Async_retVal() As Variant
    
End Function

Public Sub ExitThread(Optional ByVal ExitCode As Long = 0)
    Dim hndl As Long
    hndl = API_GetExitCodeThread(CurrentThread.Handle, ExitCode)
    Call API_ExitThread(hndl)
    RaiseEvent ThreadExit(EventArgs(Me))
End Sub
Public Sub StopThread(Optional ByVal ExitCode As Long = 0)
    Dim hndl As Long
    hndl = API_GetExitCodeThread(Handle, ExitCode)
    Call API_ExitThread(hndl)
    RaiseEvent ThreadExit(EventArgs(Me))
End Sub
Public Sub Shutdown(): Call StopThread(0): End Sub

Public Function IsCurrentThread() As Boolean
    IsCurrentThread = (CurrentThread.Handle = Handle)
End Function


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
