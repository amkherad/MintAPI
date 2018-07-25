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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Member1" ,"Stack"
Attribute VB_Ext_KEY = "Member2" ,"Process"
'@PROJECT_LICENSE

''<class name="Thread" tags="">
''  <summary>
''    Represents a thread.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>ThreadStatic , Method</using>
''  <idea>Me , Qt , .net Framework</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Thread" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Public Enum ThreadState
    tsNone = 0
    tsInitializing
    tsSuspended
    tsWaiting
    tsRunning
    tsTerminated
End Enum

'Private Type CONTEXT
'    FltF0 As Double
'    FltF1 As Double
'    FltF2 As Double
'    FltF3 As Double
'    FltF4 As Double
'    FltF5 As Double
'    FltF6 As Double
'    FltF7 As Double
'    FltF8 As Double
'    FltF9 As Double
'    FltF10 As Double
'    FltF11 As Double
'    FltF12 As Double
'    FltF13 As Double
'    FltF14 As Double
'    FltF15 As Double
'    FltF16 As Double
'    FltF17 As Double
'    FltF18 As Double
'    FltF19 As Double
'    FltF20 As Double
'    FltF21 As Double
'    FltF22 As Double
'    FltF23 As Double
'    FltF24 As Double
'    FltF25 As Double
'    FltF26 As Double
'    FltF27 As Double
'    FltF28 As Double
'    FltF29 As Double
'    FltF30 As Double
'    FltF31 As Double
'
'    IntV0 As Double
'    IntT0 As Double
'    IntT1 As Double
'    IntT2 As Double
'    IntT3 As Double
'    IntT4 As Double
'    IntT5 As Double
'    IntT6 As Double
'    IntT7 As Double
'    IntS0 As Double
'    IntS1 As Double
'    IntS2 As Double
'    IntS3 As Double
'    IntS4 As Double
'    IntS5 As Double
'    IntFp As Double
'    IntA0 As Double
'    IntA1 As Double
'    IntA2 As Double
'    IntA3 As Double
'    IntA4 As Double
'    IntA5 As Double
'    IntT8 As Double
'    IntT9 As Double
'    IntT10 As Double
'    IntT11 As Double
'    IntRa As Double
'    IntT12 As Double
'    IntAt As Double
'    IntGp As Double
'    IntSp As Double
'    IntZero As Double
'
'    Fpcr As Double
'    SoftFpcr As Double
'
'    Fir As Double
'    Psr As Long
'
'    ContextFlags As Long
'    Fill(4) As Long
'End Type
Public Type StackBoundries
    StackTop As Long
    StackBottom As Long
End Type


'Private Declare Function API_SetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Private Declare Function API_KillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Function API_WaitForInputIdle Lib "user32" Alias "WaitForInputIdle" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function API_GetThreadId Lib "ntdll.dll" Alias "GetThreadId" (ByVal Thread As Long) As Long

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables/Events>
'
Public Event ThreadInvoked(ByVal E As EventArgs)
Public Event ThreadBeginInvoked(ByVal E As EventArgs)
Public Event ThreadExit(ByVal E As EventArgs)

Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Private TID As Long
Private THNDL As Long
Private PID As Long
Private lpfn As Method
Private lpfn_Callback As Method

Private p_State As ThreadState
Private p_Name As String
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
''  <carg type="long" name="tMethod"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal tMethod As Method)
    Set lpfn = tMethod
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByVal ThreadHandle As Long)
    TID = GetThreadID(ThreadHandle)
    If TID = vbNullPtr Then throw Exps.IfError
    
    This_Constructed = True
End Sub
Friend Sub Constructor2(ByVal ThreadID As Long) 'Unsafe if thread being closed.
    If TID = vbNullPtr Then throw Exps.IfError
    TID = ThreadID
    
    This_Constructed = True
End Sub
Friend Sub Constructor3(ByVal ThreadID As Long) 'Safe anyway.
    If TID = vbNullPtr Then throw Exps.IfError
    TID = ThreadID
    
    This_Constructed = True
End Sub
Friend Sub Constructor4(TEntry As API_THREADENTRY32)
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal Thread As Thread)
    Set lpfn = Thread.Routine
    TID = Thread.ThreadID
    p_State = Thread.State
    
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

Public Property Get Name() As String
    Name = p_Name
End Property
Public Property Let Name(ByVal Value As String)
    p_Name = Value
End Property

''<summary>Represents the thread ID.</summary>
''<retval>Returns a number to control the thread.</retval>
Public Property Get ThreadID() As Long
    ThreadID = TID
End Property

''<summary>Represents the thread owner process.</summary>
''<retval>Returns the thread owner process class.</retval>
Public Property Get Process() As Process
    Set Process = New Process
    Call Process.Constructor1(ProcessID)
End Property
''<summary>Represents the thread owner process ID.</summary>
''<retval>Returns the thread owner process ID.</retval>
Friend Property Get ProcessID() As Long
    If PID = 0 Then
        ProcessID = API_GetCurrentProcessId
    Else
        ProcessID = PID
    End If
End Property

Public Property Get State() As ThreadState
    State = p_State
End Property
Public Property Get Apartment() As ThreadApartment
    
End Property
Public Property Get IsAlive() As Boolean
    
End Property

Public Property Get IsCurrentThread() As Boolean
    IsCurrentThread = (TID = Thread.CurrentThreadID)
End Property

Public Property Get IsBackground() As Boolean
    
End Property
Public Property Let IsBackground(ByVal Value As Boolean)
    
End Property

'Public Function Invoke(Optional Args) As Long
'    Dim tS As Method
'    If tS Is Nothing Then throw Exps.InvalidCallException("Target CallBack Method Is Not Set.")
'    If Not tS.Executable Then throw Exps.InvalidCallException("Target CallBack Method Is Not Executable.")
'    Dim Addr As Long
'    Addr = tS.Handle
'    t_cH = API_CreateThread(ByVal 0, 0, ByVal Addr, ByVal 0, 0, tID)
'    If t_cH = 0 Then
'        throw Exps.Exception("Failed To Invoke Thread With Error [" & API_GetLastError & "]")
'    Else
'        RaiseEvent ThreadInvoked(EventArgs(Me))
'    End If
'    Call API_WaitForSingleObject(tID, C_INFINITE)
'End Function

Public Sub SuspendThread()
    With UsingH(Thread.CreateHandle(Me, THREAD_SUSPEND_RESUME, False))
        If API_SuspendThread(.Handle) = -1 Then throw Exps.IfError
    End With
End Sub
Public Sub ResumeThread()
    With UsingH(Thread.CreateHandle(Me, THREAD_SUSPEND_RESUME, False))
        If API_ResumeThread(.Handle) = -1 Then throw Exps.IfError
    End With
End Sub

''<summary>
''  Causes the operating system to change the
''  state of the current instance to ThreadState.Running,
''  and optionally supplies an object containing
''  data to be used by the method the thread executes.
''</summary>
Public Sub Start(ByVal Args As ArgumentList)
'    Dim i As Long ' NOP ' future extention
    
    'Check for WinNT for next assembly execution compatibilities.
    'mint_thread_initialize() doesn't compatible with other OS types yet.
    If OS.Platform <> WinNT Then throw Exps.NotSupportedException("OS not supported.")
    If OS.Version.GreaterThan(Version.Create(6, 1)) Then throw Exps.NotSupportedException("OS not supported.") 'Does not support OS version higher than Win7, I dont know anything about its structure.
    'If Environment.Is64BitOperatingSystem Then throw Exps.OsNotSupportedException("There is no 64-bit compatible assembly code.") 'also VB6 can't run on 64-bit OS, this is only for sure.
    'Check if thread already started.
    If p_State <> tsNone Then throw Exps.InvalidOperationException("Thread Already Started.")
    
    'Check if running in VB6 IDE, because if running
    ' on IDE this may cause crash, and this also cause
    ' to user data-loss.
    'So prevent to create threads in the IDE.
    If IsDesignMode Then
        Call Debugger.ShowFatalErrorAndExit("Design Mode Error", "Unable to create a thread in design mode.\nBecause VB6 IDE doesn't support free threaded environment and this may cause crash and your data being loss.", CLASSALIASSTRING, Me)
        Exit Sub
    End If
    
    'Indicates that thread is initializing.
    'this is un-usefull, but if asynchronously access to thread's state it may come usefull.
    p_State = tsInitializing
    
    'mint_thread_create() will automatically set the p_State to tsRunning.
    
    Call mint_thread_create(Me, Args) 'pthread_create()
End Sub

Public Property Get Routine() As Method
    Set Routine = lpfn
End Property
Public Property Get CallbackRoutine() As Method
    Set CallbackRoutine = lpfn_Callback
End Property

Public Property Get Priority() As Long
    
End Property
Public Property Let Priority(Value As Long)
    
End Property

''<summary></summary>
''<retval></retval>
Public Sub Join()
    With UsingH(Thread.CreateHandle(Me, THREAD_SUSPEND_RESUME, False))
        If API_WaitForSingleObject(.Handle, C_INFINITE) <> NO_VALUE Then _
            throw Exps.IfError
    End With
End Sub

'Public Property Get LocalStorage() As Memory
'
'End Property

Public Sub SetIdealProcessor(ByVal Processor As Object)
    
End Sub

Public Sub Abort()
    
End Sub

Public Sub Exec()
    
End Sub

Public Sub StopThread(Optional ByVal ExitCode As Long = 0)
'    Dim hndl As Long
'    hndl = API_GetExitCodeThread(Handle, ExitCode)
'    Call API_ExitThread(hndl)
'    RaiseEvent ThreadExit(EventArgs(Me))
End Sub

Public Property Get Stack() As ThreadStack
    
End Property

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

Public Function Clone() As Thread
    Set Clone = New Thread
    Call Clone.ConstructorCopy(Me)
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
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Friend Sub SetThreadInfo(ByVal mTID As Long, mTHandle As Long, ByVal TState As ThreadState)
    TID = mTID
    THNDL = mTHandle
    p_State = TState
End Sub

Friend Function GetTebBaseAddress() As Long
    Dim TBI As API_THREAD_BASIC_INFORMATION, TBILen As Long
    
    If OS.Platform = WinNT Then
        If OS.Version.GreaterThanOrEqualTo(Version.Create(5, 0)) Then
            With UsingH(Thread.CreateHandle(Me, THREAD_QUERY_INFORMATION, False))
                If API_NtQueryInformationThread(.Handle, ticThreadBasicInformation, TBI, ticsoc_ThreadBasicInformation, TBILen) <> S_OK Then
                    throw Exps.SystemCallFailureException("NtQueryInformationThread call failed.")
                End If
            End With
        End If
    End If
    
    GetTebBaseAddress = TBI.TebBaseAddress
End Function

Private Function GetThreadID(ByVal ThreadHandle As Long) As Long
    If OS.Version.Major >= 6 Then
        GetThreadID = API_GetThreadId(ThreadHandle)
    ElseIf OS.Platform = WinNT Then
        Dim TBInfo As API_THREAD_BASIC_INFORMATION, TBInfoLen As Long
        
        If API_NtQueryInformationThread(ThreadHandle, ticThreadBasicInformation, TBInfo, ticsoc_ThreadBasicInformation, TBInfoLen) <> NO_VALUE Then _
            throw Exps.IfError 'IfError
        
        GetThreadID = TBInfo.ClientId.UniqueThread
    Else
        throw Exps.SystemCallFailureException("Unable to obtain the thread ID, thread ID can only ontain in windows NT or Longhorn(and more...).")
    End If
End Function

Friend Property Get StackBounderies() As StackBoundries
    Dim TEB As API_TEB, TebBaseAddress As Long

    If IsCurrentThread Then
        TebBaseAddress = mHelper.ReadFS
    Else
        TebBaseAddress = GetTebBaseAddress
    End If
    
    Call memcpy(TEB, GetTebBaseAddress, VLEN_PTR)
    With TEB
        StackBounderies.StackBottom = .pvStackUserBase
        StackBounderies.StackTop = .pvStackUserTop
    End With
    Call memcpy(TEB, ByVal vbNullPtr, VLEN_PTR)
End Property

'Friend Function GetContext() As API_CONTEXT
'    If API_GetThreadContext(OpenNewThread(THREAD_GET_CONTEXT Or THREAD_QUERY_INFORMATION, False), GetContext) = vbNullPtr Then _
'        throw Exps.IfError
'End Function
'Friend Sub SetContext(ByRef Context As API_CONTEXT)
'    If API_SetThreadContext(OpenNewThread(THREAD_SET_CONTEXT Or THREAD_QUERY_INFORMATION, False), Context) = vbNullPtr Then _
'        throw Exps.IfError
'End Sub
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
Private Function IObject_MetaObject() As MetaObject: End Function
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
