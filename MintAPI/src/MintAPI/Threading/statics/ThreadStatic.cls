VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ThreadStatic" static tags="">
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
Const CLASSID As String = "ThreadStatic" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const THREAD_WAITFOR_DESIREACCESS As Long = THREAD_SUSPEND_RESUME
Const THREAD_WAITFOR_INHERITHANDLES As Boolean = False
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
'<default/>
Public Function Create(ByVal Delegator As Method) As Thread
Attribute Create.VB_UserMemId = 0
    Dim Thread As New Thread
    Call Thread.Constructor0(Delegator)
    Set Create = Thread
End Function
Public Function FromHandle(ByVal Handle As Long) As Thread
    Dim Thread As New Thread
    Call Thread.Constructor1(Handle)
    Set FromHandle = Thread
End Function
Public Function FromThreadID(ByVal ThreadID As Long) As Thread
    Dim Thread As New Thread
    Call Thread.Constructor2(ThreadID)
    Set FromThreadID = Thread
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

''<summary>Represents the current executing thread context.</summary>
''<retval>Returns the current thread.</retval>
Public Property Get Current() As Thread: Set Current = FromThreadID(API_GetCurrentThreadId): End Property
Public Property Get CurrentThreadID() As Long: CurrentThreadID = API_GetCurrentThreadId: End Property

Public Property Get CurrentHandle() As Win32Handle
    Set CurrentHandle = New Win32Handle
    Call CurrentHandle.Constructor0(API_GetCurrentThread)
End Property

Public Function CreateHandle(ByVal Thread As Thread, ByVal dwDesiredAccess As Long, Optional ByVal bInheritHandle As Long = False) As Win32Handle
    If Thread Is Nothing Then throw Exps.ArgumentNullException("Thread")
    Dim pHndl As Long
    pHndl = API_OpenThread(dwDesiredAccess, bInheritHandle, Thread.ThreadID)
    If pHndl = NO_VALUE Then throw Exps.IfError
    Set CreateHandle = New Win32Handle
    Call CreateHandle.Constructor0(pHndl)
    Call GC.AddToFamilyChain(CreateHandle, Thread)
End Function

''<summary>Provides sleep for the current thread.</summary>
''<remarks>The thread freezes if the Milliseconds argument set to infinite.</remarks>
Public Sub Sleep(ByVal Milliseconds As Long): Call API_Sleep(Milliseconds): End Sub

''<summary>
''  The Yield function causes the calling thread
''  to yield execution to another thread that is
''  ready to run on the current processor.
''  The operating system selects the thread to yield to.
''</summary>
Public Function Yield() As Boolean
    Yield = CBool(API_SwitchToThread) 'Or API_Sleep(0)
End Function

Public Sub ResetAbort()

End Sub

''<summary>Starts the concurrent execution of the given delegate, the method must have no arguments and return void [Sub methodname()].</summary>
''<remarks>The process may run into crash if the delegator's given method has arguments of return value.</remarks>
''<delegats><delegate arg="Delegator" retval="void"></delegate></delegates>
Public Function Start(ByVal Delegator As Method, ByVal Args As ArgumentList) As Thread
    Set Start = Create(Delegator)
    
    Call Start.Start(Args)
End Function

'Public Sub Kill(Optional ByVal ExitCode As Long = 0)
'    Dim hndl As Long
'    'hndl = API_GetExitCodeThread(Thread.Current.Handle, ExitCode)
'    Call API_ExitThread(hndl)
'End Sub

'''<summary>
'''  Reads the value of a field. The value is the latest written by any processor
'''  in a computer, regardless of the number of processors or the state of processor
'''  cache.
'''</summary>
'Public Function VolatileRead(ByRef Value As Variant) As Variant
'
'End Function
'''<summary>
'''  Writes a value to a field immediately, so that the value is visible to all
'''  processors in the computer.
'''</summary>
'Public Function VolatileWrite(ByRef Value As Variant) As Variant
'
'End Function


'Public Sub EnableLibraryCalls()
'
'End Sub
'Public Sub DisableLibraryCalls()
'
'End Sub

''<summary>
''  Notifies a host that execution is about to enter a region of code in which
''  the effects of a thread abort or unhandled exception might jeopardize other
''  tasks in the application domain.
''</summary>
Public Sub BeginCriticalRegion()
    Call API_EnterCriticalSection(0)
End Sub
''<summary>
''  Notifies a host that execution is about to enter a region of code in which
''  the effects of a thread abort or unhandled exception are limited to the current
''  task.
''</summary>
Public Sub EndCriticalRegion()
    Call API_LeaveCriticalSection(0)
End Sub

''<summary>Enumerates all threads of the current process.</summary>
''<retval>A vector of all threads in the current process.</retval>
Public Function EnumThreads() As Vector
    Set EnumThreads = Process.Current.Threads
End Function

Public Sub WaitForSingleThread(ByVal Thread As Thread, Optional ByVal Milliseconds As Long = C_INFINITE)
    If Thread Is Nothing Then throw Exps.ArgumentNullException("Thread")
    With UsingH(CreateHandle(Thread, THREAD_WAITFOR_DESIREACCESS, THREAD_WAITFOR_INHERITHANDLES))
        If API_WaitForSingleObject(.Handle, Milliseconds) <> NO_VALUE Then _
            throw Exps.IfError
    End With
End Sub
Public Sub WaitForSingleObject(ByVal ObjectHandle As Long, Optional ByVal Milliseconds As Long = C_INFINITE)
    If ObjectHandle = 0 Then throw Exps.InvalidArgumentException("ObjectHandle")
    If API_WaitForSingleObject(ObjectHandle, Milliseconds) <> NO_VALUE Then _
        throw Exps.IfError
End Sub
Public Sub WaitForMultipleThreads(Threads() As Thread, ByVal WaitAll As Boolean, Optional ByVal Milliseconds As Long = C_INFINITE)
    If IsNullOrEmpty(Threads) Then throw Exps.ArgumentNullException("Threads")
    Dim Disps As New Disposables, Thread  As Thread, VThread As Variant
    Dim Handles() As Long, Size As Long, i As Long, ThreadHandle As Win32Handle
    
    Size = ArraySize(Threads)
    If Size = 0 Then Exit Sub
    
    Handles = Arrays.CreateSafeVector(atLongArray, Size)
    i = 0
    For Each VThread In Threads
        Set Thread = VThread
        Set ThreadHandle = MintAPI.Thread.CreateHandle(Thread, THREAD_WAITFOR_DESIREACCESS, THREAD_WAITFOR_INHERITHANDLES)
        Call Disps.Enqueue(ThreadHandle)
        Handles(i) = ThreadHandle.Handle
        i = i + 1
    Next
    
    With Using(Disps)
        If API_WaitForMultipleObjects(Size, VarPtr(Handles(0)), WaitAll, Milliseconds) <> NO_VALUE Then _
            throw Exps.IfError
    End With
End Sub
Public Sub WaitForMultipleObjects(Handles() As Long, ByVal WaitAll As Boolean, Optional ByVal Milliseconds As Long = C_INFINITE)
    If IsNullOrEmpty(Handles) Then throw Exps.ArgumentNullException("Threads")
    
    Dim Size As Long
    Size = ArraySize(Handles)
    If Size = 0 Then Exit Sub
    
    If API_WaitForMultipleObjects(Size, VarPtr(Handles(0)), WaitAll, Milliseconds) <> NO_VALUE Then _
        throw Exps.IfError
End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
