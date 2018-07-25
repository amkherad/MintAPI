Attribute VB_Name = "mint_threading"
Option Explicit

Public Const THREAD_CREATE_SUSPENDED As Long = &H4

Private Const TLS_START_ADDRESS As Long = &HE10
Private Const TLS_ARRAY As Long = &H2C
Private Const THREAD_STACKSIZE As Long = 8 * CB_KILOBYTES

Public Const THREAD_CONTEXT_i386               As Long = &H10000 ' this assumes that i386 and
Public Const THREAD_CONTEXT_i486               As Long = &H10000 ' i486 have identical context records

Public Const THREAD_CONTEXT_CONTROL            As Long = (THREAD_CONTEXT_i386 Or &H1)  ' SS:SP, CS:IP, FLAGS, BP
Public Const THREAD_CONTEXT_INTEGER            As Long = (THREAD_CONTEXT_i386 Or &H2)  ' AX, BX, CX, DX, SI, DI
Public Const THREAD_CONTEXT_SEGMENTS           As Long = (THREAD_CONTEXT_i386 Or &H4)  ' DS, ES, FS, GS
Public Const THREAD_CONTEXT_FLOATING_POINT     As Long = (THREAD_CONTEXT_i386 Or &H8)  ' 387 state
Public Const THREAD_CONTEXT_DEBUG_REGISTERS    As Long = (THREAD_CONTEXT_i386 Or &H10) ' DB 0-3,6,7
Public Const THREAD_CONTEXT_EXTENDED_REGISTERS As Long = (THREAD_CONTEXT_i386 Or &H20) ' cpu specific extensions

Public Const THREAD_CONTEXT_FULL               As Long = (THREAD_CONTEXT_CONTROL Or THREAD_CONTEXT_INTEGER Or THREAD_CONTEXT_SEGMENTS)
Public Const THREAD_CONTEXT_ALL                As Long = (THREAD_CONTEXT_CONTROL Or THREAD_CONTEXT_INTEGER Or THREAD_CONTEXT_SEGMENTS Or THREAD_CONTEXT_FLOATING_POINT Or THREAD_CONTEXT_DEBUG_REGISTERS Or THREAD_CONTEXT_EXTENDED_REGISTERS)


'Buffer to store main thread's TLS data.
Private Const TLS_SLOTS_LENGTH As Long = 256 'Maximum is 1088
Private TLSData(TLS_SLOTS_LENGTH) As Byte

Private Type ThreadCallbackDataTransformer
    Thread As Thread
    Args As ArgumentList
End Type

Private SemaphoreHandle As Long

Public Function Await(ByVal AsyncResult As IAsyncResult) As Variant
    
End Function

Public Sub Async()
    'Call mint_thread_initialize
End Sub

Public Sub LockObj(ByVal Obj As IObject)
    If Obj Is Nothing Then throw Exps.ArgumentNullException
    Call API_EnterCriticalSection(Obj.MetaObject.Synchronization.SyncHandle)
End Sub
Public Function TryLockObj(ByVal Obj As IObject) As Boolean
    If Obj Is Nothing Then throw Exps.ArgumentNullException
    If API_TryEnterCriticalSection(Obj.MetaObject.Synchronization.SyncHandle) <> NO_VALUE Then _
        Exit Function
    TryLockObj = True
End Function
Public Sub EndLockObj(ByVal Obj As IObject)
    If Obj Is Nothing Then throw Exps.ArgumentNullException
    Call API_LeaveCriticalSection(Obj.MetaObject.Synchronization.SyncHandle)
End Sub


Public Function mint_thread_create(ByVal Thread As Thread, _
        ByVal Args As ArgumentList, _
        Optional ByVal CreateSuspended As Boolean = False)
    
    Dim TID As Long, THNDL As Long
    
    Call memcpy(TLSData(0), ByVal (mHelper.ReadFS + TLS_START_ADDRESS), TLS_SLOTS_LENGTH)
    
    Call mint_thread_init_semaphore
    
    Dim CFlags As Long
    If CreateSuspended Then CFlags = THREAD_CREATE_SUSPENDED
    
    Dim Transformer As ThreadCallbackDataTransformer
    Set Transformer.Thread = Thread
    Set Transformer.Args = Args
    
    Dim TransformerPtr As Long
    TransformerPtr = Memory.FastAllocate(LenB(Transformer))
    
    Call memcpy(ByVal TransformerPtr, Transformer, VLEN_PTR)
    
    THNDL = API_CreateThread(ByVal 0&, THREAD_STACKSIZE, AddressOf mint_threadproc, ByVal TransformerPtr, CFlags, TID)
    If THNDL = vbNullPtr Then throw Exps.IfError
    
    'tsRunning must set before CreateThread() but we can ignore this,
    ' the result is about one in million querying thread's information about state may fail. :D
    Call Thread.SetThreadInfo(TID, THNDL, tsRunning)
    
    mint_thread_create = THNDL
End Function

Private Function mint_thread_init_semaphore() As Long
    If SemaphoreHandle = vbNullPtr Then
        SemaphoreHandle = API_CreateSemaphore(ByVal 0, 1, 1, Null)
        If Not SemaphoreHandle Then throw Exps.IfError
    End If
End Function
Private Function mint_thread_initialize() As Long
'Copy main thread's TLS Data into our TLS. And return TLS Address i.e. FS[E10]
    mint_thread_initialize = mHelper.ReadFS + TLS_START_ADDRESS
    Call memcpy(ByVal (mint_thread_initialize), TLSData(0), TLS_SLOTS_LENGTH)
End Function

Private Function mint_threadproc(ByRef ThreadData As ThreadCallbackDataTransformer) As Long
    'Thread.Method.Reference
    Dim TLS_Start As Long, Method As Method
    'Initialize the thread i.e. copy TLS Data
    TLS_Start = mint_thread_initialize
    
    Set Method = ThreadData.Thread.Routine
    
End Function
