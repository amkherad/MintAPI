#ifndef __SYNCHRONIZATION_H__
#define __SYNCHRONIZATION_H__

#pragma pack(4)


#pragma pack()
[
    dllname("Kernel32.dll"),
    helpstring("Access to synchronization API functions within the Kernel32.dll system file.")
]
module KernelSynch {
//========================================
[entry("CreateSemaphoreA"), usesgetlasterror]
    long API_CreateSemaphore([in] Any lpSemaphoreAttributes, [in] long lInitialCount, [in] long lMaximumCount, [in] String lpName);
[entry("OpenSemaphoreA"), usesgetlasterror]
    long API_OpenSemaphore([in] long dwDesiredAccess, [in] long bInheritHandle, [in] String lpName);
[entry("ReleaseSemaphore"), usesgetlasterror]
    long API_ReleaseSemaphore([in] long hSemaphore, [in] long lReleaseCount, [out] long* lpPreviousCount);
//========================================
[entry("CreateMutexA"), usesgetlasterror]
    long API_CreateMutex([in] Any lpMutexAttributes, [in] long bInitialOwner, [in] String lpName);
[entry("OpenMutexA"), usesgetlasterror]
    long API_OpenMutex([in] long dwDesiredAccess, [in] long bInheritHandle, [in] String lpName);
[entry("ReleaseMutex"), usesgetlasterror]
    long API_ReleaseMutex([in] long hMutex);
//========================================
[entry("EnterCriticalSection"), usesgetlasterror]
    void API_EnterCriticalSection([in] long lpCriticalSection_dummy); // CRITICAL_SECTION { dummy as long }
[entry("TryEnterCriticalSection"), usesgetlasterror]
    long API_TryEnterCriticalSection([in] long lpCriticalSection_dummy); // CRITICAL_SECTION { dummy as long }
[entry("LeaveCriticalSection"), usesgetlasterror]
    void API_LeaveCriticalSection([in] long lpCriticalSection_dummy); // CRITICAL_SECTION { dummy as long }

[entry("InitializeCriticalSection"), usesgetlasterror]
    void API_InitializeCriticalSection([in] long lpCriticalSection_dummy); // CRITICAL_SECTION { dummy as long }
[entry("DeleteCriticalSection"), usesgetlasterror]
    void API_DeleteCriticalSection([in] long lpCriticalSection_dummy); // CRITICAL_SECTION { dummy as long }

[entry("InitializeCriticalSectionAndSpinCount"), usesgetlasterror]
	long API_InitializeCriticalSectionAndSpinCount([in] long lpCriticalSection_dummy, [in] long dwSpinCount);
//========================================
[entry("InterlockedIncrement")]
    long API_InterlockedIncrement([in] long* lpAddend);
[entry("InterlockedDecrement")]
    long API_InterlockedDecrement([in] long* lpAddend);
[entry("InterlockedExchange")]
    long API_InterlockedExchange([in] long* Target, [in] long Value);
[entry("InterlockedExchangeAdd")]
    long API_InterlockedExchangeAdd([in] long* Target, [in] long Value);
[entry("InterlockedCompareExchange")]
    long API_InterlockedCompareExchange([in] Any Destination, [in] Any Exchange, [in] Any Comperand);
//========================================
[entry("WaitCommEvent"), usesgetlasterror]
    long API_WaitCommEvent([in] long hFile, [out] long* lpEvtMask, [out] Any lpOverlapped);
[entry("WaitForDebugEvent"), usesgetlasterror]
    long API_WaitForDebugEvent([out] Any lpDebugEvent, [in] long dwMilliseconds, [out] Any lpOverlapped); // 'any is DEBUG_EVENT
[entry("WaitForMultipleObjects"), usesgetlasterror]
    long API_WaitForMultipleObjects([in] long nCount, [out] long* lpHandles, [in] long bWaitAll, [in] long dwMilliseconds);
[entry("WaitForMultipleObjectsEx"), usesgetlasterror]
    long API_WaitForMultipleObjectsEx([in] long nCount, [out] long* lpHandles, [in] long bWaitAll, [in] long dwMilliseconds, [in] long bAlertable);
[entry("WaitForSingleObject"), usesgetlasterror]
    long API_WaitForSingleObject([in] long hHandle, [in] long dwMilliseconds);
[entry("WaitForSingleObjectEx"), usesgetlasterror]
    long API_WaitForSingleObjectEx([in] long hHandle, [in] long dwMilliseconds, [in] long bAlertable);

[entry("MsgWaitForMultipleObjects"), usesgetlasterror]
    long API_MsgWaitForMultipleObjects([in] long nCount, [in] long* pHandles, [in] Boolean fWaitAll, [in] long dwMilliseconds, [in] long dwWakeMask);
[entry("MsgWaitForMultipleObjectsEx"), usesgetlasterror]
    long API_MsgWaitForMultipleObjectsEx([in] long nCount, [in] long* pHandles, [in] Boolean fWaitAll, [in] long dwMilliseconds, [in] long dwWakeMask, [in] long dwFlags);
//========================================
[entry("SetEvent"), usesgetlasterror]
    Boolean API_SetEvent([in] long hEvent);
[entry("ResetEvent"), usesgetlasterror]
    Boolean API_ResetEvent([in] long hEvent);
[entry("PulseEvent"), usesgetlasterror]
    Boolean API_PulseEvent([in] long hEvent);

[entry("CreateEventA"), usesgetlasterror]
    long API_CreateEvent([in] Any lpEventAttributes, [in] Boolean bManualReset, [in] Boolean bInitialState, [in] String lpName);
[entry("CreateEventW"), usesgetlasterror]
    long API_CreateEventUnicode([in] Any lpEventAttributes, [in] Boolean bManualReset, [in] Boolean bInitialState, [in] WString lpName);

[entry("OpenEventA"), usesgetlasterror]
    long API_OpenEvent([in] long dwDesiredAccess, [in] Boolean bInheritHandle, [in] Boolean bInitialState, [in] String lpName);
[entry("OpenEventW"), usesgetlasterror]
    long API_OpenEventUnicode([in] long dwDesiredAccess, [in] Boolean bInheritHandle, [in] Boolean bInitialState, [in] WString lpName);

//========================================

};

#endif //__SYNCHRONIZATION_H__