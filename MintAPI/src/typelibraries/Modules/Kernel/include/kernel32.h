#ifndef __KERNEL32_H__
#define __KERNEL32_H__

#pragma pack(4)


typedef struct API_SYSTEM_INFO {
    Long dwOemID;
    Long dwPageSize;
    Long lpMinimumApplicationAddress;
    Long lpMaximumApplicationAddress;
    Long dwActiveProcessorMask;
    Long dwNumberOfProcessors;
    Long dwProcessorType;
    Long dwAllocationGranularity;
    Long dwReserved;
} API_SYSTEM_INFO;

typedef struct API_SYSTEMTIME {
	Integer wYear;
	Integer wMonth;
	Integer wDayOfWeek;
	Integer wDay;
	Integer wHour;
	Integer wMinute;
	Integer wSecond;
	Integer wMilliseconds;
} API_SYSTEMTIME;


#pragma pack()
[
    dllname("Kernel32.dll"),
    helpstring("Access to API functions within the Kernel32.dll system file.")
]
module KernelAPI {
[entry("GetLastError")]
    long API_GetLastError();
[entry("SetLastError")]
    void API_SetLastError([in] long dwErrCode);
//========================================
[entry("RtlMoveMemory"), usesgetlasterror]
    void API_CopyMemory([in] Any Destination, [in] Any Source, [in] long Length);
[entry("RtlMoveMemory"), usesgetlasterror]
    void memcpy([in] Any Destination, [in] Any Source, [in] long Length);
//========================================
[entry("RtlZeroMemory"), usesgetlasterror]
    void API_ZeroMemory([in] Any Destination, [in] long Length);
[entry("RtlZeroMemory"), usesgetlasterror]
    void memzero([in] Any Destination, [in] long Length);
//========================================
[entry("RtlMoveMemory"), usesgetlasterror]
    void API_MoveMemory([in] Any Destination, [in] Any Source, [in] long Length);
[entry("RtlMoveMemory"), usesgetlasterror]
    void memmove([in] Any Destination, [in] Any Source, [in] long Length);
//========================================
[entry("IsBadCodePtr"), usesgetlasterror]
    long API_IsBadCodePtr([in] Any lpfn);
[entry("IsBadStringPtrA"), usesgetlasterror]
    long API_IsBadStringPtr([in] String lpsz, [in] long ucchMax);
[entry("IsBadStringPtrW"), usesgetlasterror]
    long API_IsBadStringPtrUnicode([in] WString lpsz, [in] long ucchMax);
[entry("IsBadHugeReadPtr"), usesgetlasterror]
    long API_IsBadHugeReadPtr([in] Any lp, [in] long ucb);
[entry("IsBadHugeWritePtr"), usesgetlasterror]
    long API_IsBadHugeWritePtr([in] Any lp, [in] long ucb);
[entry("IsBadReadPtr"), usesgetlasterror]
    long API_IsBadReadPtr([in] Any lp, [in] long ucb);
[entry("IsBadWritePtr"), usesgetlasterror]
    long API_IsBadWritePtr([in] Any lp, [in] long ucb);
//========================================
[entry("VirtualProtect"), usesgetlasterror]
    long API_VirtualProtect([in] Any lpAddress, [in] long dwSize, [in] long flNewProtect, [out] long* lpflOldProtect);
[entry("VirtualProtectEx"), usesgetlasterror]
    long API_VirtualProtectEx([in] long hProcess, [in] Any lpAddress, [in] long dwSize, [in] long flNewProtect, [out] long* lpflOldProtect);
[entry("VirtualAlloc"), usesgetlasterror]
    long API_VirtualAlloc([in] long lpAddress, [in] long dwSize, [in] long flAllocationType, [in] long flProtect);
[entry("VirtualAllocEx"), usesgetlasterror]
    long API_VirtualAllocEx([in] long hProcess, [in] long lpAddress, [in] long dwSize, [in] long flAllocationType, [in] long flProtect);
//[entry("VirtualCopy"), usesgetlasterror]
//    long API_VirtualCopy([in] long lpvDest, [in] long lpvSrc, [in] long cbSize, [in] long fdwProtect);
[entry("VirtualFree"), usesgetlasterror]
    long API_VirtualFree([in] long lpAddress, [in] long dwSize, [in] long dwFreeType);
[entry("VirtualFreeEx"), usesgetlasterror]
    long API_VirtualFreeEx([in] long hProcess, [in] long lpAddress, [in] long dwSize, [in] long dwFreeType);
[entry("VirtualLock"), usesgetlasterror]
    long API_VirtualLock([in] long lpAddress, [in] long dwSize);
[entry("VirtualUnlock"), usesgetlasterror]
    long API_VirtualUnlock([in] long lpAddress, [in] long dwSize);
[entry("VirtualQuery"), usesgetlasterror]
    long API_VirtualQuery([in] long lpAddress, [in] Any lpBuffer, [in] long dwLength);
[entry("VirtualQueryEx"), usesgetlasterror]
    long API_VirtualQueryEx([in] long hProcess, [in] long lpAddress, [in] Any lpBuffer, [in] long dwLength);
//========================================
[entry("LoadLibraryA"), usesgetlasterror]
    long API_LoadLibrary([in] String lpLibFileName);
[entry("LoadLibraryW"), usesgetlasterror]
    long API_LoadLibraryUnicode([in] WString lpLibFileName);
[entry("LoadLibraryExA"), usesgetlasterror]
    long API_LoadLibraryEx([in] String lpLibFileName, [in] long hFile, [in] long dwFlags);
[entry("LoadLibraryExW"), usesgetlasterror]
    long API_LoadLibraryExUnicode([in] WString lpLibFileName, [in] long hFile, [in] long dwFlags);
[entry("GetProcAddress"), usesgetlasterror]
    long API_GetProcAddress([in] long hModule, [in] string lpProcName);
[entry("FreeLibrary"), usesgetlasterror]
    long API_FreeLibrary([in] long hLibModule);
[entry("FreeLibraryAndExitThread"), usesgetlasterror]
    long API_FreeLibraryAndExitThread([in] long hLibModule, [in] long dwExitCode);
//[entry("LoadModule"), usesgetlasterror]
//    long API_LoadModule([in] String lpModuleName, [in] Any lpParameterBlock);
//========================================
[entry("RaiseException"), usesgetlasterror]
    void API_RaiseException([in] long dwExceptionCode, [in] long dwExceptionFlags, [in] long nNumberOfArguments, [in] Any lpArguments);
[entry("SetErrorMode"), usesgetlasterror]
    long API_SetErrorMode([in] long uMode/*process error mode*/);
//========================================
[entry("CloseHandle"), usesgetlasterror]
    long API_CloseHandle([in] long hObject);
[entry("DuplicateHandle"), usesgetlasterror]
    long API_DuplicateHandle([in] long hSourceProcessHandle, [in] long hSourceHandle, [in] long hTargetProcessHandle, [out] long* lpTargetHandle, [in] long dwDesiredAccess, [in] Boolean bInheritHandle, [in] long dwOptions);
[entry("SetHandleCount"), usesgetlasterror]
    long API_SetHandleCount([in] long wNumber);
[entry("SetHandleInformation"), usesgetlasterror]
    long API_SetHandleInformation([in] long hObject, [in] long dwMask, [in] long dwFlags);
//========================================
[entry("Beep"), usesgetlasterror]
    long API_Beep([in] long dwFreq, [in] long dwDuration);
[entry("GetTickCount"), usesgetlasterror]
    long API_GetTickCount();
//========================================
[entry("GetVersionExA"), usesgetlasterror]
    long API_GetVersionEx([in] Any lpVersionInfo);
[entry("GetVersionExW"), usesgetlasterror]
    long API_GetVersionExUnicode([in] Any lpVersionInfo);
//========================================
[entry("GetSystemInfo"), usesgetlasterror]
    long API_GetSystemInfo([in] API_SYSTEM_INFO* lpSystemInfo);
[entry("GetHandleInformation"), usesgetlasterror]
    long API_GetHandleInformation([in] long hObject, [out] long* lpdwFlags);
[entry("GetUserDefaultLCID")]//no error!
    long API_GetUserDefaultLCID();
//========================================
[entry("GetSystemTime"), usesgetlasterror]
    void API_GetSystemTime([in] API_SYSTEMTIME* lpSystemTime);
[entry("GetSystemTimeAdjustment"), usesgetlasterror]
    long API_GetSystemTimeAdjustment([in] long lpTimeAdjustment, [in] long lpTimeIncrement, [in] long lpTimeAdjustmentDisabled);
//========================================
[entry("IsWow64Process"), usesgetlasterror]
    long API_IsWow64Process([in] long hProcess, [out] long* Wow64Process);
};

#endif //__KERNEL32_H__