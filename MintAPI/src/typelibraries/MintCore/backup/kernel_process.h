#ifndef __KERNEL_PROCESS_H__
#define __KERNEL_PROCESS_H__


#define MAX_MODULE_NAME32_inline 255

#pragma pack(4)
typedef struct API_PROCESSENTRY32 {
    long dwSize;
    long cntUsage;
    long th32ProcessID;
    long th32DefaultHeapID;
    long th32ModuleID;
    long cntThreads;
    long th32ParentProcessID;
    long pcPriClassBase;
    long dwFlags;
    byte szExeFile[MAX_PATH];
} API_PROCESSENTRY32;

typedef struct API_MODULEENTRY32 {
    long   dwSize;
    long   th32ModuleID;
    long   th32ProcessID;
    long   GlblcntUsage;
    long   ProccntUsage;
    long   modBaseAddr;
    long   modBaseSize;
    long   hModule;
    byte   szModule[MAX_MODULE_NAME32_inline + 1];
    byte   szExePath[MAX_PATH];
} API_MODULEENTRY32;

typedef struct API_PROCESS_MEMORY_COUNTERS {
  long    cb;
  long    PageFaultCount;
  Integer PeakWorkingSetSize;
  Integer WorkingSetSize;
  Integer QuotaPeakPagedPoolUsage;
  Integer QuotaPagedPoolUsage;
  Integer QuotaPeakNonPagedPoolUsage;
  Integer QuotaNonPagedPoolUsage;
  Integer PagefileUsage;
  Integer PeakPagefileUsage;
} API_PROCESS_MEMORY_COUNTERS;

typedef struct API_MODULEINFO {
  Long lpBaseOfDll;
  Long SizeOfImage;
  Long EntryPoint;
} API_MODULEINFO;

typedef struct API_STARTUPINFO {
  Long cb;
  Long lpReserved;
  Long lpDesktop;
  Long lpTitle;
  Long dwX;
  Long dwY;
  Long dwXSize;
  Long dwYSize;
  Long dwXCountChars;
  Long dwYCountChars;
  Long dwFillAttribute;
  Long dwFlags;
  Integer wShowWindow;
  Integer cbReserved2;
  Long lpReserved2;
  Long hStdInput;
  Long hStdOutput;
  Long hStdError;
} API_STARTUPINFO;

typedef struct API_PROCESS_INFORMATION {
  Long hProcess;
  Long hThread;
  Long dwProcessId;
  Long dwThreadId;
} API_PROCESS_INFORMATION;


typedef enum API_PROCESSENUMS {
    PROCESS_ALL_ACCESS = 0x1F0FFF,
    PROCESS_CREATE_THREAD = 0x2,
    PROCESS_CREATE_PROCESS = 0x80,
    PROCESS_DUP_HANDLE = 0x40,
    PROCESS_HEAP_ENTRY_BUSY = 0x4,
    PROCESS_HEAP_ENTRY_DDESHARE = 0x20,
    PROCESS_HEAP_ENTRY_MOVEABLE = 0x10,
    PROCESS_HEAP_REGION = 0x1,
    PROCESS_HEAP_UNCOMMITTED_RANGE = 0x2,
    PROCESS_QUERY_INFORMATION = 0x400,
    PROCESS_SET_INFORMATION = 0x200,
    PROCESS_SET_QUOTA = 0x100,
    PROCESS_TERMINATE = 0x1,
    PROCESS_VM_OPERATION = 0x8,
    PROCESS_VM_READ = 0x10,
    PROCESS_VM_WRITE = 0x20,
    
    MAX_MODULE_NAME32 = 255,
    TH32CS_INHERIT = 0x80000000,
    TH32CS_SNAPHEAPLIST = 0x00000001,
    TH32CS_SNAPPROCESS = 0x00000002,
    TH32CS_SNAPTHREAD = 0x00000004,
    TH32CS_SNAPMODULE = 0x00000008,
    TH32CS_SNAPMODULE32 = 0x00000010,
    TH32CS_SNAPALL = 0x0000000f/* TH32CS_SNAPHEAPLIST | TH32CS_SNAPPROCESS |
                     TH32CS_SNAPTHREAD | TH32CS_SNAPMODULE */,
    
    STD_ERROR_HANDLE = -12,
    STD_INPUT_HANDLE = -10,
    STD_OUTPUT_HANDLE = -11
} API_PROCESSENUMS;

#pragma pack()
[
    dllname("Kernel32.dll"),
    helpstring("Access to process API functions within the Kernel32.dll system file.")
]
module KernelProcess {
[entry("GetCurrentProcess"), usesgetlasterror]
    long API_GetCurrentProcess();
[entry("GetCurrentProcessId"), usesgetlasterror]
    long API_GetCurrentProcessId();
//========================================
[entry("GetProcessId"), usesgetlasterror]
    long API_GetProcessId([in] long Process);
//========================================
[entry("CreateToolhelp32Snapshot"), usesgetlasterror]
    long API_CreateToolhelp32Snapshot([in] long dwFlags, [in] long th32ProcessID);
[entry("Toolhelp32ReadProcessMemory"), usesgetlasterror]
    long API_Toolhelp32ReadProcessMemory([in] long th32ProcessID, [out] Any lpBaseAddress, [out] Any lpBuffer, [in] long cbRead, [out] long* lpNumberOfBytesRead);
//========================================
[entry("Process32First"), usesgetlasterror]
    long API_Process32First([in] long hSnapshot, [out] API_PROCESSENTRY32* lppe);
[entry("Process32Next"), usesgetlasterror]
    long API_Process32Next([in] long hSnapshot, [out] API_PROCESSENTRY32* lppe);
//========================================
[entry("Module32First"), usesgetlasterror]
    long API_Module32First([in] long hSnapshot, [out] API_MODULEENTRY32* lppe);
[entry("Module32Next"), usesgetlasterror]
    long API_Module32Next([in] long hSnapshot, [out] API_MODULEENTRY32* lppe);
//========================================
[entry("GetModuleHandleA"), usesgetlasterror]
    long API_GetModuleHandle([in] String lpModuleName);
[entry("GetModuleHandleW"), usesgetlasterror]
    long API_GetModuleHandleUnicode([in] WString lpModuleName);

[entry("GetModuleFileNameA"), usesgetlasterror]
    long API_GetModuleFileName([in] long hModule, [out] String lpFileName, [in] long nSize);
[entry("GetModuleFileNameW"), usesgetlasterror]
    long API_GetModuleFileNameUnicode([in] long hModule, [out] WString lpFileName, [in] long nSize);
//========================================
[entry("CreateProcessA"), usesgetlasterror]
    long API_CreateProcess([in] String lpApplicationName, [in] String lpCommandLine, [out] Any lpProcessAttributes, [out] Any lpThreadAttributes, [in] long bInheritHandles, [in] long dwCreationFlags, [out] Any lpEnvironment, [in] String lpCurrentDirectory, [out] API_STARTUPINFO* lpStartupInfo, [out] API_PROCESS_INFORMATION* lpProcessInformation);
[entry("CreateProcessW"), usesgetlasterror]
    long API_CreateProcessUnicode([in] WString lpApplicationName, [in] WString lpCommandLine, [out] Any lpProcessAttributes, [out] Any lpThreadAttributes, [in] long bInheritHandles, [in] long dwCreationFlags, [out] Any lpEnvironment, [in] WString lpCurrentDirectory, [out] API_STARTUPINFO* lpStartupInfo, [out] API_PROCESS_INFORMATION* lpProcessInformation);
//[entry("CreateProcessAsUserA"), usesgetlasterror]
//    long API_CreateProcessAsUser([in] long hToken, [in] String lpApplicationName, [in] String lpCommandLine, [out] Any lpProcessAttributes, [out] Any lpThreadAttributes, [in] long bInheritHandles, [in] long dwCreationFlags, [out] Any lpEnvironment, [in] String lpCurrentDirectory, [out] API_STARTUPINFO* lpStartupInfo, [out] API_PROCESS_INFORMATION* lpProcessInformation);
//[entry("CreateProcessAsUserW"), usesgetlasterror]
//    long API_CreateProcessAsUserUnicode([in] long hToken, [in] WString lpApplicationName, [in] WString lpCommandLine, [out] Any lpProcessAttributes, [out] Any lpThreadAttributes, [in] long bInheritHandles, [in] long dwCreationFlags, [out] Any lpEnvironment, [in] WString lpCurrentDirectory, [out] API_STARTUPINFO* lpStartupInfo, [out] API_PROCESS_INFORMATION* lpProcessInformation);
[entry("OpenProcess"), usesgetlasterror]
    long API_OpenProcess([in] long dwDesiredAccess, [in] long blnheritHandle, [in] long dwAppProcessId);
[entry("GetExitCodeProcess"), usesgetlasterror]
    long API_GetExitCodeProcess([in] long hProcess, [out] long* lpExitCode);
//========================================
[entry("ExitProcess"), usesgetlasterror]
    void API_ExitProcess([in] long uExitCode);
[entry("TerminateProcess"), usesgetlasterror]
    long API_TerminateProcess([in] long hProcess, [in] long uExitCode);
[entry("FatalExit"), usesgetlasterror]
    void API_FatalExit([in] long uExitCode);
[entry("FatalAppExitA"), usesgetlasterror]
    void API_FatalAppExit([in] long uAction, [in] String lpMessageText);
[entry("FatalAppExitW"), usesgetlasterror]
    void API_FatalAppExitUnicode([in] long uAction, [in] String lpMessageText);
//========================================
[entry("ReadProcessMemory"), usesgetlasterror]
    void API_ReadProcessMemory([in] long hProcess, [in] long lpBaseAddress, [out] Any lpBuffer, [in] long nSize, [out] long* lpNumberOfBytesReaded);
[entry("WriteProcessMemory"), usesgetlasterror]
    void API_WriteProcessMemory([in] long hProcess, [in] long lpBaseAddress, [out] Any lpBuffer, [in] long nSize, [out] long* lpNumberOfBytesWritten);
/* [entry("RegisterServiceProcess"), usesgetlasterror]
    void API_RegisterServiceProcess([in] long dwProcessId, [in] long dwType); */
//========================================
[entry("GetStartupInfoA"), usesgetlasterror]
    void API_GetStartupInfo([out] API_STARTUPINFO* lpStartupInfo);
[entry("GetStartupInfoW"), usesgetlasterror]
    void API_GetStartupInfoUnicode([out] API_STARTUPINFO* lpStartupInfo);
//========================================
[entry("GetProcessVersion"), usesgetlasterror]
    long API_GetProcessVersion([in] long dwProcessId);
};

[
    dllname("psapi.dll"),
    helpstring("Access to process API functions within the psapi.dll system file.")
]
module PSAPI {
[entry("EnumProcessModules"), usesgetlasterror]
    long API_EnumProcessModules([in] long hProcess, [out] long* lphModule, [in] long cb, [out] long* lpcbNeeded);
[entry("GetProcessMemoryInfo"), usesgetlasterror]
    long API_GetProcessMemoryInfo([in] long hProcess, [out] API_PROCESS_MEMORY_COUNTERS* ppsmemCounters, [in] long cb);
[entry("GetModuleInformation"), usesgetlasterror]
    long API_GetModuleInformation([in] long hProcess, [in] long hModule, [out] API_MODULEINFO* ppsmemCounters, [in] long cb);
[entry("GetModuleFileNameExA"), usesgetlasterror]
    long API_GetModuleFileNameEx([in] long hProcess, [in] long hModule, [out] WString lpFilename, [in] long nSize);
[entry("GetModuleFileNameExW"), usesgetlasterror]
    long API_GetModuleFileNameExUnicode([in] long hProcess, [in] long hModule, [out] WString lpFilename, [in] long nSize);
[entry("GetModuleBaseNameA"), usesgetlasterror]
    long API_GetModuleBaseName([in] long hProcess, [in] long hModule, [out] String lpBaseName, [in] long nSize);
[entry("GetModuleBaseNameW"), usesgetlasterror]
    long API_GetModuleBaseNameUnicode([in] long hProcess, [in] long hModule, [out] WString lpBaseName, [in] long nSize);
};

#endif //__KERNEL_PROCESS_H__