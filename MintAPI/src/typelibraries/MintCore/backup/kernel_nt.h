#ifndef __KERNEL_NT_H__
#define __KERNEL_NT_H__

typedef enum API_ThreadInformationClass {
    ticThreadBasicInformation               = 0,
    ticThreadTimes                          = 1,
    ticThreadPriority                       = 2,
    ticThreadBasePriority                   = 3,
    ticThreadAffinityMask                   = 4,
    ticThreadImpersonationToken             = 5,
    ticThreadDescriptorTableEntry           = 6,
    ticThreadEnableAlignmentFaultFixup      = 7,
    ticThreadEventPair                      = 8,
    ticThreadQuerySetWin32StartAddress      = 9,
    ticThreadZeroTlsCell                    = 10,
    ticThreadPerformanceCount               = 11,
    ticThreadAmILastThread                  = 12,
    ticThreadIdealProcessor                 = 13,
    ticThreadPriorityBoost                  = 14,
    ticThreadSetTlsArrayAddress             = 15,
    ticThreadIsIoPending                    = 16,
    ticThreadHideFromDebugger               = 17
} API_ThreadInformationClass;

typedef enum API_ThreadInformationClass_SizeOfClasses {
    ticsoc_ThreadBasicInformation           = 0x1C,
    ticsoc_ThreadTimes                      = 0x20,
    ticsoc_ThreadPriority                   = 0x4,
    ticsoc_ThreadBasePriority               = 0x4,
    ticsoc_ThreadAffinityMask               = 0x4,
    ticsoc_ThreadImpersonationToken         = 0x4,
    ticsoc_ThreadDescriptorTableEntry       = 0xC,
    ticsoc_ThreadEnableAlignmentFaultFixup  = 0x1,
    ticsoc_ThreadEventPair                  = 0x4,
    ticsoc_ThreadQuerySetWin32StartAddress  = 0x4,
    ticsoc_ThreadZeroTlsCell                = 0x4,
    ticsoc_ThreadPerformanceCount           = 0x8,
    ticsoc_ThreadAmILastThread              = 0x4,
    ticsoc_ThreadIdealProcessor             = 0x4,
    ticsoc_ThreadPriorityBoost              = 0x4,
    ticsoc_ThreadSetTlsArrayAddress         = 0x4,
    ticsoc_ThreadIsIoPending                = 0x0,  //Not implemented - STATUS_INVALID_INFO_CLASS.
    ticsoc_ThreadHideFromDebugger           = 0x0,  //Not implemented - STATUS_INVALID_INFO_CLASS.
} API_ThreadInformationClass_SizeOfClasses;

#pragma pack(4)

typedef struct API_CLIENT_ID {
    long            UniqueProcess;
    long            UniqueThread;
} API_CLIENT_ID;

typedef struct API_THREAD_BASIC_INFORMATION {
    long            ExitStatus;
    long            TebBaseAddress;
    API_CLIENT_ID   ClientId;
    long            AffinityMask;
    long            Priority;
    long            BasePriority;
    long            Reserved; //added by me.
} API_THREAD_BASIC_INFORMATION;

typedef struct API_THREAD_TIMES_INFORMATION {
    Int64           CreationTime;
    Int64           ExitTime;
    Int64           KernelTime;
    Int64           UserTime;
} API_THREAD_TIMES_INFORMATION;

#pragma pack()
[
    dllname("NtDll.dll"),
    helpstring("Access to console API functions within the NtDll.dll system file.")
]
module KernelNt {
[entry("NtQueryInformationThread"), usesgetlasterror]
    long API_NtQueryInformationThread([in] long ThreadHandle, [in] API_ThreadInformationClass ThreadInformationClass, [out] Any ThreadInformation, [in] long ThreadInformationLength, [out] long* ReturnLength); 
};

#endif //__KERNEL_NT_H__