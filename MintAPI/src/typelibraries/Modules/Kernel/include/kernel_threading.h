#ifndef __KERNEL_THREADING_H__
#define __KERNEL_THREADING_H__


typedef enum API_THREADING_CONSTANTS {
    API_CRITICALSECTION_SIZE = 24
} API_THREADING_CONSTANTS;

typedef enum API_THREAD_FLAGS {
    THREAD_TERMINATE              = (0x0001),
    THREAD_SUSPEND_RESUME         = (0x0002),
    THREAD_GET_CONTEXT            = (0x0008),
    THREAD_SET_CONTEXT            = (0x0010),
    THREAD_SET_INFORMATION        = (0x0020),
    THREAD_QUERY_INFORMATION      = (0x0040),
    THREAD_SET_THREAD_TOKEN       = (0x0080),
    THREAD_IMPERSONATE            = (0x0100),
    THREAD_DIRECT_IMPERSONATION   = (0x0200)
} API_THREAD_FLAGS;

#define CONTEXT_i386    0x10000    // this assumes that i386 and
#define CONTEXT_i486    0x10000    // i486 have identical context records
// end_wx86
typedef enum API_CONTEXT_FLAGS_i386 {
    i386_CONTEXT_CONTROL                 = (0x10001), // SS:SP, CS:IP, FLAGS, BP
    i386_CONTEXT_INTEGER                 = (0x10002), // AX, BX, CX, DX, SI, DI
    i386_CONTEXT_SEGMENTS                = (0x10004), // DS, ES, FS, GS
    i386_CONTEXT_FLOATING_POINT          = (0x10008), // 387 state
    i386_CONTEXT_DEBUG_REGISTERS         = (0x10010), // DB 0-3,6,7
    i386_CONTEXT_EXTENDED_REGISTERS      = (0x10020), // cpu specific extensions

  //i386_CONTEXT_FULL                    = (i386_CONTEXT_CONTROL + i386_CONTEXT_INTEGER + i386_CONTEXT_SEGMENTS),
    i386_CONTEXT_FULL                    = ((0x10001)            + (0x10002)            + (0x10004)),

  //i386_CONTEXT_ALL                     = (i386_CONTEXT_CONTROL + i386_CONTEXT_INTEGER + i386_CONTEXT_SEGMENTS + i386_CONTEXT_FLOATING_POINT + i386_CONTEXT_DEBUG_REGISTERS + i386_CONTEXT_EXTENDED_REGISTERS)
    i386_CONTEXT_ALL                     = ((0x10001)            + (0x10002)            + (0x10004)             + (0x10008)                   + (0x10010)                    + (0x10020))
} API_CONTEXT_FLAGS_i386;

#define CONTEXT_AMD64   0x100000
typedef enum API_CONTEXT_FLAGS_AMD64 {
    AMD64_CONTEXT_CONTROL                 = (0x100001),
    AMD64_CONTEXT_INTEGER                 = (0x100002),
    AMD64_CONTEXT_SEGMENTS                = (0x100004),
    AMD64_CONTEXT_FLOATING_POINT          = (0x100008),
    AMD64_CONTEXT_DEBUG_REGISTERS         = (0x100010),

  //AMD64_CONTEXT_FULL                    = (AMD64_CONTEXT_CONTROL + AMD64_CONTEXT_INTEGER + AMD64_CONTEXT_FLOATING_POINT),
    AMD64_CONTEXT_FULL                    = ((0x100001)            + (0x100002)            + (0x100008)),

  //AMD64_CONTEXT_ALL                     = (AMD64_CONTEXT_CONTROL + AMD64_CONTEXT_INTEGER + AMD64_CONTEXT_SEGMENTS + AMD64_CONTEXT_FLOATING_POINT + AMD64_CONTEXT_DEBUG_REGISTERS),
    AMD64_CONTEXT_ALL                     = ((0x100001)            + (0x100002)            + (0x100004)             + (0x100008)                   + (0x100010)),
    
    AMD64_CONTEXT_EXCEPTION_ACTIVE        = 0x8000000,
    AMD64_CONTEXT_SERVICE_ACTIVE          = 0x10000000,
    AMD64_CONTEXT_EXCEPTION_REQUEST       = 0x40000000,
    AMD64_CONTEXT_EXCEPTION_REPORTING     = 0x80000000
} API_CONTEXT_FLAGS_AMD64;

#define CONTEXT_IA64                    0x00080000
typedef enum API_CONTEXT_FLAGS_IA64 {
    IA64_CONTEXT_CONTROL                 = (0x80001),
    IA64_CONTEXT_LOWER_FLOATING_POINT    = (0x80002),
    IA64_CONTEXT_HIGHER_FLOATING_POINT   = (0x80004),
    IA64_CONTEXT_INTEGER                 = (0x80008),
    IA64_CONTEXT_DEBUG                   = (0x80010),
    IA64_CONTEXT_IA32_CONTROL            = (0x80020),  // Includes StIPSR

  //IA64_CONTEXT_FLOATING_POINT          = (IA64_CONTEXT_LOWER_FLOATING_POINT + IA64_CONTEXT_HIGHER_FLOATING_POINT),
    IA64_CONTEXT_FLOATING_POINT          = ((0x80002)                         + (0x80004)),
    
  //IA64_CONTEXT_FULL                    = (IA64_CONTEXT_CONTROL + IA64_CONTEXT_FLOATING_POINT + IA64_CONTEXT_INTEGER + IA64_CONTEXT_IA32_CONTROL),
    IA64_CONTEXT_FULL                    = ((0x80001)            + ((0x80002) + (0x80004))     + (0x80008)            + (0x80020)),
    
  //IA64_CONTEXT_ALL                     = (IA64_CONTEXT_CONTROL + IA64_CONTEXT_FLOATING_POINT + IA64_CONTEXT_INTEGER + IA64_CONTEXT_DEBUG + IA64_CONTEXT_IA32_CONTROL),
    IA64_CONTEXT_ALL                     = ((0x80001)            + ((0x80002) + (0x80004))     + (0x80008)            + (0x80010)          + (0x80020)),

    IA64_CONTEXT_EXCEPTION_ACTIVE        = 0x8000000,
    IA64_CONTEXT_SERVICE_ACTIVE          = 0x10000000,
    IA64_CONTEXT_EXCEPTION_REQUEST       = 0x40000000,
    IA64_CONTEXT_EXCEPTION_REPORTING     = 0x80000000
} API_CONTEXT_FLAGS_IA64;

#pragma pack(4)
typedef struct API_THREADENTRY32 {
    long        dwSize;
    long        cntUsage;
    long        th32ThreadID;
    long        th32OwnerProcessID;
    long        tpBasePri;
    long        tpDeltaPri;
    long        dwFlags;
} API_THREADENTRY32;

typedef struct API_LDT_ENTRY {
    Integer     LimitLow;
    Integer     BaseLow;
    long        HighWord; //Can use LDT_BYTES Type
} API_LDT_ENTRY;

#pragma pack(16)


#define SIZE_OF_80387_REGISTERS      80
typedef struct API_FLOATING_SAVE_AREA {
    long   ControlWord;
    long   StatusWord;
    long   TagWord;
    long   ErrorOffset;
    long   ErrorSelector;
    long   DataOffset;
    long   DataSelector;
    BYTE    RegisterArea[SIZE_OF_80387_REGISTERS];
    long   Cr0NpxState;
} API_FLOATING_SAVE_AREA;

#define MAXIMUM_SUPPORTED_EXTENSION     512
typedef struct API_CONTEXT_i386 {

    //
    // The flags values within this flag control the contents of
    // a CONTEXT record.
    //
    // If the context record is used as an input parameter, then
    // for each portion of the context record controlled by a flag
    // whose value is set, it is assumed that that portion of the
    // context record contains valid context. If the context record
    // is being used to modify a threads context, then only that
    // portion of the threads context will be modified.
    //
    // If the context record is used as an IN OUT parameter to capture
    // the context of a thread, then only those portions of the thread's
    // context corresponding to set flags will be returned.
    //
    // The context record is never used as an OUT only parameter.
    //

    API_CONTEXT_FLAGS_i386 ContextFlags;

    //
    // This section is specified/returned if CONTEXT_DEBUG_REGISTERS is
    // set in ContextFlags.  Note that CONTEXT_DEBUG_REGISTERS is NOT
    // included in CONTEXT_FULL.
    //

    long   Dr0;
    long   Dr1;
    long   Dr2;
    long   Dr3;
    long   Dr6;
    long   Dr7;

    //
    // This section is specified/returned if the
    // ContextFlags word contians the flag CONTEXT_FLOATING_POINT.
    //

    API_FLOATING_SAVE_AREA FloatSave;

    //
    // This section is specified/returned if the
    // ContextFlags word contians the flag CONTEXT_SEGMENTS.
    //

    long   SegGs;
    long   SegFs;
    long   SegEs;
    long   SegDs;

    //
    // This section is specified/returned if the
    // ContextFlags word contians the flag CONTEXT_INTEGER.
    //

    long   Edi;
    long   Esi;
    long   Ebx;
    long   Edx;
    long   Ecx;
    long   Eax;

    //
    // This section is specified/returned if the
    // ContextFlags word contians the flag CONTEXT_CONTROL.
    //

    long   Ebp;
    long   Eip;
    long   SegCs;              // MUST BE SANITIZED
    long   EFlags;             // MUST BE SANITIZED
    long   Esp;
    long   SegSs;

    //
    // This section is specified/returned if the ContextFlags word
    // contains the flag CONTEXT_EXTENDED_REGISTERS.
    // The format and contexts are processor specific
    //

    BYTE    ExtendedRegisters[MAXIMUM_SUPPORTED_EXTENSION];

} API_CONTEXT;

#pragma pack(16)
typedef struct API_M128A {
    Int64 Low;
    Int64 High;
} API_M128A;
typedef struct API_XMM_SAVE_AREA32 {
    Integer     ControlWord;
    Integer     StatusWord;
    BYTE        TagWord;
    BYTE        Reserved1;
    Integer     ErrorOpcode;
    long        ErrorOffset;
    Integer     ErrorSelector;
    Integer     Reserved2;
    long        DataOffset;
    Integer     DataSelector;
    Integer     Reserved3;
    long        MxCsr;
    long        MxCsr_Mask;
    API_M128A   FloatRegisters[8];
    API_M128A   XmmRegisters[16];
    BYTE        Reserved4[96];
} API_XMM_SAVE_AREA32;

typedef struct API_CONTEXT_AMD64 {

    //
    // Register parameter home addresses.
    //
    // N.B. These fields are for convience - they could be used to extend the
    //      context record in the future.
    //

    Int64 P1Home;
    Int64 P2Home;
    Int64 P3Home;
    Int64 P4Home;
    Int64 P5Home;
    Int64 P6Home;

    //
    // Control flags.
    //

    API_CONTEXT_FLAGS_AMD64 ContextFlags;
    long MxCsr;

    //
    // Segment Registers and processor flags.
    //

    Integer   SegCs;
    Integer   SegDs;
    Integer   SegEs;
    Integer   SegFs;
    Integer   SegGs;
    Integer   SegSs;
    long      EFlags;

    //
    // Debug registers
    //

    Int64 Dr0;
    Int64 Dr1;
    Int64 Dr2;
    Int64 Dr3;
    Int64 Dr6;
    Int64 Dr7;

    //
    // Integer registers.
    //

    Int64 Rax;
    Int64 Rcx;
    Int64 Rdx;
    Int64 Rbx;
    Int64 Rsp;
    Int64 Rbp;
    Int64 Rsi;
    Int64 Rdi;
    Int64 R8;
    Int64 R9;
    Int64 R10;
    Int64 R11;
    Int64 R12;
    Int64 R13;
    Int64 R14;
    Int64 R15;

    //
    // Program counter.
    //

    Int64 Rip;

    //
    // Floating point state.
    //

    API_XMM_SAVE_AREA32 FltSave;

    //
    // Vector registers.
    //

    Int128 VectorRegister[26];
    Int64  VectorControl;

    //
    // Special debug control registers.
    //

    Int64 DebugControl;
    Int64 LastBranchToRip;
    Int64 LastBranchFromRip;
    Int64 LastExceptionToRip;
    Int64 LastExceptionFromRip;
} API_CONTEXT_AMD64;

typedef struct API_CONTEXT_IA64 {

    //
    // The flags values within this flag control the contents of
    // a CONTEXT record.
    //
    // If the context record is used as an input parameter, then
    // for each portion of the context record controlled by a flag
    // whose value is set, it is assumed that that portion of the
    // context record contains valid context. If the context record
    // is being used to modify a thread's context, then only that
    // portion of the threads context will be modified.
    //
    // If the context record is used as an IN OUT parameter to capture
    // the context of a thread, then only those portions of the thread's
    // context corresponding to set flags will be returned.
    //
    // The context record is never used as an OUT only parameter.
    //

    API_CONTEXT_FLAGS_IA64 ContextFlags;
    long Fill1[3];         // for alignment of following on 16-byte boundary

    //
    // This section is specified/returned if the ContextFlags word contains
    // the flag CONTEXT_DEBUG.
    //
    // N.B. CONTEXT_DEBUG is *not* part of CONTEXT_FULL.
    //

    UInt64 DbI0;
    UInt64 DbI1;
    UInt64 DbI2;
    UInt64 DbI3;
    UInt64 DbI4;
    UInt64 DbI5;
    UInt64 DbI6;
    UInt64 DbI7;

    UInt64 DbD0;
    UInt64 DbD1;
    UInt64 DbD2;
    UInt64 DbD3;
    UInt64 DbD4;
    UInt64 DbD5;
    UInt64 DbD6;
    UInt64 DbD7;

    //
    // This section is specified/returned if the ContextFlags word contains
    // the flag CONTEXT_LOWER_FLOATING_POINT.
    //

    Float128 FltS0;
    Float128 FltS1;
    Float128 FltS2;
    Float128 FltS3;
    Float128 FltT0;
    Float128 FltT1;
    Float128 FltT2;
    Float128 FltT3;
    Float128 FltT4;
    Float128 FltT5;
    Float128 FltT6;
    Float128 FltT7;
    Float128 FltT8;
    Float128 FltT9;

    //
    // This section is specified/returned if the ContextFlags word contains
    // the flag CONTEXT_HIGHER_FLOATING_POINT.
    //

    Float128 FltS4;
    Float128 FltS5;
    Float128 FltS6;
    Float128 FltS7;
    Float128 FltS8;
    Float128 FltS9;
    Float128 FltS10;
    Float128 FltS11;
    Float128 FltS12;
    Float128 FltS13;
    Float128 FltS14;
    Float128 FltS15;
    Float128 FltS16;
    Float128 FltS17;
    Float128 FltS18;
    Float128 FltS19;

    Float128 FltF32;
    Float128 FltF33;
    Float128 FltF34;
    Float128 FltF35;
    Float128 FltF36;
    Float128 FltF37;
    Float128 FltF38;
    Float128 FltF39;

    Float128 FltF40;
    Float128 FltF41;
    Float128 FltF42;
    Float128 FltF43;
    Float128 FltF44;
    Float128 FltF45;
    Float128 FltF46;
    Float128 FltF47;
    Float128 FltF48;
    Float128 FltF49;

    Float128 FltF50;
    Float128 FltF51;
    Float128 FltF52;
    Float128 FltF53;
    Float128 FltF54;
    Float128 FltF55;
    Float128 FltF56;
    Float128 FltF57;
    Float128 FltF58;
    Float128 FltF59;

    Float128 FltF60;
    Float128 FltF61;
    Float128 FltF62;
    Float128 FltF63;
    Float128 FltF64;
    Float128 FltF65;
    Float128 FltF66;
    Float128 FltF67;
    Float128 FltF68;
    Float128 FltF69;

    Float128 FltF70;
    Float128 FltF71;
    Float128 FltF72;
    Float128 FltF73;
    Float128 FltF74;
    Float128 FltF75;
    Float128 FltF76;
    Float128 FltF77;
    Float128 FltF78;
    Float128 FltF79;

    Float128 FltF80;
    Float128 FltF81;
    Float128 FltF82;
    Float128 FltF83;
    Float128 FltF84;
    Float128 FltF85;
    Float128 FltF86;
    Float128 FltF87;
    Float128 FltF88;
    Float128 FltF89;

    Float128 FltF90;
    Float128 FltF91;
    Float128 FltF92;
    Float128 FltF93;
    Float128 FltF94;
    Float128 FltF95;
    Float128 FltF96;
    Float128 FltF97;
    Float128 FltF98;
    Float128 FltF99;

    Float128 FltF100;
    Float128 FltF101;
    Float128 FltF102;
    Float128 FltF103;
    Float128 FltF104;
    Float128 FltF105;
    Float128 FltF106;
    Float128 FltF107;
    Float128 FltF108;
    Float128 FltF109;

    Float128 FltF110;
    Float128 FltF111;
    Float128 FltF112;
    Float128 FltF113;
    Float128 FltF114;
    Float128 FltF115;
    Float128 FltF116;
    Float128 FltF117;
    Float128 FltF118;
    Float128 FltF119;

    Float128 FltF120;
    Float128 FltF121;
    Float128 FltF122;
    Float128 FltF123;
    Float128 FltF124;
    Float128 FltF125;
    Float128 FltF126;
    Float128 FltF127;

    //
    // This section is specified/returned if the ContextFlags word contains
    // the flag CONTEXT_LOWER_FLOATING_POINT | CONTEXT_HIGHER_FLOATING_POINT | CONTEXT_CONTROL.
    //

    UInt64 StFPSR;       //  FP status

    //
    // This section is specified/returned if the ContextFlags word contains
    // the flag CONTEXT_INTEGER.
    //
    // N.B. The registers gp, sp, rp are part of the control context
    //

    UInt64 IntGp;        //  r1, volatile
    UInt64 IntT0;        //  r2-r3, volatile
    UInt64 IntT1;        //
    UInt64 IntS0;        //  r4-r7, preserved
    UInt64 IntS1;
    UInt64 IntS2;
    UInt64 IntS3;
    UInt64 IntV0;        //  r8, volatile
    UInt64 IntT2;        //  r9-r11, volatile
    UInt64 IntT3;
    UInt64 IntT4;
    UInt64 IntSp;        //  stack pointer (r12), special
    UInt64 IntTeb;       //  teb (r13), special
    UInt64 IntT5;        //  r14-r31, volatile
    UInt64 IntT6;
    UInt64 IntT7;
    UInt64 IntT8;
    UInt64 IntT9;
    UInt64 IntT10;
    UInt64 IntT11;
    UInt64 IntT12;
    UInt64 IntT13;
    UInt64 IntT14;
    UInt64 IntT15;
    UInt64 IntT16;
    UInt64 IntT17;
    UInt64 IntT18;
    UInt64 IntT19;
    UInt64 IntT20;
    UInt64 IntT21;
    UInt64 IntT22;

    UInt64 IntNats;      //  Nat bits for r1-r31
                            //  r1-r31 in bits 1 thru 31.
    UInt64 Preds;        //  predicates, preserved

    UInt64 BrRp;         //  return pointer, b0, preserved
    UInt64 BrS0;         //  b1-b5, preserved
    UInt64 BrS1;
    UInt64 BrS2;
    UInt64 BrS3;
    UInt64 BrS4;
    UInt64 BrT0;         //  b6-b7, volatile
    UInt64 BrT1;

    //
    // This section is specified/returned if the ContextFlags word contains
    // the flag CONTEXT_CONTROL.
    //

    // Other application registers
    UInt64 ApUNAT;       //  User Nat collection register, preserved
    UInt64 ApLC;         //  Loop counter register, preserved
    UInt64 ApEC;         //  Epilog counter register, preserved
    UInt64 ApCCV;        //  CMPXCHG value register, volatile
    UInt64 ApDCR;        //  Default control register (TBD)

    // Register stack info
    UInt64 RsPFS;        //  Previous function state, preserved
    UInt64 RsBSP;        //  Backing store pointer, preserved
    UInt64 RsBSPSTORE;
    UInt64 RsRSC;        //  RSE configuration, volatile
    UInt64 RsRNAT;       //  RSE Nat collection register, preserved

    // Trap Status Information
    UInt64 StIPSR;       //  Interruption Processor Status
    UInt64 StIIP;        //  Interruption IP
    UInt64 StIFS;        //  Interruption Function State

    // iA32 related control registers
    UInt64 StFCR;        //  copy of Ar21
    UInt64 Eflag;        //  Eflag copy of Ar24
    UInt64 SegCSD;       //  iA32 CSDescriptor (Ar25)
    UInt64 SegSSD;       //  iA32 SSDescriptor (Ar26)
    UInt64 Cflag;        //  Cr0+Cr4 copy of Ar27
    UInt64 StFSR;        //  x86 FP status (copy of AR28)
    UInt64 StFIR;        //  x86 FP status (copy of AR29)
    UInt64 StFDR;        //  x86 FP status (copy of AR30)

    UInt64 UNUSEDPACK;   //  added to pack StFDR to 16-bytes

} API_CONTEXT_IA64;

#pragma pack(4)

typedef struct API_CRITICAL_SECTION_DEBUG {
    Integer         Type;
    Integer         CreatorBackTraceIndex;
    LONG            CriticalSection;//struct API_CRITICAL_SECTION *CriticalSection;
    API_LIST_ENTRY  ProcessLocksList;
    LONG            EntryCount;
    LONG            ContentionCount;
    LONG            Spare[ 2 ];
} API_CRITICAL_SECTION_DEBUG;

typedef struct API_CRITICAL_SECTION {
    API_CRITICAL_SECTION_DEBUG DebugInfo;

    //
    //  The following three fields control entering and exiting the critical
    //  section for the resource
    //

    LONG LockCount;
    LONG RecursionCount;
    LONG OwningThread;        // from the thread's ClientId->UniqueThread
    LONG LockSemaphore;
    LONG SpinCount;
} API_CRITICAL_SECTION;

typedef struct API_TEB {
    long   pvExcept; // 00h Head of exception record list 'PEXCEPTION_REGISTRATION_RECORD
    long   pvStackUserTop;     // 04h Top of user stack     //SP
    long   pvStackUserBase;    // 08h Base of user stack    //BP

//union                       // 0Ch (NT/Win95 differences)
//{
    //struct  // Win95 fields
    //{
    //    WORD    pvTDB;         // 0Ch TDB
    //    WORD    pvThunkSS;     // 0Eh SS selector used for thunking to 16 bits
    //    long   unknown1;      // 10h
    //} WIN95;

    //struct  // WinNT fields
    //{
        long SubSystemTib;     // 0Ch
        long FiberData;        // 10h
    //} WINNT;
//} TIB_UNION1;

long   pvArbitrary;        // 14h Available for application use
long   pTIBSelf;      // 18h Linear address of TIB structure

//union                       // 1Ch (NT/Win95 differences)
//{
//    struct  // Win95 fields
//    {
//        WORD    TIBFlags;           // 1Ch
//        WORD    Win16MutexCount;    // 1Eh
//        DWORD   DebugContext;       // 20h
//        DWORD   pCurrentPriority;   // 24h
//        DWORD   pvQueue;            // 28h Message Queue selector
//    } WIN95;

//    struct  // WinNT fields
//    {
        long Unknown1;             // 1Ch
        long ProcessID;            // 20h
        long ThreadID;             // 24h
        long Unknown2;             // 28h
//    } WINNT;
//} TIB_UNION2;

    long  pvTLSArray;         // 2Ch Thread Local Storage array

/* union                       // 30h (NT/Win95 differences)
{
    struct  // Win95 fields
    {
        PVOID*  pProcess;     // 30h Pointer to owning process database
    } WIN95;
} TIB_UNION3; */
    
} API_TEB;


#pragma pack()
[
    dllname("Kernel32.dll"),
    helpstring("Access to threading API functions within the Kernel32.dll system file.")
]
module KernelThreading {
[entry("GetCurrentThreadId"), usesgetlasterror]
    long API_GetCurrentThreadId();
[entry("GetCurrentThread"), usesgetlasterror]
    long API_GetCurrentThread();
/* [entry("GetThreadId"), usesgetlasterror]
    long API_GetThreadId([in] long Thread); */
//========================================
[entry("CreateThread"), usesgetlasterror]
    long API_CreateThread([in] Any lpThreadAttributes, [in] long dwStackSize, [in] long lpStartAddress, [in] Any lpParameter, [in] long dwCreationFlags, [out] long* lpThreadId);

[entry("OpenThread"), usesgetlasterror]
    long API_OpenThread([in] long dwDesiredAccess, [in] long bInheritHandle, [in] long dwThreadId);
//========================================
[entry("ExitThread"), usesgetlasterror]
    long API_ExitThread([in] long dwExitCode);
[entry("GetExitCodeThread"), usesgetlasterror]
    long API_GetExitCodeThread([in] long hThread, [out] long* lpExitCode);
[entry("TerminateThread"), usesgetlasterror]
    long API_TerminateThread([in] long hThread, [in] long dwExitCode);
//========================================
[entry("SuspendThread"), usesgetlasterror]
    long API_SuspendThread([in] long hThread);
[entry("ResumeThread"), usesgetlasterror]
    long API_ResumeThread([in] long hThread);
//========================================
[entry("Thread32First"), usesgetlasterror]
    long API_Thread32First([in] long hSnapshot, [out] API_THREADENTRY32* lpte);
[entry("Thread32Next"), usesgetlasterror]
    long API_Thread32Next([in] long hSnapshot, [out] API_THREADENTRY32* lpte);
//========================================
[entry("GetOwnerProcess"), usesgetlasterror]
    long API_GetOwnerProcess();
[entry("SwitchToThread"), usesgetlasterror]
    long API_SwitchToThread();
//========================================
[entry("SetThreadIdealProcessor"), usesgetlasterror]
    long API_SetThreadIdealProcessor([in] long hThread, [in] long dwIdealProcessor);
[entry("GetThreadContext"), usesgetlasterror]
    long API_GetThreadContext([in] long hThread, [out] API_CONTEXT* lpContext);
[entry("GetThreadDesktop"), usesgetlasterror]
    long API_GetThreadDesktop([in] long dwThread);
[entry("GetThreadLocale"), usesgetlasterror]
    long API_GetThreadLocale();
[entry("GetThreadPriority"), usesgetlasterror]
    long API_GetThreadPriority([in] long hThread);
[entry("GetThreadSelectorEntry"), usesgetlasterror]
    long API_GetThreadSelectorEntry([in] long hThread, [in] long dwSelector, [out] API_LDT_ENTRY* lpSelectorEntry);
[entry("GetThreadTimes"), usesgetlasterror]
    long API_GetThreadTimes([in] long hThread, [out] API_FILETIME* lpCreationTime, [out] API_FILETIME* lpExitTime, [out] API_FILETIME* lpKernelTime, [out] API_FILETIME* lpUserTime);
[entry("SetThreadAffinityMask"), usesgetlasterror]
    long API_SetThreadAffinityMask([in] long hThread, [in] long dwThreadAffinityMask);
[entry("SetThreadDesktop"), usesgetlasterror]
    long API_SetThreadDesktop([in] long hDesktop);
[entry("SetThreadContext"), usesgetlasterror]
    long API_SetThreadContext([in] long hThread, [out] API_CONTEXT* lpContext);
[entry("SetThreadLocale"), usesgetlasterror]
    long API_SetThreadLocale([in] long Locale);
[entry("SetThreadPriority"), usesgetlasterror]
    long API_SetThreadPriority([in] long hThread, [in] long nPriority);
[entry("SetThreadToken"), usesgetlasterror]
    long API_SetThreadToken([in] long hThread, [in] long Token);
//[entry("GetWindowThreadProcessId"), usesgetlasterror]
//    long API_GetWindowThreadProcessId([in] long hWnd, [out] long* lpdwProcessId);
//========================================
/*
[entry("GetThreadInformation"), usesgetlasterror] //Win 8 Only
    long API_GetThreadInformation([in] long hThread, [in] THREAD_INFORMATION_CLASS ThreadInformationClass, [out] Any ThreadInformation, [in] long ThreadInformationSize);
[entry("GetCurrentThreadStackLimits"), usesgetlasterror] //Win 8 Only
    void API_GetCurrentThreadStackLimits([out] long* LowLimit, [out] long* HighLimit);
*/
//========================================
[entry("Sleep"), usesgetlasterror]
    long API_Sleep([in] long dwMilliseconds);
//========================================
[entry("DisableThreadLibraryCalls"), usesgetlasterror]
    long API_DisableThreadLibraryCalls([in] long hLibModule);
};

#endif //__KERNEL_THREADING_H__