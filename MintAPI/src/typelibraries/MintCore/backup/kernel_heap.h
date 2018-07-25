#ifndef __KERNEL_HEAP_H__
#define __KERNEL_HEAP_H__

#pragma pack(4)
typedef struct API_HEAPENTRY32 {
    Integer    dwSize;
    long       hHandle;
    long       dwAddress;
    Integer    dwBlockSize;
    long       dwFlags;
    long       dwLockCount;
    long       dwResvd;
    long       th32ProcessID;
    long       th32HeapID;
} API_HEAPENTRY32;

typedef struct API_PROCESS_HEAP_ENTRY {
    long       lpData;
    long       cbData;
    BYTE       cbOverhead;
    BYTE       iRegionIndex;
    Integer    wFlags;
} API_PROCESS_HEAP_ENTRY;
typedef struct API_PROCESS_HEAP_ENTRY_UNION1 {
    long       hMem;
    long       dwReserved[3];
} API_PROCESS_HEAP_ENTRY_UNION1;
typedef struct API_PROCESS_HEAP_ENTRY_UNION2 {
    long       dwCommittedSize;
    long       dwUnCommittedSize;
    long       lpFirstBlock;
    long       lpLastBlock;
} API_PROCESS_HEAP_ENTRY_UNION2;

typedef struct API_HEAPLIST32 {
    Integer    dwSize;
    long       th32ProcessID;   // owning process
    long       th32HeapID;      // heap (in owning process's context!)
    long       dwFlags;
} API_HEAPLIST32;

#pragma pack()
[
    dllname("Kernel32.dll"),
    helpstring("Access to threading API functions within the Kernel32.dll system file.")
]
module KernelHeap {
[entry("GetProcessHeap"), usesgetlasterror]
    long API_GetProcessHeap();
//========================================
[entry("Heap32First"), usesgetlasterror]
    long API_Heap32First([out] API_HEAPENTRY32* lphe, [in] long th32ProcessID, [out] long* th32HeapID);
[entry("Heap32Next"), usesgetlasterror]
    long API_Heap32Next([out] API_HEAPENTRY32* lphe);
/* [entry("Heap32Next"), usesgetlasterror]
    long API_Heap32Next([in] long hHeap, [out] API_HEAPENTRY32* lphe); */
//========================================
/* [entry("Heap32ListFirst"), usesgetlasterror]
    long API_Heap32ListFirst([in] long hSnapshot, [out] API_HEAPENTRY32* lphe);
[entry("Heap32ListNext"), usesgetlasterror]
    long API_Heap32ListNext([in] long hSnapshot, [out] API_HEAPENTRY32* lphe); */
[entry("Heap32ListNext"), usesgetlasterror]
    long API_Heap32ListNext([in] long hSnapshot, [out] API_HEAPLIST32* lphl);
[entry("Heap32ListFirst"), usesgetlasterror]
    long API_Heap32ListFirst([in] long hSnapshot, [out] API_HEAPLIST32* lphl);
//========================================
[entry("HeapAlloc"), usesgetlasterror]
    long API_HeapAlloc([in] long hHeap, [in] long dwFlags, [in] long dwBytes);
[entry("HeapCompact"), usesgetlasterror]
    long API_HeapCompact([in] long hHeap, [in] long dwFlags);
[entry("HeapCreate"), usesgetlasterror]
    long API_HeapCreate([in] long flOptions, [in] long dwInitialSize, [in] long dwMaximumSize);
[entry("HeapDestroy"), usesgetlasterror]
    long API_HeapDestroy([in] long hHeap);
[entry("HeapFree"), usesgetlasterror]
    long API_HeapFree([in] long hHeap, [in] long dwFlags, [in] long lpMem);
[entry("HeapLock"), usesgetlasterror]
    long API_HeapLock([in] long hHeap);
[entry("HeapReAlloc"), usesgetlasterror]
    long API_HeapReAlloc([in] long hHeap, [in] long dwFlags, [in] long lpMem, [in] long dwBytes);
[entry("HeapSize"), usesgetlasterror]
    long API_HeapSize([in] long hHeap, [in] long dwFlags, [in] long lpMem);
[entry("HeapUnlock"), usesgetlasterror]
    long API_HeapUnlock([in] long hHeap);
[entry("HeapValidate"), usesgetlasterror]
    long API_HeapValidate([in] long hHeap, [in] long dwFlags, [in] long lpMem);
[entry("HeapWalk"), usesgetlasterror]
    long API_HeapWalk([in] long hHeap, [out] API_PROCESS_HEAP_ENTRY* lpEntry);
};

#endif //__KERNEL_HEAP_H__