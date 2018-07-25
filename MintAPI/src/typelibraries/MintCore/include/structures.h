#ifndef __STRUCTURES_H__
#define __STRUCTURES_H__

// forces the structures to be 4byte aligned
#pragma pack(4)


typedef struct API_SECURITY_ATTRIBUTES {
    Long        nLength;
    Long        lpSecurityDescriptor;
    Long        bInheritHandle;
} API_SECURITY_ATTRIBUTES;

typedef struct API_OVERLAPPED {
    long        Internal;
    long        InternalHigh;
    long        Offset;
    long        OffsetHigh;
    long        hEvent;
} API_OVERLAPPED;

typedef struct API_LIST_ENTRY {
   LONG   Flink;
   LONG   Blink;
} API_LIST_ENTRY;

typedef [uuid(00000000-7720-0045-7FFF-7ACDC6661234)]
struct API_StdGuid {
    long	    Data1;
    short	    Data2;
    short	    Data3;
    byte	    Data4[8];
} API_StdGuid;
typedef struct API_STATSTG {
    long        pwcsName;
    long        stgType;
    Int64       cbSize;
    Int64       mTime;
    Int64       cTime;
    Int64       aTime;
    long        grfMode;
    long        grfLocksSupported;
    API_StdGuid ClsID;
    long        grfStateBits;
    long        Reserved0;
} API_STATSTG;

#endif //__STRUCTURES_H__