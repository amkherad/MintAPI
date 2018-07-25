
typedef [public] short              WORD;
typedef [public] long               DWORD;
//typedef [public] long long          QWORD;
typedef [public] unsigned char      UCHAR;

typedef [public] unsigned short     UWORD;
typedef [public] unsigned long      UDWORD;
//typedef [public] unsigned long long UQWORD;

typedef [public] BSTR               MLPSTR;
typedef [public] BSTR               HLSTR;
typedef [public] BSTR               MLPWSTR;

typedef [public] long               MHANDLE;
typedef [public] long               MColor;
typedef [public] long               MERROR_ID;
typedef [public] long               MFILE_HANDLE;
typedef [public] char               MCHAR;
typedef [public] MCHAR*             MBSTR;
typedef [public] unsigned char      MUCHAR;
typedef [public] wchar_t            MWCHAR;
typedef [public] MWCHAR*            MWBSTR;
typedef [public] void               MLPVOID;
typedef [public] long               MLCID;
typedef [public] long               MDISPID;
typedef [public] long               MHRESULT;
typedef [public] void*              MAny;
typedef [public] int                MBOOL;

#pragma pack(4)

typedef struct Int32 {
    long Value;
} Int32;
typedef struct UInt32 {
    long Value;
} UInt32;

typedef struct Int64 {
    long LowerPart;
    long HigherPart;
} Int64;
typedef struct UInt64 {
    long LowerPart;
    long HigherPart;
} UInt64;

typedef struct Int128 {
    Int64 LowerPart;
    Int64 HigherPart;
} Int128;
typedef struct UInt128 {
    UInt64 LowerPart;
    UInt64 HigherPart;
} UInt128;

typedef struct Int256 {
    Int128 LowerPart;
    Int128 HigherPart;
} Int256;
typedef struct UInt256 {
    UInt128 LowerPart;
    UInt128 HigherPart;
} UInt256;

typedef struct Float64 {
    long LowerPart;
    long HigherPart;
} Float64;
typedef struct Float128 {
    Float64 LowerPart;
    Float64 HigherPart;
} Float128;
typedef struct Float256 {
    Float128 LowerPart;
    Float128 HigherPart;
} Float256;



typedef [public] Int32   MInt32;
typedef [public] Int64   MInt64;
typedef [public] Int128  MInt128;