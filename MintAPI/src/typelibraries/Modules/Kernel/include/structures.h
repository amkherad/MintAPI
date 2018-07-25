#ifndef __STRUCTURES_H__
#define __STRUCTURES_H__

// forces the structures to be 4byte aligned
#pragma pack(4)


typedef struct API_RECT
{
    long    Left;
    long    Top;
    long    Right;
    long    Bottom;
} API_RECT;


typedef struct API_IDLDESC {
    long        dwReserved;
    Integer     wIDLFlags;
} API_IDLDESC;

typedef struct API_TYPEDESC {
    long        lpValue;
    Integer     VT;
} API_TYPEDESC;

typedef struct API_EXCEPINFO {
    Integer     wCode;
    Integer     wReserved;
    BSTR        bstrSource;
    BSTR        bstrDescription;
    BSTR        bstrHelpFile;
    Long        dwHelpContext;
    Long        pvReserved;
    Long        pfnDeferredFillIn;
    Long        scode;
} API_EXCEPINFO;

typedef struct API_DISPPARAMS
{
    VARIANT     rgvarg;
    long        rgdispidNamedArgs;
    long        cArgs;
    long        cNamedArgs;
} API_DISPPARAMS;

typedef struct API_PARAMDESCEX {
    long            cBytes;
    VARIANT         varDefaultValue;
} API_PARAMDESCEX;
    
typedef struct API_PARAMDESC {
    API_PARAMDESCEX pParamDescex;
    Integer         wParamFlags;
} API_PARAMDESC;

typedef struct API_ELEMDESC {
    API_TYPEDESC    tDesc;/* the type of the element */
    API_IDLDESC     IdlDesc;        /* info for remoting the element */
    API_PARAMDESC   ParamDesc;    /* info about the parameter */
} API_ELEMDESC;

typedef struct API_FUNCDESC {
    long            MemID;
    long            lPrgSCode;
    API_ELEMDESC    lPrgElemDescParam;
    API_FuncKind    funcKind;
    API_InvokeKind  InvKind;
    API_CallConv    CallConv;
    Integer         cParams;
    Integer         cParamsOpt;
    Integer         oVft;
    Integer         cScodes;
    API_ELEMDESC    ElemdescFunc;
    Integer         wFuncFlags;
} API_FUNCDESC;

typedef struct API_VARDESC {
    long            MemID;
    long            lpStrPtr_Schema;
    long oInst;
    VARIANT         lpVarValue;
    API_ELEMDESC    ElemDescVar;
    Integer         wVarFlags;
    API_VARKIND     VarKind;
} API_VARDESC;

typedef struct API_CUSTDATAITEM {
    API_StdGuid     Guid;
    VARIANT         varValue;
} API_CUSTDATAITEM;
typedef struct API_CUSTDATA {
    long            cCustData;
    API_CUSTDATAITEM prgCustData;
} API_CUSTDATA;

typedef struct API_TLIBATTR {
    API_StdGuid     Guid;
    long            LCID;
    API_SysKind     SysKind;
    Integer         wMajorVerNum;
    Integer         wMinorVerNum;
    Integer         wLibFlags;
} API_TLIBATTR;

typedef struct API_BINDPTR {
    API_FUNCDESC    lpFuncDesc;
    API_VARDESC     lpVarDesc;
    long            lpITypeCompPtr;
} API_BINDPTR;

typedef struct API_TYPEATTR {
    API_StdGuid Guid;
    long        LCID;
    long        dwReserved;
    long        MemIDConstructor;
    long        MemIDDestructor;
    long        lpStrPtr_Schema;
    long        cbSizeInstance;
    API_TypeKind Typekind;
    long        cFuncs;
    long        cVars;
    long        cImplTypes;
    long        cbSizeVft;
    long        cbAlignment;
    long        wTypeFlags;
    long        wMajorVerNum;
    long        wMinorVerNum;
    API_TYPEDESC tDescAlias;
    API_IDLDESC IdlDescType;
} API_TYPEATTR;


typedef struct API_FILE_SEGMENT_ELEMENT {
    Int64       Buffer;
    UInt64      Alignment;
} API_FILE_SEGMENT_ELEMENT;

typedef struct API_OSVERSIONINFOA {
    long        dwOSVersionInfoSize;
    long        dwMajorVersion;
    long        dwMinorVersion;
    long        dwBuildNumber;
    long        dwPlatformId;
    byte        szCSDVersion[128];
} API_OSVERSIONINFOA;

typedef struct API_OSVERSIONINFOEXA {
    long	    dwOSVersionInfoSize;
    long	    dwMajorVersion;
    long	    dwMinorVersion;
    long	    dwBuildNumber;
    long	    dwPlatformId;
    byte	    szCSDVersion[128];
    short	    wServicePackMajor;
    short	    wServicePackMinor;
    short	    wSuiteMask;
    byte	    wProductType;
    byte	    wReserved;
} API_OSVERSIONINFOEXA;

typedef struct API_OSVERSIONINFOW {
    long        dwOSVersionInfoSize;
    long        dwMajorVersion;
    long        dwMinorVersion;
    long        dwBuildNumber;
    long        dwPlatformId;
    Char        szCSDVersion[128];
} API_OSVERSIONINFOW;

typedef struct API_OSVERSIONINFOEXW {
    long	    dwOSVersionInfoSize;
    long	    dwMajorVersion;
    long	    dwMinorVersion;
    long	    dwBuildNumber;
    long	    dwPlatformId;
    Char	    szCSDVersion[128];
    short	    wServicePackMajor;
    short	    wServicePackMinor;
    short	    wSuiteMask;
    byte	    wProductType;
    byte	    wReserved;
} API_OSVERSIONINFOEXW;

typedef struct API_LIST_ENTRY {
   LONG   Flink;
   LONG   Blink;
} API_LIST_ENTRY;

#endif //__STRUCTURES_H__