#ifndef __OLE32_H__
#define __OLE32_H__


// interface marshaling definitions
#define MARSHALINTERFACE_MIN 500 // minimum number of bytes for interface marshl

typedef enum API_REGCLS
{
    REGCLS_SINGLEUSE = 0,       // class object only generates one instance
    REGCLS_MULTIPLEUSE = 1,     // same class object genereates multiple inst.
                                // and local automatically goes into inproc tbl.
    REGCLS_MULTI_SEPARATE = 2,  // multiple use, but separate control over each
                                // context.
    REGCLS_SUSPENDED      = 4,  // register is as suspended, will be activated
                                // when app calls CoResumeClassObjects
    REGCLS_SURROGATE      = 8   // must be used when a surrogate process
                                // is registering a class object that will be
                                // loaded in the surrogate
} API_REGCLS;

typedef enum API_CLSCTX {
    CLSCTX_INPROC_SERVER	= 0x1,
	CLSCTX_INPROC_HANDLER	= 0x2,
	CLSCTX_LOCAL_SERVER	= 0x4,
	CLSCTX_INPROC_SERVER16	= 0x8,
	CLSCTX_REMOTE_SERVER	= 0x10,
	CLSCTX_INPROC_HANDLER16	= 0x20,
	CLSCTX_RESERVED1	= 0x40,
	CLSCTX_RESERVED2	= 0x80,
	CLSCTX_RESERVED3	= 0x100,
	CLSCTX_RESERVED4	= 0x200,
	CLSCTX_NO_CODE_DOWNLOAD	= 0x400,
	CLSCTX_RESERVED5	= 0x800,
	CLSCTX_NO_CUSTOM_MARSHAL	= 0x1000,
	CLSCTX_ENABLE_CODE_DOWNLOAD	= 0x2000,
	CLSCTX_NO_FAILURE_LOG	= 0x4000,
	CLSCTX_DISABLE_AAA	= 0x8000,
	CLSCTX_ENABLE_AAA	= 0x10000,
	CLSCTX_FROM_DEFAULT_CONTEXT	= 0x20000,
	CLSCTX_ACTIVATE_32_BIT_SERVER	= 0x40000,
	CLSCTX_ACTIVATE_64_BIT_SERVER	= 0x80000,
    
    CLSCTX_ALL             = 0x7,
    CLSCTX_SERVER          = 0x5
} API_CLSCTX;

// COM initialization flags; passed to CoInitialize.
typedef enum API_COINIT {
  COINIT_APARTMENTTHREADED  = 0x2,      // Apartment model
  // These constants are only valid on Windows NT 4.0
  COINIT_MULTITHREADED      = 0x0,      // OLE calls objects on any thread.
  COINIT_DISABLE_OLE1DDE    = 0x4,      // Don't use DDE for Ole1 support.
  COINIT_SPEED_OVER_MEMORY  = 0x8,      // Trade memory for speed.
} API_COINIT;

#pragma pack(4)

typedef struct API_COAUTHIDENTITY {
    Integer   *User;
    Long      UserLength;
    Integer   *Domain;
    Long      DomainLength;
    Integer   *Password;
    Long      PasswordLength;
    Long      Flags;
} API_COAUTHIDENTITY;

typedef struct API_COAUTHINFO {
    DWORD          dwAuthnSvc;
    DWORD          dwAuthzSvc;
    String         pwszServerPrincName;
    DWORD          dwAuthnLevel;
    DWORD          dwImpersonationLevel;
    API_COAUTHIDENTITY *pAuthIdentityData;
    DWORD          dwCapabilities;
} API_COAUTHINFO;

typedef struct API_COSERVERINFO {
    DWORD      dwReserved1;
    String     pwszName;
    API_COAUTHINFO *pAuthInfo;
    DWORD      dwReserved2;
} API_COSERVERINFO;

typedef struct API_MULTI_QI {
  API_StdGuid   *pIID;
  IUnknown      *pItf;
  HRESULT       HR;
} API_MULTI_QI;

#pragma pack()
[
    dllname("ole32.dll"),
    helpstring("Access to API functions within the Ole32.dll system file.")
]
module Ole32 {
[entry("CoTaskMemAlloc")]
    long	API_CoTaskMemAlloc([in] long cb);
[entry("CoTaskMemFree")]
    void	API_CoTaskMemFree([in] long pv);
[entry("CoGetMalloc")]
    HRESULT API_CoGetMalloc([in] long dwMemContext, [out, retval] IMalloc** ppMalloc);
//========================================
[entry("OleInitialize")]
    long	API_OleInitialize([in] Any pvReserved);
//========================================
[entry("GetErrorInfo")]
    long	API_GetErrorInfo([in] long dwReserved, [in] Any ppErrInfo);
//========================================
[entry("GetClassFile")]
    long	API_GetClassFile([in] long szFilename, [out] long* pClsID);
//========================================
[entry("CoCreateGuid")]
    HRESULT API_CoCreateGuid([out, retval] API_StdGuid * RetVal);
[entry("IIDFromString")]
    HRESULT API_GUIDFromString([in] LPWSTR lpsz, [out, retval] API_StdGuid * RetVal);
[entry("IsEqualGUID")]
    long    API_IsEqualGUID([in] void * rguid1, [in] void * rguid2);
[entry("StringFromGUID2")]
    long    API_StringFromGUID2([in] API_StdGuid * rguid, [in] LPWSTR lpsz, [in] long cbMax);
[entry("ProgIDFromCLSID")]
    long    API_ProgIDFromCLSID([in] API_StdGuid * CLSID, [in] long * lplpszProgID);
[entry("CLSIDFromProgID")]
    HRESULT API_CLSIDFromProgID([in] LPWSTR ProgID, [out, retval] API_StdGuid * RetVal);
[entry("CLSIDFromString")]
    HRESULT API_CLSIDFromString([in] LPWSTR lpSz, [out, retval] API_StdGuid * RetVal);
//========================================
[entry("CoGetCurrentProcess")]
    long	API_CoGetCurrentProcess();
//========================================
[entry("CoLoadLibrary")]
    long	API_CoLoadLibrary([in] String lpszLibName, [in] long bAutoFree);
[entry("CoFreeLibrary")]
    void	API_CoFreeLibrary([in] long hInst);
[entry("CoFreeAllLibraries")]
    void	API_CoFreeAllLibraries();
[entry("CoFreeUnusedLibraries")]
    void	API_CoFreeUnusedLibraries();
//========================================
[entry("CreateStreamOnHGlobal")]
    long	API_CreateStreamOnHGlobal([in] long hGlobal, [in] Boolean fDeleteOnRelease, [in, out] IStream ** Stream);
//========================================
[entry("CoCreateInstance")]
    void	API_CoCreateInstance([in] Any rClsid, [in] IUnknown* pUnkOuter, [in] long dwClsContext, [in] Any riid, [out] Any* ppv);
[entry("CoCreateInstanceEx")]
    void	API_CoCreateInstanceEx([in] Any rClsid, [in] IUnknown* pUnkOuter, [in] long dwClsCtx, [in] API_COSERVERINFO pServerInfo, [in] long dwCount, [out] API_MULTI_QI* pResults);
/*
Declare Function CoAddRefServerProcess lib "ole32" () As Long
Declare Function CoBuildVersion lib "ole32" () As Long
Declare Function CoDosDateTimeToFileTime lib "ole32" (ByVal nDosDate As Integer, ByVal nDosTime As Integer, ByRef lpFileTime As FILETIME) As Long
Declare Function CoFileTimeToDosDateTime lib "ole32" (ByRef lpFileTime As FILETIME, ByRef lpDosDate As Integer, ByRef lpDosTime As Integer) As Long

Declare Function CoIsHandlerConnected lib "ole32" (ByVal pUnk As Long) As Long
Declare Function CoIsOle1Class lib "ole32" (ByVal rclsid As Long) As Long

Declare Sub CoAllowSetForegroundWindow lib "OLE32" (ByVal pUnk As Long, lpvReserved As Any)
Declare Sub CoCancelCall lib "ole32" (ByVal dwThreadId As Long, ByVal ulTimeout As Long)
Declare Sub CoCopyProxy lib "ole32" (ByVal pProxy As Long, ByVal ppCopy As Long)
Declare Sub CoCreateFreeThreadedMarshaler lib "ole32" (ByVal punkOuter As Long, ByVal ppunkMarshal As Long)
Declare Sub CoCreateGuid lib "ole32" (ByRef pguid As GUID)
Declare Sub CoDisableCallCancellation lib "ole32" (pReserved As Any)
Declare Sub CoDisconnectObject lib "ole32" (ByVal pUnk As Long, ByVal dwReserved As Long)
Declare Sub CoEnableCallCancellation lib "ole32" (pReserved As Any)
Declare Sub CoFileTimeNow lib "ole32" (ByRef lpFileTime As FILETIME)

Declare Sub CoGetCallContext lib "ole32" (ByVal riid As Long, ppInterface As Any)
Declare Sub CoGetCancelObject lib "ole32" (ByVal dwThreadId As Long, ByVal iid As Long, ppUnk As Any)
Declare Sub CoGetClassObject lib "ole32" (ByVal rclsid As Long, ByVal dwClsContext As Long, pvReserved As Any, ByVal riid As Long, ppv As Any)
Declare Sub CoGetClassObjectFromURL lib "URLMON" (ByVal rCLASSID As Long, ByVal szCODE As String, ByVal dwFileVersionMS As Long, ByVal dwFileVersionLS As Long, ByVal szType As String, ByVal pBindCtx As Long, ByVal dwClsContext As Long, pvReserved As Any, ByVal riid As Long, ppv As Any)
Declare Sub CoGetClassVersion lib "ole32" (ByRef pClassSpec As uCLSSPEC, ByRef pdwVersionMS As Long, ByRef pdwVersionLS As Long)
Declare Sub CoGetInstanceFromFile lib "ole32" (ByRef pServerInfo As COSERVERINFO, ByRef pClsid As Long, ByVal punkOuter As Long, ByVal dwClsCtx As Long, ByVal grfMode As Long, ByRef pwszName As Byte, ByVal dwCount As Long, ByRef pResults As MULTI_QI)
Declare Sub CoGetInstanceFromIStorage lib "ole32" (ByRef pServerInfo As COSERVERINFO, ByRef pClsid As Long, ByVal punkOuter As Long, ByVal dwClsCtx As Long, ByVal pstg As Long, ByVal dwCount As Long, ByRef pResults As MULTI_QI)
Declare Sub CoGetInterfaceAndReleaseStream lib "ole32" (ByRef pStm As Long, ByVal iid As Long, ppv As Any)
Declare Sub CoGetMalloc lib "ole32" (ByVal dwMemContext As Long, ByVal ppMalloc As Long)
Declare Sub CoGetMarshalSizeMax lib "ole32" (ByRef pulSize As Long, ByVal riid As Long, ByVal pUnk As Long, ByVal dwDestContext As Long, pvDestContext As Any, ByVal mshlflags As Long)
Declare Sub CoGetObject lib "ole32" (ByVal pszName As String, ByRef pBindOptions As BIND_OPTS, ByVal riid As Long, ppv As Any)
Declare Sub CoGetObjectContext lib "ole32" (ByVal riid As Long, ppv As Any)
Declare Sub CoGetPSClsid lib "ole32" (ByVal riid As Long, ByRef pClsid As Long)
Declare Sub CoGetStandardMarshal lib "ole32" (ByVal riid As Long, ByVal pUnk As Long, ByVal dwDestContext As Long, pvDestContext As Any, ByVal mshlflags As Long, ByVal ppMarshal As Long)
Declare Sub CoGetStdMarshalEx lib "ole32" (ByVal pUnkOuter As Long, ByVal smexflags As Long, ByVal ppUnkInner As Long)
Declare Sub CoGetTreatAsClass lib "ole32" (ByVal clsidOld As Long, ByVal pClsidNew As Long)
Declare Sub CoImpersonateClient lib "ole32" ()
Declare Sub CoInitialize lib "ole32" (pvReserved As Any)
Declare Sub CoInitializeEx lib "ole32" (pvReserved As Any, ByVal dwCoInit As Long)
Declare Sub CoInitializeSecurity lib "ole32" (ByRef pSecDesc As SECURITY_DESCRIPTOR, ByVal cAuthSvc As Long, ByRef asAuthSvc As SOLE_AUTHENTICATION_SERVICE, pReserved1 As Any, ByVal dwAuthnLevel As Long, ByVal dwImpLevel As Long, pAuthList As Any, ByVal dwCapabilities As Long, pReserved3 As Any)
Declare Sub CoInstall lib "ole32" (ByVal pbc As Long, ByVal dwFlags As Long, ByRef pClassSpec As uCLSSPEC, ByRef pQuery As QUERYCONTEXT, ByVal pszCodeBase As String)
Declare Sub CoInternetCombineUrl lib "URLMON" (ByVal pwzBaseUrl As String, ByVal pwzRelativeUrl As String, ByVal dwCombineFlags As Long, ByVal pszResult As String, ByVal cchResult As Long, ByRef pcchResult As Long, ByVal dwReserved As Long)
Declare Sub CoInternetCompareUrl lib "URLMON" (ByVal pwzUrl1 As String, ByVal pwzUrl2 As String, ByVal dwFlags As Long)
Declare Sub CoInternetCreateSecurityManager lib "URLMON" (ByRef pSP As IServiceProvider, ByRef ppSM As IInternetSecurityManager, ByVal dwReserved As Long)
Declare Sub CoInternetCreateZoneManager lib "URLMON" (ByRef pSP As IServiceProvider, ByRef ppZM As IInternetZoneManager, ByVal dwReserved As Long)
Declare Sub CoInternetGetProtocolFlags lib "URLMON" (ByVal pwzUrl As String, ByRef pdwFlags As Long, ByVal dwReserved As Long)
Declare Sub CoInternetGetSecurityUrl lib "URLMON" (ByVal pwzUrl As String, ByVal ppwzSecUrl As String, ByRef psuAction As psuAction, ByVal dwReserved As Long)
Declare Sub CoInternetGetSession lib "URLMON" (ByVal dwSessionMode As Long, ByRef ppIInternetSession As IInternetSession, ByVal dwReserved As Long)
Declare Sub CoInternetParseUrl lib "URLMON" (ByVal pwzUrl As String, ByRef ParseAction As ParseAction, ByVal dwFlags As Long, ByVal pszResult As String, ByVal cchResult As Long, ByRef pcchResult As Long, ByVal dwReserved As Long)
Declare Sub CoInternetQueryInfo lib "URLMON" (ByVal pwzUrl As String, ByRef QueryOptions As QUERYOPTION, ByVal dwQueryFlags As Long, pvBuffer As Any, ByVal cbBuffer As Long, ByRef pcbBuffer As Long, ByVal dwReserved As Long)
Declare Sub CoLockObjectExternal lib "ole32" (ByVal pUnk As Long, ByVal fLock As Long, ByVal fLastUnlockReleases As Long)
Declare Sub ColorRGBToHLS lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
Declare Sub CoMarshalHresult lib "ole32" (ByRef pstm As Long, ByVal hresult As Long)
Declare Sub CoMarshalInterface lib "ole32" (ByRef pStm As Long, ByVal riid As Long, ByVal pUnk As Long, ByVal dwDestContext As Long, pvDestContext As Any, ByVal mshlflags As Long)
Declare Sub CoMarshalInterThreadInterfaceInStream lib "ole32" (ByVal riid As Long, ByVal pUnk As Long, ByRef ppStm As Long)
Declare Sub CommitUrlCacheEntry lib "wininet" Alias "CommitUrlCacheEntryA" (ByVal lpszUrlName As String, ByVal lpszLocalFileName As String, ByVal ExpireTime As Struct_MembersOf_FILETIME, ByVal LastModifiedTime As Struct_MembersOf_FILETIME, ByVal CacheEntryType As Long, ByVal lpHeaderInfo As String, ByVal dwHeaderSize As Long, ByVal lpszFileExtension As String, ByVal lpszOriginalUrl As String)
Declare Sub CompleteAuthToken lib "digest" (ByRef phContext As Long, ByRef pToken As PSecBufferDesc)
Declare Sub CopyBindInfo lib "URLMON" (ByRef pcbiSrc As longx, ByRef pbiDest As BINDINFO)
Declare Sub CopyMemory lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Sub CopyStgMedium lib "URLMON" (ByRef pcstgmedSrc As STGMEDIUM, ByRef pstgmedDest As STGMEDIUM)
Declare Sub CoQueryAuthenticationServices lib "ole32" (ByRef pcAuthSvc As Long, ByRef asAuthSvc As SOLE_AUTHENTICATION_SERVICE)
Declare Sub CoQueryClientBlanket lib "ole32" (ByRef pAuthnSvc As Long, ByRef pAuthzSvc As Long, ByRef pServerPrincName As Byte, ByRef pAuthnLevel As Long, ByRef pImpLevel As Long, ByRef pPrivs As Long, ByRef pCapabilities As Long)
Declare Sub CoQueryProxyBlanket lib "ole32" (ByVal pProxy As Long, ByRef pwAuthnSvc As Long, ByRef pAuthzSvc As Long, ByRef pServerPrincName As Byte, ByRef pAuthnLevel As Long, ByRef pImpLevel As Long, ByRef pAuthInfo As Long, ByRef pCapabilites As Long)
Declare Sub CoRegisterChannelHook lib "ole32" (ByVal ExtensionUuid As Long, ByRef pChannelHook As Long)
Declare Sub CoRegisterClassObject lib "ole32" (ByVal rclsid As Long, ByVal pUnk As Long, ByVal dwClsContext As Long, ByVal flags As Long, ByRef lpdwRegister As Long)
Declare Sub CoRegisterMallocSpy lib "ole32" (ByVal pMallocSpy As Long)
Declare Sub CoRegisterMessageFilter lib "ole32" (ByVal lpMessageFilter As Long, ByVal lplpMessageFilter As Long)
Declare Sub CoRegisterPSClsid lib "ole32" (ByVal riid As Long, ByVal rclsid As Long)
Declare Sub CoRegisterSurrogate lib "ole32" (ByRef pSurrogate As SURROGATE)
Declare Sub CoReleaseMarshalData lib "ole32" (ByRef pStm As Long)
Declare Sub CoResumeClassObjects lib "ole32" ()
Declare Sub CoRevertToSelf lib "ole32" ()
Declare Sub CoRevokeClassObject lib "ole32" (ByVal dwRegister As Long)
Declare Sub CoRevokeMallocSpy lib "ole32" ()
Declare Sub CoSetCancelObject lib "ole32" (ByVal pUnk As Long)
Declare Sub CoSetProxyBlanket lib "ole32" (ByVal pProxy As Long, ByVal dwAuthnSvc As Long, ByVal dwAuthzSvc As Long, ByRef pServerPrincName As Byte, ByVal dwAuthnLevel As Long, ByVal dwImpLevel As Long, ByVal pAuthInfo As Long, ByVal dwCapabilities As Long)
Declare Sub CoSuspendClassObjects lib "ole32" ()
Declare Sub CoSwitchCallContext lib "ole32" (ByVal pNewObject As Long, ByVal ppOldObject As Long)
Declare Sub CoTaskMemFree lib "ole32" (pv As Any)
Declare Sub CoTestCancel lib "ole32" ()
Declare Sub CoTreatAsClass lib "ole32" (ByVal clsidOld As Long, ByVal clsidNew As Long)
Declare Sub CoUninitialize lib "ole32" ()
Declare Sub CoUnmarshalHresult lib "ole32" (ByRef pstm As Long, ByRef phresult As Long)
Declare Sub CoUnmarshalInterface lib "ole32" (ByRef pStm As Long, ByVal riid As Long, ppv As Any)
Declare Sub CoWaitForMultipleHandles lib "ole32" (ByVal dwFlags As Long, ByVal dwTimeout As Long, ByVal cHandles As Long, ByRef pHandles As Long, ByRef lpdwindex As Long)
*/
}

#endif //__OLE32_H__