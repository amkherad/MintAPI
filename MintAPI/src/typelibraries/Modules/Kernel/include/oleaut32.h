#ifndef __OLEAUT32_H__
#define __OLEAUT32_H__

#pragma pack(4)

typedef struct SafeArrayBound {
    long        cElements;
    long        lLbound;
} SafeArrayBound;

typedef struct SafeArray1d {
    short       cDims;
    short       fFeatures;
    long        cbElements;
    long        cLocks;
    long        pvData;
    long        cElements;
    long        lLbound;
} SafeArray1d;

typedef struct SafeArray {
    short       cDims;
    short       fFeatures;
    long        cbElements;
    long        cLocks;
    long        pvData;
} SafeArray;


#pragma pack()
[
    dllname("OleAut32.dll"),
    helpstring("Access to API functions within the oleaut32.dll system file.")
]
module OleAuto32 {
[entry("LoadTypeLib")]
    HRESULT API_LoadTypeLib([in] String TLpszModule, [out] ITypeLib** TPpTypeLib);
[entry("LoadTypeLibEx")]
    HRESULT API_LoadTypeLibEx([in] String szFile, [in] API_RegKind regkind, [out] ITypeLib** TPpTypeLib);
[entry("LoadRegTypeLib")]
    HRESULT API_LoadRegTypeLib([in] API_StdGuid* rGuid, [in] short wVerMajor, [in] short wVerMinor, [in] long LCID, [out] ITypeLib** TPpTypeLib);

[entry("CreateTypeLib")]
    HRESULT API_CreateTypeLib([in] API_SysKind SysKind, [in] String szFile, [out] ICreateTypeLib** TPpTypeLib);
[entry("CreateTypeLib2")]
    HRESULT API_CreateTypeLibUnicode([in] API_SysKind SysKind, [in] WString szFile, [out] ICreateTypeLib** TPpTypeLib);
    
[entry("QueryPathOfRegTypeLib")]
    HRESULT API_QueryPathOfRegTypeLib([in] API_StdGuid* Guid, [in] short wMaj, [in] short wMin, [in] long LCID, [out] BSTR* lpbstrPathName);

[entry("RegisterTypeLib")]
    HRESULT API_RegisterTypeLib([in] ITypeLib *ptlib, [in] WString szFullPath, [in] WString szHelpDir);
[entry("UnRegisterTypeLib")]
    HRESULT API_UnRegisterTypeLib([in] API_StdGuid* libID, [in] short wVerMajor, [in] short wVerMinor, [in] long LCID, [in] API_SysKind syskind);
[entry("RegisterTypeLibForUser")]
    HRESULT API_RegisterTypeLibForUser([in] ITypeLib *ptlib, [in] WString szFullPath, [in] WString szHelpDir);
[entry("UnRegisterTypeLibForUser")]
    HRESULT API_UnRegisterTypeLibForUser([in] API_StdGuid* libID, [in] short wVerMajor, [in] short wVerMinor, [in] long LCID, [in] API_SysKind syskind);
//========================================
[entry("DispGetIDsOfNames")]
	HRESULT API_DispGetIDsOfNames([in] ITypeInfo*  ptinfo, [out] WString* rgszNames, [in] long cNames, [in] MDISPID*  rgdispid);
[entry("DispGetParam")]
	HRESULT API_DispGetParam([in] API_DISPPARAMS* pdispparams,[in] long position,[in] vbVARTYPE vtTarg,[out] VARIANT* pvarResult, [out] long* puArgErr);
[entry("DispInvoke")]
	HRESULT API_DispInvoke([in] Any pThis, [in] ITypeInfo *ptinfo, [in] MDISPID dispidMember, [in] long wFlags, [in] API_DISPPARAMS *pparams, [out] VARIANT *pvarResult, [in] API_EXCEPINFO pexcepinfo, [out] long *puArgErr );
//========================================
[entry("OleLoadPicture")]
    HRESULT API_OleLoadPicture([in] IStream * pStream, [in] long lSize, [in] MBOOL fRunmode, [in] API_StdGuid * riid, [in] void * ppvObj);
[entry("OleCreatePictureIndirect")]
    long    API_OleCreatePictureIndirect([in] void * lpPictDesc, [in] API_StdGuid * riid, [in] long fOwn, [in] void * lplpvObj);
//========================================
[entry("SafeArrayCopyData")]
    HRESULT API_SafeArrayCopyData([in] long psaSource, [in] long psaTarget);
[entry("SafeArrayCreate")]
    long    API_SafeArrayCreate([in] VBA.vbVarType vt, [in] long cDims, [in] SafeArrayBound * rgsaBounds);
[entry("SafeArrayCreate")]
    long    API_SafeArrayCreateN([in] VBA.vbVarType vt, [in] long cDims, [in] Any rgsaBounds);
[entry("SafeArrayCreateVector")]
    long    API_SafeArrayCreateVector([in] VBA.vbVarType vt, [in] long lLbound, [in] long cElements);
[entry("SafeArrayCreateVectorEx")]
    long    API_SafeArrayCreateVectorEx([in] VBA.vbVarType vt, [in] long lLbound, [in] long cElements, [in, defaultvalue(0)] long pvExtra);
[entry("SafeArrayDestroyData")]
    HRESULT API_SafeArrayDestroyData([in] long psa);
[entry("SafeArrayGetDim")]
    long    API_SafeArrayGetDim([in] long psa);
[entry("SafeArrayGetElemsize")]
    long    API_SafeArrayGetElemsize([in] long psa);
[entry("SafeArrayGetLBound")]
    HRESULT API_SafeArrayGetLBound([in] long psa, [in] long nDim, [out, retval] long * plLbound);
[entry("SafeArrayGetRecordInfo")]
    HRESULT API_SafeArrayGetRecordInfo([in] long psa, [out, retval] IRecordInfo ** RetVal);
[entry("SafeArrayGetUBound")]
    HRESULT API_SafeArrayGetUBound([in] long psa, [in] long nDim, [out, retval] long * RetVal);
[entry("SafeArrayGetVartype")]
    HRESULT API_SafeArrayGetVartype([in] long psa, [out, retval] VBA.vbVarType * RetVal);
[entry("SafeArrayLock")]
    HRESULT API_SafeArrayLock([in] long psa);
[entry("SafeArrayUnlock")]
    HRESULT API_SafeArrayUnlock([in] long psa);
//========================================
[entry("SysAllocString")]
    BSTR    API_SysAllocString([in] String sz);
[entry("SysAllocString")]
    BSTR    API_SysAllocStringPtr([in] long sz);
[entry("SysAllocStringLen")]
    BSTR    API_SysAllocStringLen([in] long psz, [in] long cch);
[entry("SysFreeString")]
    HRESULT API_SysFreeString([in] long sz);
[entry("SysReAllocString")]
    INT     API_SysReAllocString([in] long sz, [in] Byte* pOlechar);
[entry("SysReAllocStringLen")]
    INT     API_SysReAllocStringLen([in] long sz, [in] Byte* pOlechar, [in] long uint);
[entry("SysStringByteLen")]
    INT     API_SysStringByteLen([in] long sz);
[entry("SysStringLen")]
    INT     API_SysStringLen([in] long sz);
//========================================
[entry("VariantCopy")]
    HRESULT API_VariantCopy([in] VARIANT * pvarDest, [in] VARIANT * pvarSrc);
[entry("VariantCopyInd")]
    HRESULT API_VariantCopyInd([in] VARIANT * pvarDest, [in] VARIANT * pvarSrc);
[entry("VariantClear")]
    HRESULT API_VariantClear([in] VARIANT * pvarg);

    
/* 
Public Declare Sub VariantChangeType lib "oleaut32" (ByRef pvargDest As Variant, ByRef pvarSrc As Variant, ByVal wFlags As Integer, ByVal vt As Integer)
Public Declare Sub VariantChangeTypeEx lib "oleaut32" (ByRef pvargDest As Variant, ByRef pvarSrc As Variant, ByVal lcid As Long, ByVal wFlags As Integer, ByVal vt As Integer)
Public Declare Sub VariantInit lib "oleaut32" (ByRef pvarg As Variant)
Public Declare Sub VarAbs lib "oleaut32" (ByRef pvarIn As Variant, ByRef pvarResult As Variant)
Public Declare Sub VarAdd lib "oleaut32" (ByRef pvarLeft As Variant, ByRef pvarRight As Variant, ByRef pvarResult As Variant)
Public Declare Sub VarBoolFromCy lib "oleaut32" (ByVal cyIn As Struct_MembersOf_CY, ByRef pboolOut As Integer)

Public Declare Sub VarFormat lib "oleaut32" (ByRef pvarIn As Variant, ByVal pstrFormat As String, ByVal iFirstDay As Long, ByVal iFirstWeek As Long, ByVal dwFlags As Long, ByVal pbstrOut As Long)
Public Declare Sub VarFormatNumber lib "oleaut32" (ByRef pvarIn As Variant, ByValumDig As Long, ByVal iIncLead As Long, ByVal iUseParens As Long, ByVal iGroup As Long, ByVal dwFlags As Long, ByVal pbstrOut As Long)
Public Declare Sub VarFormatFromTokens lib "oleaut32" (ByRef pvarIn As Variant, ByVal pstrFormat As String, ByVal pbTokCur As String, ByVal dwFlags As Long, ByVal pbstrOut As Long, ByVal lcid As Long)
Public Declare Sub VarFormatDateTime lib "oleaut32" (ByRef pvarIn As Variant, ByValamedFormat As Long, ByVal dwFlags As Long, ByVal pbstrOut As Long)
Public Declare Sub VarFormatCurrency lib "oleaut32" (ByRef pvarIn As Variant, ByValumDig As Long, ByVal iIncLead As Long, ByVal iUseParens As Long, ByVal iGroup As Long, ByVal dwFlags As Long, ByVal pbstrOut As Long)
Public Declare Sub VarFormatPercent lib "oleaut32" (ByRef pvarIn As Variant, ByValumDig As Long, ByVal iIncLead As Long, ByVal iUseParens As Long, ByVal iGroup As Long, ByVal dwFlags As Long, ByVal pbstrOut As Long)
Public Declare Sub VARIANT_UserFree lib "oleaut32" (ByRef pLong As Long, ByRef pVariant As Variant)
Public Declare Sub VarMonthName lib "oleaut32" (ByVal iMonth As Long, ByVal fAbbrev As Long, ByVal dwFlags As Long, ByVal pbstrOut As Long)
Public Declare Sub VarNeg lib "oleaut32" (ByRef pvarIn As Variant, ByRef pvarResult As Variant)
Public Declare Sub VarNumFromParseNum lib "oleaut32" (ByRef pnumprs As NUMPARSE, ByVal rgbDig As String, ByVal dwVtBits As Long, ByRef pvar As Variant)
Public Declare Sub VarOr lib "oleaut32" (ByRef pvarLeft As Variant, ByRef pvarRight As Variant, ByRef pvarResult As VARIANT)
Public Declare Sub VarParseNumFromStr lib "oleaut32" (ByRef strIn As Byte, ByVal lcid As Long, ByVal dwFlags As Long, ByRef pnumprs As NUMPARSE, ByVal rgbDig As String)
Public Declare Sub VarSub lib "oleaut32" (ByRef pvarLeft As Variant, ByRef pvarRight As VARIANT, ByRef pvarResult As Variant)
Public Declare Sub VarXor lib "oleaut32" (ByRef pvarLeft As VARIANT, ByRef pvarRight As VARIANT, ByRef pvarResult As VARIANT)
 */
}


[
    dllname("OleAut32.dll"),
    helpstring("Access to math API functions within the oleaut32.dll system file.")
]
module OleAutoVar32 {
[entry("VarCmp")]
    long API_VarCmp([in] VARIANT* pvarLeft, [in] VARIANT* pvarRight, [in] long lcid, [in] long dwFlags);
[entry("VarImp")]
    HRESULT API_VarImp([in] VARIANT* pvarLeft, [in] VARIANT* pvarRight, [out, retval] VARIANT* pvarResult);
[entry("VarCat")]
    HRESULT API_VarCat([in] VARIANT* pvarLeft, [in] VARIANT* pvarRight, [out, retval] VARIANT* pvarResult);
//========================================
[entry("VarMod")]
    HRESULT API_VarMod([in] VARIANT* pvarLeft, [in] VARIANT* pvarRight, [out, retval] VARIANT* pvarResult);
[entry("VarMul")]
    HRESULT API_VarMul([in] VARIANT* pvarLeft, [in] VARIANT* pvarRight, [out, retval] VARIANT* pvarResult);
[entry("VarPow")]
    HRESULT API_VarPow([in] VARIANT* pvarLeft, [in] VARIANT* pvarRight, [out, retval] VARIANT* pvarResult);
[entry("VarDiv")]
    HRESULT API_VarDiv([in] VARIANT* pvarLeft, [in] VARIANT* pvarRight, [out, retval] VARIANT* pvarResult);
[entry("VarIdiv")]
    HRESULT API_VarIdiv([in] VARIANT* pvarLeft, [in] VARIANT* pvarRight, [out, retval] VARIANT* pvarResult);
//========================================
[entry("VarInt")]
    HRESULT API_VarInt([in] VARIANT* pvarIn, [out, retval] VARIANT* pvarResult);
[entry("VarNot")]
    HRESULT API_VarNot([in] VARIANT* pvarIn, [out, retval] VARIANT* pvarResult);
[entry("VarFix")]
    HRESULT API_VarFix([in] VARIANT* pvarIn, [out, retval] VARIANT* pvarResult);
[entry("VarRound")]
    HRESULT API_VarRound([in] VARIANT* pvarIn, [in] long cDecimals, [out, retval] VARIANT* pvarResult);
[entry("VarEqv")]
    HRESULT API_VarEqv([in] VARIANT* pvarLeft, [in] VARIANT* pvarRight, [out, retval] VARIANT* pvarResult);
}

[
    dllname("OleAut32.dll"),
    helpstring("Access to decimal API functions within the oleaut32.dll system file.")
]
module OleAutoDecimal32 {
[entry("VarDecAbs")]
    HRESULT API_VarDecAbs([in] Any pdecIn, [in] VARIANT* pvarRight, [out] Any pdecResult);
[entry("VarDecAdd")]
    HRESULT API_VarDecAdd([in] Any pdecLeft, [in] Any pdecRight, [out] Any pdecResult);
[entry("VarDecCmp")]
    HRESULT API_VarDecCmp([in] Any pdecLeft, [in] Any pdecRight, [out] Any pdecResult);
[entry("VarDecCmpR8")]
    HRESULT API_VarDecCmpR8([in] Any pdecLeft, [in] double dblRight);
[entry("VarDecDiv")]
    HRESULT API_VarDecDiv([in] Any pdecLeft, [in] Any pdecRight, [out] Any pdecResult);

/* Public Declare Sub VarDecFix lib "oleaut32" (ByRef pdecIn As DECIMAL, ByRef pdecResult As DECIMAL)
Public Declare Sub VarDecFromBool lib "oleaut32" (ByVal boolIn As Integer, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromCy lib "oleaut32" (ByVal cyIn As Struct_MembersOf_CY, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromDate lib "oleaut32" (ByVal dateIn As Double, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromDisp lib "oleaut32" (ByVal pdispIn As Long, ByVal lcid As Long, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromI1 lib "oleaut32" (ByVal cIn As Byte, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromI2 lib "oleaut32" (ByVal uiIn As Integer, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromI4 lib "oleaut32" (ByVal lIn As Long, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromR4 lib "oleaut32" (ByVal fltIn As Single, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromR8 lib "oleaut32" (ByRef dblIn As Double, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromStr lib "oleaut32" (ByRef strIn As Byte, ByVal lcid As Long, ByVal dwFlags As Long, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromUI1 lib "oleaut32" (ByVal bIn As Byte, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromUI2 lib "oleaut32" (ByVal uiIn As Integer, ByRef pdecOut As Variant)
Public Declare Sub VarDecFromUI4 lib "oleaut32" (ByVal ulIn As Long, ByRef pdecOut As Variant)
Public Declare Sub VarDecInt lib "oleaut32" (ByRef pdecIn As DECIMAL, ByRef pdecResult As DECIMAL)
Public Declare Sub VarDecMul lib "oleaut32" (ByRef pdecLeft As DECIMAL, ByRef pdecRight As DECIMAL, ByRef pdecResult As DECIMAL)
Public Declare Sub VarDecNeg lib "oleaut32" (ByRef pdecIn As DECIMAL, ByRef pdecResult As DECIMAL)
Public Declare Sub VarDecRound lib "oleaut32" (ByRef pdecIn As DECIMAL, ByVal cDecimals As Long, ByRef pdecResult As DECIMAL)
Public Declare Sub VarDecSub lib "oleaut32" (ByRef pdecLeft As DECIMAL, ByRef pdecRight As DECIMAL, ByRef pdecResult As DECIMAL)
 */}

#endif //__OLEAUT32_H__