#ifndef __INTERFACES_H__
#define __INTERFACES_H__

#pragma pack(4)

#pragma pack()
[
    uuid(20000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface IStream : IUnknown
{
    HRESULT Read([in] Any pv, [in] long cb, [out] long * pcbRead);
    HRESULT Write([in] Any pv, [in] long cb, [out] long * pcbWritten);
    HRESULT Seek([in] Int64 dLibMove, [in] long dwOrigin, [in] Int64 pLibNewPosition);
    HRESULT SetSize([in] Int64 LibNewSize);
    HRESULT CopyTo([in] IStream * pstm);
    HRESULT Commit([in] API_STGC grfCommitFlags);
    HRESULT Revert();
    HRESULT LockRegion([in] Int64 LibOffset, [in] Int64 cb, [in] long dwLockType);
    HRESULT UnlockRegion([in] Int64 LibOffset, [in] Int64 cb, [in] long dwLockType);
    HRESULT Stat([in] API_STATSTG * pstatstg, [in] long grfStatFlag);
    HRESULT Clone([out, retval] IStream ** RetVal);
};

[
    uuid(50000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITypeLib : IUnknown
{
    long GetTypeInfoCount();
    HRESULT GetTypeInfo([in] long Index,[out] Any ppITypeInfo);
    HRESULT GetTypeInfoType([in] long Index,[out,retval] API_TYPEKIND *pTKind);
    HRESULT GetTypeInfoOfGuid([in] Any guid,[out] Any ppITypeInfo);
    HRESULT GetLibAttr([out] Any API_TLIBATTR_ppTLibAttr);
    HRESULT GetTypeComp([out] Any ppITypeComp);
    HRESULT GetDocumentation([in] long Index, [out] BSTR *pBstrName, [out] BSTR *pBstrDocString, [out] long *pdwHelpContext, [out] BSTR *pBstrHelpFile);
    HRESULT IsName([in] BSTR szNameBuf, [in] long lHashVal, [out,retval] Boolean* pfName);
    HRESULT FindName([in] BSTR szNameBuf, [in] long lHashVal, [out] Any ppITypeInfo, [out] long* rgMemId, [out,retval] Integer *pcFound);
    void ReleaseTLibAttr([in] Any API_TLIBATTR_pTLibAttr);
};

[
    uuid(30000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITypeComp : IUnknown
{
    HRESULT Bind([in] BSTR szName, [in] long lHashVal, [in] Integer wFlags, [out] Any ppITypeInfo, [out] Any API_DESCKIND_pDescKind, [out] Any API_BINDPTR_pBindPtr);
    HRESULT BindType([in] BSTR szName, [in] long lHashVal, [out] Any ppTInfo, [out] ITypeComp **ppTComp);
};

[
    uuid(10000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITypeInfo : IUnknown
{
    HRESULT GetTypeAttr([out] Any API_TYPEATTR_ppTypeAttr);
    HRESULT GetTypeComp([out] ITypeComp** ppTComp);
    HRESULT GetFuncDesc([in] long Index,[out] Any API_FUNCDESC_ppFuncDesc);
    HRESULT GetVarDesc([in] long Index,[out] Any API_VARDESC_ppVarDesc);
    HRESULT GetNames([in] long memid,[out] BSTR *rgBstrNames,[in] long cMaxNames,[out] long *pcNames);
    HRESULT GetRefTypeOfImplType([in]  long Index,[out] long *pRefType);
    HRESULT GetImplTypeFlags([in] long Index,[out] long *pImplTypeFlags);
    HRESULT GetIDsOfNames([in] BSTR *rgszNames, [in] long cNames, [out] long *pMemId);
    HRESULT Invoke([in] Any pvInstance,[in] long memid,[in] Integer wFlags,[out] Any API_DISPPARAMS_pDispParams,[out] VARIANT *pVarResult,[out] Any* API_EXCEPINFO_pExcepInfo,[out] long *puArgErr);
    HRESULT GetDocumentation([in] long memid,[out] BSTR *pBstrName,[out] BSTR *pBstrDocString,[out] long *pdwHelpContext,[out] BSTR *pBstrHelpFile);
    HRESULT GetDllEntry([in] long memid,[in] API_InvokeKind invKind,[out] BSTR *pBstrDllName,[out] BSTR *pBstrName,[out] Integer *pwOrdinal);
    HRESULT GetRefTypeInfo([in] long hRefType,[out] ITypeInfo **ppTInfo);
    HRESULT AddressOfMember([in] long memid,[in] API_InvokeKind invKind,[out] Any ppv);
    HRESULT CreateInstance([in] IUnknown *pUnkOuter,[in] Any riid,[out] Any ppvObj);
    HRESULT GetMops([in] long memid,[out] BSTR *pBstrMops);
    HRESULT GetContainingTypeLib([out] ITypeLib **ppTLib,[out] long *pIndex);
    HRESULT ReleaseTypeAttr([in] Any API_TYPEATTR_pTypeAttr);
    HRESULT ReleaseFuncDesc([in] Any API_FUNCDESC_pFuncDesc);
    HRESULT ReleaseVarDesc([in] Any API_VARDESC_pVarDesc);
};

[
    uuid(10000100-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITypeInfoOverridable : IUnknown
{
    HRESULT GetTypeAttr([out] API_TYPEATTR* ppTypeAttr);
    HRESULT GetTypeComp([out] ITypeComp** ppTComp);
    HRESULT GetFuncDesc([in] long Index,[out] API_FUNCDESC* ppFuncDesc);
    HRESULT GetVarDesc([in] long Index,[out] API_VARDESC* ppVarDesc);
    HRESULT GetNames([in] long memid,[out] BSTR *rgBstrNames,[in] long cMaxNames,[out] long *pcNames);
    HRESULT GetRefTypeOfImplType([in]  long Index,[out] long *pRefType);
    HRESULT GetImplTypeFlags([in] long Index,[out] long *pImplTypeFlags);
    HRESULT GetIDsOfNames([in] BSTR *rgszNames, [in] long cNames, [out] long *pMemId);
    HRESULT Invoke([in] long pvInstance,[in] long memid,[in] Integer wFlags,[out] API_DISPPARAMS* pDispParams,[out] VARIANT *pVarResult,[out] API_EXCEPINFO** pExcepInfo,[out] long *puArgErr);
    HRESULT GetDocumentation([in] long memid,[out] BSTR **pBstrName,[out] BSTR **pBstrDocString,[out] long *pdwHelpContext,[out] BSTR **pBstrHelpFile);
    HRESULT GetDllEntry([in] long memid,[in] API_InvokeKind invKind,[out] BSTR **pBstrDllName,[out] BSTR **pBstrName,[out] Integer *pwOrdinal);
    HRESULT GetRefTypeInfo([in] long hRefType,[out] ITypeInfo **ppTInfo);
    HRESULT AddressOfMember([in] long memid,[in] API_InvokeKind invKind,[out] long* ppv);
    HRESULT CreateInstance([in] IUnknown *pUnkOuter,[in] API_StdGuid* riid,[out] long* ppvObj);
    HRESULT GetMops([in] long memid,[out] BSTR **pBstrMops);
    HRESULT GetContainingTypeLib([out] ITypeLib **ppTLib,[out] long *pIndex);
    HRESULT ReleaseTypeAttr([in] API_TYPEATTR* pTypeAttr);
    HRESULT ReleaseFuncDesc([in] API_FUNCDESC* pFuncDesc);
    HRESULT ReleaseVarDesc([in] API_VARDESC* pVarDesc);
};

[
    uuid(10000001-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITypeInfo2 : ITypeInfo
{
    HRESULT GetTypeKind([out, retval] API_TypeKind* pTypeKind);
    HRESULT GetTypeFlags([out, retval] long* pTypeFlags);
    HRESULT GetFuncIndexOfMemId([in] long memid, [in] API_InvokeKind invKind, [out, retval] long* pFuncIndex);
    HRESULT GetVarIndexOfMemId([in] long memid, [out, retval] long* pVarIndex);
    HRESULT GetCustData([in] API_StdGuid* Guid, [out] VARIANT* pVarVal);
    HRESULT GetFuncCustData([in] long Index, [in] API_StdGuid* Guid, [out] VARIANT* pVarVal);
    HRESULT GetParamCustData([in] long indexFunc, [in] long indexParam, [in] API_StdGuid* Guid, [out] VARIANT* pVarVal);
    HRESULT GetVarCustData([in] long Index, [in] API_StdGuid* Guid, [out] VARIANT* pVarVal);
    HRESULT GetImplTypeCustData([in] long Index, [in] API_StdGuid* Guid, [out] VARIANT* pVarVal);
    HRESULT GetDocumentation2([in] long memid, [in] long LCID, [out] BSTR* pbstrHelpString, [out] long* pdwHelpStringContext, [out] BSTR* pbstrHelpStringDll);
    HRESULT GetAllCustData([out] API_CUSTDATA* pCustData);
    HRESULT GetAllFuncCustData([in] long Index, [out] API_CUSTDATA* pCustData);
    HRESULT GetAllParamCustData([in] long indexFunc, [in] long indexParam, [out] API_CUSTDATA* pCustData);
    HRESULT GetAllVarCustData([in] long Index, [out] API_CUSTDATA* pCustData);
    HRESULT GetAllImplTypeCustData([in] long Index, [out] API_CUSTDATA* pCustData);
};

[
    uuid(10000101-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITypeInfo2Overridable : ITypeInfo
{
    HRESULT GetTypeKind([out, retval] API_TypeKind* pTypeKind);
    HRESULT GetTypeFlags([out, retval] long* pTypeFlags);
    HRESULT GetFuncIndexOfMemId([in] long memid, [in] API_InvokeKind invKind, [out, retval] long* pFuncIndex);
    HRESULT GetVarIndexOfMemId([in] long memid, [out, retval] long* pVarIndex);
    HRESULT GetCustData([in] API_StdGuid* Guid, [out] VARIANT* pVarVal);
    HRESULT GetFuncCustData([in] long Index, [in] API_StdGuid* Guid, [out] VARIANT* pVarVal);
    HRESULT GetParamCustData([in] long indexFunc, [in] long indexParam, [in] API_StdGuid* Guid, [out] VARIANT* pVarVal);
    HRESULT GetVarCustData([in] long Index, [in] API_StdGuid* Guid, [out] VARIANT* pVarVal);
    HRESULT GetImplTypeCustData([in] long Index, [in] API_StdGuid* Guid, [out] VARIANT* pVarVal);
    HRESULT GetDocumentation2([in] long memid, [in] long LCID, [out] BSTR* pbstrHelpString, [out] long* pdwHelpStringContext, [out] BSTR* pbstrHelpStringDll);
    HRESULT GetAllCustData([out] API_CUSTDATA* pCustData);
    HRESULT GetAllFuncCustData([in] long Index, [out] API_CUSTDATA* pCustData);
    HRESULT GetAllParamCustData([in] long indexFunc, [in] long indexParam, [out] API_CUSTDATA* pCustData);
    HRESULT GetAllVarCustData([in] long Index, [out] API_CUSTDATA* pCustData);
    HRESULT GetAllImplTypeCustData([in] long Index, [out] API_CUSTDATA* pCustData);
};

[
    uuid(50000001-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITypeLib2 : ITypeLib
{
    HRESULT GetCustData([in] API_StdGuid* Guid, [out] VARIANT* pVarVal);
    HRESULT GetLibStatistics([out] long* pcUniqueNames, [out] long* pcchUniqueNames);
    HRESULT GetDocumentation2([in] long Index, [in] long LCID, [out] BSTR* pbstrHelpString, [out] long* pdwHelpStringContext, [out] BSTR* pbstrHelpStringDll);
    HRESULT GetAllCustData([out] API_CUSTDATA* pCustData);
    
};

[
    uuid(60000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ICreateTypeInfo : IUnknown
{
    HRESULT SetGuid([in] API_StdGuid* guid);
    HRESULT SetTypeFlags([in] long uTypeFlags);
    HRESULT SetDocString([in] WString pStrDoc);
    HRESULT SetHelpContext([in] long dwHelpContext);
    HRESULT SetVersion([in] short wMajorVerNum, [in] short wMinorVerNum);
    HRESULT AddRefTypeInfo([in] ITypeInfo* pTInfo, [in] long *phRefType);
    HRESULT AddFuncDesc([in] long Index, [in] API_FUNCDESC *pFuncDesc);
    HRESULT AddImplType([in] long Index, [in] long hRefType);
    HRESULT SetImplTypeFlags([in] long Index, [in] long implTypeFlags);
    HRESULT SetAlignment([in] short cbAlignment);
    HRESULT SetSchema([in] WString pStrSchema);
    HRESULT AddVarDesc([in] long Index, [in] API_VARDESC *pVarDesc);
    HRESULT SetFuncAndParamNames([in] long Index, [in] WString *rgszNames, [in] long cNames);
    HRESULT SetVarName([in] long Index, [in] WString szName);
    HRESULT SetTypeDescAlias([in] API_TYPEDESC *pTDescAlias);
    HRESULT DefineFuncAsDllEntry([in] long Index, [in] WString szDllName, [in] WString szProcName);
    HRESULT SetFuncDocString([in] long Index, [in] WString szDocString);
    HRESULT SetVarDocString([in] long Index, [in] WString szDocString);
    HRESULT SetFuncHelpContext([in] long Index, [in] long dwHelpContext);
    HRESULT SetVarHelpContext([in] long Index, [in] long dwHelpContext);
    HRESULT SetMops([in] long Index, [in] BSTR bstrMops);
    HRESULT SetTypeIdldesc([in] API_IDLDESC *pIdlDesc);
    HRESULT LayOut();
};

[
    uuid(70000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ICreateTypeLib : IUnknown {
    HRESULT CreateTypeInfo([in] WString szName, [in] API_TypeKind tkind, [out] ICreateTypeInfo **ppCTInfo);
    HRESULT SetName([in] WString szName);
    HRESULT SetVersion([in] short wMajorVerNum, [in] short wMinorVerNum);
    HRESULT SetGuid([in] API_StdGuid* guid);
    HRESULT SetDocString([in] WString szDoc);
    HRESULT SetHelpFileName([in] WString szHelpFileName);
    HRESULT SetHelpContext([in] DWORD dwHelpContext);
    HRESULT SetLcid([in] long LCID);
    HRESULT SetLibFlags([in] long uLibFlags);
    HRESULT SaveAllChanges();
};

[
    uuid(00000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface IVBUnknown
{
    long QueryInterface([in] Any riid, [in, out] long * ppvObj);
    long AddRef();
    long Release();
}

[
    uuid(01000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface IVBDispatch : IUnknown
{
    HRESULT GetTypeInfoCount([out, retval] long* pctinfo);
    HRESULT GetTypeInfo([in] long iTInfo, [in] long LCID, [out, retval] ITypeInfo** ppTInfo);
    HRESULT GetIDsOfNames([in] Any riid, [in] String rgszNames, [in] long cNames, [in] long LCID, [out, retval] long* rgDispId);
    HRESULT Invoke([in] long dispIdMember, [in] Any riid, [in] long LCID, [in] long wFlags, [out] Any* API_DISPPARAMS_pDispParams, [out] VARIANT* pVarResult, [out] Any* API_EXCEPINFO_pExcepInfo, [out, retval] long* puArgErr);
}

[
    uuid(02000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface IRecordInfo : IUnknown
{
    HRESULT RecordInit([in] long pvNew);
    HRESULT RecordClear([in] long pvExisting);
    HRESULT RecordCopy([in] long pvExisting, [in] long pvNew);
    HRESULT GetGuid([out, retval] API_StdGuid ** RetVal);
    HRESULT GetName([out, retval] BSTR * RetVal);
    HRESULT GetSize([out, retval] long * RetVal);
    HRESULT GetTypeInfo([out, retval] long * RetVal);
    HRESULT GetField([in] long pvData, [in] LPWSTR szFieldName, [out, retval] VARIANT * RetVal);
    HRESULT GetFieldNoCopy([in] long pvData, [in] LPWSTR szFieldName, [in] VARIANT * pvarField, [in] long * ppvDataCArray);
    HRESULT PutField([in] long wFlags, [in] long pvData, [in] LPWSTR szFieldName, [in] VARIANT * pvarField);
    HRESULT PutFieldNoCopy([in] long wFlags, [in] long pvData, [in] LPWSTR szFieldName, [in] VARIANT * pvarField);
    HRESULT GetFieldNames([in] long * pcNames, [in] BSTR* rgBstrNames);
    MBOOL   IsMatchingType([in] IRecordInfo* pRecordInfo);
    long    RecordCreate();
    HRESULT RecordCreateCopy([in] long * pvSource, [out, retval] long * RetVal);
    HRESULT RecordDestroy([in] long pvRecord);
}

[
    uuid(03000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface IOneArgReturnBool : IUnknown
{
    boolean Call([in] long Arg);
}

[
    uuid(04000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface IOneArgReturnVoid : IUnknown
{
    void Call([in] long lpArg);
}

[
    uuid(05000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface IOneRefReturnLong : IUnknown
{   
    long Call([in] void* Arg);
}

[
    uuid(06000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface IProvideClassInfo : IVBUnknown
{
    long GetClassInfo([in] long * ppTI);
}

[
    uuid(07000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ISortRoutine : IUnknown
{
    void Call([in] long * pSA, [in] long Left, [in] long Right);
}

[
    uuid(08000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ISwap : IUnknown
{
    void Call([in] long Bogus, [in] void* X, [in] void* Y);
}

[
    uuid(09000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITwoArgReturnLong : IUnknown
{
    long Call([in] long X, [in] long Y);
}

[
    uuid(0A000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITwoArgReturnVoid : IUnknown
{
    void Call([in] long X, [in] long Y);
}

[
    uuid(0B000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITwoRefReturnBool : IUnknown
{
    boolean Call([in] void* X, [in] void* Y);
}

[
    uuid(0C000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITwoRefReturnLong : IUnknown
{
    long Call([in] void* X, [in] void* Y);
}

[
    uuid(0D000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface ITwoArgReturnBool : IUnknown
{
    boolean Call([in] long X, [in] long Y);
}

[
    uuid(0E000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface IBasicCall : IUnknown
{
    void Call();
}

[
    uuid(0F000000-7710-0045-7FFF-7ACDC6661234), odl
]
interface IProvideEnumerableVersion : IUnknown
{
    HRESULT GetEnumerableVersion([out, retval] long* retVal);
}

[
    uuid(1F0F0000-7710-0045-7FFF-7ACDC6661234), odl
]
interface IMalloc : IUnknown
{
    long Alloc([in] long cb);
    long Realloc([in] long pv, [in] long cb);
    void Free([in] long pv);
    long GetSize([in] long pv);
    Boolean DidAlloc([in] long pv);
    void HeapMinimize();
}

/* VERY IMPORTANT ! */
[
    uuid(7F87FFFF-7710-0045-7FFF-7ACDC6661234), odl
]
interface IMintHelper : IUnknown
{
    //void memcpy([in] long pDest, [in] long pSrc, [in] long lngSize);
    void memzero([in] long pDest, [in] long lngSize);                           /* 00 */

    long ReadEAX();                                                             /* 01 */

    long ReadFS();                                                              /* 02 */
    void WriteFS([in] long fsValue);                                            /* 03 */

    long ReadESP();                                                             /* 04 */
    void WriteESP([in] long spValue);                                           /* 05 */    

    long ReadCalleeEBP();                                                       /* 06 */
    void WriteCalleeEBP([in] long bpValue);                                     /* 07 */
    long ReadCallerEBP();                                                       /* 08 */
    void WriteCallerEBP([in] long bpValue);                                     /* 09 */

    long ShiftLeft([in] long Value, [in] long Count);                           /* 10 */
    long ShiftRight([in] long Value, [in] long Count);                          /* 11 */

    //void PushL([in] long lValue);                                               /* 12 */    
    //long PopL();                                                                /* 13 */

    //void PushD([in] double dValue);                                             /* 14 */
    //double PopD();                                                              /* 15 */

    long CalleeThis();                                                          /* 12 */
    long CallerThis();                                                          /* 13 */

    long GetIP();                                                               /* 14 */

    HRESULT Reserve([in] long Length);                                          /* 15 */

    long Return();                                                              /* 16 */

    HRESULT CallInt32([in] long FuncPtr, [out, retval] long* retVal);             /* 17 */
    HRESULT CallDbl([in] long FuncPtr, [out, retval] Double* retVal);           /* 18 */
    HRESULT CallInt64([in] long FuncPtr, [out, retval] Currency* retVal);         /* 19 */

    HRESULT Call([in] long FuncPtr);                                            /* 20 */

    long IncVar32([in] long* int32Var);                                         /* 21 */
    long VarInc32([in] long* int32Var);                                         /* 22 */
}

/* [
    uuid(0ACDC666-0000-396F-000F-00E07FE4DCA7), odl
]
interface IMintAPIWrapperInstance : IUnknown
{
    HRESULT GetCoreModule([out, retval] long* retVal);
    HRESULT GetVersion([out, retval] long* retVal);
    HRESULT Exit();
} */

#endif //__INTERFACES_H__