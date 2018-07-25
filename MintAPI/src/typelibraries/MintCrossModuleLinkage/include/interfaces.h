#ifndef __INTERFACES_H__
#define __INTERFACES_H__

#pragma pack(4)


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

[
    uuid(0ACDC666-0000-396F-000F-00E07FE4DCA7), odl
]
interface IMintAPIWrapperInstance : IUnknown
{
    HRESULT GetCoreModule([out, retval] long* retVal);
    HRESULT GetVersion([out, retval] long* retVal);
    HRESULT Exit();
}

#endif //__INTERFACES_H__