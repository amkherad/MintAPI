#ifndef __ATL_H__
#define __ATL_H__

#pragma pack(4)


#pragma pack()
[
    dllname("atl.dll"),
    helpstring("Access to API functions within the atl.dll system file.")
]
module ATL {
[entry("AtlGetVersion"), usesgetlasterror]
    void API_AtlGetVersion([in] Any pReserved);
[entry("AtlComPtrAssign"), usesgetlasterror]
    void API_AtlComPtrAssign([in] long pp, [in] long lp);
[entry("AtlComQIPtrAssign"), usesgetlasterror]
    void API_AtlComQIPtrAssign([in] long pp, [in] long lp, [in] long riid);
//[entry("AtlWaitWithMessageLoop"), usesgetlasterror]
//    void API_AtlWaitWithMessageLoop([in] long hEvent);
}
#endif //__ATL_H__