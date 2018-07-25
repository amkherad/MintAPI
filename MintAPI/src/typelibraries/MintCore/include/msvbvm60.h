#ifndef __ATL_H__
#define __ATL_H__

#pragma pack(4)


#pragma pack()
[
    dllname("msvbvm60.dll"),
    helpstring("Microsoft VisualBasic Virtual Machine 6.0 API.")
]
module MSVBVM60 {
[entry("VarPtr"), usesgetlasterror]
    long API_VarPtrArray([in] AnyArr ArrayPtr);
}

#endif //__ATL_H__