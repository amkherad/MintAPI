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
[entry("VarPtr"), usesgetlasterror]
    long API_VarPtr([in] Any ArrayPtr);
//========================================
// Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
// Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Integer)
// Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
// Declare Sub GetMem8 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Currency)

// Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
// Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
// Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
// Declare Sub PutMem8 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Currency)

// Alternatively one could use:

// Declare Sub GetMem1 Lib "msvbvm60" (Ptr As Any, RetVal As Byte)
// Declare Sub GetMem2 Lib "msvbvm60" (Ptr As Any, RetVal As Integer)
// Declare Sub GetMem4 Lib "msvbvm60" (Ptr As Any, RetVal As Long)
// Declare Sub GetMem8 Lib "msvbvm60" (Ptr As Any, RetVal As Currency)

// Declare Sub PutMem1 Lib "msvbvm60" (Ptr As Any, ByVal NewVal As Byte)
// Declare Sub PutMem2 Lib "msvbvm60" (Ptr As Any, ByVal NewVal As Integer)
// Declare Sub PutMem4 Lib "msvbvm60" (Ptr As Any, ByVal NewVal As Long)
// Declare Sub PutMem8 Lib "msvbvm60" (Ptr As Any, ByVal NewVal As Currency)

}

#endif //__ATL_H__