#ifndef ___INCLUDE__
#define ___INCLUDE__

 #define GEN_ANSI
 #define GEN_UNICODE

 #define MAX_PATH 260
    
//private types
    //typedef unsigned char      byte;
    typedef void*              Any;
    typedef void*              AnyArr;
    typedef LPSTR              String;
    #ifdef GEN_UNICODE
    typedef LPWSTR             WString;
    #endif
    typedef unsigned char      Byte;
    typedef long               Long;
    typedef unsigned long      ULong;
    typedef short              Integer;
    typedef int                Boolean;
    typedef currency           Currency;
    typedef currency           Date;
    typedef VARIANT            Variant;

#endif ___INCLUDE__