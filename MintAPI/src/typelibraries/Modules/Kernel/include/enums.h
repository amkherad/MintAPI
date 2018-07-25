#ifndef __ENUMS_H__
#define __ENUMS_H__

typedef enum API_SysKind {
    SYS_Win16	                            = 0,
	SYS_Win32	                            ,
	SYS_Mac	                                ,
	SYS_Win64
} API_SYSKIND;

typedef enum API_RegKind {
    RegKind_DEFAULT                         ,
    RegKind_REGISTER                        ,
    RegKind_NONE                           
} API_RegKind;

typedef enum API_CallConv {	
    CC_FastCall	                            = 0,
    CC_CDECL	                            = 1,
    CC_MscPascal	                        ,
    CC_Pascal	                            ,
    CC_MacPascal	                        ,
    CC_StdCall	                            ,
    CC_FpFastCall	                        ,
    CC_SysCall	                            ,
    CC_MpwCDECL	                            ,
    CC_MpwPascal	                        ,
    CC_MAX	                                
} API_CALLCONV;

typedef enum API_InvokeKind {
    IK_Function	                            = 1,
	IK_PropertyGet	                        = 2,
	IK_PropertyPut	                        = 4,
	IK_PropertyPutRef	                    = 8
} API_InvokeKind;

typedef enum API_TypeKind {
    TK_Enum	                                = 0,
	TK_Record	                            ,
	TK_Module	                            ,
	TK_Interface	                        ,
	TK_Dispatch                             ,
	TK_CoClass	                            ,
	TK_Alias	                            ,
	TK_Union	                            ,
	TK_Max	                                
} API_TypeKind;

typedef enum API_DescKind {
    DK_None	= 0,
    DK_FuncDesc	                            ,
    DK_VarDesc	                            ,
    DK_TypeComp	                            ,
    DK_ImplicitAppObj	                    ,
    DK_MAX
} API_DescKind;

typedef enum API_VarKind {
    VAR_PerInstance	                        = 0,
	VAR_Static	                            ,
	VAR_Const	                            ,
	VAR_Dispatch	                        
} API_VarKind;

typedef enum API_FuncKind {
    FUNC_Virtual	                        = 0,
    FUNC_PureVirtual	                    ,
    FUNC_NonVirtual	                        ,
    FUNC_Static	                            ,
    FUNC_Dispatch	                        
} API_FuncKind;

typedef enum API_StatFlag {
    STFlag_DEFAULT    = 0,
    STFlag_NONAME     = 1,
    STFlag_NOOPEN     = 2
} API_StatFlag;

#endif //__ENUMS_H__