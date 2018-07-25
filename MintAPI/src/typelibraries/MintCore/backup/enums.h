#ifndef __ENUMS_H__
#define __ENUMS_H__

typedef enum OperationStatus {
    Successfull	= 0,
    Failure	    = 1
} OperationStatus;

typedef enum API_STGC {
    STGC_DEFAULT                            = 0,
    STGC_OVERWRITE                          = 1,
    STGC_ONLYIFCURRENT                      = 2,
    STGC_DANGEROUSLYCOMMITMERELYTODISKCACHE = 4,
    STGC_CONSOLIDATE                        = 8
} API_STGC;

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

typedef enum API_StreamSeek {
    SS_SET = 0,
    SS_CUR = 1,
    SS_END = 2
} API_StreamSeek;

typedef enum API_WindowMessage
{
   wm_Null = 0x00,
   wm_Create = 0x01,
   wm_Destroy = 0x02,
   wm_Move = 0x03,
   wm_Size = 0x05,
   wm_Activate = 0x06,
   wm_SetFocus = 0x07,
   wm_KillFocus = 0x08,
   wm_Enable = 0x0A,
   wm_SetRedraw = 0x0B,
   wm_SetText = 0x0C,
   wm_GetText = 0x0D,
   wm_GetTextLength = 0x0E,
   wm_Paint = 0x0F,
   wm_Close = 0x10,
   wm_QueryEndSession = 0x11,
   wm_Quit = 0x12,
   wm_QueryOpen = 0x13,
   wm_EraseBackground = 0x14,
   wm_SystemColorChange = 0x15,
   wm_EndSession = 0x16,
   wm_SystemError = 0x17,
   wm_ShowWindow = 0x18,
   wm_ControlColor = 0x19,
   wm_WinIniChange = 0x1A,
   wm_SettingChange = 0x1A,
   wm_DevModeChange = 0x1B,
   wm_ActivateApplication = 0x1C,
   wm_FontChange = 0x1D,
   wm_TimeChange = 0x1E,
   wm_CancelMode = 0x1F,
   wm_SetCursor = 0x20,
   wm_MouseActivate = 0x21,
   wm_ChildActivate = 0x22,
   wm_QueueSync = 0x23,
   wm_GetMinMaxInfo = 0x24,
   wm_PaintIcon = 0x26,
   wm_IconEraseBackground = 0x27,
   wm_NextDialogControl = 0x28,
   wm_SpoolerStatus = 0x2A,
   wm_DrawItem = 0x2B,
   wm_MeasureItem = 0x2C,
   wm_DeleteItem = 0x2D,
   wm_VKeyToItem = 0x2E,
   wm_CharToItem = 0x2F,

   wm_SetFont = 0x30,
   wm_GetFont = 0x31,
   wm_SetHotkey = 0x32,
   wm_GetHotkey = 0x33,
   wm_QueryDragIcon = 0x37,
   wm_CompareItem = 0x39,
   wm_Compacting = 0x41,
   wm_WindowPositionChanging = 0x46,
   wm_WindowPositionChanged = 0x47,
   wm_Power = 0x48,
   wm_CopyData = 0x4A,
   wm_CancelJournal = 0x4B,
   wm_Notify = 0x4E,
   wm_InputLanguageChangeRequest = 0x50,
   wm_InputLanguageChange = 0x51,
   wm_TCard = 0x52,
   wm_Help = 0x53,
   wm_UserChanged = 0x54,
   wm_NotifyFormat = 0x55,
   wm_ContextMenu = 0x7B,
   wm_StyleChanging = 0x7C,
   wm_StyleChanged = 0x7D,
   wm_DisplayChange = 0x7E,
   wm_GetIcon = 0x7F,
   wm_SetIcon = 0x80,

   wm_NCCreate = 0x81,
   wm_NCDestroy = 0x82,
   wm_NCCalculateSize = 0x83,
   wm_NCHitTest = 0x84,
   wm_NCPaint = 0x85,
   wm_NCActivate = 0x86,
   wm_GetDialogCode = 0x87,
   wm_NCMouseMove = 0xA0,
   wm_NCLeftButtonDown = 0xA1,
   wm_NCLeftButtonUp = 0xA2,
   wm_NCLeftButtonDoubleClick = 0xA3,
   wm_NCRightButtonDown = 0xA4,
   wm_NCRightButtonUp = 0xA5,
   wm_NCRightButtonDoubleClick = 0xA6,
   wm_NCMiddleButtonDown = 0xA7,
   wm_NCMiddleButtonUp = 0xA8,
   wm_NCMiddleButtonDoubleClick = 0xA9,

   wm_KeyFirst = 0x100,
   wm_KeyDown = 0x100,
   wm_KeyUp = 0x101,
   wm_Char = 0x102,
   wm_DeadChar = 0x103,
   wm_SystemKeyDown = 0x104,
   wm_SystemKeyUp = 0x105,
   wm_SystemChar = 0x106,
   wm_SystemDeadChar = 0x107,
   wm_KeyLast = 0x108,

   wm_IMEStartComposition = 0x10D,
   wm_IMEEndComposition = 0x10E,
   wm_IMEComposition = 0x10F,
   wm_IMEKeyLast = 0x10F,

   wm_InitializeDialog = 0x110,
   wm_Command = 0x111,
   wm_SystemCommand = 0x112,
   wm_Timer = 0x113,
   wm_HorizontalScroll = 0x114,
   wm_VerticalScroll = 0x115,
   wm_InitializeMenu = 0x116,
   wm_InitializeMenuPopup = 0x117,
   wm_MenuSelect = 0x11F,
   wm_MenuChar = 0x120,
   wm_EnterIdle = 0x121,

   wm_CTLColorMessageBox = 0x132,
   wm_CTLColorEdit = 0x133,
   wm_CTLColorListbox = 0x134,
   wm_CTLColorButton = 0x135,
   wm_CTLColorDialog = 0x136,
   wm_CTLColorScrollBar = 0x137,
   wm_CTLColorStatic = 0x138,

   wm_MouseFirst = 0x200,
   wm_MouseMove = 0x200,
   wm_LeftButtonDown = 0x201,
   wm_LeftButtonUp = 0x202,
   wm_LeftButtonDoubleClick = 0x203,
   wm_RightButtonDown = 0x204,
   wm_RightButtonUp = 0x205,
   wm_RightButtonDoubleClick = 0x206,
   wm_MiddleButtonDown = 0x207,
   wm_MiddleButtonUp = 0x208,
   wm_MiddleButtonDoubleClick = 0x209,
   wm_MouseWheel = 0x20A,
   wm_MouseHorizontalWheel = 0x20E,

   wm_ParentNotify = 0x210,
   wm_EnterMenuLoop = 0x211,
   wm_ExitMenuLoop = 0x212,
   wm_NextMenu = 0x213,
   wm_Sizing = 0x214,
   wm_CaptureChanged = 0x215,
   wm_Moving = 0x216,
   wm_PowerBroadcast = 0x218,
   wm_DeviceChange = 0x219,

   wm_MDICreate = 0x220,
   wm_MDIDestroy = 0x221,
   wm_MDIActivate = 0x222,
   wm_MDIRestore = 0x223,
   wm_MDINext = 0x224,
   wm_MDIMaximize = 0x225,
   wm_MDITile = 0x226,
   wm_MDICascade = 0x227,
   wm_MDIIconArrange = 0x228,
   wm_MDIGetActive = 0x229,
   wm_MDISetMenu = 0x230,
   wm_EnterSizeMove = 0x231,
   wm_ExitSizeMove = 0x232,
   wm_DropFiles = 0x233,
   wm_MDIRefreshMenu = 0x234,

   wm_IMESetContext = 0x281,
   wm_IMENotify = 0x282,
   wm_IMEControl = 0x283,
   wm_IMECompositionFull = 0x284,
   wm_IMESelect = 0x285,
   wm_IMEChar = 0x286,
   wm_IMEKeyDown = 0x290,
   wm_IMEKeyUp = 0x291,

   wm_MouseHover = 0x2A1,
   wm_NCMouseLeave = 0x2A2,
   wm_MouseLeave = 0x2A3,

   wm_Cut = 0x300,
   wm_Copy = 0x301,
   wm_Paste = 0x302,
   wm_Clear = 0x303,
   wm_Undo = 0x304,

   wm_RenderFormat = 0x305,
   wm_RenderAllFormats = 0x306,
   wm_DestroyClipboard = 0x307,
   wm_DrawClipbard = 0x308,
   wm_PaintClipbard = 0x309,
   wm_VerticalScrollClipBoard = 0x30A,
   wm_SizeClipbard = 0x30B,
   wm_AskClipboardFormatname = 0x30C,
   wm_ChangeClipboardChain = 0x30D,
   wm_HorizontalScrollClipboard = 0x30E,
   wm_QueryNewPalette = 0x30F,
   wm_PaletteIsChanging = 0x310,
   wm_PaletteChanged = 0x311,

   wm_Hotkey = 0x312,
   wm_Print = 0x317,
   wm_PrintClient = 0x318,

   wm_HandHeldFirst = 0x358,
   wm_HandHeldlast = 0x35F,
   wm_PenWinFirst = 0x380,
   wm_PenWinLast = 0x38F,
   wm_CoalesceFirst = 0x390,
   wm_CoalesceLast = 0x39F,
   wm_DDE_First = 0x3E0,
   wm_DDE_Initiate = 0x3E0,
   wm_DDE_Terminate = 0x3E1,
   wm_DDE_Advise = 0x3E2,
   wm_DDE_Unadvise = 0x3E3,
   wm_DDE_Ack = 0x3E4,
   wm_DDE_Data = 0x3E5,
   wm_DDE_Request = 0x3E6,
   wm_DDE_Poke = 0x3E7,
   wm_DDE_Execute = 0x3E8,
   wm_DDE_Last = 0x3E8,

   wm_User = 0x400,
   wm_App = 0x8000
} API_WindowMessage;

#endif //__ENUMS_H__