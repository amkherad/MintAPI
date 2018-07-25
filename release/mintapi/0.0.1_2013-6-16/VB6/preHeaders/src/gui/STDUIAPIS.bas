Attribute VB_Name = "UIAPIS"
'@PROJECT_LICENSE
'    Option Explicit
''--------------------------------------------------------------------
''--------------------------------------------------------------------
''--------------------------------------------------------------------
''UI section
''--------------------------------------------------------------------
''--------------------------------------------------------------------
''--------------------------------------------------------------------
'#If Not REMOVE_API_SECTION_USERINTERFACE Then
    Public Type API_WNDCLASSEX
        cbSize As Long
        Style As Long
        lpfnwndproc As Long
        cbClsextra As Long
        cbWndExtra As Long
        hInstance As Long
        hIcon As Long
        hCursor As Long
        hbrBackground As Long
        lpszMenuName As String
        lpszClassName As String
        hIconSm As Long
    End Type
    Public Type API_WNDCLASS
        Style As Long
        lpfnwndproc As Long
        cbClsextra As Long
        cbWndExtra2 As Long
        hInstance As Long
        hIcon As Long
        hCursor As Long
        hbrBackground As Long
        lpszMenuName As String
        lpszClassName As String
    End Type
    Public Declare Function API_CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Public Declare Function API_CreateWindow Lib "user32" Alias "CreateWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
    Public Declare Function API_CreateMDIWindow Lib "user32" Alias "CreateMDIWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hInstance As Long, ByVal lParam As Long) As Long
    Public Declare Function API_IsWindow Lib "user32" Alias "IsWindow" (ByVal hWnd As Long) As Long
    Public Declare Function API_GetFocus Lib "user32" Alias "GetFocus" () As Long
    Public Declare Function API_SetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
    Public Declare Function API_GetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As API_POINTAPI) As Long
    Public Declare Function API_GetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, lpRect As API_RECT) As Long
    Public Declare Function API_GetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
    Public Declare Function API_SetParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    Public Declare Function API_GetTopWindow Lib "user32" Alias "GetTopWindow" (ByVal hWnd As Long) As Long
    Public Declare Function API_WindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Public Declare Function API_ClientToScreen Lib "user32" Alias "ClientToScreen" (ByVal hWnd As Long, lpPoint As API_POINTAPI) As Long
    Public Declare Function API_GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Public Declare Function API_SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
    Public Declare Function API_GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Public Declare Function API_GetParent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
    Public Declare Function API_GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function API_SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function API_SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Declare Function API_ShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Public Declare Function API_FlashWindow Lib "user32" Alias "FlashWindow" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
    Public Declare Function API_EnableWindow Lib "user32" Alias "EnableWindow" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
    Public Declare Function API_SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Function API_FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Public Declare Function API_EnumWindows Lib "user32" Alias "EnumWindows" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_RegisterHotKey Lib "user32" Alias "RegisterHotKey" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'    Public Declare Function API_UnregisterHotKey Lib "user32" Alias "UnregisterHotKey" (ByVal hWnd As Long, ByVal id As Long) As Long
    Public Declare Function API_CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
'    Public Declare Function API_GetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
'    Public Declare Function API_ReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Public Declare Function API_SetCursorPos Lib "user32" Alias "SetCursorPos" (ByVal x As Long, ByVal Y As Long) As Long
    Public Declare Function API_GetScrollPos Lib "user32" Alias "GetScrollPos" (ByVal hWnd As Long, ByVal nBar As Long) As Long
    Public Declare Function API_SetScrollPos Lib "user32" Alias "SetScrollPos" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Boolean) As Long
'    Public Declare Function API_PostMessageA Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Boolean
'    Public Declare Function API_GetKeyboardState Lib "user32" Alias "GetKeyboardState" (pbKeyState As Byte) As Long
'    Public Declare Function API_SetKeyboardState Lib "user32" Alias "SetKeyboardState" (lppbKeyState As Byte) As Long
'    Public Declare Function API_GetKeyState Lib "user32" Alias "GetKeyState" (ByVal nVirtKey As Long) As Integer
    Public Declare Function API_SwapMouseButton Lib "user32" Alias "SwapMouseButton" (ByVal bSwap As Long) As Long
    Public Declare Function API_GetDoubleClickTime Lib "user32" Alias "GetDoubleClickTime" () As Long
    Public Declare Function API_SetDoubleClickTime Lib "user32" Alias "SetDoubleClickTime" (ByVal wCount As Long) As Long
    Public Declare Function API_BringWindowToTop Lib "user32" Alias "BringWindowToTop" (ByVal hWnd As Long) As Long
    Public Declare Function API_IsHungWindow Lib "user32" Alias "IsHungWindow" (ByVal hWnd As Long) As Boolean
'
'    Public Declare Function API_ActivateKeyboardLayout Lib "user32" Alias "ActivateKeyboardLayout" (ByVal HKL As Long, ByVal Flags As Long) As Long
'    Public Declare Function API_AdjustWindowRect Lib "user32" Alias "AdjustWindowRect" (lpRect As rect, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
'    Public Declare Function API_AdjustWindowRectEx Lib "user32" Alias "AdjustWindowRectEx" (lpRect As rect, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
'    Public Declare Function API_AllowSetForegroundWindow Lib "user32" Alias "AllowSetForegroundWindow" (ByVal dwProcessId As Long) As Long
'    Public Declare Function API_AnimateWindow Lib "user32" Alias "AnimateWindow" (ByVal hWnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long
'    Public Declare Function API_AnyPopup Lib "user32" Alias "AnyPopup" () As Long
'    Public Declare Function API_AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
'    Public Declare Function API_AppendMenuBynum Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
'
'    Public Declare Function API_ArrangeIconicWindows Lib "user32" Alias "ArrangeIconicWindows" (ByVal hWnd As Long) As Long
'    Public Declare Function API_AttachThreadInput Lib "user32" Alias "AttachThreadInput" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
'    Public Declare Function API_BeginDeferWindowPos Lib "user32" Alias "BeginDeferWindowPos" (ByVal nNumWindows As Long) As Long
'    Public Declare Function API_BeginPaint Lib "user32" Alias "BeginPaint" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
'    Public Declare Function API_BlockInput Lib "user32" Alias "BlockInput" (ByVal fBlockIt As Long) As Long
'    Public Declare Function API_BroadcastSystemMessage Lib "user32" Alias "BroadcastSystemMessage" (ByVal dw As Long, pdw As Long, ByVal un As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_CallMsgFilter Lib "user32" Alias "CallMsgFilterA" (lpMsg As msg, ByVal nCode As Long) As Long
'    Public Declare Function API_CallNextHookEx Lib "user32" Alias "CallNextHookEx" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Integer, lParam As Any) As Long
'    Public Declare Function API_CascadeWindows Lib "user32" Alias "CascadeWindows" (ByVal hwndParent As Long, ByVal wHow As Long, lpRect As rect, ByVal cKids As Long, lpKids As Long) As Integer
'    Public Declare Function API_ChangeClipboardChain Lib "user32" Alias "ChangeClipboardChain" (ByVal hWnd As Long, ByVal hWndNext As Long) As Long
'    Public Declare Function API_ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (ByRef lpDevMode As DEVMODE, ByVal dwFlags As Long) As Long
'    Public Declare Function API_ChangeDisplaySettingsEx Lib "user32" Alias "ChangeDisplaySettingsExA" (ByVal lpszDeviceName As String, ByRef lpDevMode As DEVMODE, ByVal hWnd As Long, ByVal dwFlags As Long, lParam As Any) As Long
'    Public Declare Function API_ChangeMenu Lib "user32" Alias "ChangeMenuA" (ByVal hMenu As Long, ByVal cmd As Long, ByVal lpszNewItem As String, ByVal cmdInsert As Long, ByVal Flags As Long) As Long
'    Public Declare Function API_CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As Long
'    Public Declare Function API_CharLowerBuff Lib "user32" Alias "CharLowerBuffA" (ByVal lpsz As String, ByVal cchLength As Long) As Long
'    Public Declare Function API_CharNext Lib "user32" Alias "CharNextA" (ByVal lpsz As String) As Long
'    Public Declare Function API_CharNextEx Lib "user32" Alias "CharNextExA" (ByVal CodePage As Integer, ByVal lpCurrentChar As String, ByVal dwFlags As Long) As Long
'    Public Declare Function API_CharPrev Lib "user32" Alias "CharPrevA" (ByVal lpszStart As String, ByVal lpszCurrent As String) As Long
'    Public Declare Function API_CharPrevEx Lib "user32" Alias "CharPrevExA" (ByVal CodePage As Integer, ByVal lpStart As String, ByVal lpCurrentChar As String, ByVal dwFlags As Long) As Long
'    Public Declare Function API_CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
'    Public Declare Function API_CharToOemBuff Lib "user32" Alias "CharToOemBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
'    Public Declare Function API_CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As Long
'    Public Declare Function API_CharUpperBuff Lib "user32" Alias "CharUpperBuffA" (ByVal lpsz As String, ByVal cchLength As Long) As Long
'    Public Declare Function API_CheckDlgButton Lib "user32" Alias "CheckDlgButton" (ByVal hDlg As Long, ByVal nIDButton As Long, ByVal wCheck As Long) As Long
'
    Public Declare Function API_CheckMenuItem Lib "user32" Alias "CheckMenuItem" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
    Public Declare Function API_CheckMenuRadioItem Lib "user32" Alias "CheckMenuRadioItem" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
    Public Declare Function API_CheckRadioButton Lib "user32" Alias "CheckRadioButton" (ByVal hDlg As Long, ByVal nIDFirstButton As Long, ByVal nIDLastButton As Long, ByVal nIDCheckButton As Long) As Long
    Public Declare Function API_ChildWindowFromPoint Lib "user32" Alias "ChildWindowFromPoint" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long) As Long
    Public Declare Function API_ChildWindowFromPointEx Lib "user32" Alias "ChildWindowFromPointEx" (ByVal hwndParent As Long, ByVal ptx As Long, ByVal pty As Long, ByVal uFlags As Long) As Long
'    Public Declare Function API_ClientToScreen Lib "user32" Alias "ClientToScreen" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'    Public Declare Function API_ClipCursor Lib "user32" Alias "ClipCursor" (lpRect As rect) As Long
'    Public Declare Function API_ClipCursorBynum Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long) As Long
'    Public Declare Function API_CloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
'    Public Declare Function API_CloseDesktop Lib "user32" Alias "CloseDesktop" (ByVal hDesktop As Long) As Long
    Public Declare Function API_CloseWindow Lib "user32" Alias "CloseWindow" (ByVal hWnd As Long) As Long
'    Public Declare Function API_CloseWindowStation Lib "user32" Alias "CloseWindowStation" (ByVal hWinSta As Long) As Long
'    Public Declare Function API_CopyAcceleratorTable Lib "user32" Alias "CopyAcceleratorTableA" (ByVal hAccelSrc As Long, lpAccelDst As ACCEL, ByVal cAccelEntries As Long) As Long
'    Public Declare Function API_CopyCursor Lib "user32" Alias "CopyCursor" (ByVal hcur As Long) As Long
'    Public Declare Function API_CopyIcon Lib "user32" Alias "CopyIcon" (ByVal hIcon As Long) As Long
'    Public Declare Function API_CopyImage Lib "user32" Alias "CopyImage" (ByVal HANDLE As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'    Public Declare Function API_CopyRect Lib "user32" Alias "CopyRect" (lpDestRect As rect, lpSourceRect As rect) As Long
'    Public Declare Function API_CountClipboardFormats Lib "user32" Alias "CountClipboardFormats" () As Long
'    Public Declare Function API_CreateAcceleratorTable Lib "user32" Alias "CreateAcceleratorTableA" (lpaccl As ACCEL, ByVal cEntries As Long) As Long
'    Public Declare Function API_CreateCaret Lib "user32" Alias "CreateCaret" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'    Public Declare Function API_CreateCursor Lib "user32" Alias "CreateCursor" (ByVal hInstance As Long, ByVal nXhotspot As Long, ByVal nYhotspot As Long, ByVal nWidth As Long, ByVal nHeight As Long, lpANDbitPlane As Any, lpXORbitPlane As Any) As Long
'    Public Declare Function API_CreateDesktop Lib "user32" Alias "CreateDesktopA" (ByVal lpszDesktop As String, ByVal lpszDevice As String, pDevmode As DEVMODE, ByVal dwFlags As Long, ByVal dwDesiredAccess As Long, lpsa As SECURITY_ATTRIBUTES) As Long
'    Public Declare Function API_CreateDialogIndirectParam Lib "user32" Alias "CreateDialogIndirectParamA" (ByVal hInstance As Long, lpTemplate As DLGTEMPLATE, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
'    Public Declare Function API_CreateDialogParam Lib "user32" Alias "CreateDialogParamA" (ByVal hInstance As Long, ByVal lpName As String, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal lParamInit As Long) As Long
'    Public Declare Function API_CreateIcon Lib "user32" Alias "CreateIcon" (ByVal hInstance As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Byte, ByVal nBitsPixel As Byte, lpANDbits As Byte, lpXORbits As Byte) As Long
'    Public Declare Function API_CreateIconFromResource Lib "user32" Alias "CreateIconFromResource" (presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long) As Long
'    Public Declare Function API_CreateIconFromResourceEx Lib "user32" (presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal uFlags As Long) As Long
'    Public Declare Function API_CreateIconIndirect Lib "user32" Alias "CreateIconIndirect" (piconinfo As ICONINFO) As Long
    Public Declare Function API_CreateMenu Lib "user32" Alias "CreateMenu" () As Long
    Public Declare Function API_CreatePopupMenu Lib "user32" Alias "CreatePopupMenu" () As Long
'    Public Declare Function API_CreateWindowStation Lib "user32" Alias "CreateWindowStationA" (ByVal lpwinsta As String, ByVal dwReserved As Long, ByVal dwDesiredAccess As Long, ByRef lpsa As SECURITY_ATTRIBUTES) As Long
'    Public Declare Function API_DdeAccessData Lib "user32" Alias "DdeAccessData" (ByVal hData As Long, pcbDataSize As Long) As Long
'    Public Declare Function API_DdeAddData Lib "user32" Alias "DdeAddData" (ByVal hData As Long, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long) As Long
'    Public Declare Function API_DdeClientTransaction Lib "user32" Alias "DdeClientTransaction" (pData As Byte, ByVal cbData As Long, ByVal hConv As Long, ByVal hszItem As Long, ByVal wFmt As Long, ByVal wType As Long, ByVal dwTimeout As Long, pdwResult As Long) As Long
'    Public Declare Function API_DdeCmpStringHandles Lib "user32" Alias "DdeCmpStringHandles" (ByVal hsz1 As Long, ByVal hsz2 As Long) As Long
'    Public Declare Function API_DdeConnect Lib "user32" Alias "DdeConnect" (ByVal idInst As Long, ByVal hszService As Long, ByVal hszTopic As Long, pCC As CONVCONTEXT) As Long
'    Public Declare Function API_DdeConnectList Lib "user32" Alias "DdeConnectList" (ByVal idInst As Long, ByVal hszService As Long, ByVal hszTopic As Long, ByVal hConvList As Long, pCC As CONVCONTEXT) As Long
'    Public Declare Function API_DdeCreateDataHandle Lib "user32" Alias "DdeCreateDataHandle" (ByVal idInst As Long, pSrc As Byte, ByVal cb As Long, ByVal cbOff As Long, ByVal hszItem As Long, ByVal wFmt As Long, ByVal afCmd As Long) As Long
'    Public Declare Function API_DdeCreateStringHandle Lib "user32" Alias "DdeCreateStringHandleA" (ByVal idInst As Long, ByVal psz As String, ByVal iCodePage As Long) As Long
'    Public Declare Function API_DdeDisconnect Lib "user32" Alias "DdeDisconnect" (ByVal hConv As Long) As Long
'    Public Declare Function API_DdeDisconnectList Lib "user32" Alias "DdeDisconnectList" (ByVal hConvList As Long) As Long
'    Public Declare Function API_DdeEnableCallback Lib "user32" Alias "DdeEnableCallback" (ByVal idInst As Long, ByVal hConv As Long, ByVal wCmd As Long) As Long
'    Public Declare Function API_DdeFreeDataHandle Lib "user32" Alias "DdeFreeDataHandle" (ByVal hData As Long) As Long
'    Public Declare Function API_DdeFreeStringHandle Lib "user32" Alias "DdeFreeStringHandle" (ByVal idInst As Long, ByVal hsz As Long) As Long
'    Public Declare Function API_DdeGetData Lib "user32" Alias "DdeGetData" (ByVal hData As Long, pDst As Byte, ByVal cbMax As Long, ByVal cbOff As Long) As Long
'    Public Declare Function API_DdeGetLastError Lib "user32" Alias "DdeGetLastError" (ByVal idInst As Long) As Long
'    Public Declare Function API_DdeImpersonateClient Lib "user32" Alias "DdeImpersonateClient" (ByVal hConv As Long) As Long
'    Public Declare Function API_DdeInitialize Lib "user32" Alias "DdeInitializeA" (pidInst As Long, ByVal pfnCallback As Long, ByVal afCmd As Long, ByVal ulRes As Long) As Integer
'    Public Declare Function API_DdeKeepStringHandle Lib "user32" Alias "DdeKeepStringHandle" (ByVal idInst As Long, ByVal hsz As Long) As Long
'    Public Declare Function API_DdeNameService Lib "user32" Alias "DdeNameService" (ByVal idInst As Long, ByVal hsz1 As Long, ByVal hsz2 As Long, ByVal afCmd As Long) As Long
'    Public Declare Function API_DdePostAdvise Lib "user32" Alias "DdePostAdvise" (ByVal idInst As Long, ByVal hszTopic As Long, ByVal hszItem As Long) As Long
'    Public Declare Function API_DdeQueryConvInfo Lib "user32" Alias "DdeQueryConvInfo" (ByVal hConv As Long, ByVal idTransaction As Long, pConvInfo As CONVINFO) As Long
'    Public Declare Function API_DdeQueryNextServer Lib "user32" Alias "DdeQueryNextServer" (ByVal hConvList As Long, ByVal hConvPrev As Long) As Long
'    Public Declare Function API_DdeQueryString Lib "user32" Alias "DdeQueryStringA" (ByVal idInst As Long, ByVal hsz As Long, ByVal psz As String, ByVal cchMax As Long, ByVal iCodePage As Long) As Long
'    Public Declare Function API_DdeReconnect Lib "user32" Alias "DdeReconnect" (ByVal hConv As Long) As Long
'    Public Declare Function API_DdeSetQualityOfService Lib "user32" Alias "DdeSetQualityOfService" (ByVal hWndClient As Long, pqosNew As SECURITY_QUALITY_OF_SERVICE, pqosPrev As SECURITY_QUALITY_OF_SERVICE) As Long
'    Public Declare Function API_DdeSetUserHandle Lib "user32" Alias "DdeSetUserHandle" (ByVal hConv As Long, ByVal id As Long, ByVal hUser As Long) As Long
'    Public Declare Function API_DdeUnaccessData Lib "user32" Alias "DdeUnaccessData" (ByVal hData As Long) As Long
'    Public Declare Function API_DdeUninitialize Lib "user32" Alias "DdeUninitialize" (ByVal idInst As Long) As Long
'    Public Declare Function API_DefDlgProc Lib "user32" Alias "DefDlgProcA" (ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_DeferWindowPos Lib "user32" Alias "DeferWindowPos" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'    Public Declare Function API_DefFrameProc Lib "user32" Alias "DefFrameProcA" (ByVal hWnd As Long, ByVal hWndMDIClient As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_DefMDIChildProc Lib "user32" Alias "DefMDIChildProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Public Declare Function API_DeleteMenu Lib "user32" Alias "DeleteMenu" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'    Public Declare Function API_DestroyCaret Lib "user32" Alias "DestroyCaret" () As Long
'    Public Declare Function API_DestroyCursor Lib "user32" Alias "DestroyCursor" (ByVal hCursor As Long) As Long
'    Public Declare Function API_DestroyAcceleratorTable Lib "user32" Alias "DestroyAcceleratorTable" (ByVal haccel As Long) As Long
'    Public Declare Function API_DestroyIcon Lib "user32" Alias "DestroyIcon" (ByVal hIcon As Long) As Long
'    Public Declare Function API_DestroyMenu Lib "user32" Alias "DestroyMenu" (ByVal hMenu As Long) As Long
'    Public Declare Function API_DestroyWindow Lib "user32" Alias "DestroyWindow" (ByVal hWnd As Long) As Long
'    Public Declare Function API_DialogBoxIndirectParam Lib "user32" Alias "DialogBoxIndirectParamA" (ByVal hInstance As Long, hDialogTemplate As DLGTEMPLATE, ByVal hwndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
'    Public Declare Function API_DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
'    Public Declare Function API_DlgDirList Lib "user32" Alias "DlgDirListA" (ByVal hDlg As Long, ByVal lpPathSpec As String, ByVal nIDListBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long
'    Public Declare Function API_DlgDirListComboBox Lib "user32" Alias "DlgDirListComboBoxA" (ByVal hDlg As Long, ByVal lpPathSpec As String, ByVal nIDComboBox As Long, ByVal nIDStaticPath As Long, ByVal wFileType As Long) As Long
'    Public Declare Function API_DlgDirSelectComboBoxEx Lib "user32" Alias "DlgDirSelectComboBoxExA" (ByVal hWndDlg As Long, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idComboBox As Long) As Long
'    Public Declare Function API_DlgDirSelectEx Lib "user32" Alias "DlgDirSelectExA" (ByVal hWndDlg As Long, ByVal lpszPath As String, ByVal cbPath As Long, ByVal idListBox As Long) As Long
'    Public Declare Function API_DragDetect Lib "user32" Alias "DragDetect" (ByVal hWnd As Long, ByVal ptx As Long, ByVal pty As Long) As Long
'    Public Declare Function API_DragObject Lib "user32" Alias "DragObject" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal un As Long, ByVal dw As Long, ByVal hCursor As Long) As Long
'    Public Declare Function API_DrawAnimatedRects Lib "user32" Alias "DrawAnimatedRects" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As rect, lprcTo As rect) As Long
'    Public Declare Function API_DrawCaption Lib "user32" Alias "DrawCaption" (ByVal hWnd As Long, ByVal hDC As Long, lprc As rect, ByVal wFlags As Long) As Long
'    Public Declare Function API_DrawEdge Lib "user32" Alias "DrawEdge" (ByVal hDC As Long, qrc As rect, ByVal edge As Long, ByVal grfFlags As Long) As Long
    Public Declare Function API_DrawFocusRect Lib "user32" Alias "DrawFocusRect" (ByVal hDC As Long, lpRect As API_RECT) As Long
'    Public Declare Function API_DrawFrameControl Lib "user32" Alias "DrawFrameControl" (ByVal hDC As Long, lpRect As rect, ByVal un1 As Long, ByVal un2 As Long) As Long
'    'Public Declare Function API_DrawFocusRect Lib "user32" Alias "DrawFocusRect" (ByVal hDC As Long, lpRect As rect) As Long
'    Public Declare Function API_DrawFrameControl Lib "user32" Alias "DrawFrameControl" (ByVal hDC As Long, lpRect As rect, ByVal un1 As Long, ByVal un2 As Long) As Long
'    Public Declare Function API_DrawIcon Lib "user32" Alias "DrawIcon" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
'    Public Declare Function API_DrawIconEx Lib "user32" Alias "DrawIconEx" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'    Public Declare Function API_DrawMenuBar Lib "user32" Alias "DrawMenuBar" (ByVal hWnd As Long) As Long
'    Public Declare Function API_DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
'    Public Declare Function API_DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpstr As String, ByVal nCount As Long, lpRect As rect, ByVal wFormat As Long) As Long
'    Public Declare Function API_DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As rect, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
'    Public Declare Function API_EmptyClipboard Lib "user32" Alias "EmptyClipboard" () As Long
'    Public Declare Function API_EnableMenuItem Lib "user32" Alias "EnableMenuItem" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
'    Public Declare Function API_EnableScrollBar Lib "user32" Alias "EnableScrollBar" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
'    Public Declare Function API_EnableWindow Lib "user32" Alias "EnableWindow" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
'    Public Declare Function API_EndDeferWindowPos Lib "user32" Alias "EndDeferWindowPos" (ByVal hWinPosInfo As Long) As Long
'    Public Declare Function API_EndDialog Lib "user32" Alias "EndDialog" (ByVal hDlg As Long, ByVal nResult As Long) As Long
'    Public Declare Function API_EndMenu Lib "user32" Alias "EndMenu" () As Long
'    Public Declare Function API_EndPaint Lib "user32" Alias "EndPaint" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
'    Public Declare Function API_EnumChildWindows Lib "user32" Alias "EnumChildWindows" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam&) As Long
'    Public Declare Function API_EnumClipboardFormats Lib "user32" Alias "EnumClipboardFormats" (ByVal wFormat As Long) As Long
'    Public Declare Function API_EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hWinSta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_EnumDesktopWindows Lib "user32" Alias "EnumDesktopWindows" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_EnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesA" (ByVal lpDevice As String, ByVal iDevNum As Long, ByRef lpDisplayDevice As PDISPLAY_DEVICEA, ByVal dwFlags As Long) As Long
'    Public Declare Function API_EnumDisplayMonitors Lib "user32" Alias "EnumDisplayMonitors" (ByVal hDC As Long, ByRef lprcClip As rect, ByRef lpfnEnum As MONITORENUMPROC, ByVal dwData As Long) As Long
'    Public Declare Function API_EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, ByRef lpDevMode As DEVMODE) As Long
'    Public Declare Function API_EnumDisplaySettingsEx Lib "user32" Alias "EnumDisplaySettingsEx" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, ByRef lpDevMode As DEVMODE, ByVal dwFlags As Long) As Long
'    Public Declare Function API_EnumProps Lib "user32" Alias "EnumPropsA" (ByVal hWnd As Long, ByVal lpEnumFunc As Long) As Long
'    Public Declare Function API_EnumPropsEx Lib "user32" Alias "EnumPropsExA" (ByVal hWnd As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_EnumThreadWindows Lib "user32" Alias "EnumThreadWindows" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_EnumWindows Lib "user32" Alias "EnumWindows" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_EqualRect Lib "user32" Alias "EqualRect" (lpRect1 As rect, lpRect2 As rect) As Long
'    Public Declare Function API_ExcludeUpdateRgn Lib "user32" Alias "ExcludeUpdateRgn" (ByVal hDC As Long, ByVal hWnd As Long) As Long
'    Public Declare Function API_ExitWindows Lib "user32" Alias "ExitWindows" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
'    Public Declare Function API_ExitWindowsEx Lib "user32" Alias "ExitWindowsEx" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'    Public Declare Function API_FillRect Lib "user32" Alias "FillRect" (ByVal hDC As Long, lpRect As rect, ByVal hBrush As Long) As Long
'    Public Declare Function API_FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'    Public Declare Function API_FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'    Public Declare Function API_FlashWindow Lib "user32" Alias "FlashWindow" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
'    Public Declare Function API_FlashWindowEx Lib "user32" Alias "FlashWindowEx" (ByRef pfwi As PFLASHWINFO) As Long
'    Public Declare Function API_FrameRect Lib "user32" Alias "FrameRect" (ByVal hDC As Long, lpRect As rect, ByVal hBrush As Long) As Long
'    Public Declare Function API_FreeDDElParam Lib "user32" Alias "FreeDDElParam" (ByVal msg As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_GetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
'    Public Declare Function API_GetAltTabInfo Lib "user32" Alias "GetAltTabInfo" (ByVal hWnd As Long, ByVal iItem As Long, ByRef pati As PALTTABINFO, ByVal pszItemText As String, ByVal cchItemText As Long) As Long
'    Public Declare Function API_GetAncestor Lib "user32" Alias "GetAncestor" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
'    Public Declare Function API_GetAsyncKeyState Lib "user32" Alias "GetAsyncKeyState" (ByVal vKey As Long) As Integer
'    Public Declare Function API_GetCapture Lib "user32" Alias "GetCapture" () As Long
'    Public Declare Function API_GetCaretBlinkTime Lib "user32" Alias "GetCaretBlinkTime" () As Long
'    Public Declare Function API_GetCaretPos Lib "user32" Alias "GetCaretPos" (lpPoint As POINTAPI) As Long
'    Public Declare Function API_GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long
'    Public Declare Function API_GetClassInfoEx Lib "user32" Alias "GetClassInfoExA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClassEx As WNDCLASSEX) As Long
'    Public Declare Function API_GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'    Public Declare Function API_GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'    Public Declare Function API_GetClassWord Lib "user32" Alias "GetClassWord" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'    Public Declare Function API_GetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As Long, lpRect As rect) As Long
'    Public Declare Function API_GetClipboardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
'    Public Declare Function API_GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
'    Public Declare Function API_GetClipboardOwner Lib "user32" Alias "GetClipboardOwner" () As Long
'    Public Declare Function API_GetClipboardSequenceNumber Lib "user32" Alias "GetClipboardSequenceNumber" () As Long
'    Public Declare Function API_GetClipboardViewer Lib "user32" Alias "GetClipboardViewer" () As Long
'    Public Declare Function API_GetClipCursor Lib "user32" Alias "GetClipCursor" (lprc As rect) As Long
'    Public Declare Function API_GetComboBoxInfo Lib "user32" Alias "GetComboBoxInfo" (ByVal hwndCombo As Long, ByRef pcbi As PCOMBOBOXINFO) As Long
'    Public Declare Function API_GetCursor Lib "user32" Alias "GetCursor" () As Long
'    Public Declare Function API_GetCursorInfo Lib "user32" Alias "GetCursorInfo" (ByRef pci As PCURSORINFO) As Long
'    Public Declare Function API_GetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
'    Public Declare Function API_GetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
'    Public Declare Function API_GetDCEx Lib "user32" Alias "GetDCEx" (ByVal hWnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
'    Public Declare Function API_GetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
'    Public Declare Function API_GetDialogBaseUnits Lib "user32" Alias "GetDialogBaseUnits" () As Long
'    Public Declare Function API_GetDlgCtrlID Lib "user32" Alias "GetDlgCtrlID" (ByVal hWnd As Long) As Long
'    Public Declare Function API_GetDlgItem Lib "user32" Alias "GetDlgItem" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
'    Public Declare Function API_GetDlgItemInt Lib "user32" Alias "GetDlgItemInt" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpTranslated As Long, ByVal bSigned As Long) As Long
'    Public Declare Function API_GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
'    Public Declare Function API_GetDoubleClickTime Lib "user32" Alias "GetDoubleClickTime" () As Long
'    Public Declare Function API_GetFocus Lib "user32" Alias "GetFocus" () As Long
'    Public Declare Function API_GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Long
'    Public Declare Function API_GetGuiResources Lib "user32" Alias "GetGuiResources" (ByVal hProcess As Long, ByVal uiFlags As Long) As Long
'    Public Declare Function API_GetGUIThreadInfo Lib "user32" Alias "GetGUIThreadInfo" (ByVal idThread As Long, ByRef pgui As PGUITHREADINFO) As Long
'    Public Declare Function API_GetIconInfo Lib "user32" Alias "GetIconInfo" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
'    Public Declare Function API_GetInputState Lib "user32" Alias "GetInputState" () As Long
'    Public Declare Function API_GetKBCodePage Lib "user32" Alias "GetKBCodePage" () As Long
'    Public Declare Function API_GetKeyboardLayout Lib "user32" Alias "GetKeyboardLayout" (ByVal dwLayout As Long) As Long
'    Public Declare Function API_GetKeyboardLayoutList Lib "user32" Alias "GetKeyboardLayoutList" (ByVal nBuff As Long, lpList As Long) As Long
'    Public Declare Function API_GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
'    Public Declare Function API_GetKeyboardState Lib "user32" Alias "GetKeyboardState" (pbKeyState As Byte) As Long
'    Public Declare Function API_GetKeyboardType Lib "user32" Alias "GetKeyboardType" (ByVal nTypeFlag As Long) As Long
'    Public Declare Function API_GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
'    Public Declare Function API_GetKeyState Lib "user32" Alias "GetKeyState" (ByVal nVirtKey As Long) As Integer
'    Public Declare Function API_GetLastActivePopup Lib "user32" Alias "GetLastActivePopup" (ByVal hwndOwnder As Long) As Long
'    Public Declare Function API_GetLastInputInfo Lib "user32" Alias "GetLastInputInfo" (ByRef plii As PLASTINPUTINFO) As Long
'    Public Declare Function API_GetListBoxInfo Lib "user32" Alias "GetListBoxInfo" (ByVal hWnd As Long) As Long
'    Public Declare Function API_GetMenu Lib "user32" Alias "GetMenu" (ByVal hWnd As Long) As Long
'    Public Declare Function API_GetMenuBarInfo Lib "user32" Alias "GetMenuBarInfo" (ByVal hWnd As Long, ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As PMENUBARINFO) As Long
'    Public Declare Function API_GetMenuCheckMarkDimensions Lib "user32" Alias "GetMenuCheckMarkDimensions" () As Long
'    Public Declare Function API_GetMenuContextHelpId Lib "user32" Alias "GetMenuContextHelpId" (ByVal hMenu As Long) As Long
'    Public Declare Function API_GetMenuDefaultItem Lib "user32" Alias "GetMenuDefaultItem" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
'    Public Declare Function API_GetMenuInfo Lib "user32" Alias "GetMenuInfo" (ByVal hMenu As Long, ByRef LPMENUINFO As MENUINFO) As Long
'    Public Declare Function API_GetMenuItemCount Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Long) As Long
'    Public Declare Function API_GetMenuItemID Lib "user32" Alias "GetMenuItemID" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'    Public Declare Function API_GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpMenuItemInfo As MENUITEMINFO) As Long
'    Public Declare Function API_GetMenuItemRect Lib "user32" Alias "GetMenuItemRect" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As rect) As Long
'    Public Declare Function API_GetMenuState Lib "user32" Alias "GetMenuState" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
'    Public Declare Function API_GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
'    Public Declare Function API_GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
'    Public Declare Function API_GetMessageExtraInfo Lib "user32" Alias "GetMessageExtraInfo" () As Long
'    Public Declare Function API_GetMessagePos Lib "user32" Alias "GetMessagePos" () As Long
'    Public Declare Function API_GetMessageTime Lib "user32" Alias "GetMessageTime" () As Long
'    Public Declare Function API_GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByRef hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
'    Public Declare Function API_GetMouseMovePointsEx Lib "user32" Alias "GetMouseMovePointsEx" (ByVal cbSize As Long, ByRef lppt As MOUSEMOVEPOINT, ByRef lpptBuf As MOUSEMOVEPOINT, ByVal nBufPoints As Long, ByVal resolution As Long) As Long
'    Public Declare Function API_GetNextDlgGroupItem Lib "user32" Alias "GetNextDlgGroupItem" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
'    Public Declare Function API_GetNextDlgTabItem Lib "user32" Alias "GetNextDlgTabItem" (ByVal hDlg As Long, ByVal hCtl As Long, ByVal bPrevious As Long) As Long
'    Public Declare Function API_GetNextWindow Lib "user32" Alias "GetNextWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
'    Public Declare Function API_GetOpenClipboardWindow Lib "user32" Alias "GetOpenClipboardWindow" () As Long
'    Public Declare Function API_GetParent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
'    Public Declare Function API_GetPriorityClipboardFormat Lib "user32" Alias "GetPriorityClipboardFormat" (lpPriorityList As Long, ByVal nCount As Long) As Long
'    Public Declare Function API_GetProcessDefaultLayout Lib "user32" Alias "GetProcessDefaultLayout" (ByRef pdwDefaultLayout As Long) As Long
'    Public Declare Function API_GetProcessWindowStation Lib "user32" Alias "GetProcessWindowStation" () As Long
'    Public Declare Function API_GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'    Public Declare Function API_GetQueueStatus Lib "user32" Alias "GetQueueStatus" (ByVal fuFlags As Long) As Long
'    Public Declare Function API_GetScrollBarInfo Lib "user32" Alias "GetScrollBarInfo" (ByVal hWnd As Long, ByVal idObject As Long, ByRef psbi As PSCROLLBARINFO) As Long
'    Public Declare Function API_GetScrollInfo Lib "user32" Alias "GetScrollInfo" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
'    Public Declare Function API_GetScrollPos Lib "user32" Alias "GetScrollPos" (ByVal hWnd As Long, ByVal nBar As Long) As Long
'    Public Declare Function API_GetScrollRange Lib "user32" Alias "GetScrollRange" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
'    Public Declare Function API_GetSubMenu Lib "user32" Alias "GetSubMenu" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'    Public Declare Function API_GetSysColor Lib "user32" Alias "GetSysColor" (ByVal nIndex As Long) As Long
'    Public Declare Function API_GetSysColorBrush Lib "user32" Alias "GetSysColorBrush" (ByVal nIndex As Long) As Long
'    Public Declare Function API_GetSystemMenu Lib "user32" Alias "GetSystemMenu" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
'    Public Declare Function API_GetSystemMetrics Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
'    Public Declare Function API_GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hDC As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long
'    Public Declare Function API_GetThreadDesktop Lib "user32" Alias "GetThreadDesktop" (ByVal dwThread As Long) As Long
'    Public Declare Function API_GetTitleBarInfo Lib "user32" Alias "GetTitleBarInfo" (ByVal hWnd As Long, ByRef pti As PTITLEBARINFO) As Long
'    Public Declare Function API_GetTopWindow Lib "user32" Alias "GetTopWindow" (ByVal hWnd As Long) As Long
'    Public Declare Function API_GetUpdateRect Lib "user32" Alias "GetUpdateRect" (ByVal hWnd As Long, lpRect As rect, ByVal bErase As Long) As Long
'    Public Declare Function API_GetUpdateRgn Lib "user32" Alias "GetUpdateRgn" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal fErase As Long) As Long
'    Public Declare Function API_GetUserObjectInformation Lib "user32" Alias "GetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
'    Public Declare Function API_GetUserObjectSecurity Lib "user32" Alias "GetUserObjectSecurity" (ByVal hObj As Long, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
'    Public Declare Function API_GetWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'    Public Declare Function API_GetWindowContextHelpId Lib "user32" Alias "GetWindowContextHelpId" (ByVal hWnd As Long) As Long
'    Public Declare Function API_GetWindowDC Lib "user32" Alias "GetWindowDC" (ByVal hWnd As Long) As Long
'    Public Declare Function API_GetWindowInfo Lib "user32" Alias "GetWindowInfo" (ByVal hWnd As Long, ByRef pwi As PWINDOWINFO) As Long
'    Public Declare Function API_GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'    Public Declare Function API_GetWindowModuleFileName Lib "user32" Alias "GetWindowModuleFileNameA" (ByVal hWnd As Long, ByVal pszFileName As String, ByVal cchFileNameMax As Long) As Long
'    Public Declare Function API_GetWindowPlacement Lib "user32" Alias "GetWindowPlacement" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
'    Public Declare Function API_GetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, lpRect As rect) As Long
'    Public Declare Function API_GetWindowRgn Lib "user32" Alias "GetWindowRgn" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
'    Public Declare Function API_GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'    Public Declare Function API_GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'    Public Declare Function API_GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
'    Public Declare Function API_GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
'    Public Declare Function API_GetWindowWord Lib "user32" Alias "GetWindowWord" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
'    Public Declare Function API_GrayString Lib "user32" Alias "GrayStringA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpOutputFunc As Long, ByVal lpData As Long, ByVal nCount As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'    Public Declare Function API_GrayStringByString Lib "user32" Alias "GrayStringA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpOutputFunc As Long, ByVal lpData As String, ByVal nCount As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'    Public Declare Function API_HideCaret Lib "user32" Alias "HideCaret" (ByVal hWnd As Long) As Long
'    Public Declare Function API_HiliteMenuItem Lib "user32" Alias "HiliteMenuItem" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
'    Public Declare Function API_ImpersonateDdeClientWindow Lib "user32" Alias "ImpersonateDdeClientWindow" (ByVal hWndClient As Long, ByVal hWndServer As Long) As Long
'    Public Declare Function API_IMPQueryIME Lib "user32" Alias "IMPQueryIMEA" (ByRef LPIMEPROA As IMEPROA) As Long
'    Public Declare Function API_IMPSetIME Lib "user32" Alias "IMPSetIMEA" (ByVal hWnd As Long, ByRef LPIMEPROA As IMEPROA) As Long
'    Public Declare Function API_IMPGetIME Lib "user32" Alias "IMPGetIMEA" (ByVal hWnd As Long, ByRef LPIMEPROA As IMEPROA) As Long
'    Public Declare Function API_InSendMessage Lib "user32" Alias "InSendMessage" () As Long
'    Public Declare Function API_InSendMessageEx Lib "user32" Alias "InSendMessageEx" (lpReserved As Any) As Long
'    Public Declare Function API_InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
'    Public Declare Function API_InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
'    Public Declare Function API_IntersectRect Lib "user32" Alias "IntersectRect" (lpDestRect As rect, lpSrc1Rect As rect, lpSrc2Rect As rect) As Long
'    Public Declare Function API_InvertRect Lib "user32" Alias "InvertRect" (ByVal hDC As Long, lpRect As rect) As Long
'    Public Declare Function API_InvalidateRect Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, lpRect As rect, ByVal bErase As Long) As Long
'    Public Declare Function API_InvalidateRectBynum Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
'    Public Declare Function API_InvalidateRgn Lib "user32" Alias "InvalidateRgn" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
'    Public Declare Function API_IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
'    Public Declare Function API_IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
'    Public Declare Function API_IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
'    Public Declare Function API_IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
'    Public Declare Function API_IsChild Lib "user32" Alias "IsChild" (ByVal hwndParent As Long, ByVal hWnd As Long) As Long
'    Public Declare Function API_IsClipboardFormatAvailable Lib "user32" Alias "IsClipboardFormatAvailable" (ByVal wFormat As Long) As Long
'    Public Declare Function API_IsDialogMessage Lib "user32" Alias "IsDialogMessageA" (ByVal hDlg As Long, lpMsg As msg) As Long
'    Public Declare Function API_IsDlgButtonChecked Lib "user32" Alias "IsDlgButtonChecked" (ByVal hDlg As Long, ByVal nIDButton As Long) As Long
'    Public Declare Function API_IsIconic Lib "user32" Alias "IsIconic" (ByVal hWnd As Long) As Long
'    Public Declare Function API_IsMenu Lib "user32" Alias "IsMenu" (ByVal hMenu As Long) As Long
'    Public Declare Function API_IsRectEmpty Lib "user32" Alias "IsRectEmpty" (lpRect As rect) As Long
    Public Declare Function API_IsWindowEnabled Lib "user32" Alias "IsWindowEnabled" (ByVal hWnd As Long) As Long
'    Public Declare Function API_IsWindowUnicode Lib "user32" Alias "IsWindowUnicode" (ByVal hWnd As Long) As Long
    Public Declare Function API_IsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal hWnd As Long) As Long
'    Public Declare Function API_IsZoomed Lib "user32" Alias "IsZoomed" (ByVal hWnd As Long) As Long
'    Public Declare Function API_KillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
'    Public Declare Function API_LoadAccelerators Lib "user32" Alias "LoadAcceleratorsA" (ByVal hInstance As Long, ByVal lpTableName As String) As Long
'    Public Declare Function API_LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
'    Public Declare Function API_LoadBitmapBynum Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long
'    Public Declare Function API_LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
'    Public Declare Function API_LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
'    Public Declare Function API_LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
'    Public Declare Function API_LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
'    Public Declare Function API_LoadIconBynum Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
'    Public Declare Function API_LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'    Public Declare Function API_LoadImageBynum Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'    Public Declare Function API_LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As Long
'    Public Declare Function API_LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As String) As Long
'    Public Declare Function API_LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (ByVal lpMenuTemplate As Long) As Long
'    Public Declare Function API_LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
'    Public Declare Function API_LockSetForegroundWindow Lib "user32" Alias "LockSetForegroundWindow" (ByVal uLockCode As Long) As Long
'    Public Declare Function API_LockWindowUpdate Lib "user32" Alias "LockWindowUpdate" (ByVal hwndLock As Long) As Long
'    Public Declare Function API_LockWorkStation Lib "user32" Alias "LockWorkStation" () As Long
'    Public Declare Function API_LookupIconIdFromDirectory Lib "user32" Alias "LookupIconIdFromDirectory" (presbits As Byte, ByVal fIcon As Long) As Long
'    Public Declare Function API_LookupIconIdFromDirectoryEx Lib "user32" Alias "LookupIconIdFromDirectoryEx" (presbits As Byte, ByVal fIcon As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
'    Public Declare Function API_MapDialogRect Lib "user32" Alias "MapDialogRect" (ByVal hDlg As Long, lpRect As rect) As Long
'    Public Declare Function API_MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
'    Public Declare Function API_MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExA" (ByVal uCode As Long, ByVal uMapType As Long, ByVal dwhkl As Long) As Long
'    Public Declare Function API_MapWindowPoints Lib "user32" Alias "MapWindowPoints" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
'    Public Declare Function API_MenuItemFromPoint Lib "user32" Alias "MenuItemFromPoint" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal ptx As Long, ByVal pty As Long) As Long
'    Public Declare Function API_MessageBeep Lib "user32" Alias "MessageBeep" (ByVal wType As Long) As Long
'    Public Declare Function API_MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
'    Public Declare Function API_MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
'    Public Declare Function API_MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectA" (lpMsgBoxParams As MSGBOXPARAMS) As Long
'    Public Declare Function API_ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
'    Public Declare Function API_ModifyMenuBynum Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
'    Public Declare Function API_MoveWindow Lib "user32" Alias "" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'    Public Declare Function API_MsgWaitForMultipleObjectsEx Lib "user32" Alias "MsgWaitForMultipleObjectsEx" (ByVal nCount As Long, ByRef pHandles As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long, ByVal dwFlags As Long) As Long
'    Public Declare Function API_MsgWaitForMultipleObjects Lib "user32" Alias "MsgWaitForMultipleObjects" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
'    Public Declare Function API_OemKeyScan Lib "user32" Alias "OemKeyScan" (ByVal wOemChar As Integer) As Long
'    Public Declare Function API_OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
'    Public Declare Function API_OemToCharBuff Lib "user32" Alias "OemToCharBuffA" (ByVal lpszSrc As String, ByVal lpszDst As String, ByVal cchDstLength As Long) As Long
'    Public Declare Function API_OffsetRect Lib "user32" Alias "OffsetRect" (lpRect As rect, ByVal X As Long, ByVal Y As Long) As Long
'    Public Declare Function API_OpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hWnd As Long) As Long
'    Public Declare Function API_OpenDesktop Lib "user32" Alias "OpenDesktopA" (ByVal lpszDesktop As String, ByVal dwFlags As Long, ByVal fInherit As Long, ByVal dwDesiredAccess As Long) As Long
'    Public Declare Function API_OpenIcon Lib "user32" Alias "OpenIcon" (ByVal hWnd As Long) As Long
'    Public Declare Function API_OpenInputDesktop Lib "user32" Alias "OpenInputDesktop" (ByVal dwFlags As Long, ByVal fInherit As Long, ByVal dwDesiredAccess As Long) As Long
'    Public Declare Function API_OpenWindowStation Lib "user32" Alias "OpenWindowStationA" (ByVal lpszWinSta As String, ByVal fInherit As Long, ByVal dwDesiredAccess As Long) As Long
'    Public Declare Function API_PackDDElParam Lib "user32" Alias "PackDDElParam" (ByVal msg As Long, ByVal uiLo As Long, ByVal uiHi As Long) As Long
'    Public Declare Function API_PaintDesktop Lib "user32" Alias "PaintDesktop" (ByVal hDC As Long) As Long
'    Public Declare Function API_PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
'    Public Declare Function API_PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'    Public Declare Function API_PostMessageBynum Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_PostMessageByString Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'    Public Declare Function API_PostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_PrintWindow Lib "user32" Alias "PrintWindow" (ByVal hWnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
'    Public Declare Function API_PtInRect Lib "user32" Alias "PtInRect" (lpRect As rect, ByVal ptx As Long, ByVal pty As Long) As Long
'    Public Declare Function API_RealChildWindowFromPoint Lib "user32" Alias "RealChildWindowFromPoint" (ByVal hwndParent As Long, ByVal ptParentClientCoords As Struct_MembersOf_POINT) As Long
'    Public Declare Function API_RealGetWindowClass Lib "user32" Alias "RealGetWindowClass" (ByVal hWnd As Long, ByVal pszType As String, ByVal cchType As Long) As Long
'    Public Declare Function API_RedrawWindow Lib "user32" Alias "RedrawWindow" (ByVal hWnd As Long, lprcUpdate As rect, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
    Public Declare Function API_RegisterClass Lib "user32" Alias "RegisterClassA" (Class As API_WNDCLASS) As Long
    Public Declare Function API_RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As API_WNDCLASSEX) As Integer
'    Public Declare Function API_RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
'    Public Declare Function API_RegisterHotKey Lib "user32" Alias "RegisterHotKey" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'    Public Declare Function API_RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
'    Public Declare Function API_ReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Long
'    Public Declare Function API_ReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As Long, ByVal hDC As Long) As Long
'    Public Declare Function API_RemoveMenu Lib "user32" Alias "RemoveMenu" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'    Public Declare Function API_RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'    Public Declare Function API_ReplyMessage Lib "user32" Alias "ReplyMessage" (ByVal lReply As Long) As Long
'    Public Declare Function API_ReuseDDElParam Lib "user32" Alias "ReuseDDElParam" (ByVal lParam As Long, ByVal msgIn As Long, ByVal msgOut As Long, ByVal uiLo As Long, ByVal uiHi As Long) As Long
'    Public Declare Function API_ScreenToClient Lib "user32" Alias "ScreenToClient" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'    Public Declare Function API_ScrollDC Lib "user32" Alias "ScrollDC" (ByVal hDC As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As rect, lprcClip As rect, ByVal hrgnUpdate As Long, lprcUpdate As rect) As Long
'    Public Declare Function API_ScrollWindow Lib "user32" Alias "ScrollWindow" (ByVal hWnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As rect, lpClipRect As rect) As Long
'    Public Declare Function API_ScrollWindowEx Lib "user32" Alias "ScrollWindowEx" (ByVal hWnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As rect, lprcClip As rect, ByVal hrgnUpdate As Long, lprcUpdate As rect, ByVal fuScroll As Long) As Long
'    Public Declare Function API_SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_SendIMEMessageEx Lib "user32" Alias "SendIMEMessageExA" (ByVal hWnd As Long, ByVal lParam As Long) As Long
'    'Public Declare Function API_SendInput lib "user32" (ByVal cInputs As Long,  pInputs As INPUT, ByVal cbSize As Long) As Long
'    Public Declare Function API_SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'    Public Declare Function API_SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'    Public Declare Function API_SendMessageCallback Lib "user32" Alias "SendMessageCallbackA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lpResultCallBack As Long, ByVal dwData As Long) As Long
'    Public Declare Function API_SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
'    Public Declare Function API_SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    Public Declare Function API_SetActiveWindow Lib "user32" Alias "SetActiveWindow" (ByVal hWnd As Long) As Long
'    Public Declare Function API_SetCapture Lib "user32" Alias "SetCapture" (ByVal hWnd As Long) As Long
'    Public Declare Function API_SetCaretBlinkTime Lib "user32" Alias "SetCaretBlinkTime" (ByVal wMSeconds As Long) As Long
'    Public Declare Function API_SetCaretPos Lib "user32" Alias "SetCaretPos" (ByVal X As Long, ByVal Y As Long) As Long
'    Public Declare Function API_SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'    Public Declare Function API_SetClassWord Lib "user32" Alias "SetClassWord" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
'    Public Declare Function API_SetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As Long) As Long
'    Public Declare Function API_SetClipboardViewer Lib "user32" Alias "SetClipboardViewer" (ByVal hWnd As Long) As Long
'    Public Declare Function API_SetCursor Lib "user32" Alias "SetCursor" (ByVal hCursor As Long) As Long
'    Public Declare Function API_SetCursorPos Lib "user32" Alias "SetCursorPos" (ByVal X As Long, ByVal Y As Long) As Long
'    Public Declare Function API_SetDlgItemInt Lib "user32" Alias "SetDlgItemInt" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wValue As Long, ByVal bSigned As Long) As Long
'    Public Declare Function API_SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
'    Public Declare Function API_SetDoubleClickTime Lib "user32" Alias "SetDoubleClickTime" (ByVal wCount As Long) As Long
'    Public Declare Function API_SetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
'    Public Declare Function API_SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
'    Public Declare Function API_SetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hWnd As Long) As Long
'    Public Declare Function API_SetKeyboardState Lib "user32" Alias "SetKeyboardState" (lppbKeyState As Byte) As Long
'    Public Declare Function API_SetLayeredWindowAttributes Lib "user32" Alias "SetLayeredWindowAttributes" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'    Public Declare Function API_SetMenu Lib "user32" Alias "SetMenu" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
'    Public Declare Function API_SetMenuContextHelpId Lib "user32" Alias "SetMenuContextHelpId" (ByVal hMenu As Long, ByVal dw As Long) As Long
'    Public Declare Function API_SetMenuDefaultItem Lib "user32" Alias "SetMenuDefaultItem" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
'    Public Declare Function API_SetMenuInfo Lib "user32" Alias "SetMenuInfo" (ByVal hMenu As Long, ByRef LPCMENUINFO As CMENUINFO) As Long
'    Public Declare Function API_SetMenuItemBitmaps Lib "user32" Alias "SetMenuItemBitmaps" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
'    Public Declare Function API_SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
'    Public Declare Function API_SetMessageExtraInfo Lib "user32" Alias "SetMessageExtraInfo" (ByVal lParam As Long) As Long
'    Public Declare Function API_SetMessageQueue Lib "user32" Alias "SetMessageQueue" (ByVal cMessagesMax As Long) As Long
'    Public Declare Function API_SetParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'    Public Declare Function API_SetProcessDefaultLayout Lib "user32" Alias "SetProcessDefaultLayout" (ByVal dwDefaultLayout As Long) As Long
'    Public Declare Function API_SetProcessWindowStation Lib "user32" Alias "SetProcessWindowStation" (ByVal hWinSta As Long) As Long
'    Public Declare Function API_SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'    Public Declare Function API_SetRect Lib "user32" Alias "SetRect" (lpRect As rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'    Public Declare Function API_SetRectEmpty Lib "user32" Alias "SetRectEmpty" (lpRect As rect) As Long
'    Public Declare Function API_SetScrollInfo Lib "user32" Alias "SetScrollInfo" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Long) As Long
'    Public Declare Function API_SetScrollPos Lib "user32" Alias "SetScrollPos" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
'    Public Declare Function API_SetScrollRange Lib "user32" Alias "SetScrollRange" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
'    Public Declare Function API_SetSysColors Lib "user32" Alias "SetSysColors" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
'    Public Declare Function API_SetSystemCursor Lib "user32" Alias "SetSystemCursor" (ByVal hcur As Long, ByVal id As Long) As Long
'    Public Declare Function API_SetThreadDesktop Lib "user32" Alias "SetThreadDesktop" (ByVal hDesktop As Long) As Long
'    Public Declare Function API_SetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'    Public Declare Function API_SetUserObjectInformation Lib "user32" Alias "SetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, pvInfo As Any, ByVal nLength As Long) As Long
'    Public Declare Function API_SetUserObjectSecurity Lib "user32" Alias "SetUserObjectSecurity" (ByVal hObj As Long, pSIRequested As Long, pSd As SECURITY_DESCRIPTOR) As Long
'    Public Declare Function API_SetWindowContextHelpId Lib "user32" Alias "SetWindowContextHelpId" (ByVal hWnd As Long, ByVal dwContextHelpId As Long) As Long
'    Public Declare Function API_SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'    Public Declare Function API_SetWindowPlacement Lib "user32" Alias "SetWindowPlacement" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
'    Public Declare Function API_SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'    Public Declare Function API_SetWindowRgn Lib "user32" Alias "SetWindowRgn" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
'    Public Declare Function API_SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
'    Public Declare Function API_SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
'    Public Declare Function API_SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
'    Public Declare Function API_SetWindowWord Lib "user32" Alias "SetWindowWord" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
'    Public Declare Function API_ShowCaret Lib "user32" Alias "ShowCaret" (ByVal hWnd As Long) As Long
'    Public Declare Function API_ShowCursor Lib "user32" Alias "ShowCursor" (ByVal bShow As Long) As Long
'    Public Declare Function API_ShowOwnedPopups Lib "user32" Alias "ShowOwnedPopups" (ByVal hWnd As Long, ByVal fShow As Long) As Long
'    Public Declare Function API_ShowScrollBar Lib "user32" Alias "ShowScrollBar" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
'    Public Declare Function API_ShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'    Public Declare Function API_ShowWindowAsync Lib "user32" Alias "ShowWindowAsync" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'    Public Declare Function API_SubtractRect Lib "user32" Alias "SubtractRect" (lprcDst As rect, lprcSrc1 As rect, lprcSrc2 As rect) As Long
'    Public Declare Function API_SwapMouseButton Lib "user32" Alias "SwapMouseButton" (ByVal bSwap As Long) As Long
'    Public Declare Function API_SwitchDesktop Lib "user32" Alias "SwitchDesktop" (ByVal hDesktop As Long) As Long
'    Public Declare Function API_SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
'    Public Declare Function API_SystemParametersInfoByval Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
'    Public Declare Function API_TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long
'    Public Declare Function API_TileWindows Lib "user32" Alias "TileWindows" (ByVal hwndParent As Long, ByVal wHow As Long, lpRect As rect, ByVal cKids As Long, lpKids As Long) As Integer
'    Public Declare Function API_ToAscii Lib "user32" Alias "ToAscii" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Integer, ByVal fuState As Long) As Long
'    Public Declare Function API_ToAsciiEx Lib "user32" Alias "ToAsciiEx" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpKeyState As Byte, lpChar As Integer, ByVal uFlags As Long, ByVal dwhkl As Long) As Long
'    Public Declare Function API_ToUnicode Lib "user32" Alias "ToUnicode" (ByVal wVirtKey As Long, ByVal wScanCode As Long, lpKeyState As Byte, ByVal pwszBuff As String, ByVal cchBuff As Long, ByVal wFlags As Long) As Long
'    Public Declare Function API_ToUnicodeEx Lib "user32" Alias "ToUnicodeEx" (ByVal wVirtKey As Long, ByVal wScanCode As Long, ByVal lpKeyState As String, ByVal pwszBuff As String, ByVal cchBuff As Long, ByVal wFlags As Long, ByVal dwhkl As Long) As Long
'    Public Declare Function API_TrackPopupMenu Lib "user32" Alias "TrackPopupMenu" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As rect) As Long
'    Public Declare Function API_TrackPopupMenuBynum Lib "user32" Alias "TrackPopupMenu" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Long) As Long
'    Public Declare Function API_TrackPopupMenuEx Lib "user32" Alias "TrackPopupMenuEx" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As Long, lpTPMParams As TPMPARAMS) As Long
'    Public Declare Function API_TranslateAccelerator Lib "user32" Alias "TranslateAcceleratorA" (ByVal hWnd As Long, ByVal hAccTable As Long, lpMsg As msg) As Long
'    Public Declare Function API_TranslateMDISysAccel Lib "user32" Alias "TranslateMDISysAccel" (ByVal hWndClient As Long, lpMsg As msg) As Long
'    Public Declare Function API_TranslateMessage Lib "user32" Alias "TranslateMessage" (lpMsg As msg) As Long
'    Public Declare Function API_UnhookWindowsHook Lib "user32" Alias "UnhookWindowsHook" (ByVal nCode As Long, ByVal pfnFilterProc As Long) As Long
'    Public Declare Function API_UnhookWindowsHookEx Lib "user32" Alias "UnhookWindowsHookEx" (ByVal hHook As Long) As Long
'    Public Declare Function API_UnhookWinEvent Lib "user32" Alias "UnhookWinEvent" (ByRef hWinEventHook As Long) As Long
'    Public Declare Function API_UnionRect Lib "user32" Alias "UnionRect" (lpDestRect As rect, lpSrc1Rect As rect, lpSrc2Rect As rect) As Long
'    Public Declare Function API_UnloadKeyboardLayout Lib "user32" Alias "UnloadKeyboardLayout" (ByVal HKL As Long) As Long
'    Public Declare Function API_UnpackDDElParam Lib "user32" Alias "UnpackDDElParam" (ByVal msg As Long, ByVal lParam As Long, puiLo As Long, puiHi As Long) As Long
'    Public Declare Function API_UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
'    Public Declare Function API_UnregisterDeviceNotification Lib "user32" Alias "UnregisterDeviceNotification" (ByRef HANDLE As Long) As Long
'    Public Declare Function API_UnregisterHotKey Lib "user32" Alias "UnregisterHotKey" (ByVal hWnd As Long, ByVal id As Long) As Long
'    Public Declare Function API_UpdateLayeredWindow Lib "user32" Alias "UpdateLayeredWindow" (ByVal hWnd As Long, ByVal hdcDst As Long, ByRef pptDst As Point, ByRef psize As Size, ByVal hdcSrc As Long, ByRef pptSrc As Point, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
'    Public Declare Function API_UpdateWindow Lib "user32" Alias "UpdateWindow" (ByVal hWnd As Long) As Long
'    Public Declare Function API_UserHandleGrantAccess Lib "user32" Alias "UserHandleGrantAccess" (ByVal hUserHandle As Long, ByVal hJob As Long, ByVal bGrant As Long) As Long
'    Public Declare Function API_ValidateRect Lib "user32" Alias "ValidateRect" (ByVal hWnd As Long, lpRect As rect) As Long
'    Public Declare Function API_ValidateRectBynum Lib "user32" Alias "ValidateRect" (ByVal hWnd As Long, ByVal lpRect As Long) As Long
'    Public Declare Function API_ValidateRgn Lib "user32" Alias "ValidateRgn" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
'    Public Declare Function API_VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
'    Public Declare Function API_VkKeyScanEx Lib "user32" Alias "VkKeyScanExA" (ByVal ch As Byte, ByVal dwhkl As Long) As Integer
'    Public Declare Function API_WaitForInputIdle Lib "user32" Alias "WaitForInputIdle" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
'    Public Declare Function API_WaitMessage Lib "user32" Alias "WaitMessage" () As Long
'    Public Declare Function API_WindowFromDC Lib "user32" Alias "WindowFromDC" (ByVal hDC As Long) As Long
'    Public Declare Function API_WindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal X As Long, ByVal Y As Long) As Long
'    Public Declare Function API_WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
'    Public Declare Function API_WINNLSEnableIME Lib "user32" Alias "WINNLSEnableIME" (ByVal hWnd As Long, ByVal bool As Long) As Long
'    Public Declare Function API_WINNLSGetEnableStatus Lib "user32" Alias "WINNLSGetEnableStatus" (ByVal hWnd As Long) As Long
'    Public Declare Function API_WINNLSGetIMEHotkey Lib "user32" Alias "WINNLSGetIMEHotkey" (ByVal hWnd As Long) As Long
'    Public Declare Function API_wsprintf Lib "user32" Alias "wsprintf" (ByVal lpstr As String, ByVal lpcstr As String, OptionalArguments As Any) As Long
'
'#End If
