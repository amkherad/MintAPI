Attribute VB_Name = "mint_apideclarations"
Option Explicit

Public Type API_ANIMATIONINFO
    cbSize As Long
    iMinAnimate As Long
End Type


Public Declare Function API_GetTopWindow Lib "user32" Alias "GetTopWindow" (ByVal hWnd As Long) As Long
Public Declare Function API_SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function API_GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Long
Public Declare Function API_GetFocus Lib "user32" Alias "GetFocus" () As Long
Public Declare Function API_GetCapture Lib "user32" Alias "GetCapture" () As Long
Public Declare Function API_GetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
Public Declare Function API_EnumWindows Lib "user32" Alias "EnumWindows" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function API_WindowFromDC Lib "user32" Alias "WindowFromDC" (ByVal hdc As Long) As Long
Public Declare Function API_WindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function API_FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function API_FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function API_EnumThreadWindows Lib "user32" Alias "EnumThreadWindows" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function API_GetParent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long
Public Declare Function API_SetParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function API_GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Public Declare Function API_SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function API_SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function API_SetWindowRgn Lib "user32" Alias "SetWindowRgn" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function API_SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
Public Declare Function API_SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function API_SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function API_SetWindowWord Lib "user32" Alias "SetWindowWord" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function API_SetWindowPlacement Lib "user32" Alias "SetWindowPlacement" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function API_SetWindowOrgEx Lib "gdi32" Alias "SetWindowOrgEx" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Public Declare Function API_SetWindowExtEx Lib "gdi32" Alias "SetWindowExtEx" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Public Declare Function API_SetWindowContextHelpId Lib "user32" Alias "SetWindowContextHelpId" (ByVal hWnd As Long, ByVal dw As Long) As Long
Public Declare Function API_GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function API_GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function API_GetWindowContextHelpId Lib "user32" Alias "GetWindowContextHelpId" (ByVal hWnd As Long) As Long
Public Declare Function API_GetWindowDC Lib "user32" Alias "GetWindowDC" (ByVal hWnd As Long) As Long
Public Declare Function API_GetWindowExtEx Lib "gdi32" Alias "GetWindowExtEx" (ByVal hdc As Long, lpSize As Size) As Long
Public Declare Function API_GetWindowOrgEx Lib "gdi32" Alias "GetWindowOrgEx" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Public Declare Function API_GetWindowPlacement Lib "user32" Alias "GetWindowPlacement" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function API_GetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function API_GetWindowRgn Lib "user32" Alias "GetWindowRgn" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Public Declare Function API_GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function API_GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function API_GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function API_GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function API_GetWindowWord Lib "user32" Alias "GetWindowWord" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function API_EnumChildWindows Lib "user32" Alias "EnumChildWindows" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function API_EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hwinsta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function API_EnumDesktopWindows Lib "user32" Alias "EnumDesktopWindows" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function API_IsChild Lib "user32" Alias "IsChild" (ByVal hWndParent As Long, ByVal hWnd As Long) As Long
Public Declare Function API_IsMenu Lib "user32" Alias "IsMenu" (ByVal hMenu As Long) As Long
Public Declare Function API_ChildWindowFromPoint Lib "user32" Alias "ChildWindowFromPoint" (ByVal hWndParent As Long, ByVal pt As POINTAPI) As Long
Public Declare Function API_ChildWindowFromPointEx Lib "user32" Alias "ChildWindowFromPointEx" (ByVal hWnd As Long, ByVal pt As POINTAPI, ByVal un As Long) As Long
Public Declare Function API_GetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function API_SetCapture Lib "user32" Alias "SetCapture" (ByVal hWnd As Long) As Long
Public Declare Function API_SetCaretPos Lib "user32" Alias "SetCaretPos" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function API_AdjustWindowRect Lib "user32" Alias "AdjustWindowRect" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Public Declare Function API_AdjustWindowRectEx Lib "user32" Alias "AdjustWindowRectEx" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
Public Declare Function API_SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function API_SendMessageCallback Lib "user32" Alias "SendMessageCallbackA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lpResultCallBack As Long, ByVal dwData As Long) As Long
Public Declare Function API_SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Public Declare Function API_SendNotifyMessage Lib "user32" Alias "SendNotifyMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function API_ScreenToClient Lib "user32" Alias "ScreenToClient" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function API_ScrollWindow Lib "user32" Alias "ScrollWindow" (ByVal hWnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long
Public Declare Function API_ScrollWindowEx Lib "user32" Alias "ScrollWindowEx" (ByVal hWnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT, ByVal fuScroll As Long) As Long
Public Declare Function API_ScrollDC Lib "user32" Alias "ScrollDC" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Public Declare Function API_ClientToScreen Lib "user32" Alias "ClientToScreen" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function API_MapWindowPoints Lib "user32" Alias "MapWindowPoints" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Public Declare Function API_ShowWindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function API_ShowWindowAsync Lib "user32" Alias "ShowWindowAsync" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function API_CloseWindow Lib "user32" Alias "CloseWindow" (ByVal hWnd As Long) As Long
Public Declare Function API_UpdateWindow Lib "user32" Alias "UpdateWindow" (ByVal hWnd As Long) As Long
Public Declare Function API_RedrawWindow Lib "user32" Alias "RedrawWindow" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function API_SetMenu Lib "user32" Alias "SetMenu" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Public Declare Function API_SetMenuContextHelpId Lib "user32" Alias "SetMenuContextHelpId" (ByVal hMenu As Long, ByVal dw As Long) As Long
Public Declare Function API_SetMenuDefaultItem Lib "user32" Alias "SetMenuDefaultItem" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Public Declare Function API_SetMenuItemBitmaps Lib "user32" Alias "SetMenuItemBitmaps" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function API_SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function API_SetMessageQueue Lib "user32" Alias "SetMessageQueue" (ByVal cMessagesMax As Long) As Long
Public Declare Function API_SetMessageExtraInfo Lib "user32" Alias "SetMessageExtraInfo" (ByVal lParam As Long) As Long
Public Declare Function API_GetMenu Lib "user32" Alias "GetMenu" (ByVal hWnd As Long) As Long
Public Declare Function API_GetMenuCheckMarkDimensions Lib "user32" Alias "GetMenuCheckMarkDimensions" () As Long
Public Declare Function API_GetMenuContextHelpId Lib "user32" Alias "GetMenuContextHelpId" (ByVal hMenu As Long) As Long
Public Declare Function API_GetMenuDefaultItem Lib "user32" Alias "GetMenuDefaultItem" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Public Declare Function API_GetMenuItemCount Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Long) As Long
Public Declare Function API_GetMenuItemID Lib "user32" Alias "GetMenuItemID" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function API_GetMenuItemRect Lib "user32" Alias "GetMenuItemRect" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Public Declare Function API_GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function API_GetMenuState Lib "user32" Alias "GetMenuState" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Public Declare Function API_GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function API_GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function API_GetMessageExtraInfo Lib "user32" Alias "GetMessageExtraInfo" () As Long
Public Declare Function API_GetMessagePos Lib "user32" Alias "GetMessagePos" () As Long
Public Declare Function API_GetMessageTime Lib "user32" Alias "GetMessageTime" () As Long
Public Declare Function API_IsWindowEnabled Lib "user32" Alias "IsWindowEnabled" (ByVal hWnd As Long) As Long
Public Declare Function API_IsWindowUnicode Lib "user32" Alias "IsWindowUnicode" (ByVal hWnd As Long) As Long
Public Declare Function API_IsZoomed Lib "user32" Alias "IsZoomed" (ByVal hWnd As Long) As Long
Public Declare Function API_IsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal hWnd As Long) As Long


Public Const SPI_GETACCESSTIMEOUT = 60
Public Const SPI_GETANIMATION = 72
Public Const SPI_GETBEEP = 1
Public Const SPI_GETBORDER = 5
Public Const SPI_GETDEFAULTINPUTLANG = 89
Public Const SPI_GETDRAGFULLWINDOWS = 38
Public Const SPI_GETFASTTASKSWITCH = 35
Public Const SPI_GETFILTERKEYS = 50
Public Const SPI_GETFONTSMOOTHING = 74
Public Const SPI_GETGRIDGRANULARITY = 18
Public Const SPI_GETHIGHCONTRAST = 66
Public Const SPI_GETICONMETRICS = 45
Public Const SPI_GETICONTITLELOGFONT = 31
Public Const SPI_GETICONTITLEWRAP = 25
Public Const SPI_GETKEYBOARDDELAY = 22
Public Const SPI_GETKEYBOARDPREF = 68
Public Const SPI_GETKEYBOARDSPEED = 10
Public Const SPI_GETLOWPOWERACTIVE = 83
Public Const SPI_GETLOWPOWERTIMEOUT = 79
Public Const SPI_GETMENUDROPALIGNMENT = 27
Public Const SPI_GETMOUSE = 3
Public Const SPI_GETMOUSEKEYS = 54
Public Const SPI_GETMOUSETRAILS = 94
Public Const SPI_GETNONCLIENTMETRICS = 41
Public Const SPI_GETPOWEROFFACTIVE = 84
Public Const SPI_GETPOWEROFFTIMEOUT = 80
Public Const SPI_GETSCREENSAVEACTIVE = 16
Public Const SPI_GETSCREENREADER = 70
Public Const SPI_GETSCREENSAVETIMEOUT = 14
Public Const SPI_GETSERIALKEYS = 62
Public Const SPI_GETSHOWSOUNDS = 56
Public Const SPI_GETSOUNDSENTRY = 64
Public Const SPI_GETSTICKYKEYS = 58
Public Const SPI_GETTOGGLEKEYS = 52
Public Const SPI_GETWINDOWSEXTENSION = 92
Public Const SPI_GETWORKAREA = 48
Public Const SPI_ICONHORIZONTALSPACING = 13
Public Const SPI_ICONVERTICALSPACING = 24
Public Const SPI_LANGDRIVER = 12
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const SPI_SETACCESSTIMEOUT = 61
Public Const SPI_SETANIMATION = 73
Public Const SPI_SETBEEP = 2
Public Const SPI_SETBORDER = 6
Public Const SPI_SETCURSORS = 87
Public Const SPI_SETDEFAULTINPUTLANG = 90
Public Const SPI_SETDESKPATTERN = 21
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_SETDOUBLECLICKTIME = 32
Public Const SPI_SETDOUBLECLKHEIGHT = 30
Public Const SPI_SETDOUBLECLKWIDTH = 29
Public Const SPI_SETDRAGFULLWINDOWS = 37
Public Const SPI_SETDRAGHEIGHT = 77
Public Const SPI_SETFASTTASKSWITCH = 36
Public Const SPI_SETDRAGWIDTH = 76
Public Const SPI_SETFILTERKEYS = 51
Public Const SPI_SETFONTSMOOTHING = 75
Public Const SPI_SETGRIDGRANULARITY = 19
Public Const SPI_SETHANDHELD = 78
Public Const SPI_SETHIGHCONTRAST = 67
Public Const SPI_SETICONMETRICS = 46
Public Const SPI_SETICONS = 88
Public Const SPI_SETICONTITLELOGFONT = 34
Public Const SPI_SETICONTITLEWRAP = 26
Public Const SPI_SETKEYBOARDPREF = 69
Public Const SPI_SETKEYBOARDDELAY = 23
Public Const SPI_SETKEYBOARDSPEED = 11
Public Const SPI_SETLANGTOGGLE = 91
Public Const SPI_SETLOWPOWERACTIVE = 85
Public Const SPI_SETLOWPOWERTIMEOUT = 81
Public Const SPI_SETMENUDROPALIGNMENT = 28
Public Const SPI_SETMINIMIZEDMETRICS = 44
Public Const SPI_SETMOUSE = 4
Public Const SPI_SETMOUSEBUTTONSWAP = 33
Public Const SPI_SETMOUSEKEYS = 55
Public Const SPI_SETMOUSETRAILS = 93
Public Const SPI_SETNONCLIENTMETRICS = 42
Public Const SPI_SETPENWINDOWS = 49
Public Const SPI_SETPOWEROFFACTIVE = 86
Public Const SPI_SETPOWEROFFTIMEOUT = 82
Public Const SPI_SETSCREENREADER = 71
Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPI_SETSCREENSAVETIMEOUT = 15
Public Const SPI_SETSERIALKEYS = 63
Public Const SPI_SETSHOWSOUNDS = 57
Public Const SPI_SETSOUNDSENTRY = 65
Public Const SPI_SETSTICKYKEYS = 59
Public Const SPI_SETTOGGLEKEYS = 53
Public Const SPI_SETWORKAREA = 47
Public Const SPIF_SENDWININICHANGE = &H2

