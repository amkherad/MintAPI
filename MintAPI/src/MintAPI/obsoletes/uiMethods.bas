Attribute VB_Name = "uiMethods"
'@PROJECT_LICENSE

Option Explicit

Private Declare Function API_uiMethods_GetDesktopWindow Lib "user32" () As Long
Private Declare Function API_uiMethods_GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function API_uiMethods_ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function API_uiMethods_GetKeyState Lib "user32" Alias "GetKeyState" (ByVal nVirtKey As Long) As Integer
Private Declare Function API_uiMethods_SetCursorPos Lib "user32" Alias "SetCursorPos" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function API_uiMethods_GetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
Private Declare Function API_uiMethods_SwapMouseButton Lib "user32" Alias "SwapMouseButton" (ByVal bSwap As Long) As Long
Private Declare Function API_uiMethods_GetDoubleClickTime Lib "user32" Alias "GetDoubleClickTime" () As Long
Private Declare Function API_uiMethods_SetDoubleClickTime Lib "user32" Alias "SetDoubleClickTime" (ByVal wCount As Long) As Long
Private Declare Function API_uiMethods_SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function API_uiMethods_FlashWindow Lib "user32" Alias "FlashWindow" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Private Declare Function API_uiMethods_BringWindowToTop Lib "user32" Alias "BringWindowToTop" (ByVal hWnd As Long) As Long

Private Const SPI_GETBEEP = 1
Private Const SPI_SETBEEP = 2
Private Const SPI_GETMOUSE = 3
Private Const SPI_SETMOUSE = 4
Private Const SPI_GETBORDER = 5
Private Const SPI_SETBORDER = 6
Private Const SPI_GETKEYBOARDSPEED = 10
Private Const SPI_SETKEYBOARDSPEED = 11
Private Const SPI_LANGDRIVER = 12
Private Const SPI_ICONHORIZONTALSPACING = 13
Private Const SPI_GETSCREENSAVETIMEOUT = 14
Private Const SPI_SETSCREENSAVETIMEOUT = 15
Private Const SPI_GETSCREENSAVEACTIVE = 16
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPI_GETGRIDGRANULARITY = 18
Private Const SPI_SETGRIDGRANULARITY = 19
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPI_SETDESKPATTERN = 21
Private Const SPI_GETKEYBOARDDELAY = 22
Private Const SPI_SETKEYBOARDDELAY = 23
Private Const SPI_ICONVERTICALSPACING = 24
Private Const SPI_GETICONTITLEWRAP = 25
Private Const SPI_SETICONTITLEWRAP = 26
Private Const SPI_GETMENUDROPALIGNMENT = 27
Private Const SPI_SETMENUDROPALIGNMENT = 28
Private Const SPI_SETDOUBLECLKWIDTH = 29
Private Const SPI_SETDOUBLECLKHEIGHT = 30
Private Const SPI_GETICONTITLELOGFONT = 31
Private Const SPI_SETDOUBLECLICKTIME = 32
Private Const SPI_SETMOUSEBUTTONSWAP = 33
Private Const SPI_SETICONTITLELOGFONT = 34
Private Const SPI_GETFASTTASKSWITCH = 35
Private Const SPI_SETFASTTASKSWITCH = 36
Private Const SPI_SETDRAGFULLWINDOWS = 37
Private Const SPI_GETDRAGFULLWINDOWS = 38
Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const SPI_SETNONCLIENTMETRICS = 42
Private Const SPI_GETMINIMIZEDMETRICS = 43
Private Const SPI_SETMINIMIZEDMETRICS = 44
Private Const SPI_GETICONMETRICS = 45
Private Const SPI_SETICONMETRICS = 46
Private Const SPI_SETWORKAREA = 47
Private Const SPI_GETWORKAREA = 48
Private Const SPI_SETPENWINDOWS = 49
Private Const SPI_GETFILTERKEYS = 50
Private Const SPI_SETFILTERKEYS = 51
Private Const SPI_GETTOGGLEKEYS = 52
Private Const SPI_SETTOGGLEKEYS = 53
Private Const SPI_GETMOUSEKEYS = 54
Private Const SPI_SETMOUSEKEYS = 55
Private Const SPI_GETSHOWSOUNDS = 56
Private Const SPI_SETSHOWSOUNDS = 57
Private Const SPI_GETSTICKYKEYS = 58
Private Const SPI_SETSTICKYKEYS = 59
Private Const SPI_GETACCESSTIMEOUT = 60
Private Const SPI_SETACCESSTIMEOUT = 61
Private Const SPI_GETSERIALKEYS = 62
Private Const SPI_SETSERIALKEYS = 63
Private Const SPI_GETSOUNDSENTRY = 64
Private Const SPI_SETSOUNDSENTRY = 65
Private Const SPI_GETHIGHCONTRAST = 66
Private Const SPI_SETHIGHCONTRAST = 67
Private Const SPI_GETKEYBOARDPREF = 68
Private Const SPI_SETKEYBOARDPREF = 69
Private Const SPI_GETSCREENREADER = 70
Private Const SPI_SETSCREENREADER = 71
Private Const SPI_GETANIMATION = 72
Private Const SPI_SETANIMATION = 73
Private Const SPI_GETFONTSMOOTHING = 74
Private Const SPI_SETFONTSMOOTHING = 75
Private Const SPI_SETDRAGWIDTH = 76
Private Const SPI_SETDRAGHEIGHT = 77
Private Const SPI_SETHANDHELD = 78
Private Const SPI_GETLOWPOWERTIMEOUT = 79
Private Const SPI_GETPOWEROFFTIMEOUT = 80
Private Const SPI_SETLOWPOWERTIMEOUT = 81
Private Const SPI_SETPOWEROFFTIMEOUT = 82
Private Const SPI_GETLOWPOWERACTIVE = 83
Private Const SPI_GETPOWEROFFACTIVE = 84
Private Const SPI_SETLOWPOWERACTIVE = 85
Private Const SPI_SETPOWEROFFACTIVE = 86
Private Const SPI_SETCURSORS = 87
Private Const SPI_SETICONS = 88
Private Const SPI_GETDEFAULTINPUTLANG = 89
Private Const SPI_SETDEFAULTINPUTLANG = 90
Private Const SPI_SETLANGTOGGLE = 91
Private Const SPI_GETWINDOWSEXTENSION = 92
Private Const SPI_SETMOUSETRAILS = 93
Private Const SPI_GETMOUSETRAILS = 94
Private Const SPI_SCREENSAVERRUNNING = 97

Private Const SPIF_SENDWININICHANGE = &H2
Private Const SPIF_UPDATEINIFILE = &H1


Public Type POINTAPI
    X As Long
    Y As Long
End Type

Dim inited As Boolean
Dim hInstance As Long

Public Sub Initialize()
    If inited Then Exit Sub
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    inited = False
End Sub

Public Function GetDesktopWindowHandle() As Long
    GetDesktopWindowHandle = API_uiMethods_GetDesktopWindow
End Function

Public Function GetKeyState(KeyCode As Long) As Boolean
    GetKeyState = (API_uiMethods_GetKeyState(KeyCode) <> 0)
End Function

Public Sub SetCursorPosition(Left As Long, Top As Long)
    Call API_uiMethods_SetCursorPos(Left, Top)
End Sub
Public Function GetCursorPosition() As POINTAPI
    If API_uiMethods_GetCursorPos(GetCursorPosition) = 0 Then _
        throw Exps.SystemCallFailureException("An error occured when calling GetCursorPos() in user32.dll")
End Function

Public Sub SwapMouseButtons(bSwap As Long)
    Call API_uiMethods_SwapMouseButton(bSwap)
End Sub
Public Function GetDoubleClickTime() As Long
    GetDoubleClickTime = API_uiMethods_GetDoubleClickTime
End Function
Public Sub SetDoubleClickTime(DoubleClickTime As Long)
    Call API_uiMethods_SetDoubleClickTime(DoubleClickTime)
End Sub

Public Sub SetDesktopBackgroundPicture(Path As String)
    Call API_uiMethods_SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, Path, SPIF_UPDATEINIFILE)
End Sub
