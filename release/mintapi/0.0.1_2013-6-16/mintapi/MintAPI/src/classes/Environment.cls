VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "Environment"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Private Declare Function API_waveOutGetNumDevs Lib "winmm" Alias "waveOutGetNumDevs" () As Long
Private Declare Function API_DwmIsCompositionEnabled Lib "DwmApi.dll" () As Boolean
Private Declare Function API_DwmEnableComposition Lib "DwmApi.dll" (compositionAction As CompositionEnable) As Long
Private Enum CompositionEnable
    Disable = 0
    Enable = 1
End Enum

Private Const API_SM_ARRANGE = 56
Private Const API_SM_CLEANBOOT = 67
Private Const API_SM_CMETRICS = 44
Private Const API_SM_CMOUSEBUTTONS = 43
Private Const API_SM_CXBORDER = 5
Private Const API_SM_CXCURSOR = 13
Private Const API_SM_CXDLGFRAME = 7
Private Const API_SM_CXDOUBLECLK = 36
Private Const API_SM_CXDRAG = 68
Private Const API_SM_CXEDGE = 45
Private Const API_SM_CXFIXEDFRAME = API_SM_CXDLGFRAME
Private Const API_SM_CXFRAME = 32
Private Const API_SM_CXFULLSCREEN = 16
Private Const API_SM_CXHSCROLL = 21
Private Const API_SM_CXHTHUMB = 10
Private Const API_SM_CXICON = 11
Private Const API_SM_CXICONSPACING = 38
Private Const API_SM_CXMAXIMIZED = 61
Private Const API_SM_CXMAXTRACK = 59
Private Const API_SM_CXMENUCHECK = 71
Private Const API_SM_CXMENUSIZE = 54
Private Const API_SM_CXMIN = 28
Private Const API_SM_CXMINIMIZED = 57
Private Const API_SM_CXMINSPACING = 47
Private Const API_SM_CXMINTRACK = 34
Private Const API_SM_CXSCREEN = 0
Private Const API_SM_CXSIZE = 30
Private Const API_SM_CXSIZEFRAME = API_SM_CXFRAME
Private Const API_SM_CXSMICON = 49
Private Const API_SM_CXSMSIZE = 52
Private Const API_SM_CXVSCROLL = 2
Private Const API_SM_CYBORDER = 6
Private Const API_SM_CYCAPTION = 4
Private Const API_SM_CYCURSOR = 14
Private Const API_SM_CYDLGFRAME = 8
Private Const API_SM_CYDOUBLECLK = 37
Private Const API_SM_CYDRAG = 69
Private Const API_SM_CYEDGE = 46
Private Const API_SM_CYFIXEDFRAME = API_SM_CYDLGFRAME
Private Const API_SM_CYFRAME = 33
Private Const API_SM_CYFULLSCREEN = 17
Private Const API_SM_CYHSCROLL = 3
Private Const API_SM_CYICON = 12
Private Const API_SM_CYICONSPACING = 39
Private Const API_SM_CYKANJIWINDOW = 18
Private Const API_SM_CYMAXIMIZED = 62
Private Const API_SM_CYMAXTRACK = 60
Private Const API_SM_CYMENU = 15
Private Const API_SM_CYMENUCHECK = 72
Private Const API_SM_CYMENUSIZE = 55
Private Const API_SM_CYMIN = 29
Private Const API_SM_CYMINIMIZED = 58
Private Const API_SM_CYMINSPACING = 48
Private Const API_SM_CYMINTRACK = 35
Private Const API_SM_CYSCREEN = 1
Private Const API_SM_CYSIZE = 31
Private Const API_SM_CYSIZEFRAME = API_SM_CYFRAME
Private Const API_SM_CYSMCAPTION = 51
Private Const API_SM_CYSMICON = 50
Private Const API_SM_CYSMSIZE = 53
Private Const API_SM_CYVSCROLL = 20
Private Const API_SM_CYVTHUMB = 9
Private Const API_SM_DBCSENABLED = 42
Private Const API_SM_DEBUG = 22
Private Const API_SM_MENUDROPALIGNMENT = 40
Private Const API_SM_MIDEASTENABLED = 74
Private Const API_SM_MOUSEPRESENT = 19
Private Const API_SM_NETWORK = 63
Private Const API_SM_PENWINDOWS = 41
Private Const API_SM_RESERVED1 = 24
Private Const API_SM_RESERVED2 = 25
Private Const API_SM_RESERVED3 = 26
Private Const API_SM_RESERVED4 = 27
Private Const API_SM_SECURE = 44
Private Const API_SM_SLOWMACHINE = 73
Private Const API_SM_SWAPBUTTON = 23

Private Const DRIVE_DOESNT_EXIST = 1
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

Public Enum DriveType
    dtDoesnt_Exist = DRIVE_DOESNT_EXIST
    dtFixed = DRIVE_FIXED
    dtCDRom = DRIVE_CDROM
    dtRamDisk = DRIVE_RAMDISK
    dtRemovable = DRIVE_REMOVABLE
    dtRemote = DRIVE_REMOTE
    dtFree = &H7F
End Enum
Public Enum LocaleConstants
    Locale_Language = &H2
    Locale_AbbrevLanguageName = &H3
    Locale_NativeLanguageName = &H4
    Locale_Country = &H6
    Locale_AbbrevCountryName = &H7
    Locale_NativeCountryName = &H8
    Locale_Decimal = &HE
    Locale_Thousand = &HF
    Locale_Currency = &H14
    Locale_IntLSymbol = &H15
    Locale_ShortDate = &H1F
    Locale_LongDate = &H20
    Locale_DayName1 = &H2A
    Locale_DayName2 = &H2B
    Locale_DayName3 = &H2C
    Locale_DayName4 = &H2D
    Locale_DayName5 = &H2E
    Locale_DayName6 = &H2F
    Locale_DayName7 = &H30
    Locale_AbbrevDayName1 = &H31
    Locale_AbbrevDayName3 = &H33
    Locale_AbbrevDayName2 = &H32
    Locale_AbbrevDayName4 = &H34
    Locale_AbbrevDayName5 = &H35
    Locale_AbbrevDayName6 = &H36
    Locale_AbbrevDayName7 = &H37
    Locale_MonthName1 = &H38
    Locale_MonthName2 = &H39
    Locale_MonthName3 = &H3A
    Locale_MonthName4 = &H3B
    Locale_MonthName5 = &H3C
    Locale_MonthName6 = &H3D
    Locale_MonthName7 = &H3E
    Locale_MonthName8 = &H3F
    Locale_MonthName9 = &H40
    Locale_MonthName10 = &H41
    Locale_MonthName11 = &H42
    Locale_MonthName12 = &H43
    Locale_AbbrevMonthName1 = &H44
    Locale_AbbrevMonthName2 = &H45
    Locale_AbbrevMonthName3 = &H46
    Locale_AbbrevMonthName4 = &H47
    Locale_AbbrevMonthName5 = &H48
    Locale_AbbrevMonthName6 = &H49
    Locale_AbbrevMonthName7 = &H4A
    Locale_AbbrevMonthName8 = &H4B
    Locale_AbbrevMonthName9 = &H4C
    Locale_AbbrevMonthName10 = &H4D
    Locale_AbbrevMonthName11 = &H4E
    Locale_AbbrevMonthName12 = &H4F
    Locale_PositiveSign = &H50
    Locale_NegativeSign = &H51
    Locale_AbbrevMonthName13 = &H100F
    Locale_EngLanguage = &H1001
    Locale_EngCountry = &H1002
    Locale_TimeFormat = &H1003
End Enum
'Public Enum OSPlatformID
'    OSP32s
'    OSP95_98
'    OSPNT
'    OS_WINDOWS_3X
'    OS_WINDOWS_95
'    OS_WINDOWS_98
'    OS_WINDOWS_NT3X
'    OS_WINDOWS_NT40
'    OS_WINDOWS_2000
'End Enum
Public Enum SystemMetricsValues
    smNULL = 0
    smARRANGE = API_SM_ARRANGE
    smCLEANBOOT = API_SM_CLEANBOOT
    smCMETRICS = API_SM_CMETRICS
    smCMOUSEBUTTONS = API_SM_CMOUSEBUTTONS
    smCXBORDER = API_SM_CXBORDER
    smCXCURSOR = API_SM_CXCURSOR
    smCXDLGFRAME = API_SM_CXDLGFRAME
    smCXDOUBLECLK = API_SM_CXDOUBLECLK
    smCXDRAG = API_SM_CXDRAG
    smCXEDGE = API_SM_CXEDGE
    smCXFIXEDFRAME = API_SM_CXDLGFRAME
    smCXFRAME = API_SM_CXFRAME
    smCXFULLSCREEN = API_SM_CXFULLSCREEN
    smCXHSCROLL = API_SM_CXHSCROLL
    smCXHTHUMB = API_SM_CXHTHUMB
    smCXICON = API_SM_CXICON
    smCXICONSPACING = API_SM_CXICONSPACING
    smCXMAXIMIZED = API_SM_CXMAXIMIZED
    smCXMAXTRACK = API_SM_CXMAXTRACK
    smCXMENUCHECK = API_SM_CXMENUCHECK
    smCXMENUSIZE = API_SM_CXMENUSIZE
    smCXMIN = API_SM_CXMIN
    smCXMINIMIZED = API_SM_CXMINIMIZED
    smCXMINSPACING = API_SM_CXMINSPACING
    smCXMINTRACK = API_SM_CXMINTRACK
    smCXSCREEN = API_SM_CXSCREEN
    smCXSIZE = API_SM_CXSIZE
    smCXSIZEFRAME = API_SM_CXFRAME
    smCXSMICON = API_SM_CXSMICON
    smCXSMSIZE = API_SM_CXSMSIZE
    smCXVSCROLL = API_SM_CXVSCROLL
    smCYBORDER = API_SM_CYBORDER
    smCYCAPTION = API_SM_CYCAPTION
    smCYCURSOR = API_SM_CYCURSOR
    smCYDLGFRAME = API_SM_CYDLGFRAME
    smCYDOUBLECLK = API_SM_CYDOUBLECLK
    smCYDRAG = API_SM_CYDRAG
    smCYEDGE = API_SM_CYEDGE
    smCYFIXEDFRAME = API_SM_CYDLGFRAME
    smCYFRAME = API_SM_CYFRAME
    smCYFULLSCREEN = API_SM_CYFULLSCREEN
    smCYHSCROLL = API_SM_CYHSCROLL
    smCYICON = API_SM_CYICON
    smCYICONSPACING = API_SM_CYICONSPACING
    smCYKANJIWINDOW = API_SM_CYKANJIWINDOW
    smCYMAXIMIZED = API_SM_CYMAXIMIZED
    smCYMAXTRACK = API_SM_CYMAXTRACK
    smCYMENU = API_SM_CYMENU
    smCYMENUCHECK = API_SM_CYMENUCHECK
    smCYMENUSIZE = API_SM_CYMENUSIZE
    smCYMIN = API_SM_CYMIN
    smCYMINIMIZED = API_SM_CYMINIMIZED
    smCYMINSPACING = API_SM_CYMINSPACING
    smCYMINTRACK = API_SM_CYMINTRACK
    smCYSCREEN = API_SM_CYSCREEN
    smCYSIZE = API_SM_CYSIZE
    smCYSIZEFRAME = API_SM_CYFRAME
    smCYSMCAPTION = API_SM_CYSMCAPTION
    smCYSMICON = API_SM_CYSMICON
    smCYSMSIZE = API_SM_CYSMSIZE
    smCYVSCROLL = API_SM_CYVSCROLL
    smCYVTHUMB = API_SM_CYVTHUMB
    smDBCSENABLED = API_SM_DBCSENABLED
    smDEBUG = API_SM_DEBUG
    smMENUDROPALIGNMENT = API_SM_MENUDROPALIGNMENT
    smMIDEASTENABLED = API_SM_MIDEASTENABLED
    smMOUSEPRESENT = API_SM_MOUSEPRESENT
    smNETWORK = API_SM_NETWORK
    smPENWINDOWS = API_SM_PENWINDOWS
    smRESERVED1 = API_SM_RESERVED1
    smRESERVED2 = API_SM_RESERVED2
    smRESERVED3 = API_SM_RESERVED3
    smRESERVED4 = API_SM_RESERVED4
    smSECURE = API_SM_SECURE
    smSLOWMACHINE = API_SM_SLOWMACHINE
    smSWAPBUTTON = API_SM_SWAPBUTTON
    smDefault = API_SM_CMETRICS
End Enum
Public Type WindowClass
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
Public Enum ComputerNames
    cnNetBIOS = cn_NetBIOS
    cnDnsHostname = cn_DnsHostname
    cnDnsDomain = cn_DnsDomain
    cnDnsFullyQualified = cn_DnsFullyQualified
    cnPhysicalNetBIOS = cn_PhysicalNetBIOS
    cnPhysicalDnsHostName = cn_PhysicalDnsHostName
    cnPhysicalDnsDomain = cn_PhysicalDnsDomain
    cnPhysicalDnsFullyQualified = cn_PhysicalDnsFullyQualified
    cnMax = cn_Max
End Enum


Friend Sub Initialize(Optional Arguments)
    If Not IsMissing(Arguments) Then
        
    End If
End Sub


Public Function GetEnvironmentVariable(Name As String, Optional DefaultValue) As ByteArray
    
End Function
Public Sub SetEnvironmentVariable(Name As String, Value As ByteArray)
    
End Sub
Public Function CheckEnvironmentVariable(Name As String) As Boolean
    
End Function

Public Function GetMintAPIVariable(Name As String, Optional DefaultValue) As ByteArray
    On Error GoTo Err_Handler
    Dim varBuffer() As Byte
    varBuffer = mint_config.ReadMintAPIVariable(Name, True)
    Set GetMintAPIVariable = ByteArray(varBuffer)
    Exit Function
Err_Handler:
    Set GetMintAPIVariable = ByteArray(DefaultValue)
End Function
Public Sub SetMintAPIVariable(Name As String, Value As ByteArray)
    'Call modMain
End Sub
Public Function CheckMintAPIVariable(Name As String) As Boolean
    
End Function

Public Function NewLine() As String
    NewLine = vbCrLf
End Function


Public Function haveSoundCard() As Boolean
    haveSoundCard = (API_waveOutGetNumDevs > 0)
End Function
Public Function supportUSB() As Boolean
    
End Function



Public Function IsCompositionEnabled() As Boolean
    IsCompositionEnabled = API_DwmIsCompositionEnabled
End Function
Public Sub EnableGlassWindow()
    Call API_DwmEnableComposition(Enable)
End Sub
Public Sub DisableGlassWindow()
    Call API_DwmEnableComposition(Disable)
End Sub
Public Sub InitializeGlassWindow()
    
End Sub
Public Sub DisposeGlassWindow()
    
End Sub
Public Sub ExludeRegionFromGlass()
    
End Sub

Public Function ChangeResolution(Width As Long, Height As Long, Optional Bits As Long = 32, Optional ByVal SendMessageToWindows As Boolean = True) As Boolean
    Call gdiMethods.ChangeResolution(Width, Height, Bits, SendMessageToWindows)
End Function

Public Function GetKeyState(KeyCode As Long) As Boolean
    GetKeyState = uiMethods.GetKeyState(KeyCode)
End Function
Public Function CapsLockState() As Boolean
    CapsLockState = uiMethods.GetKeyState(vbKeyCapital)
End Function
Public Function ScrollLockState() As Boolean
    ScrollLockState = uiMethods.GetKeyState(vbKeyScrollLock)
End Function
Public Function ControlState() As Boolean
    ControlState = uiMethods.GetKeyState(vbKeyControl)
End Function
Public Function NumLockState() As Boolean
    NumLockState = uiMethods.GetKeyState(vbKeyNumlock)
End Function
Public Function ShiftState() As Boolean
    ShiftState = uiMethods.GetKeyState(vbKeyShift)
End Function
Public Function AltState() As Boolean
    AltState = uiMethods.GetKeyState(18)
End Function
Public Function PrintScreenState() As Boolean
    Const API_VK_SNAPSHOT = &H2C
    PrintScreenState = uiMethods.GetKeyState(API_VK_SNAPSHOT)
End Function
Public Function LeftMouseButtonState() As Boolean
    LeftMouseButtonState = uiMethods.GetKeyState(vbKeyLButton)
End Function
Public Function RightMouseButtonState() As Boolean
    RightMouseButtonState = uiMethods.GetKeyState(vbKeyRButton)
End Function
Public Function MiddleMouseButtonState() As Boolean
    MiddleMouseButtonState = uiMethods.GetKeyState(vbKeyMButton)
End Function

Public Sub SwapMouseButtons(Optional bSwap As Long = 1)
    Call uiMethods.SwapMouseButtons(bSwap)
End Sub
Public Sub SetCursorPosition(Pos As Position)
    Call uiMethods.SetCursorPosition(Pos.Left, Pos.Top)
End Sub
Public Function WheelScrollValue() As Long
    
End Function
Public Sub SetWheelScrollValue()
    
End Sub
Public Function DoubleClickTime() As Long
    DoubleClickTime = uiMethods.GetDoubleClickTime
End Function
Public Sub SetDoubleClickTime(DoubleClickTime As Long)
    Call uiMethods.SetDoubleClickTime(DoubleClickTime)
End Sub
Public Function MouseSpeed() As Long
    
End Function
Public Sub SetMouseSpeed()
    
End Sub

Public Sub RunMintAPIHost(CommandArguments As String, Optional ByVal Stay As Boolean = False)
    
End Sub

Public Function CaptureScreen(Optional Window = 0) As Pixmap
    
End Function

Public Function ShowInDefaultBrowser() As Long

End Function

Public Function DotNetFrameworkInstallationState() As Boolean
    DotNetFrameworkInstallationState = _
        Registry("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP").Exists
End Function
Public Function RoNOTEInstallationState() As Boolean
    RoNOTEInstallationState = _
        Directory(Registry("HKEY_LOCAL_MACHINE\SOFTWARE\RoNOTE").GetValue("InstallationPath").toString()).Exists
End Function
Public Function QtInstallationState() As Boolean
    QtInstallationState = _
        Registry("").Exists
End Function
Public Function MSOfficeInstallationState() As Boolean
    MSOfficeInstallationState = _
        Registry("").Exists
End Function
Public Function MSSQLExpressInstallationState() As Boolean
    MSSQLExpressInstallationState = _
        Registry("").Exists
End Function
Public Function JavaRuntimeInstallationState() As Boolean
    JavaRuntimeInstallationState = _
        Directory(GetEnvironmentVariable("JAVA_HOME", "N/A").toString()).Exists
End Function
Public Function FlashPlayerInstallationState() As Boolean
    FlashPlayerInstallationState = _
        Registry("").Exists
End Function
Public Function CheckInstallationState(ApplicationName As String, HowDidIKnow As String) As Boolean
    CheckInstallationState = False
End Function

Public Function VB6InstallationState() As Boolean
    VB6InstallationState = _
        Registry("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Visual Basic").Exists
If Not VB6InstallationState Then
    VB6InstallationState = _
        Registry("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Visual Basic").Exists
End If
End Function
Public Function VB6InstallationPath() As String
    Dim Reg As Registry
    Set Reg = Registry("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Visual Basic")
    If Not Reg.Exists Then _
        Exit Function
    
    VB6InstallationPath = _
        Reg.GetValue("TemplatesDirectory").toString
End Function



Public Function GetLocaleString(LocaleID As LocaleConstants) As String
    GetLocaleString = baseMethods.GetLocaleString(LocaleID)
End Function
Public Function SetLocaleString(LocaleID As LocaleConstants, Value As String) As String
    Call baseMethods.SetLocaleString(LocaleID, Value)
End Function

Public Property Get SystemLocaleDateFormat() As String
    SystemLocaleDateFormat = baseMethods.LocaleDateFormat
End Property
Public Property Let SystemLocaleDateFormat(Value As String)
    Call baseMethods.SetLocaleDateFormat(Value)
End Property
Public Property Get SystemLocaleTimeFormat() As String
    SystemLocaleTimeFormat = baseMethods.LocaleTimeFormat
End Property
Public Property Let SystemLocaleTimeFormat(Value As String)
    Call baseMethods.SetLocaleTimeFormat(Value)
End Property
Public Property Get SystemLocaleNumberFormat() As String
    SystemLocaleNumberFormat = baseMethods.LocaleNumberFormat
End Property
Public Property Let SystemLocaleNumberFormat(Value As String)
    Call baseMethods.SetLocaleNumberFormat(Value)
End Property
Public Property Get SystemLocaleThousandSpecifier() As String
    SystemLocaleDateFormat = baseMethods.LocaleThousandSpecifier
End Property
Public Property Let SystemLocaleThousandSpecifier(Value As String)
    Call baseMethods.SetLocaleThousandSpecifier(Value)
End Property
Public Property Get SystemLocaleDecimalSpecifier() As String
    SystemLocaleDecimalSpecifier = baseMethods.LocaleDecimalSpecifier
End Property
Public Property Let SystemLocaleDecimalSpecifier(Value As String)
    Call baseMethods.SetLocaleDecimalSpecifier(Value)
End Property
Public Property Get SystemLocaleCurrencySpecifier() As String
    SystemLocaleCurrencySpecifier = baseMethods.LocaleCurrencySpecifier
End Property
Public Property Let SystemLocaleCurrencySpecifier(Value As String)
    Call baseMethods.SetLocaleCurrencySpecifier(Value)
End Property
Public Property Get SystemLocaleShortMonthName(iMonthNum As Integer) As String
    SystemLocaleShortMonthName = baseMethods.LocaleShortMonthName(iMonthNum)
End Property
Public Property Let SystemLocaleShortMonthName(iMonthNum As Integer, Value As String)
    Call baseMethods.SetLocaleShortMonthName(iMonthNum, Value)
End Property
Public Property Get SystemLocaleMonthName(iMonthNum As Integer) As String
    SystemLocaleMonthName = baseMethods.LocaleMonthName(iMonthNum)
End Property
Public Property Let SystemLocaleMonthName(iMonthNum As Integer, Value As String)
    Call baseMethods.SetLocaleMonthName(iMonthNum, Value)
End Property
Public Property Get SystemLocaleShortDayName(iDayNum As Integer) As String
    SystemLocaleShortDayName = baseMethods.LocaleShortDayName(iDayNum)
End Property
Public Property Let SystemLocaleShortDayName(iDayNum As Integer, Value As String)
    Call baseMethods.SetLocaleShortDayName(iDayNum, Value)
End Property
Public Property Get SystemLocaleDayName(iDayNum As Integer) As String
    SystemLocaleDayName = baseMethods.LocaleDayName(iDayNum)
End Property
Public Property Let SystemLocaleDayName(iDayNum As Integer, Value As String)
    Call baseMethods.SetLocaleDayName(iDayNum, Value)
End Property
Public Property Get SystemLocaleCountry() As String
    SystemLocaleCountry = baseMethods.LocaleCountry
End Property
Public Property Let SystemLocaleCountry(Value As String)
    Call baseMethods.SetLocaleCountry(Value)
End Property '
Public Property Get SystemLocaleLanguageName() As String
    SystemLocaleLanguageName = baseMethods.LocaleLanguageName
End Property
Public Property Let SystemLocaleLanguageName(Value As String)
    Call baseMethods.SetLocaleLanguageName(Value)
End Property
Public Property Get SystemLocaleNativeCountryName() As String
    SystemLocaleNativeCountryName = baseMethods.LocaleNativeCountryName
End Property
Public Property Let SystemLocaleNativeCountryName(Value As String)
    Call baseMethods.SetLocaleNativeCountryName(Value)
End Property
Public Property Get SystemLocalePositiveSign() As String
    SystemLocalePositiveSign = baseMethods.LocalePositiveSign
End Property
Public Property Let SystemLocalePositiveSign(Value As String)
    Call baseMethods.SetLocalePositiveSign(Value)
End Property
Public Property Get SystemLocaleNegativeSign() As String
    SystemLocaleNegativeSign = baseMethods.LocaleNegativeSign
End Property
Public Property Let SystemLocaleNegativeSign(Value As String)
    Call baseMethods.SetLocaleNegativeSign(Value)
End Property
'-------Let nadaran!
Public Property Get SystemLocaleLanguageID() As String
    SystemLocaleLanguageID = baseMethods.LocaleSystemLanguageID
End Property
Public Property Get SystemLocaleSystemLanguageName() As String
    SystemLocaleSystemLanguageName = baseMethods.LocaleSystemLanguageName
End Property


Public Sub Shutdown(Optional ByVal Force As Boolean = False)
    Call kernelMethods.EnableShutdown
    Call kernelMethods.Shutdown(Force)
End Sub
Public Sub CancelShutdown()
    
End Sub
Public Sub Hibernate(Optional ByVal Force As Boolean = False)
    Call kernelMethods.EnableHibernate
    Call kernelMethods.Hibernate(Force)
End Sub
Public Sub CancelHibernate()
    
End Sub
Public Sub Sleep(Optional ByVal Force As Boolean = False)
    Call kernelMethods.Sleep(Force)
End Sub
Public Sub Restart(Optional ByVal Force As Boolean = False)
    Call kernelMethods.RestartSystem(Force)
End Sub
Public Sub CancelRestart()
    
End Sub
Public Sub Logoff(Optional ByVal Force As Boolean = False)
    Call kernelMethods.Logoff(Force)
End Sub
Public Sub CancelLogoff()
    
End Sub
Public Sub SwitchUser()
    Call kernelMethods.SwitchUser
End Sub

Public Property Get MousePosition() As Position
    Dim POINTAPI_ As POINTAPI
    POINTAPI_ = uiMethods.GetCursorPosition
    MousePosition.Left = POINTAPI_.X
    MousePosition.Top = POINTAPI_.Y
End Property
Public Property Let MousePosition(Value As Position)
    Call uiMethods.SetCursorPosition(Value.Left, Value.Top)
End Property

Public Sub SetDesktopBackgroundPicture(Path As String)
    Call uiMethods.SetDesktopBackgroundPicture(Path)
End Sub

Public Function CreateTrayIcon(hWnd As Long, IconHandle As Long, ToolTip As String, uID As Long, CallbackMessage As Long) As Long
    CreateTrayIcon = shellMethods.CreateTrayIcon(hWnd, IconHandle, ToolTip, uID, CallbackMessage)
End Function
Public Sub DestroyTrayIcon(tray_hWnd As Long)
    Call shellMethods.DestroyTrayIcon(tray_hWnd)
End Sub

Public Sub SendWindowMessage(hWnd As Long, uMsg As Long, uParam As Long, lParam As Long)
    
End Sub

Public Function SpecialFolder(SpecialFolders As SpecialFolders) As String
    SpecialFolder = shellMethods.GetSpecialfolder(SpecialFolders)
End Function


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
