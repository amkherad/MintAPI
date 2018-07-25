VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Types_Enums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

'This class is used to prevent the circular refrences between classes and modules.

Option Base 0
Option Explicit
Const CLASSID As String = "Types_Enums"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

'Public Type WaitHandle
'    WaitHandle As Long
'End Type
Public Enum MintConstants
    mcInfinite = C_INFINITE
    mcPI = C_PI
End Enum

Private Const CPUFamily_Intel           As Long = &H1000
Private Const CPUFamily_AMD             As Long = &H2000
Private Const CPUFamily_Intel_80        As Long = CPUFamily_Intel Or &H0
Private Const CPUFamily_Intel_Core      As Long = CPUFamily_Intel Or &H1
Public Enum CPUFamily
    CPUi80 = CPUFamily_Intel_80 + 1
    CPUi86 = CPUi80 + 1
    CPUi286 = CPUi86 + 1
    CPUi386 = CPUi286 + 1
    CPUi486 = CPUi386 + 1
    CPUi586 = CPUi486 + 1
    CPUaXXX
End Enum

Public Enum VariantTypes
    VT_EMPTY = 0
    VT_Null = 1
    VT_I2 = 2
    VT_Integer = VT_I2
    VT_Int16 = VT_I2
    VT_I4 = 3
    VT_Int32 = VT_I4
    VT_Long = VT_I4
    VT_R4 = 4
    VT_Single = VT_R4
    VT_Float = VT_R4
    VT_R8 = 5
    VT_Double = VT_R8
    VT_CY = 6
    VT_Currency = VT_CY
    VT_DATE = 7
    VT_BSTR = 8
    VT_DISPATCH = 9
    VT_Object = VT_DISPATCH
    VT_Error = 10
    VT_BOOL = 11
    VT_Boolean = VT_BOOL
    VT_VARIANT = 12
    VT_Unknown = 13
    VT_DECIMAL = 14
    VT_I1 = 16
    VT_Char = VT_I1
    VT_UI1 = 17
    VT_Byte = VT_UI1
    VT_UI2 = 18
    VT_UnsignedInteger = VT_UI2
    VT_UI4 = 19
    VT_UnsignedLong = VT_UI4
    VT_I8 = 20
    VT_Int64 = VT_I8
    VT_LongLong = VT_I8
    VT_UI8 = 21
    VT_UnsignedLongLong = VT_UI8
    VT_INT = 22
    VT_UINT = 23
    VT_VOID = 24
    VT_HRESULT = 25
    VT_Ptr = 26
    VT_Handle = VT_Ptr
    VT_SAFEARRAY = 27
    VT_CArray = 28
    VT_UserDefined = 29
    VT_UDT = VT_UserDefined
    VT_LPSTR = 30
    VT_LPWSTR = 31
    VT_Record = 36
    VT_Structure = VT_Record
    VT_INT_PTR = 37
    VT_UINT_PTR = 38
    VT_FileTime = 64
    VT_Blob = 65
    VT_Stream = 66
    VT_Storage = 67
    VT_Streamed_Object = 68
    VT_Stored_Object = 69
    VT_Blob_Object = 70
    VT_CF = 71
    VT_ClsId = 72
    VT_Versioned_Stream = 73
    VT_BSTR_BLOB = &HFFF
    VT_Vector = &H1000
    VT_ARRAY = &H2000
    VT_BYREF = &H4000
    VT_ByVal = 0
    VT_In = 0
    VT_Out = VT_BYREF
    VT_Ref = VT_BYREF
    VT_Reserved = &H8000
    VT_Illegal = &HFFFF
    VT_IllegalMasked = &HFFF
    VT_TypeMask = &HFFF
End Enum
Public Enum DialogResults
    drOK = VbMsgBoxResult.vbOK
    drCancel = VbMsgBoxResult.vbCancel
    drYes = VbMsgBoxResult.vbYes
    drNo = VbMsgBoxResult.vbNo
    drRetry = VbMsgBoxResult.vbRetry
    drIgnore = VbMsgBoxResult.vbIgnore
    drAbort = VbMsgBoxResult.vbAbort
End Enum
Public Enum MessageBoxButtons
    mbbOKOnly = VbMsgBoxStyle.vbOKOnly
    mbbOKCancel = VbMsgBoxStyle.vbOKCancel
    mbbYesNo = VbMsgBoxStyle.vbYesNo
    mbbYesNoCancel = VbMsgBoxStyle.vbYesNoCancel
    mbbRetryCancel = VbMsgBoxStyle.vbRetryCancel
    mbbAbortRetryIgnore = VbMsgBoxStyle.vbAbortRetryIgnore
End Enum
Public Enum MessageBoxIcons
    siNoIcon = 0
    siInformation = VbMsgBoxStyle.vbInformation
    siAsterisk = siInformation
    siWarning = VbMsgBoxStyle.vbExclamation
    siExclamation = siWarning
    siError = VbMsgBoxStyle.vbCritical
    siFatal = siError
    siQuestion = VbMsgBoxStyle.vbQuestion
    'additional
    'siShield =
    'siApplication = 6
End Enum
Public Enum MessageBoxStyles
    mbsNoSpecifiedStyle = 0
    mbsRtlReading = VbMsgBoxStyle.vbMsgBoxRtlReading
    mbsRight = VbMsgBoxStyle.vbMsgBoxRight
    mbsHelpButton = VbMsgBoxStyle.vbMsgBoxHelpButton
    mbsSetForeground = VbMsgBoxStyle.vbMsgBoxSetForeground
    mbsApplicationModal = VbMsgBoxStyle.vbApplicationModal
    mbsSystemModal = VbMsgBoxStyle.vbSystemModal
End Enum
Public Enum MessageBoxDefaultButton
    mbdButton1 = VbMsgBoxStyle.vbDefaultButton1
    mbdButton2 = VbMsgBoxStyle.vbDefaultButton2
    mbdButton3 = VbMsgBoxStyle.vbDefaultButton3
    mbdButton4 = VbMsgBoxStyle.vbDefaultButton4
End Enum

Public Enum SystemIcons
    siNoIcon = 0
    siInformation = 1
    siAsterisk = siInformation
    siWarning = 2
    siExclamation = siWarning
    siError = 3
    siFatal = siError
    'additional
    siQuestion = 4
    siShield = 5
    siApplication = 6
End Enum

Public Enum HTTPResponseHeaderCodes
    http_Continue = 100
    http_SwitchingProtocols = 101
    http_OK = 200
    http_Created = 201
    http_Accepted = 202
    http_NonAuthoritativeInformation = 203
    http_NoContent = 204
    http_ResetContent = 205
    http_PartialContent = 206
    http_MultipleChoices = 300
    http_MovedPermanently = 301
    http_Found = 302
    http_SeeOther = 303
    http_NotModified = 304
    http_UseProxy = 305
    http_TemporaryRedirect = 307
    http_BadRequest = 400
    http_Unauthorized = 401
    http_PaymentRequired = 402
    http_Forbidden = 403
    http_NotFound = 404
    http_MethodNotAllowed = 405
    http_NotAcceptable = 406
    http_ProxyAuthenticationRequired = 407
    http_RequestTimeout = 408
    http_Conflict = 409
    http_Gone = 410
    http_LengthRequired = 411
    http_PreconditionFailed = 412
    http_RequestEntityTooLarge = 413
    http_RequestURITooLarge = 414
    http_UnsupportedMediaType = 415
    http_RequestedRangeNotSatisfiable = 416
    http_ExpectationFailed = 417
    http_InternalServerError = 500
    http_NotImplemented = 501
    http_BadGateway = 502
    http_ServiceUnavailable = 503
    http_GatewayTimeout = 504
    http_HTTPVersionNotSupported = 505
End Enum

Public Enum TristateValues
    tvLess = -1
    tvEqual = 0
    tvMore = 1
End Enum
Public Enum CompareResults
    crLess = -1
    crEqual = 0
    crGreater = 1
    crNotEqual = &H7FFFFFFF
End Enum
Public Enum NumberSign
    nsNegative = -1
    nsZero = 0
    nsPositive = 1
    nsEqual = nsZero
End Enum
Public Enum SpecialFolders
    sfCustom = -1
    sfDesktop = &H0
    sfALL = sfDesktop
    sfPrograms = &H2
    sfUser_Start_Menu_Programs = sfPrograms
    sfControls = &H3
    sfPrinters = &H4
    sfPersonal = &H5
    sfUser_MyDocuments = sfPersonal
    sfFavorites = &H6
    sfStartup = &H7
    sfUser_Start_Menu_Programs_Startup = sfStartup
    sfRecent = &H8
    sfSendTo = &H9
    sfBitBucket = &HA
    sfStartMenu = &HB
    sfMyMusic = &HD
    sfMyVideos = &HE
    sfMyPictures = &H27
    sfUser_StartMenu = sfStartMenu
    sfUser_Desktop = &H10
    sfDesktopDirectory = sfUser_Desktop
    sfDrives = &H11
    sfMyComputer = sfDrives
    sfNetwork = &H12
    sfAll_Network = sfNetwork
    sfNethood = &H13
    sfFonts = &H14
    sfTemplates = &H15
    sfCommon_StartMenu = &H16
    sfCommon_StartMenu_Programs = &H17
    sfCommon_StartMenu_Programs_Startup = &H18
    sfCommon_Desktop = &H19

    sfApplicationData = &H1A
    sfPrintHood = &H1B
    sfLocalApplicationData = &H1C
    sfCommon_Favorites = &H1F
    sfTemp_InternetFiles = &H20
    sfCookies = &H21
    sfHistory = &H22
    sfCommon_ApplicationData = &H23

    sfWindows = &H24
    sfSystem = &H25
    sfProgram_Files = &H26
    sfUser = &H28
    sfCommon_Templates = &H2D
    sfProgramFiles_CommonFiles = &H2B
    sfCommon_Documents = &H2E
    sfCommon_AdministrativeTools = &H2F
    sfAdministrativeTools = &H30
    sfCommon_MyMusic = &H35
    sfCommon_MyPictures = &H36
    sfCommon_MyVideos = &H37
    sfResources = &H38
    sfCDBurning = &H3B

    sfWorkgroup = &H3D
    sfNetwork_Computers = sfWorkgroup
End Enum
Public Enum MouseButtons
    mbNULL = 0
    mbLeft = 1
    mbRight = 2
    mbMiddle = 3
    mbMiddle1 = mbMiddle
    mbMiddle2 = 4
    mbDefault = mbLeft
End Enum

Public Type IntPtr
    Pointer As Long
End Type
Public Type Int64
    LowerValue As Long
    HigherValue As Long
End Type
Public Type LPDecimal
    Value4 As Long 'Lower  Lower  Value
    Value3 As Long 'Lower  Higher Value
    Value2 As Long 'Higher Lower  Value
    Value1 As Long 'Higher Higher Value
End Type

Public Type Point
    X As Long
    Y As Long
End Type
Public Type FPoint
    X As Single
    Y As Single
End Type
Public Type DPoint
    X As Double
    Y As Double
End Type
Public Type Rectangle
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type
Public Type FRectangle
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type
Public Type DRectangle
    Left As Double
    Top As Double
    Width As Double
    Height As Double
End Type
Public Type Position
    Left As Long
    Top As Long
End Type
Public Type FPosition
    Left As Single
    Top As Single
End Type
Public Type DPosition
    Left As Double
    Top As Double
End Type
Public Type Size
    Width As Long
    Height As Long
End Type
Public Type FSize
    Width As Single
    Height As Single
End Type
Public Type DSize
    Width As Double
    Height As Double
End Type
Public Type Region
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type FRegion
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type
Public Type DRegion
    Left As Double
    Top As Double
    Right As Double
    Bottom As Double
End Type
Public Type Point3D
    X As Long
    Y As Long
    Z As Long
End Type
Public Type FPoint3D
    X As Single
    Y As Single
    Z As Single
End Type
Public Type DPoint3D
    X As Double
    Y As Double
    Z As Double
End Type
Public Type Margin
    Top As Long
    Right As Long
    Bottom As Long
    Left As Long
End Type
Public Type FMargin
    Top As Single
    Right As Single
    Bottom As Single
    Left As Single
End Type
Public Type DMargin
    Top As Double
    Right As Double
    Bottom As Double
    Left As Double
End Type
Public Type Padding
    Top As Long
    Right As Long
    Bottom As Long
    Left As Long
End Type
Public Type FPadding
    Top As Single
    Right As Single
    Bottom As Single
    Left As Single
End Type
Public Type DPadding
    Top As Double
    Right As Double
    Bottom As Double
    Left As Double
End Type

Public Type RangedInt8
    Start As Byte
    End As Byte
End Type
Public Type RangedInt16
    Start As Integer
    End As Integer
End Type
Public Type RangedInt32
    Start As Long
    End As Long
End Type
Public Type RangedInt64
    Start As Int64
    End As Int64
End Type
Public Type RangedFloat
    Start As Single
    End As Single
End Type
Public Type RangedDouble
    Start As Double
    End As Double
End Type
Public Type RangedCurrency
    Start As Currency
    End As Currency
End Type
Public Type RangedDate
    Start As Date
    End As Date
End Type
Public Type RangedVariant
    Start As Variant
    End As Variant
End Type
Public Type RangedLength
    Start As Long
    End As Long
End Type
Public Type RangedDateTime
    Start As DateTime
    End As DateTime
End Type
Public Type RangedTimeSpan
    Start As TimeSpan
    End As TimeSpan
End Type
Public Type RangedLetter
    Start As String
    End As String
End Type


Public Type GlobalFilters
    IncludeTemplates() As String
    ExcludeTemplates() As String
End Type
