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
Option Base 0
Option Explicit
Const CLASSID As String = "Types_Enums"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Public Enum Timeouts
    tINFINITE = INFINITE
End Enum

Public Enum TexingFormat
    tfStream = tf_Stream
    
    tfReadable = tf_Readable
    
    tfMultiLine = tf_MultiLine
    
    tfNoIndented = tf_NoIndented
    tf1Indented = tf_1Indented
    tf4Indented = tf_4Indented 'this activates under tfMultiline state.
    
    tfAuto_Editable = tfReadable Or tf1Indented Or tfMultiLine
    tfAuto_Binary = tfStream
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

Public Enum TristateValues
    tvLess = -1
    tvEqual = 0
    tvMore = 1
End Enum
Public Enum CompareValue
    cvLess = -1
    cvEqual = 0
    cvMore = 1
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

Public Type Point
    X As Long
    Y As Long
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
Public Type DPosition
    Left As Double
    Top As Double
End Type
Public Type Size
    Width As Long
    Height As Long
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
Public Type DPoint3D
    X As Double
    Y As Double
    Z As Double
End Type
Public Type Color
    Color As Colors 'long
    Alpha As Byte
End Type
Public Type Margin
    Top As Long
    Right As Long
    Bottom As Long
    Left As Long
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
Public Type DPadding
    Top As Double
    Right As Double
    Bottom As Double
    Left As Double
End Type

Public Type GlobalFilters
    IncludeTemplates() As String
    ExcludeTemplates() As String
End Type
Public Type Argument
    Name As String
    Value As Variant
End Type

Public Type IPAddress
    IPVersion As Long
    IPAddress As Long
End Type
Public Type SocketAddress
    IPVersion As Long
    IPAddress As Long
    SocketPort As Long
End Type
Public Type EndPointTag
    Refrence As Long
End Type
Public Type EndPoint
    AddressFamily As Long
    Tag As Long
    SocketPort As Long
    IPAddress As String
End Type
