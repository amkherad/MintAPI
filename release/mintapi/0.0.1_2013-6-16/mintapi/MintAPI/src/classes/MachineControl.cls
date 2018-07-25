VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MachineControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "MachineControl"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

Implements IAliasable

Public Enum PlatformID
    PIDWin32s = pltID_Win32s
    PIDWin95_98 = pltID_Win95_98
    PIDWinNT = pltID_WinNT
End Enum
Public Enum OSVersion
    OSDOS = OSV_DOS
    OSWin3x = OSV_Win3X
    OSWin95 = OSV_Win95
    OSWin98 = OSV_Win98
    OSWinNT = OSV_WinNT
    OSWinNT3X = OSV_WinNT3X
    OSWinNT40 = OSV_WinNT40
    OSWin2000 = OSV_Win2000
    OSWinXP = OSV_WinXP
    OSWinServer2003 = OSV_WinServer2003
    OSWinVista = OSV_WinVista
    OSWin7 = OSV_Win7
    OSWinServer2008 = OSV_WinServer2008
    OSWin8 = OSV_Win8
    OSWin9 = OSV_Win9
    OSHigher = OSV_Higher
    OSUnknown = OSV_Unknown
End Enum
Public Type OSVersionInfo
    PlatformID As Long
    MajorVersion As Long
    MinorVersion As Long
    BuildNumber As Long
    ServicePack As String
End Type

Private Type API_Memory
    Length As Long
    MemoryLoad As Long
    TotalPhysMemory As Long
    AvailablePhysMemory As Long
    TotalPageFile As Long
    AvailPageFile As Long
    TotalVirtualMemory As Long
    AvailableVirtualMemory As Long
End Type
Private Declare Sub API_GlobalMemoryStatus Lib "kernel32" Alias "GlobalMemoryStatus" (M As API_Memory)

Dim lmemInfo As API_Memory

Friend Sub Initialize()
    
End Sub

Public Function PlatformID() As PlatformID
    PlatformID = baseMethods.GetVersionEx.dwPlatformId
End Function
Public Function OSVersion() As OSVersion
    OSVersion = baseMethods.OSVersion
End Function
Public Function Version() As OSVersionInfo
    Dim vX As API_OSVERSIONINFO
    vX = baseMethods.GetVersionEx
    Version.PlatformID = vX.dwPlatformId
    Version.MajorVersion = vX.dwMajorVersion
    Version.MinorVersion = vX.dwMinorVersion
    Version.BuildNumber = vX.dwBuildNumber
    Version.ServicePack = Trim(vX.szCSDVersion)
End Function
Public Function MajorVersion() As Long
    MajorVersion = baseMethods.GetVersionEx.dwMajorVersion
End Function
Public Function MinorVersion() As Long
    MinorVersion = baseMethods.GetVersionEx.dwMinorVersion
End Function
Public Function BuildNumber() As Long
    BuildNumber = baseMethods.GetVersionEx.dwBuildNumber
End Function
Public Function ServicePack() As String
    ServicePack = Trim(baseMethods.GetVersionEx.szCSDVersion)
End Function
Public Function GE(CheckForOS As OSVersion) As Boolean
    GE = (baseMethods.OSVersion >= CheckForOS)
End Function
Public Function Gr(CheckForOS As OSVersion) As Boolean
    Gr = (baseMethods.OSVersion > CheckForOS)
End Function
Public Function Eq(CheckForOS As OSVersion) As Boolean
    Eq = (baseMethods.OSVersion = CheckForOS)
End Function
Public Function supportOS(CheckForOS As OSVersion) As Boolean
    If CheckForOS = OSHigher Then _
        supportOS = True: Exit Function
    If CheckForOS = OSUnknown Then throw InvalidArgumentValueException("Unknown OSVersion.")
    Dim IsHigher As Boolean
    IsHigher = ((CheckForOS And OSHigher) = OSHigher)
    If IsHigher Then
        Dim lngChk As Long
        lngChk = CheckForOS
        lngChk = (lngChk Xor OSHigher)
        supportOS = (baseMethods.OSVersion >= lngChk)
    Else
        supportOS = (baseMethods.OSVersion = CheckForOS)
    End If
End Function


Public Function ComputerName(Index As ComputerNames) As String
    ComputerName = kernelMethods.ComputerName(Index)
End Function

Public Function IsWow64Environment() As Boolean
    
End Function

Public Sub RefreshMemoryStatus(Optional RefreshInfo As Boolean = True)
    Static memInfo_IsEverRefreshed As Boolean
    If Not RefreshInfo Then _
        If memInfo_IsEverRefreshed Then Exit Sub
    Call API_GlobalMemoryStatus(lmemInfo)
    memInfo_IsEverRefreshed = True
End Sub
Public Function TotalMemorySize(Optional RefreshInfo As Boolean = False) As Double
    Call RefreshMemoryStatus(RefreshInfo)
    TotalMemorySize = CDbl(lmemInfo.TotalPhysMemory)
End Function
Public Function AvailableMemorySize(Optional RefreshInfo As Boolean = False) As Double
    Call RefreshMemoryStatus(RefreshInfo)
    AvailableMemorySize = CDbl(lmemInfo.AvailablePhysMemory)
End Function
Public Function UsedMemorySize(Optional RefreshInfo As Boolean = False) As Double
    Call RefreshMemoryStatus(RefreshInfo)
    UsedMemorySize = CDbl(lmemInfo.TotalPhysMemory) - CDbl(lmemInfo.AvailablePhysMemory)
End Function
Public Function TotalPageFileSize(Optional RefreshInfo As Boolean = False) As Double
    Call RefreshMemoryStatus(RefreshInfo)
    TotalPageFileSize = CDbl(lmemInfo.Length)
End Function
Public Function AvailablePageFileSize(Optional RefreshInfo As Boolean = False) As Double
    Call RefreshMemoryStatus(RefreshInfo)
    AvailablePageFileSize = CDbl(lmemInfo.AvailPageFile)
End Function
Public Function UsedPageFileSize(Optional RefreshInfo As Boolean = False) As Double
    Call RefreshMemoryStatus(RefreshInfo)
    UsedPageFileSize = CDbl(lmemInfo.TotalPageFile) - CDbl(lmemInfo.AvailPageFile)
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
