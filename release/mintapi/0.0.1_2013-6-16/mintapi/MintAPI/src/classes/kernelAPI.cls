VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "kernelAPI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 0
Const CLASSID As String = "kernelMethods"

Private Const API_kernelMethods_AC_LINE_OFFLINE = &H0
Private Const API_kernelMethods_AC_LINE_ONLINE = &H1
Private Const API_kernelMethods_AC_LINE_BACKUP_POWER = &H2
Private Const API_kernelMethods_AC_LINE_UNKNOWN = &HFF

Private Declare Function API_kernelMethods_GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function API_kernelMethods_GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function API_kernelMethods_GetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function API_kernelMethods_GetSystemPowerStatus Lib "kernel32" Alias "GetSystemPowerStatus" (lpSystemPowerStatus As API_kernelMethods_SYSTEM_POWER_STATUS) As Long
Private Declare Function API_kernelMethods_GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function API_kernelMethods_GetCurrentProcess Lib "kernel32" Alias "GetCurrentProcess" () As Long

Private Type API_kernelMethods_SYSTEM_POWER_STATUS
    ACLineStatus As Byte
    BatteryFlag As Byte
    BatteryLifePercent As Byte
    Reserved1 As Byte
    BatteryLifeTime As Long
    BatteryFullLifeTime As Long
End Type

Public Enum BatteryState
    bsOffline = API_kernelMethods_AC_LINE_OFFLINE
    bsOnline = API_kernelMethods_AC_LINE_ONLINE
    bsPowerSave = API_kernelMethods_AC_LINE_BACKUP_POWER
    bsUnknown = API_kernelMethods_AC_LINE_UNKNOWN
End Enum
Public Type DiskDriveSizesString
    ddssTotal As String
    ddssAvailable As String
    ddssUsed As String
End Type
Public Type MemorySizesString
    mssTotal As String
    mssAvailable As String
    mssFree As String
    mssKernel As String
    mssUsed As String
    mssPaged As String
    mssTotalPage As String
    mssAvailablePage As String
    mssUsedPage As String
End Type

Public Function KhInstance() As Long
    KhInstance = API_kernelMethods_GetModuleHandle(vbNullString)
End Function
Public Function GetCurrentThreadID() As Long
    
End Function
Public Function LoadModule(Path As String) As Long
    
End Function
Public Sub UnloadModule(HANDLE As Long)
    
End Sub
Public Function GetMethodAddress(mHandle As Long, ModuleName As String) As String
    
End Function
Public Function GetTickCount() As Long
    GetTickCount = API_kernelMethods_GetTickCount
End Function

Public Sub EnableShutdown()
    
End Sub
Public Sub Shutdown(Optional ByVal Force As Boolean = False)
    
End Sub
Public Sub EnableHibernate()
    
End Sub
Public Sub Hibernate(Optional ByVal Force As Boolean = False)
    
End Sub
Public Sub DisableHibernate()
    
End Sub
Public Sub Logoff()
    
End Sub
Public Sub SwitchUser()
    
End Sub
Public Function GetDiskDriveSizesString(Optional ByVal DriveLetter As String = "") As DiskDriveSizesString
    
End Function
Public Function GetMemorySizesString() As MemorySizesString
    
End Function
Public Function GetLogicalDrivesString() As String
    Dim buf As String * SMALLLPSTR, bufSize As Long
    bufSize = SMALLLPSTR
    buf = String(SMALLLPSTR, Chr(0))
    If Not API_kernelMethods_GetLogicalDriveStrings(buf, bufSize) = SUCCESS Then throw SystemCallFailureException
    GetLogicalDrivesString = GetLPSTR(buf)
End Function

Public Sub SetCurrentDirectory(Path As String)
    Call ChDir(Path)
End Sub
Public Function GetCurrentDirectory() As String
    GetCurrentDirectory = CurDir
End Function
Public Sub SetCurrentDrive(Drive As String)
    Call ChDrive(Drive)
End Sub
Public Function GetCurrentDrive() As String
    GetCurrentDrive = Left(CurDir, 3)
End Function

Public Function Power_BatteryMode() As BatteryState
    Dim pBattery As API_kernelMethods_SYSTEM_POWER_STATUS
    If API_kernelMethods_GetSystemPowerStatus(pBattery) <> SUCCESS Then throw SystemCallFailureException
    Power_BatteryMode = pBattery.ACLineStatus
End Function
Public Function Power_GetBatteryValue() As Long
    Dim ps As BatteryState
    ps = Power_BatteryMode
    If equalTo(SomeEqual Or NotValue, ps, bsOnline, bsPowerSave) Then throw InvalidCallException
    
End Function
Public Function Power_GetBatteryTotal() As Long
    
End Function


