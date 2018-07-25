Attribute VB_Name = "kernelMethods"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "kernelMethods"

Private Const INVALID_HANDLE_VALUE = -1
Private Const STD_INPUT_HANDLE As Integer = -10
Private Const STD_OUTPUT_HANDLE As Integer = -11
Private Const STD_ERROR_HANDLE As Integer = -12

Private Type COORD
    X As Integer
    Y As Integer
End Type
Private Type SMALL_RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type
Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD
    dwCursorPosition As COORD
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type
Public Type CHAR_INFO
    Char As Integer
    Attributes As Integer
End Type


Private Const API_kernelMethods_AC_LINE_OFFLINE = &H0
Private Const API_kernelMethods_AC_LINE_ONLINE = &H1
Private Const API_kernelMethods_AC_LINE_BACKUP_POWER = &H2
Private Const API_kernelMethods_AC_LINE_UNKNOWN = &HFF

Private Const API_kernelMethods_DDD_REMOVE_DEFINITION As Long = &H2

Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6
Public Enum API_DriveType
    dt_Fixed = DRIVE_FIXED
    dt_CDRom = DRIVE_CDROM
    dt_RamDisk = DRIVE_RAMDISK
    dt_Removable = DRIVE_REMOVABLE
    dt_Remote = DRIVE_REMOTE
    dt_Free = &H7F
End Enum


Private Const API_kernelMethods_ICC_USEREX_CLASSES = &H200

Private Type API_kernelMethods_tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Public Declare Function API_GetLastError Lib "kernel32" Alias "GetLastError" () As Long
Private Declare Sub API_kernelMethods_CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hPass As Long)
Private Declare Sub API_kernelMethods_InitCommonControls Lib "comctl32" Alias "InitCommonControls" ()
Public Declare Function API_kernelMethods_InitCommonControlsEx Lib "comctl32" Alias "InitCommonControlsEx" (Iccex As API_kernelMethods_tagInitCommonControlsEx) As Boolean
Private Declare Function API_kernelMethods_EnumThreadWindows Lib "user32" Alias "EnumThreadWindows" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Private Declare Function API_kernelMethods_DisableThreadLibraryCalls Lib "kernel32" Alias "DisableThreadLibraryCalls" (ByVal hLibModule As Long) As Long
Private Declare Function API_kernelMethods_GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function API_kernelMethods_GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function API_kernelMethods_GetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function API_kernelMethods_GetSystemPowerStatus Lib "kernel32" Alias "GetSystemPowerStatus" (lpSystemPowerStatus As API_kernelMethods_SYSTEM_POWER_STATUS) As Long
Private Declare Function API_kernelMethods_GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function API_kernelMethods_GetCurrentProcess Lib "kernel32" Alias "GetCurrentProcess" () As Long
Private Declare Function API_kernelMethods_GetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Function API_kernelMethods_CreateProcessAsUser Lib "kernel32" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As API_kernelMethods_SECURITY_ATTRIBUTES, lpThreadAttributes As API_kernelMethods_SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, lpStartupInfo As API_kernelMethods_STARTUPINFO, lpProcessInformation As API_kernelMethods_PROCESS_INFORMATION) As Long
Private Declare Function API_kernelMethods_GetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Long
Private Declare Function API_kernelMethods_GetCurrentThread Lib "kernel32" Alias "GetCurrentThread" () As Long
Private Declare Function API_kernelMethods_LoadModule Lib "kernel32" Alias "LoadModule" (ByVal lpModuleName As String, lpParameterBlock As Any) As Long
Private Declare Function API_kernelMethods_GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function API_kernelMethods_SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Declare Function API_kernelMethods_DefineDosDevice Lib "kernel32" Alias "DefineDosDeviceA" (ByVal dwFlags As Long, ByVal lpDeviceName As String, Optional ByVal lpTargetPath As String = vbNullString) As Long
Private Declare Function API_kernelMethods_GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function API_kernelMethods_GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function API_kernelMethods_SetStdHandle Lib "kernel32" (ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
Private Declare Function API_kernelMethods_GetComputerNameEx Lib "kernel32" Alias "GetComputerNameExA" (ByVal NameType As API_ComputerNames, ByVal lpBuffer As String, ByRef nSize As Long) As Long
    
Private Declare Function API_kernelMethods_WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function API_kernelMethods_WriteConsoleUnicode Lib "kernel32" Alias "WriteConsoleW" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function API_kernelMethods_WriteConsoleOutput Lib "kernel32" Alias "WriteConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpWriteRegion As SMALL_RECT) As Long
Private Declare Function API_kernelMethods_WriteConsoleOutputAttribute Lib "kernel32" Alias "WriteConsoleOutputAttribute" (ByVal hConsoleOutput As Long, lpAttribute As Integer, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
Private Declare Function API_kernelMethods_WriteConsoleOutputCharacter Lib "kernel32" Alias "WriteConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long

Private Type API_kernelMethods_SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type API_kernelMethods_STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type API_kernelMethods_PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type


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

Public Enum API_ComputerNames
    cn_NetBIOS = 0
    cn_DnsHostname = 1
    cn_DnsDomain = 2
    cn_DnsFullyQualified = 3
    cn_PhysicalNetBIOS = 4
    cn_PhysicalDnsHostName = 5
    cn_PhysicalDnsDomain = 6
    cn_PhysicalDnsFullyQualified = 7
    cn_Max = 8
End Enum

Dim inited As Boolean

Dim lErr As Long

Public Sub Initialize()
    If inited Then Exit Sub
    Call baseConstants.Initialize
    Call baseExceptions.Initialize
    Call baseMethods.Initialize
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    inited = False
End Sub

Public Sub rLastError()
    lErr = 0
End Sub
Public Function IfError() As Exception
'    If VarType(lErr) = vbObject Then
'        Set IfError = SystemCallFailureException
'    Else
'            IfError = SystemCallFailureException
'    End If
    lErr = API_GetLastError
    If lErr = 0 Then
        IfError.ExceptionType = ExceptionType.EXP_et_NoError
    End If
End Function

Public Function KhInstance() As Long
    KhInstance = API_kernelMethods_GetModuleHandle(vbNullString)
End Function
Public Function GetCurrentThreadId() As Long
    GetCurrentThreadId = API_kernelMethods_GetCurrentThreadId
End Function
Public Function GetCurrentProcessId() As Long
    GetCurrentProcessId = API_kernelMethods_GetCurrentProcessId
End Function
Public Function LoadModule(Path As String) As Long

End Function
Public Sub UnloadModule(Handle As Long)

End Sub
Public Function GetMethodAddress(mHandle As Long) As Long
    GetMethodAddress = mHandle
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
Public Sub Logoff(Optional ByVal Force As Boolean = False)

End Sub
Public Sub SwitchUser()

End Sub
Public Sub RestartSystem(Optional ByVal Force As Boolean = False)
    
End Sub
Public Sub Sleep(Optional ByVal Force As Boolean = False)
    
End Sub
Public Function GetMemorySizesString() As MemorySizesString

End Function

Public Sub InitializeCommonControls()
    'Call API_kernelMethods_InitCommonControls
    Dim Iccex As API_kernelMethods_tagInitCommonControlsEx
    With Iccex
        .lngSize = LenB(Iccex)
        .lngICC = API_kernelMethods_ICC_USEREX_CLASSES
    End With
    Call API_kernelMethods_InitCommonControlsEx(Iccex)
End Sub

Public Function ComputerName(Index As API_ComputerNames) As String
    Dim buf As String * SMALLLPSTR, bufSize As Long
    bufSize = SMALLLPSTR
    buf = String(SMALLLPSTR, Chr(0))
    If API_kernelMethods_GetComputerNameEx(Index, buf, bufSize) = 0 Then _
        throw SystemCallFailureException
    ComputerName = GetLPSTR(buf)
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
Public Function GetDrives() As String
    Dim buf As String * SMALLLPSTR, bufSize As Long
    bufSize = SMALLLPSTR
    buf = String(SMALLLPSTR, " ")
    If Not (API_kernelMethods_GetLogicalDriveStrings(bufSize, buf) = SUCCESS) Then throw SystemCallFailureException
    GetDrives = Trim$(buf)
End Function
Private Function ValidateDriveLetter(DriveLetter As String) As String
    DriveLetter = Trim(DriveLetter)
    Dim ln As Long
    ln = Len(DriveLetter)
    If ln <= 0 Then GoTo raiseError
    If ln = 1 Then
        ' do nothing
    ElseIf ln = 2 Then
        If Not (DriveLetter Like "?:") Then GoTo raiseError
    ElseIf ln = 3 Then
        If Not (DriveLetter Like "?:[/,\,|]") Then GoTo raiseError
    Else
        If Not (Left(DriveLetter, 3) Like "?:[/,\,|]") Then GoTo raiseError
    End If
    ValidateDriveLetter = Left(DriveLetter, 1)
    Exit Function
raiseError:
    throw Exception("Invalid Drive Letter.")
End Function
Public Function GetDriveType(DriveLetter As String) As API_DriveType
    DriveLetter = ValidateDriveLetter(DriveLetter)
    GetDriveType = API_kernelMethods_GetDriveType(DriveLetter & ":\")
End Function
Public Function GetDiskDriveSizesString(Optional ByVal DriveLetter As String = "") As DiskDriveSizesString

End Function
Public Function GetLogicalDrivesString() As String
    Dim buf As String * SMALLLPSTR, bufSize As Long
    bufSize = SMALLLPSTR
    buf = String(SMALLLPSTR, " ")
    If API_kernelMethods_GetLogicalDriveStrings(bufSize, buf) = 0 Then throw SystemCallFailureException
    GetLogicalDrivesString = Trim$(Replace(buf, Chr(0), " "))
End Function
Public Function VolumeName(ByVal DriveLetter As String) As String
    Dim sBuffer As String

    sBuffer = String(SMALLLPSTR, Chr(0))
    'fix bad parameter values
    DriveLetter = ValidateDriveLetter(DriveLetter) & ":\"
    If API_kernelMethods_GetVolumeInformation(DriveLetter, sBuffer, Len(sBuffer), 0, 0, 0, Space$(SMALLLPSTR), SMALLLPSTR) = 0 Then
        throw SystemCallFailureException("An error occured while trying to call GetVolumeInformation in kernel.")
    Else
        VolumeName = GetLPSTR(sBuffer) 'Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
    End If
End Function
Public Sub SetVolumeName(ByVal DriveLetter As String, VolumeName As String)
    DriveLetter = ValidateDriveLetter(DriveLetter) & ":\"
    If API_kernelMethods_SetVolumeLabel(DriveLetter, VolumeName) = 0 Then
        throw SystemCallFailureException("An error occured while trying to call SetVolumeLabel in kernel.")
    End If
End Sub
Public Sub EjectDrive(DriveLetter As String)
    
End Sub
Public Sub MountFolder(DriveLetter As String, FolderPath As String, Optional VolumeName As String = "")
    Call MountVitualDirectory(ValidateDriveLetter(DriveLetter), FolderPath)
    On Error GoTo Err
    If VolumeName <> "" Then
        Call SetVolumeName(DriveLetter, VolumeName)
    End If
Err:
End Sub
Public Sub UnmountFolder(DriveLetter As String, FolderPath As String)
    Call MountVitualDirectory(ValidateDriveLetter(DriveLetter), FolderPath, True)
End Sub
Private Function MountVitualDirectory(ByVal sDriveLetter As String, ByVal sMountPath As String, Optional ByVal bUnmount As Boolean = False) As Boolean
    On Error GoTo MountVD_Error
    Dim lDriveType As Long
    'Remove any white spaces ...
    sDriveLetter = ValidateDriveLetter(sDriveLetter)
    'Check if specified Folder Path is correct & exists ...
    If Not DirectoryExists(sMountPath) Then
        throw Exception("Specified mount path is wrong or does not point to a valid Windows Folder item.")
        MountVitualDirectory = False
    End If
    'DefineDosDevice requires ':' at the end of drive letter ...
    sDriveLetter = sDriveLetter & ":"
    'Oops ! For GetDriveType, we need to append a \ to Drive Letter. Let us check if specified Drive Letter is available ...
    lDriveType = API_kernelMethods_GetDriveType(sDriveLetter & "\")
    'Only Unknown type of Drive letters are allowed to use for Virtual Mount for obvious reasons ...
    Select Case lDriveType
        Case DRIVE_CDROM
                throw Exception("Specified Drive letter is not available to mount virtual drive.")
                MountVitualDirectory = False
        'Virtual Drive, when mounted successfully, is recognized as Fixed Drive. So, here we will implement the code for Unmount ...
        Case DRIVE_FIXED
                If bUnmount = False Then
                    throw Exception("Specified Drive letter is not available to mount virtual drive.")
                    MountVitualDirectory = False
                Else
                    MountVitualDirectory = CBool(API_kernelMethods_DefineDosDevice(API_kernelMethods_DDD_REMOVE_DEFINITION, sDriveLetter, sMountPath))
                    MountVitualDirectory = True
                End If
        Case DRIVE_RAMDISK
                throw Exception("Specified Drive letter is not available to mount virtual drive.")
                MountVitualDirectory = False
        Case DRIVE_REMOVABLE
                throw Exception("Specified Drive letter is not available to mount virtual drive.")
                MountVitualDirectory = False
        Case DRIVE_REMOTE:
                throw Exception("Specified Drive letter is not available to mount virtual drive.")
                MountVitualDirectory = False
        'Here it means that the Drive Letter is available for us to mount Virtual Drive ...
        Case Else:
                If bUnmount = False Then
                    MountVitualDirectory = CBool(API_kernelMethods_DefineDosDevice(0, sDriveLetter, sMountPath))
                    MountVitualDirectory = True
                End If
    End Select
    MountVitualDirectory = True
    'This will avoid empty error window to appear.
    Exit Function
MountVD_Error:
On Error GoTo 0
    throw Exception(Err.Description)
    MountVitualDirectory = False
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


Public Function IsWow64OperatingSystem() As Boolean
    
End Function

Public Function GetApplicationStdOutHandle() As Long
    GetApplicationStdOutHandle = API_kernelMethods_GetStdHandle(STD_OUTPUT_HANDLE)
End Function
Public Function GetApplicationStdInHandle() As Long
    GetApplicationStdInHandle = API_kernelMethods_GetStdHandle(STD_INPUT_HANDLE)
End Function
Public Function GetApplicationStdErrHandle() As Long
    GetApplicationStdErrHandle = API_kernelMethods_GetStdHandle(STD_ERROR_HANDLE)
End Function

Public Sub CloseHandle(Handle As Long)
    Call API_kernelMethods_CloseHandle(Handle)
End Sub

Public Sub PrintToHandle(strPrintable As String)
    
End Sub
Public Function ReadFromHandle() As String
    
End Function

Public Sub PrintToStdOutput(strPrintable As String)
    
End Sub
Public Sub PrintToStdError(strPrintable As String)
    
End Sub
Public Sub PrintToConsole(ConsoleHandle As Long, strPrintable As String)
    
End Sub

Public Function AppConsole() As Long
    
End Function
Public Function CreateConsole() As Long
    
End Function
Public Sub DistroyConsole(ConsoleHandle As Long)
    
End Sub

Public Function ReadFromConsole(ConsoleHandle As Long) As String
    
End Function
Public Function ReadFromStdInput() As String
    
End Function
