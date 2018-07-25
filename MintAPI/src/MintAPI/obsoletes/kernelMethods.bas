Attribute VB_Name = "kernelMethods"
''@PROJECT_LICENSE
'
'Option Explicit
'Option Base 0
'Const CLASSID As String = "kernelMethods"
'
'Private Const INVALID_HANDLE_VALUE = -1
'Private Const STD_INPUT_HANDLE As Integer = -10
'Private Const STD_OUTPUT_HANDLE As Integer = -11
'Private Const STD_ERROR_HANDLE As Integer = -12
'
'Private Type COORD
'    X As Integer
'    Y As Integer
'End Type
'Private Type SMALL_RECT
'    Left As Integer
'    Top As Integer
'    Right As Integer
'    Bottom As Integer
'End Type
'Private Type CONSOLE_SCREEN_BUFFER_INFO
'    dwSize As COORD
'    dwCursorPosition As COORD
'    wAttributes As Integer
'    srWindow As SMALL_RECT
'    dwMaximumWindowSize As COORD
'End Type
'Public Type CHAR_INFO
'    Char As Integer
'    Attributes As Integer
'End Type
'
'
'Private Const API_kernelMethods_AC_LINE_OFFLINE = &H0
'Private Const API_kernelMethods_AC_LINE_ONLINE = &H1
'Private Const API_kernelMethods_AC_LINE_BACKUP_POWER = &H2
'Private Const API_kernelMethods_AC_LINE_UNKNOWN = &HFF
'
''Public Enum API_DriveType
''    dt_Fixed = DRIVE_FIXED
''    dt_CDRom = DRIVE_CDROM
''    dt_RamDisk = DRIVE_RAMDISK
''    dt_Removable = DRIVE_REMOVABLE
''    dt_Remote = DRIVE_REMOTE
''    dt_Free = &H7F
''End Enum
'
'
'Private Declare Function API_GetLastError Lib "Kernel32" Alias "GetLastError" () As Long
'Private Declare Sub API_kernelMethods_CloseHandle Lib "Kernel32" Alias "CloseHandle" (ByVal hPass As Long)
'Private Declare Function API_kernelMethods_EnumThreadWindows Lib "user32" Alias "EnumThreadWindows" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
'Private Declare Function API_kernelMethods_DisableThreadLibraryCalls Lib "Kernel32" Alias "DisableThreadLibraryCalls" (ByVal hLibModule As Long) As Long
'Private Declare Function API_kernelMethods_GetModuleFileName Lib "Kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
'Private Declare Function API_kernelMethods_GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'Private Declare Function API_kernelMethods_GetTickCount Lib "Kernel32" Alias "GetTickCount" () As Long
'Private Declare Function API_kernelMethods_GetSystemPowerStatus Lib "Kernel32" Alias "GetSystemPowerStatus" (lpSystemPowerStatus As API_kernelMethods_SYSTEM_POWER_STATUS) As Long
'Private Declare Function API_kernelMethods_GetLogicalDriveStrings Lib "Kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Private Declare Function API_kernelMethods_GetCurrentProcess Lib "Kernel32" Alias "GetCurrentProcess" () As Long
'Private Declare Function API_kernelMethods_GetCurrentProcessId Lib "Kernel32" Alias "GetCurrentProcessId" () As Long
'Private Declare Function API_kernelMethods_CreateProcessAsUser Lib "Kernel32" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As API_kernelMethods_SECURITY_ATTRIBUTES, lpThreadAttributes As API_kernelMethods_SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, ByVal lpCurrentDirectory As String, lpStartupInfo As API_kernelMethods_STARTUPINFO, lpProcessInformation As API_kernelMethods_PROCESS_INFORMATION) As Long
'Private Declare Function API_kernelMethods_GetCurrentThreadId Lib "Kernel32" Alias "GetCurrentThreadId" () As Long
'Private Declare Function API_kernelMethods_GetCurrentThread Lib "Kernel32" Alias "GetCurrentThread" () As Long
'Private Declare Function API_kernelMethods_LoadModule Lib "Kernel32" Alias "LoadModule" (ByVal lpModuleName As String, lpParameterBlock As Any) As Long
'Private Declare Function API_kernelMethods_GetVolumeInformation Lib "Kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'Private Declare Function API_kernelMethods_SetVolumeLabel Lib "Kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
'Private Declare Function API_kernelMethods_GetDriveType Lib "Kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'Private Declare Function API_kernelMethods_GetStdHandle Lib "Kernel32" (ByVal nStdHandle As Long) As Long
'Private Declare Function API_kernelMethods_SetStdHandle Lib "Kernel32" (ByVal nStdHandle As Long, ByVal nHandle As Long) As Long
'
'Private Declare Function API_kernelMethods_WriteConsole Lib "Kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
'Private Declare Function API_kernelMethods_WriteConsoleUnicode Lib "Kernel32" Alias "WriteConsoleW" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
'Private Declare Function API_kernelMethods_WriteConsoleOutput Lib "Kernel32" Alias "WriteConsoleOutputA" (ByVal hConsoleOutput As Long, lpBuffer As CHAR_INFO, dwBufferSize As COORD, dwBufferCoord As COORD, lpWriteRegion As SMALL_RECT) As Long
'Private Declare Function API_kernelMethods_WriteConsoleOutputAttribute Lib "Kernel32" Alias "WriteConsoleOutputAttribute" (ByVal hConsoleOutput As Long, lpAttribute As Integer, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfAttrsWritten As Long) As Long
'Private Declare Function API_kernelMethods_WriteConsoleOutputCharacter Lib "Kernel32" Alias "WriteConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As String, ByVal nLength As Long, dwWriteCoord As COORD, lpNumberOfCharsWritten As Long) As Long
'
'Private Type API_kernelMethods_SECURITY_ATTRIBUTES
'    nLength As Long
'    lpSecurityDescriptor As Long
'    bInheritHandle As Long
'End Type
'Private Type API_kernelMethods_STARTUPINFO
'    cb As Long
'    lpReserved As String
'    lpDesktop As String
'    lpTitle As String
'    dwX As Long
'    dwY As Long
'    dwXSize As Long
'    dwYSize As Long
'    dwXCountChars As Long
'    dwYCountChars As Long
'    dwFillAttribute As Long
'    dwFlags As Long
'    wShowWindow As Integer
'    cbReserved2 As Integer
'    lpReserved2 As Long
'    hStdInput As Long
'    hStdOutput As Long
'    hStdError As Long
'End Type
'Private Type API_kernelMethods_PROCESS_INFORMATION
'    hProcess As Long
'    hThread As Long
'    dwProcessId As Long
'    dwThreadId As Long
'End Type
'
'
'Private Type API_kernelMethods_SYSTEM_POWER_STATUS
'    ACLineStatus As Byte
'    BatteryFlag As Byte
'    BatteryLifePercent As Byte
'    Reserved1 As Byte
'    BatteryLifeTime As Long
'    BatteryFullLifeTime As Long
'End Type
'
''Public Enum BatteryState
''    bsOffline = API_kernelMethods_AC_LINE_OFFLINE
''    bsOnline = API_kernelMethods_AC_LINE_ONLINE
''    bsPowerSave = API_kernelMethods_AC_LINE_BACKUP_POWER
''    bsUnknown = API_kernelMethods_AC_LINE_UNKNOWN
''End Enum
''Public Type DiskDriveSizesString
''    ddssTotal As String
''    ddssAvailable As String
''    ddssUsed As String
''End Type
''Public Type MemorySizesString
''    mssTotal As String
''    mssAvailable As String
''    mssFree As String
''    mssKernel As String
''    mssUsed As String
''    mssPaged As String
''    mssTotalPage As String
''    mssAvailablePage As String
''    mssUsedPage As String
''End Type
''
'
''Dim inited As Boolean
''
''Dim lErr As Long
'
''Public Sub Initialize()
''    If inited Then Exit Sub
''    inited = True
''End Sub
''Public Sub Dispose(Optional ByVal Force As Boolean = False)
''    If Not inited Then Exit Sub
''    inited = False
''End Sub
''
''Public Sub rLastError()
''    lErr = 0
''End Sub
''Public Function IfError() As Exception
'''    If VarType(lErr) = vbObject Then
'''        Set IfError = SystemCallFailureException
'''    Else
'''            IfError = SystemCallFailureException
'''    End If
''    lErr = API_GetLastError
''    If lErr = 0 Then
''        IfError.ExceptionType = ExceptionType.EXP_et_NoError
''    Else
''        IfError.ExceptionType = ExceptionType.EXP_et_SystemCallFailure
''    End If
''End Function
'
''Public Function KhInstance() As Long
''    KhInstance = API_kernelMethods_GetModuleHandle(vbNullString)
''End Function
''Public Function GetCurrentThreadId() As Long
''    GetCurrentThreadId = API_kernelMethods_GetCurrentThreadId
''End Function
''Public Function GetCurrentProcessId() As Long
''    GetCurrentProcessId = API_kernelMethods_GetCurrentProcessId
''End Function
''Public Function GetTickCount() As Long
''    GetTickCount = API_kernelMethods_GetTickCount
''End Function
'
''Public Sub EnableShutdown()
''
''End Sub
''Public Sub Shutdown(Optional ByVal Force As Boolean = False)
''
''End Sub
''Public Sub EnableHibernate()
''
''End Sub
''Public Sub Hibernate(Optional ByVal Force As Boolean = False)
''
''End Sub
''Public Sub DisableHibernate()
''
''End Sub
''Public Sub Logoff(Optional ByVal Force As Boolean = False)
''
''End Sub
''Public Sub SwitchUser()
''
''End Sub
''Public Sub RestartSystem(Optional ByVal Force As Boolean = False)
''
''End Sub
''Public Sub Sleep(Optional ByVal Force As Boolean = False)
''
''End Sub
''Public Function GetMemorySizesString() As MemorySizesString
''
''End Function
''
''Public Function ComputerName(Index As API_ComputerNames) As String
''
''End Function
'
''Public Sub SetCurrentDirectory(Path As String)
''    Call ChDir(Path)
''End Sub
''Public Function GetCurrentDirectory() As String
''    GetCurrentDirectory = CurDir
''End Function
''Public Sub SetCurrentDrive(Drive As String)
''    Call ChDrive(Drive)
''End Sub
''Public Function GetCurrentDrive() As String
''    GetCurrentDrive = Left(CurDir, 3)
''End Function
''Public Function GetDrives() As String
''    Dim Buf As String * SMALLLPSTR, bufSize As Long
''    bufSize = SMALLLPSTR
''    Buf = String(SMALLLPSTR, " ")
''    If Not (API_kernelMethods_GetLogicalDriveStrings(bufSize, Buf) = SUCCESS) Then throw Exps.SystemCallFailureException
''    GetDrives = Trim$(Buf)
''End Function
''Public Function GetDriveType(DriveLetter As String) As API_DriveType
''    'DriveLetter = ValidateDriveLetter(DriveLetter)
''    GetDriveType = API_kernelMethods_GetDriveType(DriveLetter & ":\")
''End Function
''Public Function GetDiskDriveSizesString(Optional ByVal DriveLetter As String = "") As DiskDriveSizesString
''
''End Function
''Public Function GetLogicalDrivesString() As String
''    Dim Buf As String * SMALLLPSTR, bufSize As Long
''    bufSize = SMALLLPSTR
''    Buf = String(SMALLLPSTR, " ")
''    If API_kernelMethods_GetLogicalDriveStrings(bufSize, Buf) = 0 Then throw Exps.SystemCallFailureException
''    GetLogicalDrivesString = Trim$(Replace(Buf, Chr(0), " "))
''End Function
''Public Function VolumeName(ByVal DriveLetter As String) As String
''    Dim sBuffer As String
''
''    sBuffer = String(SMALLLPSTR, Chr(0))
''    'fix bad parameter values
''    'DriveLetter = ValidateDriveLetter(DriveLetter) & ":\"
''    If API_kernelMethods_GetVolumeInformation(DriveLetter, sBuffer, Len(sBuffer), 0, 0, 0, Space$(SMALLLPSTR), SMALLLPSTR) = 0 Then
''        throw Exps.SystemCallFailureException("An error occured while trying to call GetVolumeInformation in kernel.")
''    Else
''        VolumeName = GetLPSTR(sBuffer) 'Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
''    End If
''End Function
''Public Sub SetVolumeName(ByVal DriveLetter As String, VolumeName As String)
''    'DriveLetter = ValidateDriveLetter(DriveLetter) & ":\"
''    If API_kernelMethods_SetVolumeLabel(DriveLetter, VolumeName) = 0 Then
''        throw Exps.SystemCallFailureException("An error occured while trying to call SetVolumeLabel in kernel.")
''    End If
''End Sub
''Public Sub EjectDrive(DriveLetter As String)
''
''End Sub
''Public Sub MountFolder(DriveLetter As String, FolderPath As String, Optional VolumeName As String = "")
''    'Call MountVitualDirectory(ValidateDriveLetter(DriveLetter), FolderPath)
''    On Error GoTo Err
''    If VolumeName <> "" Then
''        Call SetVolumeName(DriveLetter, VolumeName)
''    End If
''Err:
''End Sub
''Public Sub UnmountFolder(DriveLetter As String, FolderPath As String)
''    'Call MountVitualDirectory(ValidateDriveLetter(DriveLetter), FolderPath, True)
''End Sub
''Private Function MountVitualDirectory(ByVal sDriveLetter As String, ByVal sMountPath As String, Optional ByVal bUnmount As Boolean = False) As Boolean
'
''End Function
''
''Public Function Power_BatteryMode() As BatteryState
''    Dim pBattery As API_kernelMethods_SYSTEM_POWER_STATUS
''    If API_kernelMethods_GetSystemPowerStatus(pBattery) <> SUCCESS Then throw Exps.SystemCallFailureException
''    Power_BatteryMode = pBattery.ACLineStatus
''End Function
''Public Function Power_GetBatteryValue() As Long
''    Dim ps As BatteryState
''    ps = Power_BatteryMode
''    'If EqualTo(SomeEqual Or NotValue, ps, bsOnline, bsPowerSave) Then throw Exps.InvalidCallException
''
''End Function
''Public Function Power_GetBatteryTotal() As Long
''
''End Function
