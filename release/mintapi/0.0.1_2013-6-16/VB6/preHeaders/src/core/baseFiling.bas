Attribute VB_Name = "baseFiling"
'@PROJECT_LICENSE

Option Explicit



Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_SHARE_DELETE = &H4

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_DEVICE = &H40
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_SPARSE_FILE = &H200
Private Const FILE_ATTRIBUTE_REPARSE_POINT = &H400
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000
Private Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
Private Const FILE_ATTRIBUTE_ENCRYPTED = &H4000

Private Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1
Private Const FILE_NOTIFY_CHANGE_DIR_NAME = &H2
Private Const FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
Private Const FILE_NOTIFY_CHANGE_SIZE = &H8
Private Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
Private Const FILE_NOTIFY_CHANGE_LAST_ACCESS = &H20
Private Const FILE_NOTIFY_CHANGE_CREATION = &H40
Private Const FILE_NOTIFY_CHANGE_SECURITY = &H100

Private Const FILE_ACTION_ADDED = &H1
Private Const FILE_ACTION_REMOVED = &H2
Private Const FILE_ACTION_MODIFIED = &H3
Private Const FILE_ACTION_RENAMED_OLD_NAME = &H4
Private Const FILE_ACTION_RENAMED_NEW_NAME = &H5

Private Const MAILSLOT_NO_MESSAGE = -1
Private Const MAILSLOT_WAIT_FOREVER = -1

Private Const FILE_CASE_SENSITIVE_SEARCH = &H1
Private Const FILE_CASE_PRESERVED_NAMES = &H2
Private Const FILE_UNICODE_ON_DISK = &H4
Private Const FILE_PERSISTENT_ACLS = &H8
Private Const FILE_FILE_COMPRESSION = &H10
Private Const FILE_VOLUME_QUOTAS = &H20
Private Const FILE_SUPPORTS_SPARSE_FILES = &H40
Private Const FILE_SUPPORTS_REPARSE_POINTS = &H80
Private Const FILE_SUPPORTS_REMOTE_STORAGE = &H100
Private Const FILE_VOLUME_IS_COMPRESSED = &H8000
Private Const FILE_SUPPORTS_OBJECT_IDS = &H10000
Private Const FILE_SUPPORTS_ENCRYPTION = &H20000
Private Const FILE_NAMED_STREAMS = &H40000
Private Const FILE_READ_ONLY_VOLUME = &H80000



Private Const OF_READ = &H0
Private Const OF_WRITE = &H1
Private Const OF_READWRITE = &H2
Private Const OF_SHARE_COMPAT = &H0
Private Const OF_SHARE_EXCLUSIVE = &H10
Private Const OF_SHARE_DENY_WRITE = &H20
Private Const OF_SHARE_DENY_READ = &H30
Private Const OF_SHARE_DENY_NONE = &H40
Private Const OF_PARSE = &H100
Private Const OF_DELETE = &H200
Private Const OF_VERIFY = &H400
Private Const OF_CANCEL = &H800
Private Const OF_CREATE = &H1000
Private Const OF_PROMPT = &H2000
Private Const OF_EXIST = &H4000
Private Const OF_REOPEN = &H8000

Private Const OFS_MAXPATHNAME = 128

Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Private Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
Private Const FILE_FLAG_OPEN_REPARSE_POINT = &H200000
Private Const FILE_FLAG_OPEN_NO_RECALL = &H100000
Private Const FILE_FLAG_FIRST_PIPE_INSTANCE = &H80000

Private Const INVALID_HANDLE_VALUE = -1
Private Const INVALID_FILE_SIZE = -1 '&HFFFFFFFF
Private Const INVALID_SET_FILE_POINTER = -1
Private Const INVALID_FILE_ATTRIBUTES = -1


Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_ALL = &H10000000

Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2

Private Const CREATE_NEW = 1
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const TRUNCATE_EXISTING = 5


'Private Const SECURITY_ANONYMOUS         = ( SecurityAnonymous      << 16 )
'Private Const SECURITY_IDENTIFICATION   =  ( SecurityIdentification << 16 )
'Private Const  SECURITY_IMPERSONATION    =  ( SecurityImpersonation  << 16 )
'Private Const  SECURITY_DELEGATION       =  ( SecurityDelegation     << 16 )

Private Const COPY_FILE_FAIL_IF_EXISTS = &H1
Private Const COPY_FILE_RESTARTABLE = &H2
Private Const COPY_FILE_OPEN_SOURCE_FOR_WRITE = &H4
Private Const COPY_FILE_ALLOW_DECRYPTED_DESTINATION = &H8

Private Const PROGRESS_CONTINUE = 0
Private Const PROGRESS_CANCEL = 1
Private Const PROGRESS_STOP = 2
Private Const PROGRESS_QUIET = 3

Private Const SECURITY_CONTEXT_TRACKING = &H40000
Private Const SECURITY_EFFECTIVE_ONLY = &H80000

Private Const SECURITY_SQOS_PRESENT = &H100000
Private Const SECURITY_VALID_SQOS_FLAGS = &H1F0000

Private Declare Function API_baseFiling_OpenFile Lib "kernel32" Alias "OpenFile" (ByVal lpFileName As String, lpReOpenBuff As API_OFSTRUCT, ByVal wStyle As Long) As Long
'Private Declare Function API_baseFiling_OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
'#If Win32_WINNT >= &H502 Then '&H502
'Private Declare Function API_baseFiling_ReOpenFile Lib "kernel32" Alias "ReOpenFile" (hOriginalFile As Long, dwDesiredAccess As Long, dwShareMode As Long, dwFlagsAndAttributes As Long) As Long
'#End If

Private Declare Function API_baseFiling_CreateFile_SEC Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As API_SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function API_baseFiling_CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Private Declare Function API_baseFiling_CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappingAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long

Private Declare Function API_baseFiling_CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long

Private Declare Function API_baseFiling_SetFilePointer Lib "kernel32" Alias "SetFilePointer" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function API_baseFiling_SetFilePointerEx Lib "kernel32" Alias "SetFilePointerEx" (ByVal hFile As Long, ByVal liDistanceToMove As API_LARGE_INTEGER, ByRef lpNewFilePointer As API_LARGE_INTEGER, ByVal dwMoveMethod As Long) As Long

Private Declare Function API_baseFiling_GetFileSize Lib "kernel32" Alias "GetFileSize" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function API_baseFiling_GetFileSizeEx Lib "kernel32" Alias "GetFileSizeEx" (ByVal hFile As Long, lpFileSizeHigh As API_File_BigInt) As Long
'Private Declare Function API_baseFiling_GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, ByRef lpFileSize As LARGE_INTEGER) As Long
Private Declare Function API_baseFiling_GetFileTime Lib "kernel32" Alias "GetFileTime" (ByVal hFile As Long, lpCreationTime As API_FILE_TIME, lpLastAccessTime As API_FILE_TIME, lpLastWriteTime As API_FILE_TIME) As Long
'Private Declare Function API_baseFiling_GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Private Declare Function API_baseFiling_GetFileType Lib "kernel32" Alias "GetFileType" (ByVal hFile As Long) As Long
'Private Declare Function API_baseFiling_GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Byte) As Long
'Private Declare Function API_baseFiling_GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function API_baseFiling_GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
'Private Declare Function API_baseFiling_GetFileAttributesEx Lib "kernel32" Alias "GetFileAttributesExA" (ByVal lpFileName As String, ByVal fInfoLevelId As Struct_MembersOf_GET_FILEEX_INFO_LEVELS, lpFileInformation As Any) As Long
'Private Declare Function API_baseFiling_GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
'Private Declare Function API_baseFiling_GetFilePatchSignature Lib "MSPATCHA" (ByVal FileName As String, ByVal OptionFlags As Long, OptionData As Any, ByVal IgnoreRangeCount As Long, ByRef IgnoreRangeArray As PPATCH_IGNORE_RANGE, ByVal RetainRangeCount As Long, ByRef RetainRangeArray As PPATCH_RETAIN_RANGE, ByVal SignatureBufferSize As Long, SignatureBuffer As Any) As Long
'Private Declare Function API_baseFiling_GetFilePatchSignatureByHandle Lib "MSPATCHA" (ByVal FileHandle As Long, ByVal OptionFlags As Long, OptionData As Any, ByVal IgnoreRangeCount As Long, ByRef IgnoreRangeArray As PPATCH_IGNORE_RANGE, ByVal RetainRangeCount As Long, ByRef RetainRangeArray As PPATCH_RETAIN_RANGE, ByVal SignatureBufferSize As Long, SignatureBuffer As Any) As Long
'Private Declare Function API_baseFiling_GetFileSecurity Lib "advapi32" Alias "GetFileSecurityA" (ByVal lpFileName As String, ByVal RequestedInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR, ByVal nLength As Long, lpnLengthNeeded As Long) As Long
Private Declare Function API_baseFiling_SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
'Private Declare Function API_baseFiling_SetFileSecurity Lib "advapi32" Alias "SetFileSecurityA" (ByVal lpFileName As String, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Private Declare Function API_baseFiling_SetFileTime Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, lpCreationTime As API_FILE_TIME, lpLastAccessTime As API_FILE_TIME, lpLastWriteTime As API_FILE_TIME) As Long

'Private Declare Function API_baseFiling_FileEncryptionStatus Lib "advapi32" Alias "FileEncryptionStatusA" (ByVal lpFileName As String, ByRef lpStatus As Long) As Long
'Private Declare Function API_baseFiling_FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, lpFatDate As Integer, lpFatTime As Integer) As Long
'Private Declare Function API_baseFiling_FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
'Private Declare Function API_baseFiling_FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
'Private Declare Function API_baseFiling_FileSaveRestoreOnINF Lib "advpack" (ByVal hWnd As Long, ByVal pszTitle As String, ByVal pszINF As String, ByVal pszSection As String, ByVal pszBackupDir As String, ByVal pszBaseBackupFile As String, ByVal dwFlags As Long) As Long
'Private Declare Function API_baseFiling_FileSaveRestore Lib "advpack" (ByVal hDlg As Long, ByVal lpFileList As String, ByVal lpDir As String, ByVal lpBaseName As String, ByVal dwFlags As Long) As Long
'Private Declare Function API_baseFiling_FileSaveMarkNotExist Lib "advpack" (ByVal lpFileList As String, ByVal lpDir As String, ByVal lpBaseName As String) As Long

Private Declare Function API_baseFiling_WriteFileOVR Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function API_baseFiling_WriteFileOVREx Lib "kernel32" Alias "WriteFileEx" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long

Private Declare Function API_baseFiling_WriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function API_baseFiling_WriteFileEx Lib "kernel32" Alias "WriteFileEx" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpOverlapped As Any, ByVal lpCompletionRoutine As Long) As Long

Private Declare Function API_baseFiling_FlushFileBuffers Lib "kernel32" Alias "FlushFileBuffers" (ByVal hFile As Long) As Long

Private Declare Function API_baseFiling_ReadFileOVR Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function API_baseFiling_ReadFileOVREx Lib "kernel32" Alias "ReadFileEx" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long

Private Declare Function API_baseFiling_ReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function API_baseFiling_ReadFileEx Lib "kernel32" Alias "ReadFileEx" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpOverlapped As Any, ByVal lpCompletionRoutine As Long) As Long
'Private Declare Function API_baseFiling_ReadFileScatter Lib "kernel32" Alias "ReadFileScatter" (ByVal hFile As Long, ByRef aSegmentArray As FILE_SEGMENT_ELEMENT, ByVal nNumberOfBytesToRead As Long, ByRef lpReserved As Long, ByRef lpOverlapped As OVERLAPPED) As Long


Public Enum API_FileMode
    API_fmAppend = OPEN_EXISTING
    API_fmCreate = CREATE_ALWAYS
    API_fmCreateNew = CREATE_NEW
    API_fmOpen = OPEN_ALWAYS
    API_fmOpenOrCreate = API_fmCreate Or API_fmOpen
    API_fmTruncate = TRUNCATE_EXISTING
End Enum
Public Enum API_FileAccess
    API_faRead = GENERIC_READ
    API_faWrite = GENERIC_WRITE
    API_faReadWrite = (GENERIC_READ Or GENERIC_WRITE)
    API_faExecute = GENERIC_EXECUTE
    API_faAll = GENERIC_ALL
End Enum
Public Enum API_FileShare
    API_fshNone = 0
    API_fshRead = 1
    API_fshWrite = 2
    API_fshReadWrite = 3
    API_fshDelete = 4
    API_fshInheritable = 16
End Enum
Public Enum API_FileAttributes
    API_fNormal = FILE_ATTRIBUTE_NORMAL
    API_fSystem = FILE_ATTRIBUTE_SYSTEM
    API_fReadOnly = FILE_ATTRIBUTE_READONLY
    API_fHidden = FILE_ATTRIBUTE_HIDDEN
    API_fDirectory = FILE_ATTRIBUTE_DIRECTORY
    API_fArchive = FILE_ATTRIBUTE_ARCHIVE
    API_fCompressed = FILE_ATTRIBUTE_COMPRESSED
    API_fEncrypted = FILE_ATTRIBUTE_ENCRYPTED
    API_fOffline = FILE_ATTRIBUTE_OFFLINE
    API_fDevice = FILE_ATTRIBUTE_DEVICE
    API_fTemporary = FILE_ATTRIBUTE_TEMPORARY
    API_fSparseFile = FILE_ATTRIBUTE_SPARSE_FILE
    API_fReparsePoint = FILE_ATTRIBUTE_REPARSE_POINT
    API_fContentIndexed = FILE_ATTRIBUTE_NOT_CONTENT_INDEXED
End Enum

Public Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type
Private Type API_OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Type API_LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
Private Type API_FILE_TIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type API_File_BigInt
    LowPart As Long
    HighPart As Long
End Type
Public Type BIGFILE_TIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type fOSFILE
    fHandle As Long
    Path As String
    Position As Long
    LLPosition As API_File_BigInt
    Length As Long
    LLLength As API_File_BigInt
    
    FileMode As API_FileMode
    FileAccess As API_FileAccess
    FileShare As API_FileShare
    FileAttributes As API_FileAttributes
    
    IsBigFile As Boolean
End Type

Public Type API_FILEBUFFER
    fBufferLength As Long
    fBuffer() As Byte
End Type

Dim inited As Boolean

Public Sub Initialize()
    If inited Then Exit Sub
    Call baseConstants.Initialize
    Call baseExceptions.Initialize
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    inited = False
End Sub

Public Sub EnsureFileExists(Path As String)
    Dim Created As Boolean
    Call MakeTreeDirectories(GetFilePath(Path), Created)
    Call CloseFile(CreateFile(Path, API_fmCreateNew, API_fNormal, API_faRead, API_fshNone))
End Sub

Private Function CheckFile(f As fOSFILE) As Boolean
    CheckFile = ((f.fHandle <> 0) And (f.fHandle <> INVALID_HANDLE_VALUE))
End Function

Public Function FileExists(Path As String) As Boolean
    FileExists = (Dir(Path, vbNormal) <> "")
End Function

Public Function CreateFile(Path As String, fOpenMode As API_FileMode, fAttributes As API_FileAttributes, fAccess As API_FileAccess, fShare As API_FileShare) As fOSFILE
    'HANDLE hFile = CreateFile(szFileName, GENERIC_READ, 0, NULL,
    '            OPEN_EXISTING,  0, NULL);
    Dim hFile As Long
    hFile = API_baseFiling_CreateFile(Path, fAccess, fShare, ByVal 0, fOpenMode, fAttributes, 0)
    If hFile = INVALID_HANDLE_VALUE Then throw InvalidFileException("Unable to create target file.")
    CreateFile.Length = API_baseFiling_GetFileSize(hFile, 0)
    If CreateFile.Length = INVALID_FILE_SIZE Then
        Dim fBS As API_File_BigInt
        If API_baseFiling_GetFileSizeEx(hFile, fBS) = 0 Then throw InvalidFileException("Invalid file size, maybe it's too small.")
        CreateFile.LLLength = fBS
        CreateFile.IsBigFile = True
    Else
        CreateFile.LLLength.LowPart = CreateFile.Length
        CreateFile.LLLength.HighPart = 0
        CreateFile.IsBigFile = False
    End If
    CreateFile.Path = Path
    CreateFile.fHandle = hFile
    CreateFile.Position = 0
    CreateFile.FileMode = fOpenMode
    CreateFile.FileAccess = fAccess
    CreateFile.FileAttributes = fAttributes
    CreateFile.FileShare = fShare
    
    If fOpenMode = API_fmAppend Then _
        Call SetEndOfFile(CreateFile)
End Function
Public Sub CloseFile(f As fOSFILE)
     If f.fHandle Then
        Call API_baseFiling_CloseHandle(f.fHandle)
        f.fHandle = 0
        f.Length = 0
        f.Path = ""
        f.Position = 0
        f.IsBigFile = False
        f.LLLength.HighPart = 0
        f.LLLength.LowPart = 0
        f.LLPosition.HighPart = 0
        f.LLPosition.LowPart = 0
        f.FileMode = 0
        f.FileAccess = 0
        f.FileAttributes = 0
        f.FileShare = 0
     End If
End Sub

Public Function GetFilePosition(f As fOSFILE) As Long
    If Not CheckFile(f) Then throw InvalidHandleException
    GetFilePosition = API_baseFiling_SetFilePointer(f.fHandle, 0, ByVal 0, FILE_CURRENT)
    
    If GetFilePosition = INVALID_SET_FILE_POINTER Then
        f.IsBigFile = True
        throw InvalidFileException("Invalid file position, maybe it's too large.")
    End If
    f.Position = GetFilePosition
End Function
Public Sub SetFilePosition(f As fOSFILE, Position As Long)
    If Not CheckFile(f) Then throw InvalidHandleException
    Dim sfpRetVal As Long
    sfpRetVal = API_baseFiling_SetFilePointer(f.fHandle, Position, ByVal 0, FILE_BEGIN)
    
    If sfpRetVal = INVALID_SET_FILE_POINTER Then _
        throw SystemCallFailureException("Unable to set file pointer.")
    f.Position = sfpRetVal
End Sub
Public Function GetFilePositionLL(f As fOSFILE) As API_File_BigInt
    If Not CheckFile(f) Then throw InvalidHandleException
    
End Function
Public Sub SetFilePositionLL(f As fOSFILE, Position As API_File_BigInt)
    If Not CheckFile(f) Then throw InvalidHandleException
    
End Sub

Public Sub TranslateFilePosition(f As fOSFILE, Position As Long)
    If Not CheckFile(f) Then throw InvalidHandleException
    Dim sfpRetVal As Long
    sfpRetVal = API_baseFiling_SetFilePointer(f.fHandle, Position, ByVal 0, FILE_CURRENT)
    
    If sfpRetVal = INVALID_SET_FILE_POINTER Then _
        throw SystemCallFailureException("Unable to set file pointer.")
    f.Position = sfpRetVal
End Sub

'    DWORD dwRC = ::SetFilePointer(hFile,        // handle to file
'                                              0,            // bytes to move pointer
'                                              NULL,         // bytes to move pointer
'                                              FILE_END);    // starting point
Public Sub SetEndOfFile(f As fOSFILE)
    If Not CheckFile(f) Then throw InvalidHandleException
    Dim sfpRetVal As Long
    sfpRetVal = API_baseFiling_SetFilePointer(f.fHandle, 0, ByVal 0, FILE_END)
    
    If sfpRetVal = INVALID_SET_FILE_POINTER Then _
        throw SystemCallFailureException("Unable to set file pointer.")
    f.Position = sfpRetVal
End Sub
Public Sub SetBeginOfFile(f As fOSFILE)
    If Not CheckFile(f) Then throw InvalidHandleException
    Dim sfpRetVal As Long
    sfpRetVal = API_baseFiling_SetFilePointer(f.fHandle, 0, ByVal 0, FILE_BEGIN)
    
    If sfpRetVal = INVALID_SET_FILE_POINTER Then _
        throw SystemCallFailureException("Unable to set file pointer.")
    f.Position = sfpRetVal
End Sub

Public Function IsEndOfFile(f As fOSFILE) As Boolean
    If Not CheckFile(f) Then throw InvalidHandleException
    If f.IsBigFile Then
        Dim bg1 As API_File_BigInt, bg2 As API_File_BigInt
        bg1 = GetFileLengthLL(f)
        bg2 = GetFilePositionLL(f)
        IsEndOfFile = ((bg1.LowPart = bg2.LowPart) And (bg1.HighPart = bg2.HighPart))
    Else
        IsEndOfFile = (GetFilePosition(f) = GetFileLength(f))
    End If
End Function

Public Function GetFileLength(f As fOSFILE) As Long
    If f.fHandle = 0 Or f.fHandle = INVALID_HANDLE_VALUE Then _
        throw InvalidHandleException
    GetFileLength = API_baseFiling_GetFileSize(f.fHandle, 0)
    If GetFileLength = INVALID_FILE_SIZE Then
        f.IsBigFile = True
        throw InvalidFileException("Invalid file size, maybe it's too large.")
    End If
    f.Length = GetFileLength
End Function
Public Sub SetFileLength(f As fOSFILE, Length As Long)
    If Not CheckFile(f) Then throw InvalidHandleException
    
End Sub
Public Function GetFileLengthLL(f As fOSFILE) As API_File_BigInt
    If Not CheckFile(f) Then throw InvalidHandleException
    Dim fBS As API_File_BigInt
    If API_baseFiling_GetFileSizeEx(f.fHandle, fBS) = 0 Then throw InvalidFileException("Invalid file size, maybe it's too small.")
    GetFileLengthLL = fBS
    f.LLLength = fBS
End Function
Public Sub SetFileLengthLL(f As fOSFILE, Length As API_File_BigInt)
    If Not CheckFile(f) Then throw InvalidHandleException
    
End Sub

Public Sub FlushFile(f As fOSFILE)
    If Not CheckFile(f) Then throw InvalidHandleException
    If API_baseFiling_FlushFileBuffers(f.fHandle) = 0 Then _
        throw SystemCallFailureException("Unable to flush file buffers.")
End Sub

Public Function GetStaticFileLength(Path As String) As Long
    
End Function

Public Function GetFileAttributes(f As fOSFILE) As API_FileAttributes
    If Not CheckFile(f) Then throw InvalidHandleException
    
End Function
Public Sub SetFileAttributes(f As fOSFILE, Attributes As API_FileAttributes)
    If Not CheckFile(f) Then throw InvalidHandleException
    
End Sub
Public Sub SetStaticFileAttributes(Path As String, Attributes As API_FileAttributes)
    
End Sub

Public Sub ReadByteArrayFromFile(f As fOSFILE, B() As Byte, Optional Length As Long = -1)
    If Not CheckFile(f) Then throw InvalidHandleException
    Dim NumberOfBytesRead As Long
    Dim ReadLength As Long
    
    If Length <> -1 Then
        ReadLength = Length
    Else
        ReadLength = API_baseFiling_GetFileSize(f.fHandle, 0)
        If ReadLength = INVALID_FILE_SIZE Then
            f.IsBigFile = True
            throw InvalidFileException("Invalid file size, maybe it's too large.")
        End If
        f.Length = ReadLength
        
        NumberOfBytesRead = API_baseFiling_SetFilePointer(f.fHandle, 0, ByVal 0, FILE_CURRENT)
        If NumberOfBytesRead = INVALID_SET_FILE_POINTER Then
            f.IsBigFile = True
            throw InvalidFileException("Invalid file position, maybe it's too large.")
        End If
        f.Position = NumberOfBytesRead
        
        ReadLength = ReadLength - NumberOfBytesRead
        NumberOfBytesRead = 0
    End If
    
    ReDim B(ReadLength - 1)
    
    If API_baseFiling_ReadFile(f.fHandle, B(0), ReadLength, NumberOfBytesRead, ByVal 0) = 0 Then _
        throw ReadFileException("Unable to read from file.")
    f.Position = f.Position + NumberOfBytesRead
End Sub
Public Sub WriteByteArrayToFile(f As fOSFILE, B() As Byte, Optional Length As Long = -1)
    If Not CheckFile(f) Then throw InvalidHandleException
    Dim outLen As Long, NumberOfBytesWritten As Long
    outLen = ArraySize(B)
    If outLen <= 0 Then Exit Sub
    If Length > -1 Then
        If Length = 0 Then Exit Sub
        If outLen < Length Then
            Dim sB() As Byte
            ReDim sB(Length - 1)
            Dim i As Long
            For i = 0 To outLen - 1
                sB(i) = B(i)
            Next
            If API_baseFiling_WriteFile(f.fHandle, sB(0), Length, NumberOfBytesWritten, ByVal 0) = 0 Then _
                throw WriteFileException("Unable to write in file.")
            f.Position = f.Position + NumberOfBytesWritten
            Exit Sub
        Else
            outLen = Length
        End If
    End If
    If API_baseFiling_WriteFile(f.fHandle, B(0), outLen, NumberOfBytesWritten, ByVal 0) = 0 Then _
        throw WriteFileException("Unable to write in file.")
    f.Position = f.Position + NumberOfBytesWritten
End Sub

Public Sub ReadRecordFromFile(f As fOSFILE, bRecord() As Byte, Index As Long, Optional RecordLength As Long = -1)
    If Not CheckFile(f) Then throw InvalidHandleException
    
End Sub
Public Sub WriteRecordToFile(f As fOSFILE, bRecord() As Byte, Index As Long, Optional RecordLength As Long = -1)
    If Not CheckFile(f) Then throw InvalidHandleException
    
End Sub
