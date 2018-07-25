VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "File1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "File is a advanced object to manage files and open or read/write to them.\r\n\r\nITargetStream<File>"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Directory"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"Path"


'Ideas from My earlier applications

Option Base 0
Option Explicit
Const CLASSID As String = "File"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Const SIGNAL_WRITE As String = "write"
Const SIGNAL_READ As String = "read"
Const SIGNAL_FLUSH As String = "flush"
Const SIGNAL_CLEAR As String = "clear"
Const SIGNAL_CHANGED As String = "changed"
Const SIGNAL_POSITIONCHANGED As String = "positionchanged"
Const SIGNAL_LENGTHCHANGED As String = "lengthchanged"
Const SIGNAL_MOVED As String = "moved"
Const SIGNAL_COPIED As String = "copied"
Const SIGNAL_DELETED As String = "deleted"
Const SIGNAL_CREATED As String = "created"
Const SIGNAL_OPEN As String = "open"
Const SIGNAL_CLOSE As String = "close"
Const SIGNAL_SEEK As String = "seek"
Const SIGNAL_RENAMED As String = "renamed"
Const SIGNAL_FILECHOOSED As String = "filechoosed"
Const SIGNAL_BUFFERSIZECHANGED As String = "buffersizechanged"

Const SLOT_CLEAR As String = "clear"
Const SLOT_ENDL As String = "endl"
Const SLOT_NEWLINE As String = "newline"
Const SLOT_FLUSH As String = "flush"
Const SLOT_DELETE As String = "delete"
Const SLOT_CLOSE As String = "close"
Const SLOT_ENDOFFILE As String = "endoffile"
Const SLOT_BEGINOFFILE As String = "beginoffile"

Public Event WriteData(E As EventArgs)
Public Event ReadData(E As EventArgs)
Public Event Flush(E As EventArgs)
Public Event Clear(E As EventArgs)
Public Event Changed(E As EventArgs)
Public Event PositionChanged(E As EventArgs)
Public Event LengthChanged(E As EventArgs)
Public Event Moved(E As EventArgs)
Public Event Copied(E As EventArgs)
Public Event Deleted(E As EventArgs)
Public Event Created(E As EventArgs)
Public Event FileClose(E As EventArgs)
Public Event FileOpen(E As EventArgs)
Public Event FileSeek(E As EventArgs)
Public Event Renamed(E As EventArgs)
Public Event FileChoosed(E As EventArgs)
Public Event BufferSizeChanged(E As EventArgs)

'Implements IData
'Implements ITargetStream
'Implements ISupportSignal
'Implements ISupportSlot
Implements IDevice
'Implements IAliasable


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

Private Enum API_FileMode
    API_fmAppend = OPEN_EXISTING
    API_fmCreate = CREATE_ALWAYS
    API_fmCreateNew = CREATE_NEW
    API_fmOpen = OPEN_ALWAYS
    API_fmOpenOrCreate = API_fmCreate Or API_fmOpen
    API_fmTruncate = TRUNCATE_EXISTING
End Enum
Private Enum API_FileAccess
    API_faRead = GENERIC_READ
    API_faWrite = GENERIC_WRITE
    API_faReadWrite = (GENERIC_READ Or GENERIC_WRITE)
    API_faExecute = GENERIC_EXECUTE
    API_faAll = GENERIC_ALL
End Enum
Private Enum API_FileShare
    API_fshNone = 0
    API_fshRead = 1
    API_fshWrite = 2
    API_fshReadWrite = 3
    API_fshDelete = 4
    API_fshInheritable = 16
End Enum
Private Enum API_FileAttributes
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
Private Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    Offset As Long
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

Private Type API_File_BigInt
    LowPart As Long
    HighPart As Long
End Type
Private Type BIGFILE_TIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type fOSFILE
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
Private Type API_FILEBUFFER
    fBufferLength As Long
    fBuffer() As Byte
End Type

Public Enum FileOpenModes
    fomRead = API_faRead
    fomWrite = API_faWrite
    fomExecute = API_faExecute
    fomAccessAll = API_faAll
    fomReadWrite = API_faReadWrite

    fomOpen = API_fmOpen
    fomCreate = API_fmCreate
    fomDeleteExistFile = API_fmCreateNew
    fomCreateNew = API_fmCreateNew
    fomAppend = API_fmAppend
    fomAppendCreate = fomAppend Or fomCreate
    fomTruncate = API_fmTruncate

    fomOpenCreate = fomReadWrite Or fomCreate
    fomOpenCreateNew = fomReadWrite Or fomCreateNew
    fomWriteCreate = fomWrite Or fomCreate
    fomWriteCreateNew = fomWrite Or fomCreateNew

    fomSafeControl = fomOpen Or fomCreate
    fomFullControl = fomOpen
    fomSafeControlNew = fomOpen Or fomCreateNew
    fomDefault = fomOpen
End Enum
Public Enum FileMode
    fmAppend = API_fmAppend
    fmCreate = API_fmCreate
    fmCreateNew = API_fmCreateNew
    fmOpen = API_fmOpen
    fmOpenOrCreate = API_fmOpenOrCreate
    fmTruncate = API_fmTruncate
End Enum
Public Enum FileAccess
    faRead = API_faRead
    faWrite = API_faWrite
    faReadWrite = API_faReadWrite
    faExecute = API_faExecute
    faAll = API_faAll
End Enum
Public Enum FileShare
    fshNone = API_fshNone
    fshRead = API_fshRead
    fshWrite = API_fshWrite
    fshReadWrite = API_fshReadWrite
    fshDelete = API_fshDelete
    fshInheritable = API_fshInheritable
End Enum
Public Enum FileAttributes
    fNormal = API_fNormal
    fSystem = API_fSystem
    fReadOnly = API_fReadOnly
    fHidden = API_fHidden
    fDirectory = API_fDirectory
    fArchive = API_fArchive
    fCompressed = API_fCompressed
    fEncrypted = API_fEncrypted
    fOffline = API_fOffline
    fDevice = API_fDevice
    fTemporary = API_fTemporary
    fSparseFile = API_fSparseFile
    fReparsePoint = API_fReparsePoint
    fContentIndexed = API_fContentIndexed
End Enum
Public Enum FileDialogTypes
    fdWindowsStyle
    fdNativeStyle
    fdShellStyle
    fdOpen
    fdSave
End Enum

Const CURRENT_POSITION As Long = -1

Dim p As String
Dim fFile As fOSFILE

Dim f_IsOpened As Boolean

Const DEFAULTBUFFERSIZE As Long = 2048
Const SECTORSIZE As Long = 512

Dim buffSize As Long
Dim Buff() As Byte
Dim buffStartIndex As Long
Dim buffEndIndex As Long

Dim fOpenMode As FileOpenModes

Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String

Dim mySignalEmitter As New SignalEmitter


Private Sub EnsureFileExists(Path As String)
    Dim Created As Boolean
    Call MakeTreeDirectories(GetFilePath(Path), Created)
    Call baseCloseFile(CreateFile(Path, API_fmCreateNew, API_fNormal, API_faRead, API_fshNone))
End Sub

Private Function CheckFile(f As fOSFILE) As Boolean
    CheckFile = ((f.fHandle <> 0) And (f.fHandle <> INVALID_HANDLE_VALUE))
End Function

Private Function FileExists(Path As String) As Boolean
    FileExists = (Dir(Path, vbNormal) <> "")
End Function

Private Function CreateFile(Path As String, fOpenMode As API_FileMode, fAttributes As API_FileAttributes, fAccess As API_FileAccess, fShare As API_FileShare) As fOSFILE
    'HANDLE hFile = CreateFile(szFileName, GENERIC_READ, 0, NULL,
    '            OPEN_EXISTING,  0, NULL);
    Dim hFile As Long
    hFile = API_baseFiling_CreateFile(Path, fAccess, fShare, ByVal 0, fOpenMode, fAttributes, 0)
    If hFile = INVALID_HANDLE_VALUE Then throw Exps.InvalidFileException("Unable to create target file.")
    CreateFile.Length = API_baseFiling_GetFileSize(hFile, 0)
    If CreateFile.Length = INVALID_FILE_SIZE Then
        Dim fBS As API_File_BigInt
        If API_baseFiling_GetFileSizeEx(hFile, fBS) = 0 Then throw Exps.InvalidFileException("Invalid file size, maybe it's too small.")
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
        Call baseSetEndOfFile(CreateFile)
End Function
Private Sub baseCloseFile(f As fOSFILE)
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

Private Function GetFilePosition(f As fOSFILE) As Long
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    GetFilePosition = API_baseFiling_SetFilePointer(f.fHandle, 0, ByVal 0, FILE_CURRENT)
    
    If GetFilePosition = INVALID_SET_FILE_POINTER Then
        f.IsBigFile = True
        throw Exps.InvalidFileException("Invalid file position, maybe it's too large.")
    End If
    f.Position = GetFilePosition
End Function
Private Sub SetFilePosition(f As fOSFILE, Position As Long)
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    Dim sfpRetVal As Long
    sfpRetVal = API_baseFiling_SetFilePointer(f.fHandle, Position, ByVal 0, FILE_BEGIN)
    
    If sfpRetVal = INVALID_SET_FILE_POINTER Then _
        throw Exps.SystemCallFailureException("Unable to set file pointer.")
    f.Position = sfpRetVal
End Sub
Private Function GetFilePositionLL(f As fOSFILE) As API_File_BigInt
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    
End Function
Private Sub SetFilePositionLL(f As fOSFILE, Position As API_File_BigInt)
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    
End Sub

Private Sub TranslateFilePosition(f As fOSFILE, Position As Long)
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    Dim sfpRetVal As Long
    sfpRetVal = API_baseFiling_SetFilePointer(f.fHandle, Position, ByVal 0, FILE_CURRENT)
    
    If sfpRetVal = INVALID_SET_FILE_POINTER Then _
        throw Exps.SystemCallFailureException("Unable to set file pointer.")
    f.Position = sfpRetVal
End Sub

'    DWORD dwRC = ::SetFilePointer(hFile,        // handle to file
'                                              0,            // bytes to move pointer
'                                              NULL,         // bytes to move pointer
'                                              FILE_END);    // starting point
Private Sub baseSetEndOfFile(f As fOSFILE)
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    Dim sfpRetVal As Long
    sfpRetVal = API_baseFiling_SetFilePointer(f.fHandle, 0, ByVal 0, FILE_END)
    
    If sfpRetVal = INVALID_SET_FILE_POINTER Then _
        throw Exps.SystemCallFailureException("Unable to set file pointer.")
    f.Position = sfpRetVal
End Sub
Private Sub baseSetBeginOfFile(f As fOSFILE)
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    Dim sfpRetVal As Long
    sfpRetVal = API_baseFiling_SetFilePointer(f.fHandle, 0, ByVal 0, FILE_BEGIN)
    
    If sfpRetVal = INVALID_SET_FILE_POINTER Then _
        throw Exps.SystemCallFailureException("Unable to set file pointer.")
    f.Position = sfpRetVal
End Sub

Private Function baseIsEndOfFile(f As fOSFILE) As Boolean
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    If f.IsBigFile Then
        Dim bg1 As API_File_BigInt, bg2 As API_File_BigInt
        bg1 = GetFileLengthLL(f)
        bg2 = GetFilePositionLL(f)
        baseIsEndOfFile = ((bg1.LowPart = bg2.LowPart) And (bg1.HighPart = bg2.HighPart))
    Else
        baseIsEndOfFile = (GetFilePosition(f) = GetFileLength(f))
    End If
End Function

Private Function GetFileLength(f As fOSFILE) As Long
    If f.fHandle = 0 Or f.fHandle = INVALID_HANDLE_VALUE Then _
        throw Exps.InvalidHandleException
    GetFileLength = API_baseFiling_GetFileSize(f.fHandle, 0)
    If GetFileLength = INVALID_FILE_SIZE Then
        f.IsBigFile = True
        throw Exps.InvalidFileException("Invalid file size, maybe it's too large.")
    End If
    f.Length = GetFileLength
End Function
Private Sub SetFileLength(f As fOSFILE, Length As Long)
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    
End Sub
Private Function GetFileLengthLL(f As fOSFILE) As API_File_BigInt
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    Dim fBS As API_File_BigInt
    If API_baseFiling_GetFileSizeEx(f.fHandle, fBS) = 0 Then throw Exps.InvalidFileException("Invalid file size, maybe it's too small.")
    GetFileLengthLL = fBS
    f.LLLength = fBS
End Function
Private Sub SetFileLengthLL(f As fOSFILE, Length As API_File_BigInt)
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    
End Sub

Private Sub FlushFile(f As fOSFILE)
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    If API_baseFiling_FlushFileBuffers(f.fHandle) = 0 Then _
        throw Exps.SystemCallFailureException("Unable to flush file buffers.")
End Sub

Private Function GetStaticFileLength(Path As String) As Long
    
End Function

Private Function GetFileAttributes(f As fOSFILE) As API_FileAttributes
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    
End Function
Private Sub SetFileAttributes(f As fOSFILE, Attributes As API_FileAttributes)
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    
End Sub
Private Sub SetStaticFileAttributes(Path As String, Attributes As API_FileAttributes)
    
End Sub

Private Sub ReadByteArrayFromFile(f As fOSFILE, B() As Byte, Optional Length As Long = -1)
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
    Dim NumberOfBytesRead As Long
    Dim ReadLength As Long
    
    If Length <> -1 Then
        ReadLength = Length
    Else
        ReadLength = API_baseFiling_GetFileSize(f.fHandle, 0)
        If ReadLength = INVALID_FILE_SIZE Then
            f.IsBigFile = True
            throw Exps.InvalidFileException("Invalid file size, maybe it's too large.")
        End If
        f.Length = ReadLength
        
        NumberOfBytesRead = API_baseFiling_SetFilePointer(f.fHandle, 0, ByVal 0, FILE_CURRENT)
        If NumberOfBytesRead = INVALID_SET_FILE_POINTER Then
            f.IsBigFile = True
            throw Exps.InvalidFileException("Invalid file position, maybe it's too large.")
        End If
        f.Position = NumberOfBytesRead
        
        ReadLength = ReadLength - NumberOfBytesRead
        NumberOfBytesRead = 0
    End If
    
    ReDim B(ReadLength - 1)
    
    If API_baseFiling_ReadFile(f.fHandle, B(0), ReadLength, NumberOfBytesRead, ByVal 0) = 0 Then _
        throw Exps.ReadFileException("Unable to read from file.")
    f.Position = f.Position + NumberOfBytesRead
End Sub
Private Sub WriteByteArrayToFile(f As fOSFILE, B() As Byte, Optional Length As Long = -1)
    If Not CheckFile(f) Then throw Exps.InvalidHandleException
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
                throw Exps.WriteFileException("Unable to write in file.")
            f.Position = f.Position + NumberOfBytesWritten
            Exit Sub
        Else
            outLen = Length
        End If
    End If
    If API_baseFiling_WriteFile(f.fHandle, B(0), outLen, NumberOfBytesWritten, ByVal 0) = 0 Then _
        throw Exps.WriteFileException("Unable to write in file.")
    f.Position = f.Position + NumberOfBytesWritten
End Sub




'======================================
'======================================
'======================================
'======================================
'======================================
'======================================
'======================================
'======================================
'======================================
'======================================

Private Sub Class_Initialize()
    'Call mySignalEmitter.Initialize(Me)
    inState = True
    outState = False
End Sub
Private Sub Class_Terminate()
    Call CloseFile
End Sub

Friend Sub Initialize(Optional Path As String = "")
    Call SetPath(Path)
End Sub
Friend Sub CreateCustomPlace(Place As String)

End Sub

Friend Sub SetOSHandle(pfFile As fOSFILE)
    fFile = pfFile
End Sub

Public Sub SetPath(Path As String)
    If IsOpen Then throw Exps.InvalidStatusException("File already opened and cannot change the path.")
    p = Path
End Sub
Private Function me_SelectPath(Optional OptionalPath) As String
    If IsOpen Then
        me_SelectPath = fFile.Path
    Else
        me_SelectPath = SelectString(p, OptionalPath)
    End If
End Function

Public Function VirtualPath(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidPathException, Me
    If Not CheckPathValidation(VL, True) Then throw Exps.InvalidPathException, Me
    VirtualPath = VL
End Function
Public Function AbsolutePath(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidPathException, Me
    If Not CheckPathValidation(VL, True, True) Then throw Exps.InvalidPathException, Me
    AbsolutePath = VL
End Function

Public Function Location(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidPathException, Me
    Location = GetFilePath(VL)
End Function
Public Function Path(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidPathException, Me
    Path = VL
End Function
Public Function Name(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidPathException, Me
    Name = GetFileName(VL)
End Function
Public Function NameOnly(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidPathException, Me
    NameOnly = GetFileNameOnly(VL)
End Function
Public Function Extention(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidArgumentValueException, Me
    Extention = GetFileExtention(VL)
End Function

Public Function OpenNewFile(Optional fOpenMode As FileMode = fmOpen, Optional fAttributes As FileAttributes = fNormal, Optional fAccess As FileAccess = faRead, Optional fShare As FileShare = fshNone, Optional OptionalPath) As File
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    Dim f As New File
    Call f.OpenFile(fOpenMode, fAttributes, fAccess, fShare, OptionalPath)
    Set OpenNewFile = f
End Function
Public Function OpenFile(Optional fOpenMode As FileMode = fmOpen, Optional fAttributes As FileAttributes = fNormal, Optional fAccess As FileAccess = faRead, Optional fShare As FileShare = fshNone, Optional OptionalPath) As File
    If IsOpen Then _
        throw Exps.OpenFileException("File already opened.")

    Dim VL As String, Created As Boolean
    VL = me_SelectPath(OptionalPath)

    If Not (((fOpenMode And fmCreate) = fmCreate) Or fOpenMode = fmCreateNew) Then
        If Not FileExists(VL) Then
            throw Exps.FileNotFoundException
        End If
    Else
        If Not FileExists(VL) Then
            Created = True
        End If
    End If

    fFile = CreateFile(VL, fOpenMode, fAttributes, fAccess, fShare)
    If fFile.fHandle = 0 Then throw Exps.SystemCallFailureException
    p = fFile.Path
    f_IsOpened = True
    Set OpenFile = Me

    If Created Then
        'Call mySignalEmitter.Emit(SIGNAL_CREATED)
        RaiseEvent Created(EventArgs(Me))
    End If
    'Call mySignalEmitter.Emit(SIGNAL_OPEN)
    RaiseEvent FileOpen(EventArgs(Me))
End Function
Public Function Create(Optional fOpenMode As FileMode = fmCreate, Optional fAttributes As FileAttributes = fNormal, Optional fAccess As FileAccess = faRead, Optional fShare As FileShare = fshNone, Optional OptionalPath) As File
    Dim VL As String, Created As Boolean
    VL = me_SelectPath(OptionalPath)
    If IsOpen Then
        If fFile.Path = VL Then
            throw Exps.OpenFileException("File already opened.")
        Else
            Set Create = New File
            Call Create.Create(fOpenMode, fAttributes, fAccess, fShare, OptionalPath)
            Exit Function
        End If
    End If

    If Not FileExists(VL) Then _
        Created = True

    fFile = CreateFile(VL, fOpenMode, fAttributes, fAccess, fShare)
    If fFile.fHandle = 0 Then throw Exps.SystemCallFailureException
    p = fFile.Path
    f_IsOpened = True
    Set Create = Me

    If Created Then
        'Call mySignalEmitter.Emit(SIGNAL_CREATED)
        RaiseEvent Created(EventArgs(Me))
    End If
    'Call mySignalEmitter.Emit(SIGNAL_OPEN)
    RaiseEvent FileOpen(EventArgs(Me))
End Function
Public Function EnsureOpen(Optional OptionalPath) As File
    If IsOpen Then
        Set EnsureOpen = Me
        Exit Function
    End If
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    Call OpenFile(fmOpenOrCreate, fNormal, faReadWrite, fshNone)
    Set EnsureOpen = Me
End Function
Public Function TryOpen(Optional fOpenMode As FileMode = fmOpen, Optional fAttributes As FileAttributes = fNormal, Optional fAccess As FileAccess = faRead, Optional fShare As FileShare = fshNone, Optional OptionalPath) As Boolean
On Error GoTo Err
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    Call OpenFile(fOpenMode, fAccess, fAttributes, fShare, OptionalPath)
    TryOpen = True
    Exit Function
Err:
    TryOpen = False
End Function
Public Sub CloseFile()
    If Not IsOpen Then Exit Sub 'throw Exps.InvalidStatusException("File not opened to close.")
    Call baseCloseFile(fFile)
    'Call mySignalEmitter.Emit(SIGNAL_CLOSE)
    RaiseEvent FileClose(EventArgs(Me))
End Sub
Public Sub Flush()
    If Not IsOpen Then throw Exps.InvalidStatusException("File not opened.")
    Call FlushFile(fFile)
    'Call mySignalEmitter.Emit(SIGNAL_FLUSH)
    RaiseEvent Flush(EventArgs(Me))
End Sub
Friend Sub Clear()

    'Call mySignalEmitter.Emit(SIGNAL_CLEAR)
    RaiseEvent Clear(EventArgs(Me))
End Sub

Public Function Exists(Optional OptionalPath) As Boolean
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then
        Exists = False
    Else
        Exists = FileExists(VL)
    End If
End Function
Public Function EnsureExists(Optional OptionalPath) As File
    If IsOpen Then EnsureExists = True: Exit Function
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    Call EnsureFileExists(VL)
    Set EnsureExists = Me
End Function
Public Sub Delete(Optional OptionalPath)
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If Not Exists(VL) Then throw Exps.InvalidPathException
    Call Kill(VL)
    'Call mySignalEmitter.Emit(SIGNAL_DELETED)
    RaiseEvent Deleted(EventArgs(Me))
End Sub
Public Function Directory(Optional OptionalPath) As Directory
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidPathException, Me
    Set Directory = mint_constructor.Directory(GetFilePath(VL))
End Function
Public Sub Rename(ByVal NewName As String, Optional OptionalPath)
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidPathException, Me

    'Call mySignalEmitter.Emit(SIGNAL_RENAMED)
    RaiseEvent Renamed(EventArgs(Me))
End Sub
Public Sub Move(ByVal NewPath As String, Optional OptionalPath)
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidPathException, Me

    'Call mySignalEmitter.Emit(SIGNAL_MOVED)
    RaiseEvent Moved(EventArgs(Me))
End Sub
Public Sub ConcurrentMove(ByVal NewPath As String, Optional OptionalPath)
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidPathException, Me

    'Call mySignalEmitter.Emit(SIGNAL_MOVED)
    RaiseEvent Moved(EventArgs(Me))
End Sub
Public Sub Copy(ByVal NewPath As String, Optional OptionalPath)
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw Exps.InvalidPathException, Me

    'Call mySignalEmitter.Emit(SIGNAL_COPIED)
    RaiseEvent Copied(EventArgs(Me))
End Sub
'Public Function ConcurrentCopy(ByVal NewPath As String, Optional OptionalPath) As Thread
'    Dim VL As String
'    VL = me_SelectPath(OptionalPath)
'    If VL = "" Then throw Exps.InvalidPathException, Me
'
'    Call mySignalEmitter.Emit(SIGNAL_COPIED)
'    RaiseEvent Copied(EventArgs(Me))
'End Function

Public Function ChooseFile(Optional OwnerForm As Long, Optional DialogType As FileDialogTypes = fdOpen, Optional Title As String, Optional Filters As String, Optional StaticFileName As String, Optional ByVal InitialDirectory As Directory, Optional CheckPathExists As Boolean = False, Optional CheckFileExists As Boolean = False, Optional OpenAsReadOnly As Boolean = False, Optional AllowChangingDirectory As Boolean = True, Optional Flags As Long) As File
    Dim S As String, c_path As String, AllowMultiSelect As Boolean
    AllowMultiSelect = False
    If InitialDirectory Is Nothing Then
        c_path = IIf(fFile.Path = "", CurDir$, Path(fFile.Path))
    Else
        c_path = InitialDirectory.AbsolutePath
    End If
    Dim RetVal As API_FileDialogReturn
    RetVal = shellMethods.Dialogs_BrowseFile(IIf((DialogType And fdOpen) = fdOpen, OpenDialog, SaveDialog), OwnerForm, Title, c_path, Filters, StaticFileName, Flags, CheckPathExists, CheckFileExists, AllowMultiSelect, OpenAsReadOnly, AllowChangingDirectory)
    'mint_api_dialogs_last_choose_file_read_only_flag_state = retVal.ReadOnlyCheckState
    Dim f As New File
    Call f.Initialize(RetVal.Path)
    Set ChooseFile = f
    'Call mySignalEmitter.Emit(SIGNAL_FILECHOOSED)
    RaiseEvent FileChoosed(EventArgs(Me))
End Function
Public Function ChooseFiles(Optional OwnerForm As Long, Optional DialogType As FileDialogTypes = fdOpen, Optional Title As String, Optional Filters As String, Optional ByVal InitialDirectory As Directory, Optional CheckPathExists As Boolean = False, Optional CheckFileExists As Boolean = False, Optional OpenAsReadOnly As Boolean = False, Optional AllowChangingDirectory As Boolean = True, Optional Flags As Long) As File()
    Dim S As String, c_path As String, AllowMultiSelect As Boolean
    AllowMultiSelect = True
    If InitialDirectory Is Nothing Then
        c_path = IIf(fFile.Path = "", CurDir$, Path(fFile.Path))
    Else
        c_path = InitialDirectory.AbsolutePath
    End If
    Dim RetVal As API_FileDialogReturn
    RetVal = shellMethods.Dialogs_BrowseFile(IIf((DialogType And fdOpen) = fdOpen, OpenDialog, SaveDialog), OwnerForm, Title, c_path, Filters, "", Flags, CheckPathExists, CheckFileExists, AllowMultiSelect, OpenAsReadOnly, AllowChangingDirectory)
    'mint_api_dialogs_last_choose_file_read_only_flag_state = retVal.ReadOnlyCheckState
    Dim fls() As File
    Dim flsCount As Long
    Dim i As Long
    ReDim fls(RetVal.subFilesCount - 1)
    For i = 0 To RetVal.subFilesCount - 1
        Set fls(i) = New File
        Call fls(i).Initialize(baseMethods.ConcatPath(RetVal.Path, RetVal.subFiles(i)))
    Next
    ChooseFiles = fls
    'Call mySignalEmitter.Emit(SIGNAL_FILECHOOSED)
    RaiseEvent FileChoosed(EventArgs(Me))
End Function

Public Property Get IsOpen() As Boolean
    IsOpen = ((f_IsOpened) And (fFile.fHandle <> 0))
End Property
Public Property Get Writable() As Boolean
    If Not IsOpen Then Exit Property
    If (fFile.FileAccess = API_faAll) Or ((fFile.FileAccess And API_faWrite) = API_faWrite) Then
        Writable = True
    End If
End Property
Public Property Get Readable() As Boolean
    If Not IsOpen Then Exit Property
    If (fFile.FileAccess = API_faAll) Or ((fFile.FileAccess And API_faRead) = API_faRead) Then
        Readable = True
    End If
End Property

Public Property Get ReadOnly() As Boolean
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    ReadOnly = ((fFile.FileAttributes And API_fReadOnly) = API_fReadOnly)
End Property
Public Property Let ReadOnly(Value As Boolean)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")

End Property

Public Property Get BufferSize() As Long
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")

End Property
Public Property Let BufferSize(Value As Long)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")

End Property


Public Function constData() As Byte()
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Dim retB() As Byte
    Call baseSetBeginOfFile(fFile)
    Call ReadByteArrayFromFile(fFile, retB)
    constData = retB
    'Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
End Function
Public Function readByteArray(Optional Length As Long = -1, Optional StartAt) As Byte()
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Dim retB() As Byte
    If Not IsMissing(StartAt) Then _
        Call SeekFile(CLng(StartAt))
    Call ReadByteArrayFromFile(fFile, retB, Length)
    readByteArray = retB
    'Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
End Function
Public Sub FillData(Data)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")

End Sub

Public Sub WriteData(Data, Optional Length As Long = -1)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Dim DataBA() As Byte
    DataBA = mint_get_byte_array_of(Data, Length)
    If ArraySize(DataBA) <= 0 Then Exit Sub
    Call WriteByteArrayToFile(fFile, DataBA, Length)
    'Call mySignalEmitter.Emit(SIGNAL_WRITE)
    RaiseEvent WriteData(EventArgs(Me))
    'Call mySignalEmitter.Emit(SIGNAL_CHANGED)
    RaiseEvent Changed(EventArgs(Me))
    'Call mySignalEmitter.Emit(SIGNAL_LENGTHCHANGED)
    RaiseEvent LengthChanged(EventArgs(Me))
    'Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Sub
'Public Function ReadData(Data)
'    Call ReadVar(Data)
'End Function
Public Sub ReadData(Data, Optional Length)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")

    Dim VT As VbVarType
    Dim i As Long, retBA As ByteArray
    'Call mySignalEmitter.BlockSignals
    Call Err.Clear
    On Error GoTo iErr

    VT = VarType(Data)

    Select Case VT
        Case VbVarType.vbString
            If Not IsMissing(Length) Then
                Set retBA = ReadBlock(CLng(Length))
            Else
                Set retBA = ReadBlock
            End If
            Data = retBA.ToString
        Case VbVarType.vbBoolean
            Call EnsureForwardLength(2)
            Set retBA = ReadBlock(2)
            Data = retBA.ToBoolean
        Case VbVarType.vbByte
            Set retBA = ReadBlock(1)
            Data = retBA.ByteAt(0)
        Case VbVarType.vbInteger
            Call EnsureForwardLength(2)
            Set retBA = ReadBlock(2)
            Data = retBA.ToInt
        Case VbVarType.vbLong
            Call EnsureForwardLength(4)
            Set retBA = ReadBlock(4)
            Data = retBA.ToLong
        Case VbVarType.vbDouble
            Call EnsureForwardLength(8)
            Set retBA = ReadBlock(8)
            Data = retBA.ToDouble
        Case VbVarType.vbSingle
            Call EnsureForwardLength(4)
            Set retBA = ReadBlock(4)
            Data = retBA.ToSingle
        Case VbVarType.vbCurrency
            Call EnsureForwardLength(8)
            Set retBA = ReadBlock(8)
            Data = retBA.ToCurrency
        Case VbVarType.vbDate
            Call EnsureForwardLength(8)
            Set retBA = ReadBlock(8)
            Data = retBA.ToDate
        Case VbVarType.VBObject
            Dim Obj As Object
            Set Obj = Data
            If TypeOf Obj Is BigNumber Then
                Set retBA = ReadBlock
                Set Data = retBA.ToBigNumber
            End If
        Case Else
            If (VT And vbArray) = vbArray Then
                If (VT And vbByte) = vbByte Then
                    i = ArraySize(Data)
                    If i > 0 Then
                        Set retBA = ReadBlock(i)
                        Data = retBA.constData
                    Else
                        Set retBA = ReadBlock
                        Data = retBA.constData
                    End If
                Else
                    For i = 0 To ArraySize(Data) - 1
                        Call ReadData(Data(i))
                    Next
                End If
            Else
                throw Exps.UnknownValueException("Var is unknown.")
            End If
    End Select

iErr:
    'Call mySignalEmitter.UnblockSignals
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    'Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
    'Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Sub
Public Sub WriteLine(Data)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Dim DataBA() As Byte
    DataBA = mint_get_byte_array_of(Data)
    If ArraySize(DataBA) <= 0 Then Exit Sub
    Call WriteByteArrayToFile(fFile, DataBA)
    Call WriteByteArrayToFile(fFile, StringToByteArray(vbCrLf))
    'Call mySignalEmitter.Emit(SIGNAL_WRITE)
    RaiseEvent WriteData(EventArgs(Me))
    'Call mySignalEmitter.Emit(SIGNAL_CHANGED)
    RaiseEvent Changed(EventArgs(Me))
    'Call mySignalEmitter.Emit(SIGNAL_LENGTHCHANGED)
    RaiseEvent LengthChanged(EventArgs(Me))
    'Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Sub
Public Function ReadLine() As ByteArray
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Dim retBA As ByteArray, retBA2 As ByteArray, relIndex As Long
    Dim crlfCharBA As ByteArray, indexOfCRLF As Long
    Dim fPos As Long
    Set crlfCharBA = ByteArray(vbCrLf)

    fPos = Position

    Set retBA = TryReadBlock(SECTORSIZE)
    indexOfCRLF = retBA.IndexOf(crlfCharBA, 0)
    If indexOfCRLF >= 0 Then
        Set ReadLine = retBA.SubArray(0, indexOfCRLF)
        relIndex = 2
    Else
        While (indexOfCRLF = -1 And (Not baseIsEndOfFile(fFile)))
            Set retBA2 = TryReadBlock(SECTORSIZE, fPos + retBA.Length)
            indexOfCRLF = retBA2.IndexOf(crlfCharBA)
            If indexOfCRLF = -1 Then
                Call retBA.Append(retBA2)
            Else
                Call retBA.Append(retBA2.SubArray(0, indexOfCRLF))
                relIndex = 2
            End If
        Wend
        Set ReadLine = retBA
    End If
    Position = fPos + ReadLine.Length + relIndex

    'Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
End Function
'Public Function ReadLines() As ByteArray()
'    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
'
'End Function
Public Function TryReadBlock(Optional Length As Long = -1, Optional StartAt) As ByteArray
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    If Not IsMissing(StartAt) Then Call Me.SeekFile(CLng(StartAt))
    Dim Buff() As Byte
    Dim ReadLength As Long

    ReadLength = ForwardLength

    If Length <> -1 Then _
        If Length <= ReadLength Then _
            ReadLength = Length

    Call ReadByteArrayFromFile(fFile, Buff(), ReadLength)
    Set TryReadBlock = New ByteArray
    Call TryReadBlock.SetBuffer(Buff)
    If TryReadBlock.IsEmpty Then Exit Function
    'Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
    'Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Function
Public Function ReadBlock(Optional Length As Long = -1, Optional StartAt) As ByteArray
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    If Not IsMissing(StartAt) Then Call Me.SeekFile(CLng(StartAt))
    Dim Buff() As Byte
    Call ReadByteArrayFromFile(fFile, Buff(), Length)
    Set ReadBlock = New ByteArray
    Call ReadBlock.SetBuffer(Buff)
    If ReadBlock.IsEmpty Then Exit Function
    'Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
    'Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Function
Public Sub WriteRecord(Data, Optional Index)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Dim DataBA() As Byte
    DataBA = mint_get_byte_array_of(Data)
    'Call SetFilePosition(fFile, Index * ArraySize(DataBA))
End Sub
Public Function ReadRecord(Data, Optional Index) As ByteArray
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")

End Function
Public Sub WriteFormatted(Data)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")

End Sub
Public Function ReadFormatted(Data) As ByteArray
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")

End Function
Public Sub PrintF(Template As String, ParamArray Args() As Variant)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Dim RArgs() As Variant
    RArgs = Args
    Call WriteData(funcwArgs(Template, RArgs))
End Sub
Public Function ScanF(Template As String, ParamArray Args() As Variant) As ByteArray
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")

End Function

Public Property Get Position() As Long
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Position = GetFilePosition(fFile)
End Property
Public Property Let Position(Value As Long)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Dim cPos As Long
    cPos = fFile.Position
    Call SetFilePosition(fFile, Value)
    If Value <> cPos Then
        'Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
        RaiseEvent PositionChanged(EventArgs(Me))
    End If
End Property
'Public Property Get LargePosition() As BigNumber
'    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
'    Dim fBP As API_File_BigInt
'    fBP = GetFilePositionLL(fFile)
'    Set LargePosition = BigNumber(Array_Long(fBP.LowPart, fBP.HighPart))
'End Property
'Public Property Let LargePosition(Value As BigNumber)
'    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
'
'
'End Property
Public Sub SeekFile(Position As Long)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Dim cPos As Long
    cPos = fFile.Position
    Call SetFilePosition(fFile, Position)
    If Position <> cPos Then
        Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
        RaiseEvent PositionChanged(EventArgs(Me))
    End If
End Sub
Public Sub RelativeSeekFile(RelativePosition As Long)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Dim cPos As Long
    cPos = fFile.Position
    Call TranslateFilePosition(fFile, Position)
    If RelativePosition <> cPos Then
        Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
        RaiseEvent PositionChanged(EventArgs(Me))
    End If
End Sub

Public Sub SetEndOfFile()
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Call baseSetEndOfFile(fFile)
    Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Sub
Public Sub SetBeginOfFile()
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Call baseSetBeginOfFile(fFile)
    Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Sub


Public Property Get Length() As Long
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Length = GetFileLength(fFile)
End Property
Public Property Let Length(Value As Long)
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Call SetFileLength(fFile, Value)
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
    RaiseEvent Changed(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_LENGTHCHANGED)
    RaiseEvent LengthChanged(EventArgs(Me))
End Property
'Public Property Get LargeLength() As BigNumber
'    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
'    Dim fBS As API_File_BigInt
'    fBS = GetFileLengthLL(fFile)
'    Set LargeLength = BigNumber(Array_Long(fBS.LowPart, fBS.HighPart))
'End Property
'Public Property Let LargeLength(Value As BigNumber)
'    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
'
'End Property

Public Property Get ForwardLength() As Long
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    Dim Length As Long
    Length = GetFileLength(fFile)
    Length = Length - GetFilePosition(fFile)
    ForwardLength = Length
End Property


Public Property Get IsEndOfFile() As Boolean
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    IsEndOfFile = baseIsEndOfFile(fFile)
End Property
Public Property Get IsBeginOfFile() As Boolean
    If Not IsOpen Then throw Exps.InvalidStatusException("File is not opened.")
    IsBeginOfFile = (GetFilePosition(fFile) = 0)
End Property
Public Property Get IsMintAPICustomPlace() As Boolean

End Property

Public Function Available(Length As Long) As Boolean
    Available = (ForwardLength >= Length)
End Function
Public Sub EnsureForwardLength(Length As Long)
    If (ForwardLength < Length) Then _
        throw Exps.InvalidStatusException("Not enough data available.")
End Sub

Public Function ReadAllText(Optional OptionalPath) As String
    If IsOpen Then
        Call baseSetBeginOfFile(fFile)
        ReadAllText = ReadBlock.ToString
    ElseIf Not IsMissing(OptionalPath) Then
        ReadAllText = File(me_SelectPath(OptionalPath)).OpenFile.ReadAllText
    Else
        throw Exps.InvalidPathException
    End If
End Function
Public Function ReadAllStringLines(Optional OptionalPath) As String()
    If IsOpen Then
        Call baseSetBeginOfFile(fFile)
        ReadAllStringLines = Split(ReadBlock.ToString, vbCrLf)
    ElseIf Not IsMissing(OptionalPath) Then
        ReadAllStringLines = File(me_SelectPath(OptionalPath)).ReadAllStringLines
    Else
        throw Exps.InvalidPathException
    End If
End Function
Public Function ReadAllLines(Optional OptionalPath) As ByteArray()
    If IsOpen Then
        Call baseSetBeginOfFile(fFile)
        ReadAllLines = ReadBlock.Split(vbCrLf)
    ElseIf Not IsMissing(OptionalPath) Then
        ReadAllLines = File(me_SelectPath(OptionalPath)).OpenFile.ReadAllLines
    Else
        throw Exps.InvalidPathException
    End If
End Function
Public Function ReadToEnd(Optional OptionalPath) As ByteArray
    If IsOpen Then
        Call baseSetBeginOfFile(fFile)
        Set ReadToEnd = ReadBlock
    ElseIf Not IsMissing(OptionalPath) Then
        Set ReadToEnd = File(me_SelectPath(OptionalPath)).OpenFile.ReadToEnd
    Else
        throw Exps.InvalidPathException
    End If
End Function

Public Sub WriteAllText(Value As String, Optional OptionalPath)
    If IsOpen Then
        Call baseSetBeginOfFile(fFile)
        Call WriteData(Value)
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmCreate).WriteAllText(Value)
    Else
        throw Exps.InvalidPathException
    End If
End Sub
Public Sub WriteAllLines(Value() As String, Optional OptionalPath)
    If IsOpen Then
        Call baseSetBeginOfFile(fFile)
        Dim i As Long
        For i = 0 To ArraySize(Value) - 1
            Call WriteLine(Value(i))
        Next
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmCreate).WriteAllLines(Value)
    Else
        throw Exps.InvalidPathException
    End If
End Sub
Public Sub WriteByteArray(ByVal Value As ByteArray, Optional OptionalPath)
    If IsOpen Then
        Call baseSetBeginOfFile(fFile)
        Call WriteData(Value)
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmCreate).WriteByteArray(Value)
    Else
        throw Exps.InvalidPathException
    End If
End Sub

Public Sub AppendAllText(Value As String, Optional OptionalPath)
    If IsOpen Then
        Call baseSetEndOfFile(fFile)
        Call WriteData(Value)
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmAppend, fNormal, faWrite, fshRead).AppendAllText(Value)
    Else
        throw Exps.InvalidPathException
    End If
End Sub
Public Sub AppendAllLines(Value() As String, Optional OptionalPath)
    If IsOpen Then
        Call baseSetEndOfFile(fFile)
        Dim i As Long
        For i = 0 To ArraySize(Value) - 1
            Call WriteLine(Value(i))
        Next
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmAppend, fNormal, faWrite, fshRead).AppendAllLines(Value)
    Else
        throw Exps.InvalidPathException
    End If
End Sub
Public Sub AppendByteArray(Value As ByteArray, Optional OptionalPath)
    If IsOpen Then
        Call baseSetEndOfFile(fFile)
        Call WriteData(Value)
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmAppend, fNormal, faWrite, fshRead).AppendByteArray(Value)
    Else
        throw Exps.InvalidPathException
    End If
End Sub

Public Function Clone() As File
    Dim f As New File
    Call f.SetOSHandle(fFile)
    Set Clone = f
End Function


Private Sub IData_Append(Target As Variant)
    Call AppendByteArray(ByteArray(Target))
End Sub
Private Sub IData_Clear()
    Call Clear
End Sub
Private Sub IData_CopyByteArray(targetByteArray() As Byte)
    targetByteArray = Me.constData
End Sub
Private Function IData_GetByteArray() As Byte()
    IData_GetByteArray = Me.constData
End Function
Private Property Get IData_Length() As Long
    IData_Length = Length
End Property
Private Sub IData_SetByteArray(B() As Byte)
    Call FillData(B)
End Sub

Private Property Get IDevice_DeviceType() As DeviceType
    IDevice_DeviceType = dtDevice_File
End Property
Private Property Get IDevice_Handle() As Long
    IDevice_Handle = fFile.fHandle
End Property

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter.ExternEmitter
End Function

Private Property Get ISupportSlot_AutoManage() As Boolean
    ISupportSlot_AutoManage = False
End Property
Private Sub ISupportSlot_ManualManageCall(SlotName As String)
    Select Case SlotName
        Case SLOT_FLUSH
            Call Flush
        Case SLOT_BEGINOFFILE
            Me.Position = 0
        Case SLOT_ENDOFFILE
            Me.Position = Me.Length
        Case SLOT_CLEAR
            Call Me.Clear
        Case SLOT_CLOSE
            Call Me.CloseFile
        Case SLOT_ENDL, SLOT_NEWLINE
            Call Me.WriteData(vbCrLf)
        Case Else
            throw Exps.InvalidCallException
    End Select
End Sub
Private Function ISupportSlot_Methods() As String
    ISupportSlot_Methods = _
        "0" & SLOT_CLOSE & _
        " 0" & SLOT_FLUSH & _
        " 0" & SLOT_BEGINOFFILE & _
        " 0" & SLOT_ENDOFFILE & _
        " 0" & SLOT_CLEAR & _
        " 0" & SLOT_ENDL & _
        " 0" & SLOT_NEWLINE
End Function

Private Function ITargetStream_Available(Dir As StreamDirection) As Long
    Dim Val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        Val = ((Length > 0) And inState)
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        Val = Val Or (outState)
    End If
    ITargetStream_Available = Val
End Function

Private Sub ITargetStream_CloseStream(Optional Direction As StreamDirection = 3&)
    If (Direction And sdInStream) = sdInStream Then
        inState = False
    End If
    If (Direction And sdOutStream) = sdOutStream Then
        outState = False
    End If
    If Not (inState Or outState) Then Call CloseFile
End Sub

Private Sub ITargetStream_OpenStream(Optional Direction As StreamDirection = StreamDirection.sdBoth)
    If (Direction And sdInStream) = sdInStream Then
        inState = True
    End If
    If (Direction And sdOutStream) = sdOutStream Then
        outState = True
    End If
End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean
    Dim Val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        Val = inState
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        Val = Val Or outState
    End If
    ITargetStream_getState = Val
End Function
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    ITargetStream_readBytes = Me.readByteArray(Length, StartAt)
End Function

Private Sub ITargetStream_outStream(Data, Optional Length As Long = -1)
    If Not outState Then throw Exps.InvalidStatusException
    If IsObject(Data) Then
        If TypeOf Data Is ObjectBuffer Then
            Dim objBuffer As ObjectBuffer
            Set objBuffer = Data
            If objBuffer.IsStreamingValue Then
                Call objBuffer.thrownotargs(1)
                If objBuffer.ArgType(0) = vbString Then
                    Dim strFirst As String
                    strFirst = objBuffer.First
                    If strFirst = OBJECTBUFFER_VALUES_STREAMING_CLEAR Then
                        Call Me.Clear
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_NEWLINE Then
                        Call Me.WriteData(vbCrLf)
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_FLUSH Then
                        Call Me.Flush
                        Exit Sub
                    Else
                        throw Exps.InvalidArgumentTypeException("Streaming Command '" & strFirst & "' Is Not Available For File."), Me
                    End If
                Else
                    throw Exps.InvalidArgumentTypeException("Streaming Type Is Not Available For File.")
                End If
            End If
        End If
    End If
    Call Me.WriteData(Data)
End Sub
Private Sub ITargetStream_inStream(Data, Optional Length As Long = -1)
    If Not inState Then throw Exps.InvalidStatusException
    If IsObject(Data) Then
        If TypeOf Data Is ObjectBuffer Then
            Dim objBuffer As ObjectBuffer
            Set objBuffer = Data
            If objBuffer.IsStreamingValue Then
                Call objBuffer.thrownotargs(1)
                If objBuffer.ArgType(0) = vbString Then
                    Dim strFirst As String
                    strFirst = objBuffer.First
                    If strFirst = OBJECTBUFFER_VALUES_STREAMING_CLEAR Then
                        Call Clear
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_NEWLINE Then
                        Call WriteData(vbCrLf)
                        Exit Sub
                    ElseIf strFirst = OBJECTBUFFER_VALUES_STREAMING_FLUSH Then
                        'Do nothing...
                        Exit Sub
                    Else
                        throw Exps.InvalidArgumentTypeException("Streaming Command '" & strFirst & "' Is Not Available For ByteArray.")
                    End If
                Else
                    throw Exps.InvalidArgumentTypeException("Streaming Type Is Not Available For ByteArray.")
                End If
            Else
                throw Exps.InvalidArgumentTypeException
            End If
        Else
            throw Exps.InvalidArgumentTypeException
        End If
        throw Exps.InvalidArgumentTypeException
    End If
    Call ReadData(Data)
End Sub
Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved): Call mint_setstream_state(State, inState, inState_LOCK, Reserved): End Sub
Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved): Call mint_setstream_state(State, outState, outState_LOCK, Reserved): End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property


