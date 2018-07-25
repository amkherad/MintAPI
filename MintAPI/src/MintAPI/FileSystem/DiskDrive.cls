VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DiskDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="DiskDrive" tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
'Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "DiskDrive" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Private Const DRIVE_DOESNT_EXIST = 1
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6


Private Const FILE_CASE_SENSITIVE_SEARCH      As Long = &H1
Private Const FILE_CASE_PRESERVED_NAMES       As Long = &H2
Private Const FILE_UNICODE_ON_DISK            As Long = &H4
Private Const FILE_PERSISTENT_ACLS            As Long = &H8
Private Const FILE_FILE_COMPRESSION           As Long = &H10
Private Const FILE_VOLUME_QUOTAS              As Long = &H20
Private Const FILE_SUPPORTS_SPARSE_FILES      As Long = &H40
Private Const FILE_SUPPORTS_REPARSE_POINTS    As Long = &H80
Private Const FILE_SUPPORTS_REMOTE_STORAGE    As Long = &H100
Private Const FILE_VOLUME_IS_COMPRESSED       As Long = &H8000
Private Const FILE_SUPPORTS_OBJECT_IDS        As Long = &H10000
Private Const FILE_SUPPORTS_ENCRYPTION        As Long = &H20000


Private Const FS_CASE_IS_PRESERVED      As Long = FILE_CASE_PRESERVED_NAMES
Private Const FS_CASE_SENSITIVE         As Long = FILE_CASE_SENSITIVE_SEARCH
Private Const FS_UNICODE_STORED_ON_DISK As Long = FILE_UNICODE_ON_DISK
Private Const FS_PERSISTENT_ACLS        As Long = FILE_PERSISTENT_ACLS
Private Const FS_VOL_IS_COMPRESSED      As Long = FILE_VOLUME_IS_COMPRESSED
Private Const FS_FILE_COMPRESSION       As Long = FILE_FILE_COMPRESSION

Public Enum DiskDriveType
    ddtDoesntExist = DRIVE_DOESNT_EXIST
    ddtFixed = DRIVE_FIXED
    ddtCDRom = DRIVE_CDROM
    ddtRamDisk = DRIVE_RAMDISK
    ddtRemovable = DRIVE_REMOVABLE
    ddtRemote = DRIVE_REMOTE
    ddtFree = &H7F
End Enum
Public Enum FileSystemFlags
    fsfCaseIsReserved = FS_CASE_IS_PRESERVED                'The file system preserves the case of file names when it places a name on disk.
    fsfCaseSensitive = FS_CASE_SENSITIVE                    'The file system supports case-sensitive file names.
    fsfUnicodeStoredOnDisk = FS_UNICODE_STORED_ON_DISK      'The file system supports Unicode in file names as they appear on disk.
    fsfRersistentACLs = FS_PERSISTENT_ACLS                  'The file system preserves and enforces ACLs. For example, NTFS preserves and enforces ACLs, and FAT does not.
    fsfFileCompression = FS_FILE_COMPRESSION                'The file system supports file-based compression.
    fsfVolumeIsCompressed = FS_VOL_IS_COMPRESSED            'The specified volume is a compressed volume; for example, a DoubleSpace volume.
    'fsfNamedStreams = FILE_NAMED_STREAMS                    'The file system supports named streams.
    'fsfReadOnlyVolume = FILE_READ_ONLY_VOLUME               'Windows XP: The specified volume is read-only.
    fsfSupportsEncryption = FILE_SUPPORTS_ENCRYPTION        'The file system supports the Encrypted File System (EFS).
    fsfSupportsObjectIDs = FILE_SUPPORTS_OBJECT_IDS         'The file system supports object identifiers.
    fsfSupportsReparsePoints = FILE_SUPPORTS_REPARSE_POINTS 'The file system supports reparse points.
    fsfSupportsSparseFiles = FILE_SUPPORTS_SPARSE_FILES     'The file system supports sparse files.
    fsfVolumeQuotas = FILE_VOLUME_QUOTAS                    'The file system supports disk quotas.
End Enum
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables/Events>
'
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Dim p_Path As String
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Constructors>
'
''<summary>
''  Basic constructor to use by MintAPI class provider,
''  When creating new instance of this class.
''</summary>
''<params>
''  <param name="Arguments">
''      A list of construction arguments.
''  </param>
''</params>
''<cargs order="0" count="1">
''  <carg type="long" name="ID"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal RootPath As String, Optional ByVal Removable As Boolean)
    p_Path = DiskDrive.ValidateRootPath(RootPath)
    
    This_Constructed = True
End Sub
'Friend Sub ConstructorCopy(ByVal That As {My Object Type})
'
'    This_Constructed = True
'End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

Public Property Get RootPath() As String
THROW_IfNull
    RootPath = p_Path
End Property

Public Property Get VolumeName() As String
THROW_IfNull
    Dim sBuffer As String
    sBuffer = String(SMALLLPSTR, 0)
    'fix bad parameter values
    If API_GetVolumeInformation(RootPath, sBuffer, SMALLLPSTR, 0, 0, 0, vbNullString, vbNullPtr) = NO_VALUE Then _
        throw Exps.IfError
    
    VolumeName = API_SysAllocString(sBuffer)
End Property
Public Property Let VolumeName(ByVal Value As String)
THROW_IfNull
    If API_SetVolumeLabel(RootPath, Value) = NO_VALUE Then _
        throw Exps.IfError
End Property

Public Property Get DriveType() As DiskDriveType
    DriveType = API_GetDriveType(RootPath)
End Property

Public Property Get FileSystemFlags() As FileSystemFlags
THROW_IfNull
    If API_GetVolumeInformation(RootPath, vbNullString, NO_VALUE, 0, 0, FileSystemFlags, vbNullString, NO_VALUE) = NO_VALUE Then _
        throw Exps.IfError
End Property

Public Property Get SerialNumber() As Long
THROW_IfNull
    If API_GetVolumeInformation(RootPath, vbNullString, NO_VALUE, SerialNumber, 0, 0, vbNullString, vbNullPtr) = NO_VALUE Then _
        throw Exps.IfError
End Property

Public Property Get FileSystemName() As String
THROW_IfNull
    Dim sBuffer As String
    sBuffer = String(SMALLLPSTR, 0)
    'fix bad parameter values
    If API_GetVolumeInformation(RootPath, vbNullString, NO_VALUE, 0, 0, 0, sBuffer, SMALLLPSTR) = NO_VALUE Then _
        throw Exps.IfError
    
    FileSystemName = GetLPSTR(sBuffer)
End Property

Public Property Get TotalSize() As Long
    
End Property
Public Property Get TotalSize64() As Int64
    
End Property

Public Property Get EmptySize() As Long
    
End Property
Public Property Get EmptySize64() As Int64
    
End Property

Public Property Get Permissions()
    
End Property

Public Sub Eject()
    
End Sub

Public Sub Dispose()
    
End Sub

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Basic Members>
'
Public Function Equals(Value As Variant) As Boolean
    Equals = Objects.Equals(Me, Value)
End Function
Public Function ToString() As String
    ToString = CLASSALIAS
End Function
Public Function GetHashCode() As Long
    GetHashCode = Objects.GetHashCode(Me)
End Function
Public Function MetaObject() As MetaObject
    If myMetaObject Is Nothing Then
        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    End If
    Set MetaObject = myMetaObject
End Function

'Public Function Clone() As Object
'
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
'Private Sub Class_Initialize()
'    ...
'End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
    
    Call Dispose
End Sub

Private Sub THROW_IfNull()
    If Not This_Constructed Then throw Exps.ClassNotInitializedException
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo): Call Constructor(CInfo): End Sub
Private Function IObject_Equals(Value As Variant) As Boolean: IObject_Equals = Equals(Value): End Function
Private Function IObject_GetHashCode() As Long: IObject_GetHashCode = GetHashCode: End Function
Private Function IObject_MetaObject() As MetaObject: Set IObject_MetaObject = MetaObject: End Function
Private Function IObject_ToString() As String: IObject_ToString = ToString: End Function
Private Function IObject_Stub(ByVal ProxyInfo As ObjectProxyInfo) As Boolean
    
    IObject_Stub = False
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
'Private Function ICloneable_Clone() As Object
'    Set ICloneable_Clone = Clone
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
