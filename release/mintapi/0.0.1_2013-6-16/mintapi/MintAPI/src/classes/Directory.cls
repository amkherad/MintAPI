VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Directory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Strong directory managing class."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "Directory"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IClassClonable
Implements IClassTexer
Implements IAliasable

Private Type SPCPath
    Name As String
    Path As String
End Type

Dim p As String

Friend Sub Initialize(Optional ByVal Path As String = "")
    Call SetPath(Path)
End Sub

Public Function SetPath(Path As String) As Directory
    p = Path
    If p = "" Then p = CurDir$
    Set SetPath = Me
End Function

Public Function CurrentDirectory() As Directory
    Dim d As New Directory
    Call d.Initialize(FileSystem.CurDir)
    Set CurrentDirectory = d
End Function
Public Function SetCurrentDirectory(ByVal Dir As Directory) As Directory
    If Dir Is Nothing Then throw ArgumentNullException
    Call FileSystem.ChDir(Dir.AbsolutePath)
    Set SetCurrentDirectory = Me
End Function
Public Function SetCurrentDirectoryS(ByVal Path As String) As Directory
    Call FileSystem.ChDir(Path)
    Set SetCurrentDirectoryS = Me
End Function

Public Function VirtualPath(Optional OptionalPath) As String
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    If VL = "" Then VL = CurDir$ 'current path
    If Not CheckPathValidation(VL, True, False) Then throw InvalidPathException, Me
    VirtualPath = VL
End Function
Public Function AbsolutePath(Optional OptionalPath) As String
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    If VL = "" Then VL = CurDir$ 'current path
    If Not CheckPathValidation(VL, True, True) Then throw InvalidPathException, Me
    AbsolutePath = VL
End Function

Public Function Validate(Optional CheckForDrive As Boolean = True, Optional OptionalPath) As Boolean
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    Validate = CheckPathValidation(VL, False, CheckForDrive)
End Function

Public Function Name(Optional OptionalPath) As String
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    If VL = "" Then VL = CurDir$
    Name = GetFileName(VL)
End Function
Public Function Location(Optional OptionalPath) As String
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    If VL = "" Then VL = CurDir$
    If Not CheckPathValidation(VL, True, True) Then throw InvalidPathException, Me
    Location = GetFilePath(VL)
End Function
Public Function Path(Optional OptionalPath) As String
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    If VL = "" Then VL = CurDir$
    If Not CheckPathValidation(VL, True) Then throw InvalidPathException, Me
    Path = VL
End Function
Public Function Exists(Optional OptionalPath) As Boolean
On Error GoTo Err
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    If VL = "" Then
        Exists = False
    Else
        Exists = DirectoryExists(VL)
    End If
    Exit Function
Err:
    Exists = False
End Function
Public Function Create(Optional OptionalPath) As Directory
    Dim VL As String
    Set Create = Me
    If IsMissing(OptionalPath) Then
        VL = p
    Else
        VL = OptionalPath
        Dim d As New Directory
        Call d.Initialize(VL)
        Set Create = d
    End If
    Dim Created As Boolean
    Call MakeTreeDirectories(VL, Created)
End Function
Public Function Remove(Optional OptionalPath) As Directory
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    'DONT SET CURRENT DIRECTORY.
    Call FileSystem.RmDir(VL)
    Set Remove = Me
End Function
Public Function RemoveTree(Optional steps As Long = -1) As Directory

    Set RemoveTree = Me
End Function
Public Function cd(ByVal Path As String) As Directory
    '. .. [~ / \]
    If Not CheckPathValidation(Path, False, False) Then throw InvalidPathException, Me
    Set cd = Me
    If Path = "~" Then
        p = CurDir$
        Exit Function
    ElseIf Path = "/" Or Path = "\" Then
        p = Drive
        Exit Function
    End If
    p = ConcatPath(p, Path)
    If Not CheckPathValidation(p, False, False) Then throw InvalidPathException, Me
End Function
Public Function cdup() As Directory
    Dim VL As String
    Call cd("..")
    Set cdup = Me
End Function
Public Function Drive(Optional OptionalPath) As String
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    If VL = "" Then throw ClassNotInitializedException
    If Not (Left(VL, 2) Like "?:") Then throw InvalidPathException, Me
    Drive = Left(VL, 1) & ":\"
End Function
Public Sub SetDrive(ByVal DriveLetter As String)
    Dim ln As Long
    ln = Len(DriveLetter)
    If ln > 1 Then
        throw InvalidPathException: Exit Sub
    ElseIf ln <= 0 Then
        DriveLetter = Directory(GetSystemPath).Drive
    End If
    If Left(p, 2) Like "?:" Then
        Mid(p, 1, 1) = DriveLetter
    Else
        p = baseMethods.ConcatPath(p, DriveLetter)
        Call CheckPathValidation(p, True)
    End If
End Sub
Public Function Rename(ByVal NewName As String, Optional OptionalPath) As Directory

    Set Rename = Me
End Function
Public Function Move(ByVal NewPath As String, Optional OptionalPath) As Directory

    Set Move = Me
End Function
Public Function Copy(ByVal NewPath As String, Optional OptionalPath) As Directory

    Set Copy = Me
End Function

Public Function SubDirectories(Optional Attributes As VbFileAttribute = VbFileAttribute.vbNormal, Optional ByVal IncludePath As Boolean = False, Optional OptionalPath) As String()
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    Dim CP As String, Count As Long, retVal() As String
    If Not (Right(VL, 1) = "/" Or Right(VL, 1) = "\") Then VL = VL & "\"
    CP = Dir(VL, vbDirectory Or Attributes)
    If CP <> "" Then CP = Dir
    If CP <> "" Then CP = Dir
    Dim cFile As String
    While CP <> ""
        cFile = baseMethods.ConcatPath(VL, CP)
        On Local Error GoTo cantOpenFile
        If (GetAttr(cFile) And vbDirectory) = vbDirectory Then
            On Error GoTo 0
            ReDim Preserve retVal(Count)
            If IncludePath Then
                retVal(Count) = cFile
            Else
                retVal(Count) = CP
            End If
            Count = Count + 1
        End If
cantOpenFile:
        CP = Dir
    Wend
    SubDirectories = retVal
End Function
Public Function Files(Optional Attributes As VbFileAttribute = VbFileAttribute.vbNormal, Optional ByVal IncludePath As Boolean = False, Optional OptionalPath) As File()
    Dim fls() As String
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    fls = FileNames(Attributes, IncludePath, VL)
    Dim tas As Long
    tas = ArraySize(fls)
    Dim flsC() As File
    ReDim flsC(tas - 1)
    Dim i As Long
    For i = 0 To tas - 1
        Set flsC(i) = New File
        Call flsC(i).Initialize(fls(i))
    Next
    Call API_CopyMemory(ByVal API_VarPtrArray(Files), ByVal API_VarPtrArray(flsC), 4)
End Function
Public Function FileNames(Optional Attributes As VbFileAttribute = VbFileAttribute.vbNormal, Optional ByVal IncludePath As Boolean = False, Optional OptionalPath) As String()
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    Dim CP As String, Count As Long, retVal() As String
    If Not (Right(VL, 1) = "/" Or Right(VL, 1) = "\") Then VL = VL & "\"
    CP = Dir(baseMethods.ConcatPath(VL, "", cpeEndsWithSlash), vbNormal Or Attributes)
    While CP <> ""
        ReDim Preserve retVal(Count)
        If IncludePath Then
            retVal(Count) = baseMethods.ConcatPath(VL, CP)
        Else
            retVal(Count) = CP
        End If
        Count = Count + 1
        CP = Dir
    Wend
    FileNames = retVal
End Function

Public Function FilteredSubDirectories(Filter As GlobalFilters, Optional Attributes As VbFileAttribute = VbFileAttribute.vbNormal, Optional ByVal IncludePath As Boolean = False, Optional OptionalPath) As String()
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    Dim CP As String, Count As Long, retVal() As String
    If Not (Right(VL, 1) = "/" Or Right(VL, 1) = "\") Then VL = VL & "\"
    CP = Dir(VL, vbDirectory Or Attributes)
    If CP <> "" Then CP = Dir 'drop .
    If CP <> "" Then CP = Dir 'drop ..
    Dim cFile As String
    While CP <> ""
        cFile = baseMethods.ConcatPath(VL, CP)
        On Local Error GoTo cantOpenFile
        If (GetAttr(cFile) And vbDirectory) = vbDirectory Then
            On Error GoTo 0
            If IsFilterIncluded(Filter, CP) Then
                ReDim Preserve retVal(Count)
                If IncludePath Then
                    retVal(Count) = cFile
                Else
                    retVal(Count) = CP
                End If
                Count = Count + 1
            End If
        End If
cantOpenFile:
        CP = Dir
    Wend
    FilteredSubDirectories = retVal
End Function
Public Function FilteredFiles(Filter As GlobalFilters, Optional Attributes As VbFileAttribute = VbFileAttribute.vbNormal, Optional ByVal IncludePath As Boolean = False, Optional OptionalPath) As File()
    Dim fls() As String
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    fls = FilteredFileNames(Filter, Attributes, IncludePath, VL)
    Dim tas As Long
    tas = ArraySize(fls)
    Dim flsC() As File
    ReDim flsC(tas - 1)
    Dim i As Long
    For i = 0 To tas - 1
        Set flsC(i) = New File
        Call flsC(i).Initialize(fls(i))
    Next
    Call API_CopyMemory(ByVal API_VarPtrArray(FilteredFiles), ByVal API_VarPtrArray(flsC), 4)
End Function
Public Function FilteredFileNames(Filter As GlobalFilters, Optional Attributes As VbFileAttribute = VbFileAttribute.vbNormal, Optional ByVal IncludePath As Boolean = False, Optional OptionalPath) As String()
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    Dim CP As String, Count As Long, retVal() As String
    If Not (Right(VL, 1) = "/" Or Right(VL, 1) = "\") Then VL = VL & "\"
    CP = Dir(baseMethods.ConcatPath(VL, "", cpeEndsWithSlash), vbNormal Or Attributes)
    While CP <> ""
        If IsFilterIncluded(Filter, CP) Then
            ReDim Preserve retVal(Count)
            If IncludePath Then
                retVal(Count) = baseMethods.ConcatPath(VL, CP)
            Else
                retVal(Count) = CP
            End If
            Count = Count + 1
        End If
        CP = Dir
    Wend
    FilteredFileNames = retVal
End Function


Public Function Attributes(Optional OptionalPath) As VbFileAttribute
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    Attributes = FileSystem.GetAttr(VL)
End Function
Public Sub SetAttributes(Attrs As VbFileAttribute, Optional OptionalPath)
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    Call FileSystem.SetAttr(VL, Attrs Or vbDirectory)
End Sub

Public Function ConcatPath(ByVal Path1 As String, Optional ByVal OptionalPath2 As String) As String
    Dim VL As String
    VL = SelectString(p, OptionalPath2)
    ConcatPath = baseMethods.ConcatPath(Path1, VL)
End Function

Public Property Get IsSystem(Optional OptionalPath) As Boolean
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    IsSystem = ((FileSystem.GetAttr(VL) And vbSystem) = vbSystem)
End Property
Public Property Let IsSystem(Optional OptionalPath, Value As Boolean)
    Dim VL As String, dAttrs As VbFileAttribute
    VL = SelectString(p, OptionalPath)
    dAttrs = FileSystem.GetAttr(VL)
    If ((dAttrs And vbSystem) <> vbSystem) Then Call FileSystem.SetAttr(VL, dAttrs Or vbSystem)
End Property
Public Property Get IsReadOnly(Optional OptionalPath) As Boolean
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    IsSystem = ((FileSystem.GetAttr(VL) And vbReadOnly) = vbReadOnly)
End Property
Public Property Let IsReadOnly(Optional OptionalPath, Value As Boolean)
    Dim VL As String, dAttrs As VbFileAttribute
    VL = SelectString(p, OptionalPath)
    dAttrs = FileSystem.GetAttr(VL)
    If ((dAttrs And vbReadOnly) <> vbReadOnly) Then Call FileSystem.SetAttr(VL, dAttrs Or vbReadOnly)
End Property
Public Property Get IsArchive(Optional OptionalPath) As Boolean
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    IsSystem = ((FileSystem.GetAttr(VL) And vbArchive) = vbArchive)
End Property
Public Property Let IsArchive(Optional OptionalPath, Value As Boolean)
    Dim VL As String, dAttrs As VbFileAttribute
    VL = SelectString(p, OptionalPath)
    dAttrs = FileSystem.GetAttr(VL)
    If ((dAttrs And vbArchive) <> vbArchive) Then Call FileSystem.SetAttr(VL, dAttrs Or vbArchive)
End Property
Public Property Get IsHidden(Optional OptionalPath As String) As Boolean
    Dim VL As String
    VL = SelectString(p, OptionalPath)
    IsSystem = ((FileSystem.GetAttr(VL) And vbHidden) = vbHidden)
End Property
Public Property Let IsHidden(Optional OptionalPath As String, Value As Boolean)
    Dim VL As String, dAttrs As VbFileAttribute
    VL = SelectString(p, OptionalPath)
    dAttrs = FileSystem.GetAttr(VL)
    If ((dAttrs And vbHidden) <> vbHidden) Then Call FileSystem.SetAttr(VL, dAttrs Or vbHidden)
End Property
'InitialDirectory As Directory,
Public Function ChooseDirectory(Optional ByVal OwnerForm As Long = 0, Optional Title As String = "", _
            Optional ByVal RootDir As SpecialFolders = SpecialFolders.sfALL, _
            Optional ByVal CustomRootDir As Directory = Nothing, _
            Optional ByVal StartDir As Directory = Nothing, _
            Optional NewStyle As Boolean = True, _
            Optional IncludeFiles As Boolean = False, _
            Optional ShowNewFolderButton As Boolean = True, _
            Optional OkButtonText As String = "") As Directory
    Dim s As String
    Dim custom_Dir As String, start_Dir As String
    If CustomRootDir Is Nothing Then
        custom_Dir = IIf(p = "", CurDir$, p)
    Else
        custom_Dir = CustomRootDir.AbsolutePath
    End If
    If Not StartDir Is Nothing Then start_Dir = StartDir.AbsolutePath
    s = Dialogs_Browse(OwnerForm, Title, RootDir, custom_Dir, start_Dir, NewStyle, IncludeFiles, ShowNewFolderButton, OkButtonText)
    Dim d As New Directory
    Call d.Initialize(s)
    Set ChooseDirectory = d
End Function

Public Function GetMintAPISharedPath(Optional CreateSpecifiedDirectory As Boolean = False, Optional ByVal ApplicationName As String = "default") As String
    Dim retVal As String
    retVal = baseMethods2.GetAppDataPath(True)
    retVal = ConcatPath(retVal, "MintAPI")                                          'user/MintAPI
    retVal = ConcatPath(retVal, "shared")                                           'user/MintAPI/shared
    If CreateSpecifiedDirectory Then retVal = ConcatPath(retVal, ApplicationName)   'user/MintAPI/shared/APPLICATIONNAME
    GetMintAPISharedPath = retVal
End Function
Public Function GetDesktopPath(Optional ByVal System As Boolean = False) As String
    GetDesktopPath = shellMethods.GetSpecialfolder(IIf(System, sf_Common_Desktop, sf_Desktop))
End Function
Public Function GetLocalUserPath(Optional ByVal System As Boolean = False) As String
    GetLocalUserPath = baseMethods2.GetLocalUserPath
End Function
Public Function GetApplicationSpecifiedTempPath(Optional ByVal System As Boolean = False, Optional UseRevision As Boolean = True) As String
    If tApplication Is Nothing Then throw InvalidStatusException("Unspecified target application.")
    GetApplicationSpecifiedTempPath = baseMethods2.GetApplicationSpecifiedTempPath_specified(tApplication.CompanyName, tApplication.ProductName, tApplication.VersionMajor, tApplication.VersionMinor, tApplication.VersionRevision, UseRevision)
End Function
Public Function GetApplicationDataPath(Optional ByVal System As Boolean = False, Optional UseRevision As Boolean = True) As String
    If tApplication Is Nothing Then throw InvalidStatusException("Unspecified target application.")
    GetApplicationDataPath = baseMethods2.GetApplicationDataPath_specified(tApplication.CompanyName, tApplication.ProductName, tApplication.VersionMajor, tApplication.VersionMinor, tApplication.VersionRevision, System, UseRevision)
End Function
Public Function GetMyDocumentsPath(Optional ByVal System As Boolean = False) As String
    GetMyDocumentsPath = shellMethods.GetSpecialfolder(IIf(System, sf_Common_Documents, sf_User_MyDocuments))
End Function
Public Function GetMyMusicPath(Optional ByVal System As Boolean = False) As String
    GetMyMusicPath = shellMethods.GetSpecialfolder(IIf(System, sf_Common_MyMusic, sf_MyMusic))
End Function
Public Function GetMyPicturesPath(Optional ByVal System As Boolean = False) As String
    GetMyPicturesPath = shellMethods.GetSpecialfolder(IIf(System, sf_Common_MyPictures, sf_MyPictures))
End Function
Public Function GetMyVideosPath(Optional ByVal System As Boolean = False) As String
    GetMyVideosPath = shellMethods.GetSpecialfolder(IIf(System, sf_Common_MyVideos, sf_MyVideos))
End Function
Public Function GetSharedDocumentPath() As String
    GetSharedDocumentPath = shellMethods.GetSpecialfolder(sf_Common_Documents)
End Function
Public Function GetUserPath() As String
    GetUserPath = baseMethods.GetUserPath
End Function
Public Function GetSystemDrive() As String
    GetSystemDrive = Drive(baseMethods.GetWindowsPath)
End Function
Public Function GetWindowsPath() As String
    GetWindowsPath = baseMethods.GetWindowsPath
End Function
Public Function GetSystemPath() As String
    Dim d As New Directory, p As String
    p = baseMethods.GetSystemPath
    Call d.Initialize(p) 'C:\Windows\System32
    Call d.cd("..\System")
    If d.Exists() Then
        'throw PathNotFoundException("Unable to find 'System' directory.")
        GetSystemPath = d.AbsolutePath
    Else
        GetSystemPath = p
    End If
End Function
Public Function GetSystem32Path() As String
    Dim d As New Directory, p As String
    p = baseMethods.GetSystemPath
    Call d.Initialize(p) 'C:\Windows\System32
    Call d.cd("..\System32")
    If d.Exists() Then
        'throw PathNotFoundException("Unable to find 'System' directory.")
        GetSystem32Path = d.AbsolutePath
    Else
        GetSystem32Path = p
    End If
End Function
Public Function GetSystem64Path() As String
    Dim d As New Directory, p As String
    p = baseMethods.GetSystemPath
    Call d.Initialize(p)   'C:\Windows\System32
    Call d.cd("..\SysWow64")
    If d.Exists() Then
        'throw PathNotFoundException("Unable to find 'SysWow64' directory.")
        GetSystem64Path = d.AbsolutePath
    Else
        GetSystem64Path = p
    End If
End Function
Public Function GetProgramFilesPath(Optional ByVal is_x86 As Boolean = True) As String
    If is_x86 Then
        GetProgramFilesPath = GetProgramFilesx86Path
    Else
        GetProgramFilesPath = GetProgramFilesx64Path
    End If
End Function
Public Function GetProgramFilesx86Path() As String
    Dim p As String
    p = shellMethods.GetSpecialfolder(sf_Program_Files)
    If kernelMethods.IsWow64OperatingSystem Then
        Dim d As New Directory
        Call d.Initialize(p)
        Call d.cd("../Program Files (x86)")
        If d.Exists Then _
            p = d.AbsolutePath
    End If
    GetProgramFilesx86Path = p
End Function
Public Function GetProgramFilesx64Path() As String
    GetProgramFilesx64Path = shellMethods.GetSpecialfolder(sf_Program_Files)
End Function

Private Function IClassClonable_Clone() As Object
    Dim d As New Directory
    Call d.Initialize(p)
    Set IClassClonable_Clone = d
End Function

Private Sub IClassExceptionManager_setAllExceptionsState(Optional States As Boolean = True)

End Sub

Private Sub IClassExceptionManager_setExceptionState(exp As ExceptionType, Optional State As Boolean = True)

End Sub

Private Sub IClassTexer_fromByteArray(Value() As Byte)

End Sub

Private Sub IClassTexer_fromString(Value As String)
    
End Sub
Private Property Get IClassTexer_texHeader(texType As mint_TexingType) As String
    
End Property
Private Function IClassTexer_toByteArray() As Byte()
    
End Function

Private Function IClassTexer_toString() As String
    
End Function



Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
