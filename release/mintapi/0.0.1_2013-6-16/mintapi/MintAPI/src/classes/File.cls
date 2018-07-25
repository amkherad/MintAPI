VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "File is a advanced object to manage files and open or read/write to them.\r\n\r\nITargetStream<File>"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Directory"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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

Implements IData
Implements ITargetStream
Implements ISupportSignal
Implements ISupportSlot
Implements IDevice
Implements IAliasable

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
Dim buff() As Byte
Dim buffStartIndex As Long
Dim buffEndIndex As Long

Dim fOpenMode As FileOpenModes

Dim inState As Boolean
Dim outState As Boolean
Dim inState_LOCK As String
Dim outState_LOCK As String

Dim mySignalEmitter As New SignalEmitter

Private Sub Class_Initialize()
    Call mySignalEmitter.Initialize(Me)
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

Public Sub SetPath(Path As String)
    If IsOpen Then throw InvalidStatusException("File already opened and cannot change the path.")
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
    If VL = "" Then throw InvalidPathException, Me
    If Not CheckPathValidation(VL, True) Then throw InvalidPathException, Me
    VirtualPath = VL
End Function
Public Function AbsolutePath(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidPathException, Me
    If Not CheckPathValidation(VL, True, True) Then throw InvalidPathException, Me
    AbsolutePath = VL
End Function

Public Function Location(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidPathException, Me
    Location = GetFilePath(VL)
End Function
Public Function Path(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidPathException, Me
    Path = VL
End Function
Public Function Name(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidPathException, Me
    Name = GetFileName(VL)
End Function
Public Function NameOnly(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidPathException, Me
    NameOnly = GetFileNameOnly(VL)
End Function
Public Function Extention(Optional OptionalPath) As String
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidArgumentValueException, Me
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
        throw OpenFileException("File already opened.")

    Dim VL As String, Created As Boolean
    VL = me_SelectPath(OptionalPath)

    If Not (((fOpenMode And fmCreate) = fmCreate) Or fOpenMode = fmCreateNew) Then
        If Not baseFiling.FileExists(VL) Then
            throw FileNotFoundException
        End If
    Else
        If Not baseFiling.FileExists(VL) Then
            Created = True
        End If
    End If

    fFile = baseFiling.CreateFile(VL, fOpenMode, fAttributes, fAccess, fShare)
    If fFile.fHandle = 0 Then throw SystemCallFailureException
    p = fFile.Path
    f_IsOpened = True
    Set OpenFile = Me

    If Created Then
        Call mySignalEmitter.Emit(SIGNAL_CREATED)
        RaiseEvent Created(EventArgs(Me))
    End If
    Call mySignalEmitter.Emit(SIGNAL_OPEN)
    RaiseEvent FileOpen(EventArgs(Me))
End Function
Public Function Create(Optional fOpenMode As FileMode = fmCreate, Optional fAttributes As FileAttributes = fNormal, Optional fAccess As FileAccess = faRead, Optional fShare As FileShare = fshNone, Optional OptionalPath) As File
    Dim VL As String, Created As Boolean
    VL = me_SelectPath(OptionalPath)
    If IsOpen Then
        If fFile.Path = VL Then
            throw OpenFileException("File already opened.")
        Else
            Set Create = New File
            Call Create.Create(fOpenMode, fAttributes, fAccess, fShare, OptionalPath)
            Exit Function
        End If
    End If

    If Not baseFiling.FileExists(VL) Then _
        Created = True

    fFile = baseFiling.CreateFile(VL, fOpenMode, fAttributes, fAccess, fShare)
    If fFile.fHandle = 0 Then throw SystemCallFailureException
    p = fFile.Path
    f_IsOpened = True
    Set Create = Me

    If Created Then
        Call mySignalEmitter.Emit(SIGNAL_CREATED)
        RaiseEvent Created(EventArgs(Me))
    End If
    Call mySignalEmitter.Emit(SIGNAL_OPEN)
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
    If Not IsOpen Then Exit Sub 'throw InvalidStatusException("File not opened to close.")
    Call baseFiling.CloseFile(fFile)
    Call mySignalEmitter.Emit(SIGNAL_CLOSE)
    RaiseEvent FileClose(EventArgs(Me))
End Sub
Public Sub Flush()
    If Not IsOpen Then throw InvalidStatusException("File not opened.")
    Call baseFiling.FlushFile(fFile)
    Call mySignalEmitter.Emit(SIGNAL_FLUSH)
    RaiseEvent Flush(EventArgs(Me))
End Sub
Friend Sub Clear()

    Call mySignalEmitter.Emit(SIGNAL_CLEAR)
    RaiseEvent Clear(EventArgs(Me))
End Sub

Public Function Exists(Optional OptionalPath) As Boolean
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then
        Exists = False
    Else
        Exists = baseFiling.FileExists(VL)
    End If
End Function
Public Function EnsureExists(Optional OptionalPath) As File
    If IsOpen Then EnsureExists = True: Exit Function
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    Call baseFiling.EnsureFileExists(VL)
    Set EnsureExists = Me
End Function
Public Sub Delete(Optional OptionalPath)
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If Not Exists(VL) Then throw InvalidPathException
    Call Kill(VL)
    Call mySignalEmitter.Emit(SIGNAL_DELETED)
    RaiseEvent Deleted(EventArgs(Me))
End Sub
Public Function Directory(Optional OptionalPath) As Directory
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidPathException, Me
    Set Directory = mint_constructor.Directory(GetFilePath(VL))
End Function
Public Sub Rename(ByVal NewName As String, Optional OptionalPath)
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidPathException, Me

    Call mySignalEmitter.Emit(SIGNAL_RENAMED)
    RaiseEvent Renamed(EventArgs(Me))
End Sub
Public Sub Move(ByVal NewPath As String, Optional OptionalPath)
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidPathException, Me

    Call mySignalEmitter.Emit(SIGNAL_MOVED)
    RaiseEvent Moved(EventArgs(Me))
End Sub
Public Sub ConcurrentMove(ByVal NewPath As String, Optional OptionalPath)
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidPathException, Me

    Call mySignalEmitter.Emit(SIGNAL_MOVED)
    RaiseEvent Moved(EventArgs(Me))
End Sub
Public Sub Copy(ByVal NewPath As String, Optional OptionalPath)
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidPathException, Me

    Call mySignalEmitter.Emit(SIGNAL_COPIED)
    RaiseEvent Copied(EventArgs(Me))
End Sub
Public Function ConcurrentCopy(ByVal NewPath As String, Optional OptionalPath) As Thread
    Dim VL As String
    VL = me_SelectPath(OptionalPath)
    If VL = "" Then throw InvalidPathException, Me

    Call mySignalEmitter.Emit(SIGNAL_COPIED)
    RaiseEvent Copied(EventArgs(Me))
End Function

Public Function ChooseFile(Optional OwnerForm As Long, Optional DialogType As FileDialogTypes = fdOpen, Optional Title As String, Optional Filters As String, Optional StaticFileName As String, Optional ByVal InitialDirectory As Directory, Optional CheckPathExists As Boolean = False, Optional CheckFileExists As Boolean = False, Optional OpenAsReadOnly As Boolean = False, Optional AllowChangingDirectory As Boolean = True, Optional Flags As Long) As File
    Dim s As String, c_path As String, AllowMultiSelect As Boolean
    AllowMultiSelect = False
    If InitialDirectory Is Nothing Then
        c_path = IIf(fFile.Path = "", CurDir$, Path(fFile.Path))
    Else
        c_path = InitialDirectory.AbsolutePath
    End If
    Dim retVal As API_FileDialogReturn
    retVal = shellMethods.Dialogs_BrowseFile(IIf((DialogType And fdOpen) = fdOpen, OpenDialog, SaveDialog), OwnerForm, Title, c_path, Filters, StaticFileName, Flags, CheckPathExists, CheckFileExists, AllowMultiSelect, OpenAsReadOnly, AllowChangingDirectory)
    mint_api_dialogs_last_choose_file_read_only_flag_state = retVal.ReadOnlyCheckState
    Dim f As New File
    Call f.Initialize(retVal.Path)
    Set ChooseFile = f
    Call mySignalEmitter.Emit(SIGNAL_FILECHOOSED)
    RaiseEvent FileChoosed(EventArgs(Me))
End Function
Public Function ChooseFiles(Optional OwnerForm As Long, Optional DialogType As FileDialogTypes = fdOpen, Optional Title As String, Optional Filters As String, Optional ByVal InitialDirectory As Directory, Optional CheckPathExists As Boolean = False, Optional CheckFileExists As Boolean = False, Optional OpenAsReadOnly As Boolean = False, Optional AllowChangingDirectory As Boolean = True, Optional Flags As Long) As File()
    Dim s As String, c_path As String, AllowMultiSelect As Boolean
    AllowMultiSelect = True
    If InitialDirectory Is Nothing Then
        c_path = IIf(fFile.Path = "", CurDir$, Path(fFile.Path))
    Else
        c_path = InitialDirectory.AbsolutePath
    End If
    Dim retVal As API_FileDialogReturn
    retVal = shellMethods.Dialogs_BrowseFile(IIf((DialogType And fdOpen) = fdOpen, OpenDialog, SaveDialog), OwnerForm, Title, c_path, Filters, "", Flags, CheckPathExists, CheckFileExists, AllowMultiSelect, OpenAsReadOnly, AllowChangingDirectory)
    mint_api_dialogs_last_choose_file_read_only_flag_state = retVal.ReadOnlyCheckState
    Dim fls() As File
    Dim flsCount As Long
    Dim i As Long
    ReDim fls(retVal.subFilesCount - 1)
    For i = 0 To retVal.subFilesCount - 1
        Set fls(i) = New File
        Call fls(i).Initialize(baseMethods.ConcatPath(retVal.Path, retVal.subFiles(i)))
    Next
    ChooseFiles = fls
    Call mySignalEmitter.Emit(SIGNAL_FILECHOOSED)
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
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    ReadOnly = ((fFile.FileAttributes And API_fReadOnly) = API_fReadOnly)
End Property
Public Property Let ReadOnly(Value As Boolean)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")

End Property

Public Property Get BufferSize() As Long
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")

End Property
Public Property Let BufferSize(Value As Long)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")

End Property


Public Function constData() As Byte()
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim retB() As Byte
    Call baseFiling.SetBeginOfFile(fFile)
    Call baseFiling.ReadByteArrayFromFile(fFile, retB)
    constData = retB
    Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
End Function
Public Function readByteArray(Optional Length As Long = -1, Optional StartAt) As Byte()
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim retB() As Byte
    If Not IsMissing(StartAt) Then _
        Call SeekFile(CLng(StartAt))
    Call baseFiling.ReadByteArrayFromFile(fFile, retB, Length)
    readByteArray = retB
    Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
End Function
Public Sub FillData(Data)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")

End Sub

Public Sub WriteData(Data, Optional Length As Long = -1)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim DataBA() As Byte
    DataBA = mint_get_byte_array_of(Data, Length)
    If ArraySize(DataBA) <= 0 Then Exit Sub
    Call baseFiling.WriteByteArrayToFile(fFile, DataBA, Length)
    Call mySignalEmitter.Emit(SIGNAL_WRITE)
    RaiseEvent WriteData(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
    RaiseEvent Changed(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_LENGTHCHANGED)
    RaiseEvent LengthChanged(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Sub
'Public Function ReadData(Data)
'    Call ReadVar(Data)
'End Function
Public Sub ReadData(Data, Optional Length)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")

    Dim vt As VbVarType
    Dim i As Long, retBA As ByteArray
    Call mySignalEmitter.BlockSignals
    Call Err.Clear
    On Error GoTo iErr

    vt = VarType(Data)

    Select Case vt
        Case VbVarType.vbString
            If Not IsMissing(Length) Then
                Set retBA = ReadBlock(CLng(Length))
            Else
                Set retBA = ReadBlock
            End If
            Data = retBA.toString
        Case VbVarType.vbBoolean
            Call EnsureForwardLength(2)
            Set retBA = ReadBlock(2)
            Data = retBA.toBoolean
        Case VbVarType.vbByte
            Set retBA = ReadBlock(1)
            Data = retBA.ByteAt(0)
        Case VbVarType.vbInteger
            Call EnsureForwardLength(2)
            Set retBA = ReadBlock(2)
            Data = retBA.toInt
        Case VbVarType.vbLong
            Call EnsureForwardLength(4)
            Set retBA = ReadBlock(4)
            Data = retBA.toLong
        Case VbVarType.vbDouble
            Call EnsureForwardLength(8)
            Set retBA = ReadBlock(8)
            Data = retBA.toDouble
        Case VbVarType.vbSingle
            Call EnsureForwardLength(4)
            Set retBA = ReadBlock(4)
            Data = retBA.toSingle
        Case VbVarType.vbCurrency
            Call EnsureForwardLength(8)
            Set retBA = ReadBlock(8)
            Data = retBA.toCurrency
        Case VbVarType.vbDate
            Call EnsureForwardLength(8)
            Set retBA = ReadBlock(8)
            Data = retBA.toDate
        Case VbVarType.VBObject
            Dim obj As Object
            Set obj = Data
            If TypeOf obj Is BigNumber Then
                Set retBA = ReadBlock
                Set Data = retBA.toBigNumber
            End If
        Case Else
            If (vt And vbArray) = vbArray Then
                If (vt And vbByte) = vbByte Then
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
                throw UnknownValueException("Var is unknown.")
            End If
    End Select

iErr:
    Call mySignalEmitter.UnblockSignals
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Sub
Public Sub WriteLine(Data)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim DataBA() As Byte
    DataBA = mint_get_byte_array_of(Data)
    If ArraySize(DataBA) <= 0 Then Exit Sub
    Call baseFiling.WriteByteArrayToFile(fFile, DataBA)
    Call baseFiling.WriteByteArrayToFile(fFile, StringToByteArray(vbCrLf))
    Call mySignalEmitter.Emit(SIGNAL_WRITE)
    RaiseEvent WriteData(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
    RaiseEvent Changed(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_LENGTHCHANGED)
    RaiseEvent LengthChanged(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Sub
Public Function ReadLine() As ByteArray
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
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
        While (indexOfCRLF = -1 And (Not baseFiling.IsEndOfFile(fFile)))
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

    Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
End Function
'Public Function ReadLines() As ByteArray()
'    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
'
'End Function
Public Function TryReadBlock(Optional Length As Long = -1, Optional StartAt) As ByteArray
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    If Not IsMissing(StartAt) Then Call Me.SeekFile(CLng(StartAt))
    Dim buff() As Byte
    Dim ReadLength As Long

    ReadLength = ForwardLength

    If Length <> -1 Then _
        If Length <= ReadLength Then _
            ReadLength = Length

    Call baseFiling.ReadByteArrayFromFile(fFile, buff(), ReadLength)
    Set TryReadBlock = New ByteArray
    Call TryReadBlock.setBuffer(buff)
    If TryReadBlock.IsEmpty Then Exit Function
    Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Function
Public Function ReadBlock(Optional Length As Long = -1, Optional StartAt) As ByteArray
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    If Not IsMissing(StartAt) Then Call Me.SeekFile(CLng(StartAt))
    Dim buff() As Byte
    Call baseFiling.ReadByteArrayFromFile(fFile, buff(), Length)
    Set ReadBlock = New ByteArray
    Call ReadBlock.setBuffer(buff)
    If ReadBlock.IsEmpty Then Exit Function
    Call mySignalEmitter.Emit(SIGNAL_READ)
    RaiseEvent ReadData(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Function
Public Sub WriteRecord(Data, Optional Index)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim DataBA() As Byte
    DataBA = mint_get_byte_array_of(Data)
    'Call baseFiling.SetFilePosition(fFile, Index * ArraySize(DataBA))
End Sub
Public Function ReadRecord(Data, Optional Index) As ByteArray
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")

End Function
Public Sub WriteFormatted(Data)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")

End Sub
Public Function ReadFormatted(Data) As ByteArray
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")

End Function
Public Sub printf(Template As String, ParamArray Args() As Variant)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim rArgs() As Variant
    rArgs = Args
    Call WriteData(funcwArgs(Template, rArgs))
End Sub
Public Function scanf(Template As String, ParamArray Args() As Variant) As ByteArray
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")

End Function

Public Property Get Position() As Long
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Position = baseFiling.GetFilePosition(fFile)
End Property
Public Property Let Position(Value As Long)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim cPos As Long
    cPos = fFile.Position
    Call baseFiling.SetFilePosition(fFile, Value)
    If Value <> cPos Then
        Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
        RaiseEvent PositionChanged(EventArgs(Me))
    End If
End Property
Public Property Get LargePosition() As BigNumber
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim fBP As API_File_BigInt
    fBP = baseFiling.GetFilePositionLL(fFile)
    Set LargePosition = BigNumber(Array_Long(fBP.LowPart, fBP.HighPart))
End Property
Public Property Let LargePosition(Value As BigNumber)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")


End Property
Public Sub SeekFile(Position As Long)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim cPos As Long
    cPos = fFile.Position
    Call baseFiling.SetFilePosition(fFile, Position)
    If Position <> cPos Then
        Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
        RaiseEvent PositionChanged(EventArgs(Me))
    End If
End Sub
Public Sub RelativeSeekFile(RelativePosition As Long)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim cPos As Long
    cPos = fFile.Position
    Call baseFiling.TranslateFilePosition(fFile, Position)
    If RelativePosition <> cPos Then
        Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
        RaiseEvent PositionChanged(EventArgs(Me))
    End If
End Sub

Public Sub SetEndOfFile()
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Call baseFiling.SetEndOfFile(fFile)
    Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Sub
Public Sub SetBeginOfFile()
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Call baseFiling.SetBeginOfFile(fFile)
    Call mySignalEmitter.Emit(SIGNAL_POSITIONCHANGED)
    RaiseEvent PositionChanged(EventArgs(Me))
End Sub


Public Property Get Length() As Long
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Length = baseFiling.GetFileLength(fFile)
End Property
Public Property Let Length(Value As Long)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Call baseFiling.SetFileLength(fFile, Value)
    Call mySignalEmitter.Emit(SIGNAL_CHANGED)
    RaiseEvent Changed(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_LENGTHCHANGED)
    RaiseEvent LengthChanged(EventArgs(Me))
End Property
Public Property Get LargeLength() As BigNumber
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim fBS As API_File_BigInt
    fBS = baseFiling.GetFileLengthLL(fFile)
    Set LargeLength = BigNumber(Array_Long(fBS.LowPart, fBS.HighPart))
End Property
Public Property Let LargeLength(Value As BigNumber)
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")

End Property

Public Property Get ForwardLength() As Long
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    Dim Length As Long
    Length = baseFiling.GetFileLength(fFile)
    Length = Length - baseFiling.GetFilePosition(fFile)
    ForwardLength = Length
End Property


Public Property Get IsEndOfFile() As Boolean
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    IsEndOfFile = baseFiling.IsEndOfFile(fFile)
End Property
Public Property Get IsBeginOfFile() As Boolean
    If Not IsOpen Then throw InvalidStatusException("File is not opened.")
    IsBeginOfFile = (baseFiling.GetFilePosition(fFile) = 0)
End Property
Public Property Get IsMintAPICustomPlace() As Boolean

End Property

Public Function Available(Length As Long) As Boolean
    Available = (ForwardLength >= Length)
End Function
Public Sub EnsureForwardLength(Length As Long)
    If (ForwardLength < Length) Then _
        throw InvalidStatusException("Not enough data available.")
End Sub

Public Function ReadAllText(Optional OptionalPath) As String
    If IsOpen Then
        Call baseFiling.SetBeginOfFile(fFile)
        ReadAllText = ReadBlock.toString
    ElseIf Not IsMissing(OptionalPath) Then
        ReadAllText = File(me_SelectPath(OptionalPath)).OpenFile.ReadAllText
    Else
        throw InvalidPathException
    End If
End Function
Public Function ReadAllStringLines(Optional OptionalPath) As String()
    If IsOpen Then
        Call baseFiling.SetBeginOfFile(fFile)
        ReadAllStringLines = Split(ReadBlock.toString, vbCrLf)
    ElseIf Not IsMissing(OptionalPath) Then
        ReadAllStringLines = File(me_SelectPath(OptionalPath)).ReadAllStringLines
    Else
        throw InvalidPathException
    End If
End Function
Public Function ReadAllLines(Optional OptionalPath) As ByteArray()
    If IsOpen Then
        Call baseFiling.SetBeginOfFile(fFile)
        ReadAllLines = ReadBlock.Split(vbCrLf)
    ElseIf Not IsMissing(OptionalPath) Then
        ReadAllLines = File(me_SelectPath(OptionalPath)).OpenFile.ReadAllLines
    Else
        throw InvalidPathException
    End If
End Function
Public Function ReadToEnd(Optional OptionalPath) As ByteArray
    If IsOpen Then
        Call baseFiling.SetBeginOfFile(fFile)
        Set ReadToEnd = ReadBlock
    ElseIf Not IsMissing(OptionalPath) Then
        Set ReadToEnd = File(me_SelectPath(OptionalPath)).OpenFile.ReadToEnd
    Else
        throw InvalidPathException
    End If
End Function

Public Sub WriteAllText(Value As String, Optional OptionalPath)
    If IsOpen Then
        Call baseFiling.SetBeginOfFile(fFile)
        Call WriteData(Value)
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmCreate).WriteAllText(Value)
    Else
        throw InvalidPathException
    End If
End Sub
Public Sub WriteAllLines(Value() As String, Optional OptionalPath)
    If IsOpen Then
        Call baseFiling.SetBeginOfFile(fFile)
        Dim i As Long
        For i = 0 To ArraySize(Value) - 1
            Call WriteLine(Value(i))
        Next
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmCreate).WriteAllLines(Value)
    Else
        throw InvalidPathException
    End If
End Sub
Public Sub WriteByteArray(Value As ByteArray, Optional OptionalPath)
    If IsOpen Then
        Call baseFiling.SetBeginOfFile(fFile)
        Call WriteData(Value)
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmCreate).WriteByteArray(Value)
    Else
        throw InvalidPathException
    End If
End Sub

Public Sub AppendAllText(Value As String, Optional OptionalPath)
    If IsOpen Then
        Call baseFiling.SetEndOfFile(fFile)
        Call WriteData(Value)
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmAppend, fNormal, faWrite, fshRead).AppendAllText(Value)
    Else
        throw InvalidPathException
    End If
End Sub
Public Sub AppendAllLines(Value() As String, Optional OptionalPath)
    If IsOpen Then
        Call baseFiling.SetEndOfFile(fFile)
        Dim i As Long
        For i = 0 To ArraySize(Value) - 1
            Call WriteLine(Value(i))
        Next
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmAppend, fNormal, faWrite, fshRead).AppendAllLines(Value)
    Else
        throw InvalidPathException
    End If
End Sub
Public Sub AppendByteArray(Value As ByteArray, Optional OptionalPath)
    If IsOpen Then
        Call baseFiling.SetEndOfFile(fFile)
        Call WriteData(Value)
    ElseIf Not IsMissing(OptionalPath) Then
        Call File(me_SelectPath(OptionalPath)).OpenFile(fmAppend, fNormal, faWrite, fshRead).AppendByteArray(Value)
    Else
        throw InvalidPathException
    End If
End Sub

Private Sub IData_Append(target As Variant)
    Call AppendByteArray(ByteArray(target))
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
Private Function IData_toBinary() As BinaryData

End Function

Private Property Get IDevice_DeviceType() As DeviceType
    IDevice_DeviceType = dtDevice_File
End Property
Private Property Get IDevice_Handle() As Long
    IDevice_Handle = fFile.fHandle
End Property

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter
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
            throw InvalidCallException
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
    Dim val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        val = ((Length > 0) And inState)
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        val = val Or (outState)
    End If
    ITargetStream_Available = val
End Function

Private Sub ITargetStream_CloseStream(Optional direction As StreamDirection = 3&)
    If (direction And sdInStream) = sdInStream Then
        inState = False
    End If
    If (direction And sdOutStream) = sdOutStream Then
        outState = False
    End If
    If Not (inState Or outState) Then Call CloseFile
End Sub

Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = StreamDirection.sdBoth)
    If (direction And sdInStream) = sdInStream Then
        inState = True
    End If
    If (direction And sdOutStream) = sdOutStream Then
        outState = True
    End If
End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean
    Dim val As Boolean
    If (Dir And sdInStream) = sdInStream Then
        val = inState
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        val = val Or outState
    End If
    ITargetStream_getState = val
End Function
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    ITargetStream_readBytes = Me.readByteArray(Length, StartAt)
End Function

Private Sub ITargetStream_outStream(Data, Optional Length As Long = -1)
    If Not outState Then throw InvalidStatusException
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
                        throw InvalidArgumentTypeException("Streaming Command '" & strFirst & "' Is Not Available For File."), Me
                    End If
                Else
                    throw InvalidArgumentTypeException("Streaming Type Is Not Available For File.")
                End If
            End If
        End If
    End If
    Call Me.WriteData(Data)
End Sub
Private Sub ITargetStream_inStream(Data, Optional Length As Long = -1)
    If Not inState Then throw InvalidStatusException
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
                        throw InvalidArgumentTypeException("Streaming Command '" & strFirst & "' Is Not Available For ByteArray.")
                    End If
                Else
                    throw InvalidArgumentTypeException("Streaming Type Is Not Available For ByteArray.")
                End If
            Else
                throw InvalidArgumentTypeException
            End If
        Else
            throw InvalidArgumentTypeException
        End If
        throw InvalidArgumentTypeException
    End If
    Call ReadData(Data)
End Sub
Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved): Call mint_setstream_state(State, inState, inState_LOCK, Reserved): End Sub
Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved): Call mint_setstream_state(State, outState, outState_LOCK, Reserved): End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
