Attribute VB_Name = "modExceptions"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "Exceptions"

#If Not APPLICATIONIDDEFINED Then
Const APPLICATIONID As String = "Exception"
#End If

Private Declare Function API_exceptions_GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Sub API_exceptions_RaiseException Lib "kernel32" Alias "RaiseException" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

Public Enum ExceptionType
    EXP_et_SystemException = &H80000000
    EXP_et_NoError = 0
    EXP_et_Exception = 1
    EXP_et_SystemCallFailure = 2
    EXP_et_InvalidCall = 3
    
    EXP_et_OsNotSuported = 4
    EXP_et_OutOfRange
    EXP_et_Disposed
    EXP_et_NotImplemented
    EXP_et_ClassNotInitialized
    EXP_et_ArgumentLengthOverload
    EXP_et_ArgumentNull
    EXP_et_InvalidArgumentValue
    EXP_et_StringLengthOverload
    EXP_et_InvalidArgumentType
    EXP_et_TooManyArguments
    EXP_et_NotEnoughArguments
    EXP_et_InvalidArrayLBound
    EXP_et_InvalidArrayUBound
    EXP_et_InvalidVarType
    EXP_et_AccessDenied
    EXP_et_UnknownValue
    EXP_et_ZeroArgument
    EXP_et_NegativeArgument
    EXP_et_ZeroNegativeArgument
    EXP_et_ValueIsTooHight
    EXP_et_ValueIsTooLow
    EXP_et_FileNotFound
    EXP_et_FileExists
    EXP_et_PathNotFound
    EXP_et_InvalidPath
    EXP_et_InvalidFile
    EXP_et_ItemExists
    EXP_et_ItemNotExists
    EXP_et_ListIsEmpty
    EXP_et_BufferIsEmpty
    EXP_et_OperationCanceled
    EXP_et_OpenFile
    EXP_et_CloseFile
    EXP_et_ReadFile
    EXP_et_WriteFile
    EXP_et_TargetNotReady
    EXP_et_InvalidStatus
End Enum

Public Type Exception
    ExceptionDescription As String
    ExceptionLocation As String
    ExceptionType As ExceptionType
End Type

Dim inited As Boolean
Dim exc_Counter As Long

Dim hInstance As Long

Public Sub Initialize()
    If inited Then Exit Sub
    hInstance = API_exceptions_GetModuleHandle(vbNullString)
    Call baseConstants.Initialize
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    Call baseConstants.Dispose(Force)
    inited = False
End Sub

Public Property Get InstanceID() As Long
    If Not inited Then Call Initialize
    InstanceID = hInstance
End Property

Public Sub throw0(ByRef LibraryName As String, ByRef ModuleName As String, ByVal SourceMethodName As String, ByVal ExceptionDescription As String, Optional ByVal arguments As String = "", Optional ByVal ErrorNumber As Long, Optional ByVal ExceptionType As Long, Optional HelpFile, Optional HelpContext)
    ErrorNumber = IIf(Not CBool(ErrorNumber), DEFAULTERROR, ErrorNumber)
    SourceMethodName = IIf(LibraryName = "", APPLICATIONID, LibraryName) & IIf(ModuleName = "", "", "::" & ModuleName) & IIf(SourceMethodName = "", IIf(arguments = "", "", " =" & arguments), "::" & SourceMethodName & "(" & IIf(arguments = "", "", arguments) & ")")
    exc_Counter = exc_Counter + 1
    Call Err.Raise(ErrorNumber, _
                   SourceMethodName, _
                   ExceptionDescription, _
                   HelpFile, _
                   HelpContext)
End Sub
Public Sub throw(Exception As Exception)
    exc_Counter = exc_Counter + 1
    If (Exception.ExceptionType And ExceptionType.EXP_et_SystemException) = ExceptionType.EXP_et_SystemException Then
        Call API_exceptions_RaiseException(0, 0, 0, 0)
    Else
        Debug.Print Exception.ExceptionDescription
        Call Err.Raise(DEFAULTERROR, IIf(Exception.ExceptionLocation = "", "An Exception Occured.", "Error At:" & Exception.ExceptionLocation), Exception.ExceptionDescription)
    End If
End Sub
Public Sub throwT(Description As String, Optional Location As String = ""): throw Exception(Description, Location): End Sub
Public Sub throwN(): throw Exception: End Sub

Public Function Location(Library As String, Module As String, Func As String, Optional arguments As String = "") As String
    Location = IIf(Library = "", APPLICATIONID, Library) & IIf(Module = "", "", "::" & Module) & IIf(Func = "", IIf(arguments = "", "", " =" & arguments), "::" & Func & "(" & IIf(arguments = "", "", arguments) & ")")
End Function

Public Function CountAllExceptions() As Long
    CountAllExceptions = exc_Counter
End Function
Public Sub ClearAllExceptions()
    exc_Counter = 0
End Sub

Public Function SystemException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        SystemException.ExceptionDescription = "A System Exception Occured."
    Else
        SystemException.ExceptionDescription = Description
    End If
    SystemException.ExceptionLocation = Location
    SystemException.ExceptionType = EXP_et_SystemException
End Function
Public Function Exception(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        Exception.ExceptionDescription = ERRORS_ERROR
    Else
        Exception.ExceptionDescription = Description
    End If
    Exception.ExceptionLocation = Location
    Exception.ExceptionType = EXP_et_Exception
End Function
Public Function SystemCallFailureException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        SystemCallFailureException.ExceptionDescription = ERRORS_SYSTEMCALLFAILURE
    Else
        SystemCallFailureException.ExceptionDescription = Description
    End If
    SystemCallFailureException.ExceptionLocation = Location
    SystemCallFailureException.ExceptionType = EXP_et_SystemCallFailure
End Function
Public Function InvalidCallException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        InvalidCallException.ExceptionDescription = ERRORS_INVALIDCALL
    Else
        InvalidCallException.ExceptionDescription = Description
    End If
    InvalidCallException.ExceptionLocation = Location
    InvalidCallException.ExceptionType = EXP_et_InvalidCall
End Function

Public Function OsNotSupported(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        OsNotSupported.ExceptionDescription = ERRORS_NOTSUPPORTEDOS
    Else
        OsNotSupported.ExceptionDescription = Description
    End If
    OsNotSupported.ExceptionLocation = Location
    OsNotSupported.ExceptionType = EXP_et_OsNotSuported
End Function
Public Function ArgumentLengthOverloadException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        ArgumentLengthOverloadException.ExceptionDescription = ERRORS_INVALIDARGUMENTTYPE
    Else
        ArgumentLengthOverloadException.ExceptionDescription = Description
    End If
    ArgumentLengthOverloadException.ExceptionLocation = Location
    ArgumentLengthOverloadException.ExceptionType = EXP_et_ArgumentLengthOverload
End Function
Public Function StringLengthOverloadException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        StringLengthOverloadException.ExceptionDescription = ERRORS_INVALIDARGUMENTTYPE
    Else
        StringLengthOverloadException.ExceptionDescription = Description
    End If
    StringLengthOverloadException.ExceptionLocation = Location
    StringLengthOverloadException.ExceptionType = EXP_et_StringLengthOverload
End Function
Public Function InvalidArgumentTypeException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        InvalidArgumentTypeException.ExceptionDescription = ERRORS_INVALIDARGUMENTTYPE
    Else
        InvalidArgumentTypeException.ExceptionDescription = Description
    End If
    InvalidArgumentTypeException.ExceptionLocation = Location
    InvalidArgumentTypeException.ExceptionType = EXP_et_InvalidArgumentType
End Function
Public Function TooManyArgumentsException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        TooManyArgumentsException.ExceptionDescription = ERRORS_TOOMANYARGUMENTS
    Else
        TooManyArgumentsException.ExceptionDescription = Description
    End If
    TooManyArgumentsException.ExceptionLocation = Location
    TooManyArgumentsException.ExceptionType = EXP_et_TooManyArguments
End Function
Public Function NotEnoughArgumentsException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        NotEnoughArgumentsException.ExceptionDescription = ERRORS_LITTLEARGUMENTS
    Else
        NotEnoughArgumentsException.ExceptionDescription = Description
    End If
    NotEnoughArgumentsException.ExceptionLocation = Location
    NotEnoughArgumentsException.ExceptionType = EXP_et_NotEnoughArguments
End Function
Public Function InvalidArrayLBoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        InvalidArrayLBoundException.ExceptionDescription = ERRORS_INVALIDLBOUND
    Else
        InvalidArrayLBoundException.ExceptionDescription = Description
    End If
    InvalidArrayLBoundException.ExceptionLocation = Location
    InvalidArrayLBoundException.ExceptionType = EXP_et_InvalidArrayLBound
End Function
Public Function InvalidArrayUBoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        InvalidArrayUBoundException.ExceptionDescription = ERRORS_INVALIDUBOUND
    Else
        InvalidArrayUBoundException.ExceptionDescription = Description
    End If
    InvalidArrayUBoundException.ExceptionLocation = Location
    InvalidArrayUBoundException.ExceptionType = EXP_et_InvalidArrayUBound
End Function
Public Function InvalidVarTypeException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        InvalidVarTypeException.ExceptionDescription = ERRORS_INVALIDVARTYPE
    Else
        InvalidVarTypeException.ExceptionDescription = Description
    End If
    InvalidVarTypeException.ExceptionLocation = Location
    InvalidVarTypeException.ExceptionType = EXP_et_InvalidVarType
End Function
Public Function InvalidArgumentValueException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        InvalidArgumentValueException.ExceptionDescription = ERRORS_INVALIDARGUMENTVALUE
    Else
        InvalidArgumentValueException.ExceptionDescription = Description
    End If
    InvalidArgumentValueException.ExceptionLocation = Location
    InvalidArgumentValueException.ExceptionType = EXP_et_InvalidArgumentValue
End Function
Public Function ArgumentNullException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        ArgumentNullException.ExceptionDescription = ERRORS_ARGUMENTNULL
    Else
        ArgumentNullException.ExceptionDescription = Description
    End If
    ArgumentNullException.ExceptionLocation = Location
    ArgumentNullException.ExceptionType = EXP_et_ArgumentNull
End Function
Public Function OperationCanceledException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        OperationCanceledException.ExceptionDescription = ERRORS_CANCELEDERROR
    Else
        OperationCanceledException.ExceptionDescription = Description
    End If
    OperationCanceledException.ExceptionLocation = Location
    OperationCanceledException.ExceptionType = EXP_et_OperationCanceled
End Function
Public Function AccessDeniedException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        AccessDeniedException.ExceptionDescription = ERRORS_ACCESSDENIED
    Else
        AccessDeniedException.ExceptionDescription = Description
    End If
    AccessDeniedException.ExceptionLocation = Location
    AccessDeniedException.ExceptionType = EXP_et_AccessDenied
End Function
Public Function UnknownValueException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        UnknownValueException.ExceptionDescription = ERRORS_UNKNOWNVALUE
    Else
        UnknownValueException.ExceptionDescription = Description
    End If
    UnknownValueException.ExceptionLocation = Location
    UnknownValueException.ExceptionType = EXP_et_UnknownValue
End Function
Public Function ZeroArgumentException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        ZeroArgumentException.ExceptionDescription = ERRORS_ZERONUMBER
    Else
        ZeroArgumentException.ExceptionDescription = Description
    End If
    ZeroArgumentException.ExceptionLocation = Location
    ZeroArgumentException.ExceptionType = EXP_et_ZeroArgument
End Function
Public Function NegativeArgumentException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        NegativeArgumentException.ExceptionDescription = ERRORS_NEGATIVENUMBER
    Else
        NegativeArgumentException.ExceptionDescription = Description
    End If
    NegativeArgumentException.ExceptionLocation = Location
    NegativeArgumentException.ExceptionType = EXP_et_NegativeArgument
End Function
Public Function ZeroNegativeArgumentException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        ZeroNegativeArgumentException.ExceptionDescription = ERRORS_NEGATIVEZERONUMBER
    Else
        ZeroNegativeArgumentException.ExceptionDescription = Description
    End If
    ZeroNegativeArgumentException.ExceptionLocation = Location
    ZeroNegativeArgumentException.ExceptionType = EXP_et_ZeroNegativeArgument
End Function
Public Function ValueIsTooHightException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        ValueIsTooHightException.ExceptionDescription = ERRORS_VALUETOOHIGH
    Else
        ValueIsTooHightException.ExceptionDescription = Description
    End If
    ValueIsTooHightException.ExceptionLocation = Location
    ValueIsTooHightException.ExceptionType = EXP_et_ValueIsTooHight
End Function
Public Function ValueIsTooLowException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        ValueIsTooLowException.ExceptionDescription = ERRORS_VALUETOOLOW
    Else
        ValueIsTooLowException.ExceptionDescription = Description
    End If
    ValueIsTooLowException.ExceptionLocation = Location
    ValueIsTooLowException.ExceptionType = EXP_et_ValueIsTooLow
End Function
Public Function PathNotFoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        PathNotFoundException.ExceptionDescription = ERRORS_PATHNOTEXISTS
    Else
        PathNotFoundException.ExceptionDescription = Description
    End If
    PathNotFoundException.ExceptionLocation = Location
    PathNotFoundException.ExceptionType = EXP_et_PathNotFound
End Function
Public Function InvalidPathException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        InvalidPathException.ExceptionDescription = ERRORS_INVALIDPATH
    Else
        InvalidPathException.ExceptionDescription = Description
    End If
    InvalidPathException.ExceptionLocation = Location
    InvalidPathException.ExceptionType = EXP_et_InvalidPath
End Function
Public Function InvalidFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        InvalidFileException.ExceptionDescription = ERRORS_INVALIDFILE
    Else
        InvalidFileException.ExceptionDescription = Description
    End If
    InvalidFileException.ExceptionLocation = Location
    InvalidFileException.ExceptionType = EXP_et_InvalidFile
End Function
Public Function FileNotFoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        FileNotFoundException.ExceptionDescription = ERRORS_FILENOTFOUND
    Else
        FileNotFoundException.ExceptionDescription = Description
    End If
    FileNotFoundException.ExceptionLocation = Location
    FileNotFoundException.ExceptionType = EXP_et_FileNotFound
End Function
Public Function FileExistsException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        FileExistsException.ExceptionDescription = ERRORS_FILEEXISTS
    Else
        FileExistsException.ExceptionDescription = Description
    End If
    FileExistsException.ExceptionLocation = Location
    FileExistsException.ExceptionType = EXP_et_FileExists
End Function
Public Function OutOfRangeException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        OutOfRangeException.ExceptionDescription = ERRORS_OUTOFRANGE
    Else
        OutOfRangeException.ExceptionDescription = Description
    End If
    OutOfRangeException.ExceptionLocation = Location
    OutOfRangeException.ExceptionType = EXP_et_OutOfRange
End Function
Public Function ClassNotInitializedException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        ClassNotInitializedException.ExceptionDescription = ERRORS_CLASSNOTINITIALIZED
    Else
        ClassNotInitializedException.ExceptionDescription = Description
    End If
    ClassNotInitializedException.ExceptionLocation = Location
    ClassNotInitializedException.ExceptionType = EXP_et_ClassNotInitialized
End Function
Public Function NotImplementedException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        NotImplementedException.ExceptionDescription = ERRORS_NOTIMPLEMENTED
    Else
        NotImplementedException.ExceptionDescription = Description
    End If
    NotImplementedException.ExceptionLocation = Location
    NotImplementedException.ExceptionType = EXP_et_NotImplemented
End Function
Public Function ItemNotExistsException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        ItemNotExistsException.ExceptionDescription = ERRORS_ITEMDOESNOTEXISTS
    Else
        ItemNotExistsException.ExceptionDescription = Description
    End If
    ItemNotExistsException.ExceptionLocation = Location
    ItemNotExistsException.ExceptionType = EXP_et_ItemNotExists
End Function
Public Function ItemExistsException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        ItemExistsException.ExceptionDescription = ERRORS_ITEMEXISTS
    Else
        ItemExistsException.ExceptionDescription = Description
    End If
    ItemExistsException.ExceptionLocation = Location
    ItemExistsException.ExceptionType = EXP_et_ItemExists
End Function
Public Function ListIsEmptyException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        ListIsEmptyException.ExceptionDescription = ERRORS_LISTISEMPTY
    Else
        ListIsEmptyException.ExceptionDescription = Description
    End If
    ListIsEmptyException.ExceptionLocation = Location
    ListIsEmptyException.ExceptionType = EXP_et_ListIsEmpty
End Function
Public Function BufferIsEmptyException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        BufferIsEmptyException.ExceptionDescription = ERRORS_BUFFERISEMPTY
    Else
        BufferIsEmptyException.ExceptionDescription = Description
    End If
    BufferIsEmptyException.ExceptionLocation = Location
    BufferIsEmptyException.ExceptionType = EXP_et_BufferIsEmpty
End Function
Public Function DisposedException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        DisposedException.ExceptionDescription = ERRORS_DISPOSED
    Else
        DisposedException.ExceptionDescription = Description
    End If
    DisposedException.ExceptionLocation = Location
    DisposedException.ExceptionType = EXP_et_Disposed
End Function
Public Function OpenFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        OpenFileException.ExceptionDescription = ERRORS_CANTOPENFILE
    Else
        OpenFileException.ExceptionDescription = Description
    End If
    OpenFileException.ExceptionLocation = Location
    OpenFileException.ExceptionType = EXP_et_OpenFile
End Function
Public Function CloseFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        CloseFileException.ExceptionDescription = ERRORS_CANTCLOSEFILE
    Else
        CloseFileException.ExceptionDescription = Description
    End If
    CloseFileException.ExceptionLocation = Location
    CloseFileException.ExceptionType = EXP_et_CloseFile
End Function
Public Function ReadFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        ReadFileException.ExceptionDescription = ERRORS_CANTREADFILE
    Else
        ReadFileException.ExceptionDescription = Description
    End If
    ReadFileException.ExceptionLocation = Location
    ReadFileException.ExceptionType = EXP_et_ReadFile
End Function
Public Function WriteFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        WriteFileException.ExceptionDescription = ERRORS_CANTWRITEFILE
    Else
        WriteFileException.ExceptionDescription = Description
    End If
    WriteFileException.ExceptionLocation = Location
    WriteFileException.ExceptionType = EXP_et_WriteFile
End Function
Public Function TargetNotReadyException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        TargetNotReadyException.ExceptionDescription = ERRORS_TARGETNOTREADY
    Else
        TargetNotReadyException.ExceptionDescription = Description
    End If
    TargetNotReadyException.ExceptionLocation = Location
    TargetNotReadyException.ExceptionType = EXP_et_TargetNotReady
End Function
Public Function InvalidStatusException(Optional Description As String = "", Optional Location As String = "") As Exception
    If Description = "" Then
        InvalidStatusException.ExceptionDescription = ERRORS_INVALIDSTATUS
    Else
        InvalidStatusException.ExceptionDescription = Description
    End If
    InvalidStatusException.ExceptionLocation = Location
    InvalidStatusException.ExceptionType = EXP_et_InvalidStatus
End Function


