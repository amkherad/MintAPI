Attribute VB_Name = "mint_api_exceptions"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "Exceptions"

Private Declare Function API_exceptions_GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Sub API_exceptions_RaiseException Lib "kernel32" Alias "RaiseException" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

Public Enum pExceptionType
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
    EXP_et_TargetNotOpened
    EXP_et_InvalidStatus
    EXP_et_InvalidHandle
End Enum

Dim exc_Counter As Long
Public LastException As Exception

Public Sub throw0(ByRef LibraryName As String, ByRef ModuleName As String, ByVal SourceMethodName As String, ByVal ExceptionDescription As String, Optional ByVal Arguments As String = "", Optional ByVal ErrorNumber As Long, Optional ByVal ExceptionType As Long, Optional HelpFile, Optional HelpContext)
    ErrorNumber = IIf(Not CBool(ErrorNumber), DEFAULTERROR, ErrorNumber)
    SourceMethodName = IIf(LibraryName = "", APPLICATIONID, LibraryName) & IIf(ModuleName = "", "", "::" & ModuleName) & IIf(SourceMethodName = "", IIf(Arguments = "", "", " =" & Arguments), "::" & SourceMethodName & "(" & IIf(Arguments = "", "", Arguments) & ")")
    exc_Counter = exc_Counter + 1
    Call Err.Raise(ErrorNumber, _
                   SourceMethodName, _
                   ExceptionDescription, _
                   HelpFile, _
                   HelpContext)
End Sub
Public Sub throw(Exception As Exception, Optional Object As Object, Optional AtMethod As String, Optional Details As String)
    If Exception Is Nothing Then throw ArgumentNullException("Exception is null.", "Exceptions.throw()")
    'If Exception = Nothing Then Set Exception = Me.Exception
    If Exception.ExceptionType = EXP_et_NoError Then Exit Sub
    exc_Counter = exc_Counter + 1
    Set LastException = Exception
    Debug.Print Exception.Message
    If Not tApplication Is Nothing Then
        On Error GoTo Err
        'Call tApplication.Error(ExceptionOccuredEventArgs(Nothing, Exception))
Err:
    End If
    On Error GoTo 0
    If (Exception.ExceptionType And ExceptionType.EXP_et_SystemException) = ExceptionType.EXP_et_SystemException Then
        Call API_exceptions_RaiseException(0, 0, 0, 0)
    Else
        Call Err.Raise(DEFAULTERROR, IIf(Exception.Location = "", "An Exception Occured.", "Error At:" & Exception.Location), Exception.Message)
    End If
End Sub
Public Sub rethrow()
    If Err.LastDllError = DEFAULTERROR Then
        If Not LastException Is Nothing Then
            throw LastException
        End If
    Else
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub
Public Sub throwD(Description As String, Optional Location As String = ""): throw Exception(Description, Location): End Sub
Public Sub throwN(): throw Exception: End Sub

Public Function Location(Library As String, Module As String, Func As String, Optional Arguments As String = "") As String
    Location = IIf(Library = "", APPLICATIONID, Library) & IIf(Module = "", "", "::" & Module) & IIf(Func = "", IIf(Arguments = "", "", " =" & Arguments), "::" & Func & "(" & IIf(Arguments = "", "", Arguments) & ")")
End Function

Public Function CountAllExceptions() As Long
    CountAllExceptions = exc_Counter
End Function
Public Sub ClearAllExceptions()
    exc_Counter = 0
End Sub

Public Function SystemException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = "A System Exception Occured."
    Call exp.Initialize(Description, Location, EXP_et_SystemException, DEFAULTERROR, Nothing)
    Set SystemException = exp
End Function
Public Function Exception(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = "A System Exception Occured."
    Call exp.Initialize(Description, Location, EXP_et_Exception, DEFAULTERROR, Nothing)
    Set Exception = exp
End Function
Public Function SystemCallFailureException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_SYSTEMCALLFAILURE
    Call exp.Initialize(Description, Location, EXP_et_SystemCallFailure, DEFAULTERROR, Nothing)
    Set SystemCallFailureException = exp
End Function
Public Function InvalidCallException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDCALL
    Call exp.Initialize(Description, Location, EXP_et_InvalidCall, DEFAULTERROR, Nothing)
    Set InvalidCallException = exp
End Function

Public Function OsNotSupported(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_NOTSUPPORTEDOS
    Call exp.Initialize(Description, Location, EXP_et_OsNotSuported, DEFAULTERROR, Nothing)
    Set OsNotSupported = exp
End Function
Public Function ArgumentLengthOverloadException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDARGUMENTTYPE
    Call exp.Initialize(Description, Location, EXP_et_ArgumentLengthOverload, DEFAULTERROR, Nothing)
    Set ArgumentLengthOverloadException = exp
End Function
Public Function StringLengthOverloadException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDARGUMENTTYPE
    Call exp.Initialize(Description, Location, EXP_et_StringLengthOverload, DEFAULTERROR, Nothing)
    Set StringLengthOverloadException = exp
End Function
Public Function InvalidArgumentTypeException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDARGUMENTTYPE
    Call exp.Initialize(Description, Location, EXP_et_InvalidArgumentType, DEFAULTERROR, Nothing)
    Set InvalidArgumentTypeException = exp
End Function
Public Function TooManyArgumentsException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_TOOMANYARGUMENTS
    Call exp.Initialize(Description, Location, EXP_et_TooManyArguments, DEFAULTERROR, Nothing)
    Set TooManyArgumentsException = exp
End Function
Public Function NotEnoughArgumentsException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_AFEWARGUMENTS
    Call exp.Initialize(Description, Location, EXP_et_NotEnoughArguments, DEFAULTERROR, Nothing)
    Set NotEnoughArgumentsException = exp
End Function
Public Function InvalidArrayLBoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDLBOUND
    Call exp.Initialize(Description, Location, EXP_et_InvalidArrayLBound, DEFAULTERROR, Nothing)
    Set InvalidArrayLBoundException = exp
End Function
Public Function InvalidArrayUBoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDUBOUND
    Call exp.Initialize(Description, Location, EXP_et_InvalidArrayUBound, DEFAULTERROR, Nothing)
    Set InvalidArrayUBoundException = exp
End Function
Public Function InvalidVarTypeException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDVARTYPE
    Call exp.Initialize(Description, Location, EXP_et_InvalidVarType, DEFAULTERROR, Nothing)
    Set InvalidVarTypeException = exp
End Function
Public Function InvalidArgumentValueException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDARGUMENTVALUE
    Call exp.Initialize(Description, Location, EXP_et_InvalidArgumentValue, DEFAULTERROR, Nothing)
    Set InvalidArgumentValueException = exp
End Function
Public Function ArgumentNullException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_ARGUMENTNULL
    Call exp.Initialize(Description, Location, EXP_et_ArgumentNull, DEFAULTERROR, Nothing)
    Set ArgumentNullException = exp
End Function
Public Function OperationCanceledException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_CANCELEDERROR
    Call exp.Initialize(Description, Location, EXP_et_OperationCanceled, DEFAULTERROR, Nothing)
    Set OperationCanceledException = exp
End Function
Public Function AccessDeniedException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_ACCESSDENIED
    Call exp.Initialize(Description, Location, EXP_et_AccessDenied, DEFAULTERROR, Nothing)
    Set AccessDeniedException = exp
End Function
Public Function UnknownValueException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_UNKNOWNVALUE
    Call exp.Initialize(Description, Location, EXP_et_UnknownValue, DEFAULTERROR, Nothing)
    Set UnknownValueException = exp
End Function
Public Function ZeroArgumentException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_ZERONUMBER
    Call exp.Initialize(Description, Location, EXP_et_ZeroArgument, DEFAULTERROR, Nothing)
    Set ZeroArgumentException = exp
End Function
Public Function NegativeArgumentException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_NEGATIVENUMBER
    Call exp.Initialize(Description, Location, EXP_et_NegativeArgument, DEFAULTERROR, Nothing)
    Set NegativeArgumentException = exp
End Function
Public Function ZeroNegativeArgumentException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_NEGATIVEZERONUMBER
    Call exp.Initialize(Description, Location, EXP_et_ZeroNegativeArgument, DEFAULTERROR, Nothing)
    Set ZeroNegativeArgumentException = exp
End Function
Public Function ValueIsTooHightException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_VALUETOOHIGH
    Call exp.Initialize(Description, Location, EXP_et_ValueIsTooHight, DEFAULTERROR, Nothing)
    Set ValueIsTooHightException = exp
End Function
Public Function ValueIsTooLowException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_VALUETOOLOW
    Call exp.Initialize(Description, Location, EXP_et_ValueIsTooLow, DEFAULTERROR, Nothing)
    Set ValueIsTooLowException = exp
End Function
Public Function PathNotFoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_PATHNOTEXISTS
    Call exp.Initialize(Description, Location, EXP_et_PathNotFound, DEFAULTERROR, Nothing)
    Set PathNotFoundException = exp
End Function
Public Function InvalidPathException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDPATH
    Call exp.Initialize(Description, Location, EXP_et_InvalidPath, DEFAULTERROR, Nothing)
    Set InvalidPathException = exp
End Function
Public Function InvalidFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDFILE
    Call exp.Initialize(Description, Location, EXP_et_InvalidFile, DEFAULTERROR, Nothing)
    Set InvalidFileException = exp
End Function
Public Function FileNotFoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_FILENOTFOUND
    Call exp.Initialize(Description, Location, EXP_et_FileNotFound, DEFAULTERROR, Nothing)
    Set FileNotFoundException = exp
End Function
Public Function FileExistsException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_FILEEXISTS
    Call exp.Initialize(Description, Location, EXP_et_FileExists, DEFAULTERROR, Nothing)
    Set FileExistsException = exp
End Function
Public Function OutOfRangeException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_OUTOFRANGE
    Call exp.Initialize(Description, Location, EXP_et_OutOfRange, DEFAULTERROR, Nothing)
    Set OutOfRangeException = exp
End Function
Public Function ClassNotInitializedException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_CLASSNOTINITIALIZED
    Call exp.Initialize(Description, Location, EXP_et_ClassNotInitialized, DEFAULTERROR, Nothing)
    Set ClassNotInitializedException = exp
End Function
Public Function NotImplementedException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_NOTIMPLEMENTED
    Call exp.Initialize(Description, Location, EXP_et_NotImplemented, DEFAULTERROR, Nothing)
    Set NotImplementedException = exp
End Function
Public Function ItemNotExistsException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_ITEMDOESNOTEXISTS
    Call exp.Initialize(Description, Location, EXP_et_ItemNotExists, DEFAULTERROR, Nothing)
    Set ItemNotExistsException = exp
End Function
Public Function ItemExistsException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_ITEMEXISTS
    Call exp.Initialize(Description, Location, EXP_et_ItemExists, DEFAULTERROR, Nothing)
    Set ItemExistsException = exp
End Function
Public Function ListIsEmptyException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_LISTISEMPTY
    Call exp.Initialize(Description, Location, EXP_et_ListIsEmpty, DEFAULTERROR, Nothing)
    Set ListIsEmptyException = exp
End Function
Public Function BufferIsEmptyException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_BUFFERISEMPTY
    Call exp.Initialize(Description, Location, EXP_et_BufferIsEmpty, DEFAULTERROR, Nothing)
    Set BufferIsEmptyException = exp
End Function
Public Function DisposedException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_DISPOSED
    Call exp.Initialize(Description, Location, EXP_et_Disposed, DEFAULTERROR, Nothing)
    Set DisposedException = exp
End Function
Public Function OpenFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_CANTOPENFILE
    Call exp.Initialize(Description, Location, EXP_et_OpenFile, DEFAULTERROR, Nothing)
    Set OpenFileException = exp
End Function
Public Function CloseFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_CANTCLOSEFILE
    Call exp.Initialize(Description, Location, EXP_et_CloseFile, DEFAULTERROR, Nothing)
    Set CloseFileException = exp
End Function
Public Function ReadFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_CANTREADFILE
    Call exp.Initialize(Description, Location, EXP_et_ReadFile, DEFAULTERROR, Nothing)
    Set ReadFileException = exp
End Function
Public Function WriteFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_CANTWRITEFILE
    Call exp.Initialize(Description, Location, EXP_et_WriteFile, DEFAULTERROR, Nothing)
    Set WriteFileException = exp
End Function
Public Function TargetNotReadyException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_TARGETNOTREADY
    Call exp.Initialize(Description, Location, EXP_et_TargetNotReady, DEFAULTERROR, Nothing)
    Set TargetNotReadyException = exp
End Function
Public Function TargetNotOpenedException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_TARGETNOTOPENED
    Call exp.Initialize(Description, Location, EXP_et_TargetNotOpened, DEFAULTERROR, Nothing)
    Set TargetNotOpenedException = exp
End Function
Public Function InvalidStatusException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDSTATUS
    Call exp.Initialize(Description, Location, EXP_et_InvalidStatus, DEFAULTERROR, Nothing)
    Set InvalidStatusException = exp
End Function
Public Function InvalidHandleException(Optional Description As String = "", Optional Location As String = "") As Exception
    Dim exp As New Exception
    If Description = "" Then Description = ERRORS_INVALIDHANDLE
    Call exp.Initialize(Description, Location, EXP_et_InvalidHandle, DEFAULTERROR, Nothing)
    Set InvalidHandleException = exp
End Function
