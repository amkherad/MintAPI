VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Exceptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Exception"
'@PROJECT_LICENSE

''<class name="Exceptions" static tags="">
''  <summary>
''    Provides some methods to handle exceptions.
''  </summary>
''  <using>Exception</using>
''  <idea>Me</idea>
''  <sample>
''      throw Exceptions.Exception("Some description.")
''
''      Exceptions.Clear() 'Clears the last occured exception(error).
''      ... Do Something ...
''      throw Exceptions.IfError 'Throws an exception if there is any errors in GetLastError().
''  </sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Exceptions" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Public Function Exception( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception, _
            Optional ByVal DefaultMessage As String) As Exception
    If mString.IsEmptyOrNull(Message) Then Message = DefaultMessage
    Set Exception = New Exception
    Call Exception.Constructor0(Message, InnerException)
End Function
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
Public Function LastException() As Exception
    If mint_exceptions.LastException.Catched Then Exit Function
    If mint_exceptions.LastException.Exception Is Nothing Then Exit Function
    Set LastException = mint_exceptions.LastException.Exception
End Function

Public Sub Clear(): Call API_SetLastError(vbNullPtr): End Sub
Public Function IfError(Optional ByVal MethodName As String) As Exception
    Dim dwError As Long
    dwError = API_GetLastError
    
    Select Case dwError
        Case 0
            Set IfError = NoError
        Case Else
            Dim Message As String, InnerException As Exception
            If Not mString.IsEmptyOrNull(MethodName) Then
                Set InnerException = SystemCallFailureException(MethodName)
            End If
            Message = GetSystemMessageString(dwError)
            Set IfError = Exception(Message, InnerException, ERRORS_SYSTEMCALLFAILURE)
    End Select
End Function

'Public Function CountAllExceptions() As Long
'    CountAllExceptions = mint_exceptions.CountAllExceptions
'End Function
'Public Sub ClearAllExceptions()
'    Call mint_exceptions.ClearAllExceptions
'End Sub

''<summary></summary>
Public Function NoError(Optional ByVal InnerException As Exception) As Exception
    Set NoError = Create(ERRORS_NOERROR, InnerException)
End Function

Public Function SystemCallFailureException( _
            Optional ByVal SystemCallName As String, _
            Optional ByVal InnerException As Exception) As Exception
    If Not mString.IsEmptyOrNull(SystemCallName) Then _
        SystemCallName = mString.Format(ERRORS_SYSTEMCALLFAILURESPECIFIED, SystemCallName)
    Set SystemCallFailureException = Create(SystemCallName, ERRORS_SYSTEMCALLFAILURE, InnerException)
End Function
Public Function InvalidCallException( _
            Optional ByVal MethodName As String, _
            Optional ByVal InnerException As Exception) As Exception
    If Not mString.IsEmptyOrNull(MethodName) Then _
        MethodName = mString.Format(ERRORS_INVALIDCALLSPECIFIED, MethodName)
    Set InvalidCallException = Create(MethodName, ERRORS_INVALIDCALL, InnerException)
End Function

Public Function InvalidCastException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception) As Exception
    Set InvalidCastException = Create(Message, ERRORS_INVALIDCAST, InnerException)
End Function
Public Function OsNotSupportedException(Optional ByVal Message As String) As Exception
    Set OsNotSupportedException = Create(-1, _
        EXP_et_OsNotSuported, _
        Message, ERRORS_OSNOTSUPPORTED)
End Function
Public Function GetIONumberException(ByVal ErrorNumber As Long, _
    ByVal Direction As StreamDirection, ByVal Argument As String) As Exception
    
    Select Case ErrorNumber
        Case ERROR_PATH_NOT_FOUND
            Set GetIONumberException = Exps.DirectoryNotFoundException("The directory '" & Argument & "' could not be found.")
        Case ERROR_FILE_NOT_FOUND
            Set GetIONumberException = Exps.FileNotFoundException(Argument)
        Case ERROR_ACCESS_DENIED
            Set GetIONumberException = Exps.InvalidOperationException("Permission to the specified file is denied.")
        Case Else
            Set GetIONumberException = Exps.Exception(ErrorNumber)
    End Select
'    Ex.HResult = ErrorNumber
'    Ex.Source = Src
'    Set GetIONumberException = CreateException(-1, _
'        EXP_et_IOException, _
'        Description, ERRORS_IOEXCEPTION, Argument, Location, _
'        tObject, HelpLink)
End Function
Public Function IOException(Optional ByVal Message As String) As Exception
    Set IOException = Create(-1, _
        EXP_et_IOException, _
        Message, ERRORS_IOEXCEPTION)
End Function
Public Function DirectoryNotFoundException(Optional ByVal Message As String) As Exception
    Set DirectoryNotFoundException = Create(-1, _
        EXP_et_DirectoryNotFoundException, _
        Message, ERRORS_DIRECTORYNOTFOUND)
End Function

Public Function InvalidArgumentsException(Optional ByVal Message As String) As Exception
    Set InvalidArgumentsException = Create(-1, _
        EXP_et_InvalidArgument, _
        Message, ERRORS_ARGUMENTEXCEPTION)
End Function
Public Function ArgumentLengthOverflowException(Optional ByVal Message As String) As Exception
    Set ArgumentLengthOverflowException = Create(-1, _
        EXP_et_ArgumentLengthOverflow, _
        Message, ERRORS_ARGUMENTLENGTHOVERFLOW)
End Function
Public Function StringLengthOverflowException(Optional ByVal Message As String) As Exception
    Set StringLengthOverflowException = Create(-1, _
        EXP_et_StringLengthOverload, _
        Message, ERRORS_STRINGLENGTHOVERFLOWED)
End Function
Public Function InvalidArgumentTypeException(Optional ByVal ArgumentName As String) As Exception
    Set InvalidArgumentTypeException = Create(-1, _
        EXP_et_InvalidArgumentType, _
        ArgumentName, ERRORS_INVALIDARGUMENTTYPE)
End Function
Public Function TooManyArgumentsException(Optional ByVal Message As String) As Exception
    Set TooManyArgumentsException = Create(-1, _
        EXP_et_TooManyArguments, _
        Message, ERRORS_TOOMANYARGUMENTS)
End Function
Public Function NotEnoughArgumentsException(Optional ByVal Message As String) As Exception
    Set NotEnoughArgumentsException = Create(-1, _
        EXP_et_NotEnoughArguments, _
        Message, ERRORS_NOTENOUGHARGUMENTS)
End Function
Public Function InvalidArrayLBoundException(Optional ByVal Message As String) As Exception
    Set InvalidArrayLBoundException = Create(-1, _
        EXP_et_InvalidArrayLBound, _
        Message, ERRORS_INVALIDLBOUND)
End Function
Public Function InvalidArrayUBoundException(Optional ByVal Message As String) As Exception
    Set InvalidArrayUBoundException = Create(-1, _
        EXP_et_InvalidArrayUBound, _
        Message, ERRORS_INVALIDUBOUND)
End Function
Public Function InvalidVarTypeException(Optional ByVal Message As String) As Exception
    Set InvalidVarTypeException = Create(-1, _
        EXP_et_InvalidVarType, _
        Message, ERRORS_INVALIDVARTYPE)
End Function
Public Function InvalidArgumentValueException(Optional ByVal Message As String) As Exception
    Set InvalidArgumentValueException = Create(-1, _
        EXP_et_InvalidArgumentValue, _
        Message, ERRORS_INVALIDARGUMENTVALUE)
End Function
Public Function ArgumentNullException(Optional ByVal ArgumentName As String) As Exception
    Set ArgumentNullException = Create(-1, _
        EXP_et_ArgumentNull, _
        ArgumentName, ERRORS_ARGUMENTNULL)
End Function
Public Function OptionalArgumentNotPassed(Optional ByVal Message As String) As Exception
    Set OptionalArgumentNotPassed = Create(-1, _
        EXP_et_OptionalArgumentNotPassed, _
        Message, ERRORS_OPTIONALARGUMENTNOTPASSED)
End Function
Public Function ArgumentTypeMismatch(ByVal Argument1 As String, ByVal Argument2 As String) As Exception
    Set ArgumentNullException = Create(-1, _
        EXP_et_ArgumentNull, _
        mString.Format(ERRORS_ARGUMENTSTYPEMISMATCH, Argument1, Argument2))
End Function

Public Function OperationCanceledException(Optional ByVal Message As String) As Exception
    Set OperationCanceledException = Create(-1, _
        EXP_et_OperationCanceled, _
        Message, ERRORS_OPERATIONCANCELED)
End Function
Public Function InvalidOperationException(Optional ByVal Message As String) As Exception
    Set InvalidOperationException = Create(-1, _
        EXP_et_InvalidOperation, _
        Message, ERRORS_INVALIDOPERATION)
End Function
Public Function AccessDeniedException(Optional ByVal Message As String) As Exception
    Set AccessDeniedException = Create(-1, _
        EXP_et_AccessDenied, _
        Message, ERRORS_ACCESSDENIED)
End Function
Public Function UnknownValueException(Optional ByVal Message As String) As Exception
    Set UnknownValueException = Create(-1, _
        EXP_et_UnknownValue, _
        Message, ERRORS_UNKNOWNVALUE)
End Function
Public Function ZeroArgumentException(Optional ByVal Message As String) As Exception
    Set ZeroArgumentException = Create(-1, _
        EXP_et_ZeroArgument, _
        Message, ERRORS_ZERONUMBER)
End Function
Public Function NegativeArgumentException(Optional ByVal Message As String) As Exception
    Set NegativeArgumentException = Create(-1, _
        EXP_et_NegativeArgument, _
        Message, ERRORS_NEGATIVENUMBER)
End Function
Public Function ZeroNegativeArgumentException(Optional ByVal Message As String) As Exception
    Set ZeroNegativeArgumentException = Create(-1, _
        EXP_et_ZeroNegativeArgument, _
        Message, ERRORS_ZERONEGATIVENUMBER)
End Function
Public Function ValueIsTooHightException(Optional ByVal Message As String) As Exception
    Set ValueIsTooHightException = Create(-1, _
        EXP_et_ValueIsTooHight, _
        Message, ERRORS_VALUETOOHIGH)
End Function
Public Function ValueIsTooLowException(Optional ByVal Message As String) As Exception
    Set ValueIsTooLowException = Create(-1, _
        EXP_et_ValueIsTooLow, _
        Message, ERRORS_VALUETOOLOW)
End Function
Public Function PathNotFoundException(Optional ByVal Message As String) As Exception
    Set PathNotFoundException = Create(-1, _
        EXP_et_PathNotFound, _
        Message, ERRORS_PATHNOTFOUND)
End Function
Public Function InvalidPathException(Optional ByVal Message As String) As Exception
    Set InvalidPathException = Create(-1, _
        EXP_et_InvalidPath, _
        Message, ERRORS_INVALIDPATH)
End Function
Public Function InvalidFileException(Optional ByVal Message As String) As Exception
    Set InvalidFileException = Create(-1, _
        EXP_et_InvalidFile, _
        Message, ERRORS_INVALIDFILE)
End Function
Public Function FileNotFoundException(Optional ByVal Message As String) As Exception
    Set FileNotFoundException = Create(-1, _
        EXP_et_FileNotFound, _
        Message, ERRORS_FILENOTFOUND)
End Function
Public Function FileExistsException(Optional ByVal Message As String) As Exception
    Set FileExistsException = Create(-1, _
        EXP_et_FileExists, _
        Message, ERRORS_FILEEXISTS)
End Function
Public Function OutOfRangeException(Optional ByVal Message As String) As Exception
    Set OutOfRangeException = Create(-1, _
        EXP_et_OutOfRange, _
        Message, ERRORS_OUTOFRANGE)
End Function
Public Function ClassNotInitializedException(Optional ByVal Message As String) As Exception
    Set ClassNotInitializedException = Create(-1, _
        EXP_et_ClassNotInitialized, _
        Message, ERRORS_CLASSNOTINITIALIZED)
End Function
Public Function NotImplementedException(Optional ByVal Message As String) As Exception
    Set NotImplementedException = Create(-1, _
        EXP_et_NotImplemented, _
        Message, ERRORS_NOTIMPLEMENTED)
End Function
Public Function ItemNotExistsException(Optional ByVal Message As String) As Exception
    Set ItemNotExistsException = Create(-1, _
        EXP_et_ItemNotExists, _
        Message, ERRORS_ITEMNOTEXISTS)
End Function
Public Function ItemExistsException(Optional ByVal Message As String) As Exception
    Set ItemExistsException = Create(-1, _
        EXP_et_ItemExists, _
        Message, ERRORS_ITEMEXISTS)
End Function
Public Function ListIsEmptyException(Optional ByVal Message As String) As Exception
    Set ListIsEmptyException = Create(-1, _
        EXP_et_ListIsEmpty, _
        Message, ERRORS_LISTISEMPTY)
End Function
Public Function BufferIsEmptyException(Optional ByVal Message As String) As Exception
    Set BufferIsEmptyException = Create(-1, _
        EXP_et_BufferIsEmpty, _
        Message, ERRORS_BUFFERISEMPTY)
End Function
Public Function DisposedException(Optional ByVal Message As String) As Exception
    Set DisposedException = Create(-1, _
        EXP_et_Disposed, _
        Message, ERRORS_DISPOSED)
End Function
Public Function OpenFileException(Optional ByVal Message As String) As Exception
    Set OpenFileException = Create(-1, _
        EXP_et_OpenFile, _
        Message, ERRORS_OPENFILE)
End Function
Public Function CloseFileException(Optional ByVal Message As String) As Exception
    Set CloseFileException = Create(-1, _
        EXP_et_CloseFile, _
        Message, ERRORS_CLOSEFILE)
End Function
Public Function ReadFileException(Optional ByVal Message As String) As Exception
    Set ReadFileException = Create(-1, _
        EXP_et_ReadFile, _
        Message, ERRORS_READFILE)
End Function
Public Function WriteFileException(Optional ByVal Message As String) As Exception
    Set WriteFileException = Create(-1, _
        EXP_et_WriteFile, _
        Message, ERRORS_WRITEFILE)
End Function
Public Function TargetNotReadyException(Optional ByVal Message As String) As Exception
    Set TargetNotReadyException = Create(-1, _
        EXP_et_TargetNotReady, _
        Message, ERRORS_TARGETNOTREADY)
End Function
Public Function TargetNotOpenedException(Optional ByVal Message As String) As Exception
    Set TargetNotOpenedException = Create(-1, _
        EXP_et_TargetNotOpened, _
        Message, ERRORS_TARGETNOTOPENED)
End Function
Public Function InvalidStatusException(Optional ByVal Message As String) As Exception
    Set InvalidStatusException = Create(-1, _
        EXP_et_InvalidStatus, _
        Message, ERRORS_INVALIDSTATUS)
End Function
Public Function InvalidHandleException(Optional ByVal Message As String) As Exception
    Set InvalidHandleException = Create(-1, _
        EXP_et_InvalidHandle, _
        Message, ERRORS_INVALIDHANDLE)
End Function
Public Function UnableToCreateMetaObjectException(Optional ByVal Message As String) As Exception
    Set UnableToCreateMetaObjectException = Create(-1, _
        EXP_et_UnableToCreateMetaObject, _
        Message, ERRORS_UNABLETOCREATEMETAOBJECT)
End Function
Public Function UnknownNotificationException(Optional ByVal Message As String) As Exception
    Set UnknownNotificationException = Create(-1, _
        EXP_et_UnknownNotification, _
        Message, ERRORS_UNKNOWNNOTIFICATION)
End Function
Public Function UnableToCloneObjectException(Optional ByVal Message As String) As Exception
    Set UnableToCloneObjectException = Create(-1, _
        EXP_et_UnableToCloneObject, _
        Message, ERRORS_UNABLETOCLONEOBJECT)
End Function
Public Function OutOfMemoryException(Optional ByVal Message As String) As Exception
    Set OutOfMemoryException = Create(-1, _
        EXP_et_OutOfMemory, _
        Message, ERRORS_OUTOFMEMORY)
End Function
Public Function OnlyArraysAcceptedException(Optional ByVal Message As String) As Exception
    Set OnlyArraysAcceptedException = Create(-1, _
        EXP_et_OnlyArraysAccepted, _
        Message, ERRORS_ONLYARRAYSACCEPTED)
End Function
Public Function NotSupportedException(Optional ByVal Message As String) As Exception
    Set NotSupportedException = Create(-1, _
        EXP_et_NotSupported, _
        Message, ERRORS_NOTSUPPORTED)
End Function
Public Function MultiDimentionException(Optional ByVal Message As String) As Exception
    Set MultiDimentionException = Create(-1, _
        EXP_et_MultiDimention, _
        Message, ERRORS_MULTIDIMENTION)
End Function
Public Function EnumerationBrokeException(Optional ByVal Message As String) As Exception
    Set EnumerationBrokeException = Create(-1, _
        EXP_et_EnumerationBroke, _
        Message, ERRORS_ENUMERATIONBROKE)
End Function
Public Function QueryInformationFailedException(Optional ByVal Message As String) As Exception
    Set QueryInformationFailedException = Create(-1, _
        EXP_et_QueryInformationFailed, _
        Message, ERRORS_QUERYINFORMATIONFAILED)
End Function
'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
