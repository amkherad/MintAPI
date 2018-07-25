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
Attribute Exception.VB_UserMemId = 0
    If mString.IsEmptyOrNull(DefaultMessage) Then DefaultMessage = ERRORS_EXCEPTION
    If mString.IsEmptyOrNull(Message) Then Message = DefaultMessage
    Set Exception = New Exception
    Call Exception.Constructor0(Message, InnerException)
End Function
Public Function IOException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception, _
            Optional ByVal DefaultMessage As String) As IOException
    If mString.IsEmptyOrNull(DefaultMessage) Then DefaultMessage = ERRORS_IOEXCEPTION
    If mString.IsEmptyOrNull(Message) Then Message = DefaultMessage
    Set IOException = New IOException
    Call IOException.Constructor0(Message, InnerException)
End Function
Public Function ArgumentException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception, _
            Optional ByVal DefaultMessage As String) As ArgumentException
    If mString.IsEmptyOrNull(DefaultMessage) Then DefaultMessage = ERRORS_ARGUMENTEXCEPTION
    If mString.IsEmptyOrNull(Message) Then Message = DefaultMessage
    Set ArgumentException = New ArgumentException
    Call ArgumentException.Constructor0(Message, InnerException)
End Function
Public Function IndexOutOfRangeException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception, _
            Optional ByVal DefaultMessage As String) As IndexOutOfRangeException
    If mString.IsEmptyOrNull(DefaultMessage) Then DefaultMessage = ERRORS_INDEXOUTOFRANGE
    If mString.IsEmptyOrNull(Message) Then Message = DefaultMessage
    Set IndexOutOfRangeException = New IndexOutOfRangeException
    Call IndexOutOfRangeException.Constructor0(Message, InnerException)
End Function
Public Function ThreadAbortException( _
            Optional ByVal Thread As String, _
            Optional ByVal InnerException As Exception, _
            Optional ByVal DefaultMessage As String) As ThreadAbortException
    Dim Message As String
    If mString.IsEmptyOrNull(Thread) Then
        If mString.IsEmptyOrNull(DefaultMessage) Then DefaultMessage = ERRORS_THREADABORTEXCEPTON
        Message = DefaultMessage
    Else
        Message = mString.Format(ERRORS_THREADABORTEXCEPTONSPECIFIED, Thread)
    End If
    Set ThreadAbortException = New ThreadAbortException
    Call ThreadAbortException.Constructor0(Message, InnerException)
End Function
Public Function InvalidOperationException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception, _
            Optional ByVal DefaultMessage As String) As InvalidOperationException
    If mString.IsEmptyOrNull(DefaultMessage) Then DefaultMessage = ERRORS_INVALIDOPERATION
    If mString.IsEmptyOrNull(Message) Then Message = DefaultMessage
    Set InvalidOperationException = New InvalidOperationException
    Call InvalidOperationException.Constructor0(Message, InnerException)
End Function
Public Function OutOfMemoryException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception, _
            Optional ByVal DefaultMessage As String) As OutOfMemoryException
    If mString.IsEmptyOrNull(DefaultMessage) Then DefaultMessage = ERRORS_OUTOFMEMORY
    If mString.IsEmptyOrNull(Message) Then Message = DefaultMessage
    Set OutOfMemoryException = New OutOfMemoryException
    Call OutOfMemoryException.Constructor0(Message, InnerException)
End Function
Public Function AccessDeniedException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception, _
            Optional ByVal DefaultMessage As String) As AccessDeniedException
    If mString.IsEmptyOrNull(DefaultMessage) Then DefaultMessage = ERRORS_ACCESSDENIED
    If mString.IsEmptyOrNull(Message) Then Message = DefaultMessage
    Set AccessDeniedException = New AccessDeniedException
    Call AccessDeniedException.Constructor0(Message, InnerException)
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

Public Sub RaiseSystemException(ByVal ExceptionCode As Long, _
                                ByVal ExceptionFlags As Long, _
                                Optional ByVal Arguments As ArgumentList)
    
    If Arguments Is Nothing Then
        Call API_RaiseException(ExceptionCode, ExceptionFlags, NO_VALUE, vbNullPtr)
    Else
        If Arguments.Count > SYSTEMEXCEPTION_MAXIMUM_PARAMETERS Then _
            throw Exps.IndexOutOfRangeException
        Dim Vec As Vector
        Set Vec = Arguments.ToVector
        Call API_RaiseException(ExceptionCode, ExceptionFlags, Vec.Length, Vec.DataHandle)
    End If
End Sub

Public Function LastException() As Exception
    If mint_exceptions.LastException.Catched Then Exit Function
    If mint_exceptions.LastException.Exception Is Nothing Then Exit Function
    Set LastException = mint_exceptions.LastException.Exception
End Function

Public Sub Clear()
    Call API_SetLastError(NO_VALUE)
    'Set mint_exceptions.LastException.Exception = Nothing
    mint_exceptions.LastException.Catched = True
End Sub
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
    Dim Exp As New EmptyException
    Set Exp.InnerException = InnerException
    Set NoError = Exp
End Function

'Basic exceptions.
Public Function SystemCallFailureException( _
            Optional ByVal SystemCallName As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    If Not mString.IsEmptyOrNull(SystemCallName) Then _
        SystemCallName = mString.Format(ERRORS_SYSTEMCALLFAILURESPECIFIED, SystemCallName)
    Set SystemCallFailureException = InvalidOperationException(SystemCallName, InnerException, ERRORS_SYSTEMCALLFAILURE)
End Function
Public Function InvalidCallException( _
            Optional ByVal MethodName As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    If Not mString.IsEmptyOrNull(MethodName) Then _
        MethodName = mString.Format(ERRORS_INVALIDCALLSPECIFIED, MethodName)
    Set InvalidCallException = InvalidOperationException(MethodName, InnerException, ERRORS_INVALIDCALL)
End Function
Public Function InvalidCastException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    Set InvalidCastException = InvalidOperationException(Message, InnerException, ERRORS_INVALIDCAST)
End Function
Public Function NotSupportedException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    Set NotSupportedException = InvalidOperationException(Message, InnerException, ERRORS_NOTSUPPORTED)
End Function
Public Function ObjectDisposedException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    Set ObjectDisposedException = InvalidOperationException(Message, InnerException, ERRORS_DISPOSED)
End Function
Public Function ClassNotInitializedException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    Set ClassNotInitializedException = InvalidOperationException(Message, InnerException, ERRORS_CLASSNOTINITIALIZED)
End Function
Public Function NotImplementedException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    Set NotImplementedException = InvalidOperationException(Message, InnerException, ERRORS_CLASSNOTIMPLEMENTED)
End Function
Public Function ObjectNullReferenceException( _
            Optional ByVal Object As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    If Not mString.IsEmptyOrNull(Object) Then _
        Object = mString.Format(ERRORS_OBJECTNULLREFERENCESPECIFIED, Object)
    Set ObjectNullReferenceException = InvalidOperationException(Object, InnerException, ERRORS_OBJECTNULLREFERENCE)
End Function
'Public Function AccessViolationException( _
'            Optional ByVal Object As String, _
'            Optional ByVal InnerException As Exception) As AccessDeniedException
'    If Not mString.IsEmptyOrNull(Object) Then _
'        Object = mString.Format(ERRORS_OBJECTNULLREFERENCESPECIFIED, Object)
'    Set AccessDeniedException = AccessDeniedException(Object, InnerException, ERRORS_OBJECTNULLREFERENCE)
'End Function

Public Function ItemExistsException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    Set ItemExistsException = InvalidOperationException(Message, InnerException, ERRORS_ITEMEXISTS)
End Function
Public Function ItemNotExistsException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    Set ItemNotExistsException = InvalidOperationException(Message, InnerException, ERRORS_ITEMNOTEXISTS)
End Function
Public Function ArrayExpectedException( _
            Optional ByVal Argument As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    If Not mString.IsEmptyOrNull(Argument) Then _
        Argument = mString.Format(ERRORS_ARRAYEXPECTEDSPECIFIED, Argument)
    Set ArrayExpectedException = InvalidOperationException(Argument, InnerException, ERRORS_ARRAYEXPECTED)
End Function
Public Function MultiDimentionException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    Set MultiDimentionException = InvalidOperationException(Message, InnerException, ERRORS_MULTIDIMENTION)
End Function
Public Function EnumerationBrokeException( _
            Optional ByVal Message As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    Set EnumerationBrokeException = InvalidOperationException(Message, InnerException, ERRORS_ENUMERATIONBROKE)
End Function
Public Function OperationCanceledException( _
            Optional ByVal Operation As String, _
            Optional ByVal InnerException As Exception) As InvalidOperationException
    If Not mString.IsEmptyOrNull(Operation) Then _
        Operation = mString.Format(ERRORS_ARGUMENTTYPEMISMATCHSPECIFIED, Operation)
    Set OperationCanceledException = InvalidOperationException(Operation, InnerException, ERRORS_OPERATIONCANCELED)
End Function

'Argument exceptions.
Public Function ArgumentNullException( _
            Optional ByVal Argument As String, _
            Optional ByVal InnerException As Exception) As ArgumentException
    If Not mString.IsEmptyOrNull(Argument) Then _
        Argument = mString.Format(ERRORS_ARGUMENTNULLSPECIFIED, Argument)
    Set ArgumentNullException = ArgumentException(Argument, InnerException, ERRORS_ARGUMENTNULL)
End Function
Public Function InvalidArgumentException( _
            Optional ByVal Argument As String, _
            Optional ByVal InnerException As Exception) As ArgumentException
    If Not mString.IsEmptyOrNull(Argument) Then _
        Argument = mString.Format(ERRORS_INVALIDARGUMENTSPECIFIED, Argument)
    Set InvalidArgumentException = ArgumentException(Argument, InnerException, ERRORS_INVALIDARGUMENT)
End Function
Public Function ArgumentTypeMismatch( _
            Optional ByVal Argument As String, _
            Optional ByVal InnerException As Exception) As Exception
    If Not mString.IsEmptyOrNull(Argument) Then _
        Argument = mString.Format(ERRORS_ARGUMENTTYPEMISMATCHSPECIFIED, Argument)
    Set ArgumentTypeMismatch = ArgumentException(Argument, InnerException, ERRORS_ARGUMENTTYPEMISMATCH)
End Function
Public Function ArgumentCountException( _
            Optional ByVal LegalCount As Long, _
            Optional ByVal InnerException As Exception) As ArgumentException
    Dim Message As String
    If LegalCount > 0 Then _
        Message = mString.Format(ERRORS_ARGUMENTCOUNTSPECIFIED, LegalCount)
    Set ArgumentCountException = ArgumentException(Message, InnerException, ERRORS_ARGUMENTCOUNT)
End Function
Public Function OptionalArgumentNotPassed( _
            Optional ByVal Argument As String, _
            Optional ByVal InnerException As Exception) As Exception
    If Not mString.IsEmptyOrNull(Argument) Then _
        Argument = mString.Format(ERRORS_OPTIONALARGUMENTNOTPASSEDSPECIFIED, Argument)
    Set OptionalArgumentNotPassed = ArgumentException(Argument, InnerException, ERRORS_OPTIONALARGUMENTNOTPASSED)
End Function

'Invalid operation exceptions.
'Public Function InvalidStatusException( _
'            Optional ByVal Message As String, _
'            Optional ByVal InnerException As Exception) As InvalidOperationException
'    Set InvalidStatusException = InvalidOperationException(Message, InnerException, ERRORS_INVALIDSTATUS)
'End Function
'Public Function InvalidHandleException( _
'            Optional ByVal Message As String, _
'            Optional ByVal InnerException As Exception) As InvalidOperationException
'    Set InvalidHandleException = InvalidOperationException(Message, InnerException, ERRORS_INVALIDHANDLE)
'End Function


'IO Exceptions....
'Public Function DirectoryNotFoundException( _
'            Optional ByVal Directory As String, _
'            Optional ByVal InnerException As Exception) As IOException
'    If Not mString.IsEmptyOrNull(Directory) Then _
'        Directory = mString.Format(ERRORS_DIRECTORYNOTFOUNDSPECIFIED, Directory)
'    Set DirectoryNotFoundException = IOException(Message, InnerException, ERRORS_DIRECTORYNOTFOUND)
'End Function
Public Function PathNotFoundException( _
            Optional ByVal Path As String, _
            Optional ByVal InnerException As Exception) As IOException
    If Not mString.IsEmptyOrNull(Path) Then _
        Path = mString.Format(ERRORS_PATHNOTFOUNDSPECIFIED, Path)
    Set PathNotFoundException = IOException(Path, InnerException, ERRORS_PATHNOTFOUND)
End Function
Public Function FileNotFoundException( _
            Optional ByVal FileName As String, _
            Optional ByVal InnerException As Exception) As IOException
    If Not mString.IsEmptyOrNull(FileName) Then _
        FileName = mString.Format(ERRORS_FILENOTFOUNDSPECIFIED, FileName)
    Set FileNotFoundException = IOException(FileName, InnerException, ERRORS_FILENOTFOUND)
End Function
Public Function FileAlreadyExistsException( _
            Optional ByVal FileName As String, _
            Optional ByVal InnerException As Exception) As IOException
    If Not mString.IsEmptyOrNull(FileName) Then _
        FileName = mString.Format(ERRORS_FILEEXISTSSPECIFIED, FileName)
    Set FileAlreadyExistsException = IOException(FileName, InnerException, ERRORS_FILEEXISTS)
End Function
Public Function InvalidPathException( _
            Optional ByVal Path As String, _
            Optional ByVal InnerException As Exception) As IOException
    If Not mString.IsEmptyOrNull(Path) Then _
        Path = mString.Format(ERRORS_INVALIDPATHSPECIFIED, Path)
    Set InvalidPathException = IOException(Path, InnerException, ERRORS_INVALIDPATH)
End Function
Public Function GetIOExceptionFromNumber(ByVal IOErrorNumber As Long, Optional ByVal Argument As String) As Exception
    Select Case IOErrorNumber
        Case ERRORID_NOERROR
            Set GetIOExceptionFromNumber = Exps.NoError
        Case ERRORID_PATH_NOT_FOUND
            Set GetIOExceptionFromNumber = Exps.PathNotFoundException(Argument)
        Case ERRORID_FILE_NOT_FOUND
            Set GetIOExceptionFromNumber = Exps.FileNotFoundException(Argument)
        Case ERRORID_ACCESS_DENIED
            Set GetIOExceptionFromNumber = Exps.AccessDeniedException(Argument)
        Case Else
            Set GetIOExceptionFromNumber = Exps.IOException(GetSystemMessageString(IOErrorNumber))
    End Select
End Function
'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
