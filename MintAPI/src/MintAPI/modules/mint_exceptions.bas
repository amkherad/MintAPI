Attribute VB_Name = "mint_exceptions"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "Exceptions"

Private Declare Function API_FormatMessage Lib "Kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
'Private Declare Sub API_SetLastError Lib "kernel32" Alias "SetLastError" (ByVal dwErrCode As Long)


Private Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
Private Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Private Const EXCEPTION_BREAKPOINT = &H80000003
Private Const EXCEPTION_CONTINUE_EXECUTION = -1
Private Const EXCEPTION_CONTINUE_SEARCH = 0
Private Const EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
Private Const EXCEPTION_DEBUG_EVENT = 1
Private Const EXCEPTION_EXECUTE_HANDLER = 1
Private Const EXCEPTION_FLT_DENORMAL_OPERAND = &HC000008D
Private Const EXCEPTION_FLT_DIVIDE_BY_ZERO = &HC000008E
Private Const EXCEPTION_FLT_INEXACT_RESULT = &HC000008F
Private Const EXCEPTION_FLT_INVALID_OPERATION = &HC0000090
Private Const EXCEPTION_FLT_OVERFLOW = &HC0000091
Private Const EXCEPTION_FLT_STACK_CHECK = &HC0000092
Private Const EXCEPTION_FLT_UNDERFLOW = &HC0000093
Private Const EXCEPTION_GUARD_PAGE = &H80000001
Private Const EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
Private Const EXCEPTION_IN_PAGE_ERROR = &HC0000006
Private Const EXCEPTION_INT_DIVIDE_BY_ZERO = &HC0000094
Private Const EXCEPTION_INT_OVERFLOW = &HC0000095
Private Const EXCEPTION_INVALID_DISPOSITION = &HC0000026
Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15
Private Const EXCEPTION_NONCONTINUABLE_EXCEPTION = &HC0000025
Private Const EXCEPTION_PRIV_INSTRUCTION = &HC0000096
Private Const EXCEPTION_SINGLE_STEP = &H80000004
Private Const EXCEPTION_STACK_OVERFLOW = &HC00000FD

'-------------------------
'---- Error messages. ----
'-------------------------

'- Base exceptions.
    Public Const ERRORS_NOERROR As String = "No error occured."
    Public Const ERRORS_ERROR As String = "An error occured."
    Public Const ERRORS_EXCEPTION As String = "An exception has been thrown."
    Public Const ERRORS_INVALIDHANDLE As String = "Invalid handle."
    Public Const ERRORS_INVALIDCALL As String = "Invalid method call."
    Public Const ERRORS_INVALIDCALLSPECIFIED As String = "Invalid method call on '{0}'."
    Public Const ERRORS_INVALIDCAST As String = "Invalid cast."
    Public Const ERRORS_OBJECTNULLREFERENCE As String = "Object null reference encountered."
    Public Const ERRORS_OBJECTNULLREFERENCESPECIFIED As String = "Object '{0}' null reference encountered."
    Public Const ERRORS_OUTOFMEMORY As String = "Out of memory."
    Public Const ERRORS_THREADABORTEXCEPTON As String = "Thread abort exception."
    Public Const ERRORS_THREADABORTEXCEPTONSPECIFIED As String = "Thread '{0}' aborted."
    Public Const ERRORS_CLASSNOTIMPLEMENTED As String = "Class not implemented."
    Public Const ERRORS_INVALIDSTATUS As String = "Invalid target status."
    Public Const ERRORS_INDEXOUTOFRANGE As String = "Index was out of range."
    Public Const ERRORS_INDEXOUTOFRANGESPECIFIED As String = "Index was out of range. Legal range is '{0}'."
    Public Const ERRORS_INVALIDOPERATION As String = "Invalid operation."
    Public Const ERRORS_INVALIDOPERATIONSPECIFIED As String = "Invalid operation '{0}'."
    Public Const ERRORS_NOTSUPPORTED As String = "Operation not supported."
    
'- System exceptions.
    Public Const ERRORS_SYSTEMCALLFAILURE As String = "An error occured when trying to call a system method."
    Public Const ERRORS_SYSTEMCALLFAILURESPECIFIED As String = "An error occured when trying to call a system method '{0}'."

'- IO exceptions.
    Public Const ERRORS_IOEXCEPTION As String = "An IO error occured."
'    Public Const ERRORS_OPENFILE As String = "Error In Opening File."
'    Public Const ERRORS_CLOSEFILE As String = "Error In Closing File."
'    Public Const ERRORS_READFILE As String = "Error Reading From File."
'    Public Const ERRORS_WRITEFILE As String = "Error Writing In File."
'    Public Const ERRORS_INVALIDADDRESS As String = "Invalid Address."
    Public Const ERRORS_INVALIDPATH As String = "Invalid path."
    Public Const ERRORS_INVALIDPATHSPECIFIED As String = "Invalid path '{0}'."
'    Public Const ERRORS_INVALIDFILE As String = "Invalid File Type."
    Public Const ERRORS_FILENOTFOUND As String = "File does not found."
    Public Const ERRORS_FILENOTFOUNDSPECIFIED As String = "File '{0}' does not found."
    Public Const ERRORS_FILEEXISTS As String = "File already exists."
    Public Const ERRORS_FILEEXISTSSPECIFIED As String = "File '{0}' already exists."
    Public Const ERRORS_PATHNOTFOUND As String = "Path does not found."
    Public Const ERRORS_PATHNOTFOUNDSPECIFIED As String = "Path '{0}' does not found."

'- Argument exceptions.
    Public Const ERRORS_ARGUMENTEXCEPTION As String = "Invalid arguments are passed to function."
    
'- Variable exceptions.
    Public Const ERRORS_INVALIDARGUMENT As String = "Invalid argument."
    Public Const ERRORS_INVALIDARGUMENTSPECIFIED As String = "Invalid argument '{0}'."
    Public Const ERRORS_ARGUMENTNULL As String = "Argument is null."
    Public Const ERRORS_ARGUMENTNULLSPECIFIED As String = "Argument '{0}' is null."
    Public Const ERRORS_ARGUMENTCOUNT As String = "Invalid argument count."
    Public Const ERRORS_ARGUMENTCOUNTSPECIFIED As String = "Invalid argument count. Legal count is '{0}'."
    Public Const ERRORS_OPTIONALARGUMENTNOTPASSED As String = "Optional argument does not passed."
    Public Const ERRORS_OPTIONALARGUMENTNOTPASSEDSPECIFIED As String = "Optional argument '{0}' does not passed."
    Public Const ERRORS_ARGUMENTTYPEMISMATCH As String = "Argument type mismatch."
    Public Const ERRORS_ARGUMENTTYPEMISMATCHSPECIFIED As String = "Argument '{0}' type mismatch."
    
'- String exceptions.
    'Public Const ERRORS_STRINGLENGTHOVERFLOWED As String = "String Length Is Too Much."

'- Entity exceptions.
    Public Const ERRORS_ITEMNOTEXISTS As String = "Item does not exists."
    Public Const ERRORS_ITEMEXISTS As String = "Item already exists."
'    Public Const ERRORS_LISTISEMPTY As String = "List is empty."
'    Public Const ERRORS_BUFFERISEMPTY As String = "Buffer is empty."

'- Permission exceptions.
    Public Const ERRORS_ACCESSDENIED As String = "Access is denied."
    Public Const ERRORS_ACCESSVIOLATION As String = "Access violation encountered."
    Public Const ERRORS_ACCESSVIOLATIONSPECIFIED As String = "Access to '{0}' failed."
    
'- Array exceptions.
    Public Const ERRORS_ARRAYEXPECTED As String = "Only arrays accepted."
    Public Const ERRORS_ARRAYEXPECTEDSPECIFIED As String = "Argument '{0}' is not an array. Only arrays accepted."
    Public Const ERRORS_MULTIDIMENTION As String = "Multi dimention arrays are not accepted."
    
'- UI exceptions.
    Public Const ERRORS_DISPOSED As String = "Object has been disposed."
    Public Const ERRORS_DISPOSEDSPECIFIED As String = "Object '{0}' has been disposed."
    
'- Generic exceptions.
    Public Const ERRORS_ENUMERATIONBROKE As String = "Enumeration broke because enumerable has been changed during enumeration."
    Public Const ERRORS_CLASSNOTINITIALIZED As String = "Class has not been initialized."
    Public Const ERRORS_OPERATIONCANCELED As String = "Operation has been canceled."
    Public Const ERRORS_OPERATIONCANCELEDSPECIFIED As String = "Operation '{0}' has been canceled."
    
'-------------------------
'----   Error ID's.   ----
'-------------------------

    Public Const ERRORID_NOERROR              As Long = 0
    Public Const ERRORID_PATH_NOT_FOUND       As Long = 3
    Public Const ERRORID_ACCESS_DENIED        As Long = 5
    Public Const ERRORID_FILE_NOT_FOUND       As Long = 2
    Public Const ERRORID_FILE_EXISTS          As Long = 80
    Public Const ERRORID_INSUFFICIENT_BUFFER  As Long = 122
    
'----- Private constants
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Public Type MintExceptions_InternalExceptionInfo
    Exception As Exception
    Catched As Boolean
End Type

Dim exc_Counter As Long
'~[ThreadStatic]
Public LastException As MintExceptions_InternalExceptionInfo
Dim exc_suspend_emitation_of_signals As Boolean

Public Sub throw(ByVal Exception As Exception, Optional ByVal AtMethod As String, Optional ByVal Details As String)
    Dim Obj As Object, ObjPtr As Long
    ObjPtr = mHelper.CallerThis
    If ObjPtr <> vbNullPtr Then
        Call memcpy(Obj, ObjPtr, VLEN_PTR)
        Call IUnknown.AddRef(Obj)
    End If
    Call throw0(Exception, Obj, AtMethod, Details)
End Sub
Public Sub throw0(ByVal Exception As Exception, Optional ByVal Object As Object, Optional ByVal AtMethod As String, Optional ByVal Details As String)
    If Exception Is Nothing Then throw Exps.ArgumentNullException("Exception")
    If TypeOf Exception Is EmptyException Then Exit Sub
    Set Exception.Object = Object
    exc_Counter = exc_Counter + 1
    Set LastException.Exception = Exception
    Call API_SetLastError(C_MINTAPI_INTERNAL_ERROR Or ERROR_USERERROR)
    
    Dim MessageStr As String
    If Object Is Nothing Then
        MessageStr = Exception.Message
    Else
        MessageStr = Exception.Message & vbNewLine & _
                     "at: " & TypeName(Object) & " " & AtMethod
    End If
    
    Debug.Print MessageStr
    If Not AppInfo.TargetApplication Is Nothing Then
        On Error GoTo Err
        Call AppInfo.TargetApplication.Error(ExceptionOccuredEventArgs(Object, Exception))
Err:
    End If
    On Error GoTo EmitError
    If Not exc_suspend_emitation_of_signals Then
        exc_suspend_emitation_of_signals = True
        Call mint_application.EmitForApp("error")
        exc_suspend_emitation_of_signals = False
    End If
EmitError:
    On Error GoTo 0
    Dim AppError As IClassStream
    Set AppError = AppInfo.Streams.AppError
    If Not AppError Is Nothing Then
        out AppError, Exception.Message & vbNewLine & Exception.Details
    End If
    
    LastException.Catched = False
    Call Err.Raise(C_MINTAPI_INTERNAL_ERROR, "Unknown", MessageStr)
End Sub
Public Sub rethrow()
    Dim LastError As Long
    LastError = Err.LastDllError
    If LastError = C_MINTAPI_INTERNAL_ERROR Then
        Dim lException As Exception
        Set lException = LastException.Exception
        If Not lException Is Nothing Then
            Set lException = lException.Clone
            throw lException
        End If
    Else
        Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End If
End Sub

Public Function CountAllExceptions() As Long
    CountAllExceptions = exc_Counter
End Function
Public Sub ClearAllExceptions()
    exc_Counter = 0
End Sub

Public Function GetSystemMessageString(ByVal MessageID As Long) As String
    Dim BuffStr As String, Length As Long
    
    BuffStr = String$(XLARGELPSTR, vbNullChar)
    Length = API_FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, MessageID, 0, BuffStr, Len(BuffStr), ByVal 0&)
    If Length > 0 Then
        GetSystemMessageString = Left$(BuffStr, Length - 2)
    Else
        GetSystemMessageString = "Unknown Error."
    End If
End Function
