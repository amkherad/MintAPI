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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Provides exceptioning methods such as throw."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Exception"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "Exceptions"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


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

Public Enum ErrorType
    etException = -999999999
    etReflection = -999999998
    etAbort = -999999997
End Enum

Friend Sub Initialize(Optional Arguments)
    
End Sub


Public Function IsException(ByVal Error As ErrObject) As Boolean
    Dim ErrNumber As Long
    ErrNumber = Error.Number
    IsException = (ErrNumber = etException)
End Function
Public Function IsReflection(ByVal Error As ErrObject) As Boolean
    Dim ErrNumber As Long
    ErrNumber = Error.Number
    IsReflection = (ErrNumber = etReflection)
End Function
Public Function IsAbort(ByVal Error As ErrObject) As Boolean
    Dim ErrNumber As Long
    ErrNumber = Error.Number
    IsAbort = (ErrNumber = etAbort)
End Function
Public Function GetErrorType(ByVal Error As ErrObject) As ErrorType
    GetErrorType = Error.Number
End Function


Public Function LastException() As Exception
    Set LastException = mint_api_exceptions.LastException
End Function

Public Function CountAllExceptions() As Long
    CountAllExceptions = mint_api_exceptions.CountAllExceptions
End Function
Public Sub ClearAllExceptions()
    Call mint_api_exceptions.ClearAllExceptions
End Sub

Public Function SystemException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set SystemException = mint_api_exceptions.SystemException(Description, Location)
End Function
Public Function Exception(Optional Description As String = "", Optional Location As String = "") As Exception
    Set Exception = mint_api_exceptions.Exception(Description, Location)
End Function
Public Function SystemCallFailureException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set SystemCallFailureException = mint_api_exceptions.SystemCallFailureException(Description, Location)
End Function
Public Function InvalidCallException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set InvalidCallException = mint_api_exceptions.InvalidCallException(Description, Location)
End Function

Public Function OsNotSupported(Optional Description As String = "", Optional Location As String = "") As Exception
    Set OsNotSupported = mint_api_exceptions.OsNotSupported(Description, Location)
End Function
Public Function ArgumentLengthOverloadException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set ArgumentLengthOverloadException = mint_api_exceptions.ArgumentLengthOverloadException(Description, Location)
End Function
Public Function StringLengthOverloadException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set StringLengthOverloadException = mint_api_exceptions.StringLengthOverloadException(Description, Location)
End Function
Public Function InvalidArgumentTypeException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set InvalidArgumentTypeException = mint_api_exceptions.InvalidArgumentTypeException(Description, Location)
End Function
Public Function TooManyArgumentsException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set TooManyArgumentsException = mint_api_exceptions.TooManyArgumentsException(Description, Location)
End Function
Public Function NotEnoughArgumentsException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set NotEnoughArgumentsException = mint_api_exceptions.NotEnoughArgumentsException(Description, Location)
End Function
Public Function InvalidArrayLBoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set InvalidArrayLBoundException = mint_api_exceptions.InvalidArrayLBoundException(Description, Location)
End Function
Public Function InvalidArrayUBoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set InvalidArrayUBoundException = mint_api_exceptions.InvalidArrayUBoundException(Description, Location)
End Function
Public Function InvalidVarTypeException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set InvalidVarTypeException = mint_api_exceptions.InvalidVarTypeException(Description, Location)
End Function
Public Function InvalidArgumentValueException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set InvalidArgumentValueException = mint_api_exceptions.InvalidArgumentValueException(Description, Location)
End Function
Public Function ArgumentNullException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set ArgumentNullException = mint_api_exceptions.ArgumentNullException(Description, Location)
End Function
Public Function OperationCanceledException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set OperationCanceledException = mint_api_exceptions.OperationCanceledException(Description, Location)
End Function
Public Function AccessDeniedException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set AccessDeniedException = mint_api_exceptions.AccessDeniedException(Description, Location)
End Function
Public Function UnknownValueException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set UnknownValueException = mint_api_exceptions.UnknownValueException(Description, Location)
End Function
Public Function ZeroArgumentException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set ZeroArgumentException = mint_api_exceptions.ZeroArgumentException(Description, Location)
End Function
Public Function NegativeArgumentException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set NegativeArgumentException = mint_api_exceptions.NegativeArgumentException(Description, Location)
End Function
Public Function ZeroNegativeArgumentException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set ZeroNegativeArgumentException = mint_api_exceptions.ZeroNegativeArgumentException(Description, Location)
End Function
Public Function ValueIsTooHightException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set ValueIsTooHightException = mint_api_exceptions.ValueIsTooHightException(Description, Location)
End Function
Public Function ValueIsTooLowException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set ValueIsTooLowException = mint_api_exceptions.ValueIsTooLowException(Description, Location)
End Function
Public Function PathNotFoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set PathNotFoundException = mint_api_exceptions.PathNotFoundException(Description, Location)
End Function
Public Function InvalidPathException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set InvalidPathException = mint_api_exceptions.InvalidPathException(Description, Location)
End Function
Public Function InvalidFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set InvalidFileException = mint_api_exceptions.InvalidFileException(Description, Location)
End Function
Public Function FileNotFoundException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set FileNotFoundException = mint_api_exceptions.FileNotFoundException(Description, Location)
End Function
Public Function FileExistsException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set FileExistsException = mint_api_exceptions.FileExistsException(Description, Location)
End Function
Public Function OutOfRangeException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set OutOfRangeException = mint_api_exceptions.OutOfRangeException(Description, Location)
End Function
Public Function ClassNotInitializedException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set ClassNotInitializedException = mint_api_exceptions.ClassNotInitializedException(Description, Location)
End Function
Public Function NotImplementedException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set NotImplementedException = mint_api_exceptions.NotImplementedException(Description, Location)
End Function
Public Function ItemNotExistsException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set ItemNotExistsException = mint_api_exceptions.ItemNotExistsException(Description, Location)
End Function
Public Function ItemExistsException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set ItemExistsException = mint_api_exceptions.ItemExistsException(Description, Location)
End Function
Public Function ListIsEmptyException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set ListIsEmptyException = mint_api_exceptions.ListIsEmptyException(Description, Location)
End Function
Public Function BufferIsEmptyException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set BufferIsEmptyException = mint_api_exceptions.BufferIsEmptyException(Description, Location)
End Function
Public Function DisposedException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set DisposedException = mint_api_exceptions.DisposedException(Description, Location)
End Function
Public Function OpenFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set OpenFileException = mint_api_exceptions.OpenFileException(Description, Location)
End Function
Public Function CloseFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set CloseFileException = mint_api_exceptions.CloseFileException(Description, Location)
End Function
Public Function ReadFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set ReadFileException = mint_api_exceptions.ReadFileException(Description, Location)
End Function
Public Function WriteFileException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set WriteFileException = mint_api_exceptions.WriteFileException(Description, Location)
End Function
Public Function TargetNotReadyException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set TargetNotReadyException = mint_api_exceptions.TargetNotReadyException(Description, Location)
End Function
Public Function InvalidStatusException(Optional Description As String = "", Optional Location As String = "") As Exception
    Set InvalidStatusException = mint_api_exceptions.InvalidStatusException(Description, Location)
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
