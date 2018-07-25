VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "AppIO"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements ITargetStream
Implements IAliasable


Public Enum AppIOType
    AppIO_StdOutput = &H0
    AppIO_StdInput = &H1
    AppIO_StdError = &H2
    AppIO_StdInnerException = &H4
    AppIO_StdOSException = &H8
    AppIO_Console = &H10

    AppIO_Debug = AppIO_StdOutput
    AppIO_Warning = AppIO_StdOutput
    AppIO_Error = AppIO_StdOutput Or AppIO_StdError
    AppIO_Critical = AppIO_Error Or AppIO_StdInnerException
    AppIO_Fatal = AppIO_Error Or AppIO_StdOSException
    AppIO_System = AppIO_Fatal

    AppIO_StdStream = AppIO_Console
End Enum

Dim app_iotype As AppIOType
Dim inState As Boolean
Dim outState As Boolean

Dim inState_LOCK As String
Dim outState_LOCK As String

Dim isOpened As Boolean

Dim output_File As File

Friend Sub Initialize(AppIOType As AppIOType)
    app_iotype = AppIOType
End Sub
Friend Sub SetFile(File As File)
    Set output_File = File
End Sub

Friend Function IsAppIOStdInput() As Boolean
    IsAppIOStdInput = ((app_iotype And AppIO_StdInput) = AppIO_StdInput)
End Function
Friend Function IsAppIOStdOutput() As Boolean
    IsAppIOStdOutput = ((app_iotype And AppIO_StdOutput) = AppIO_StdOutput)
End Function
Friend Function IsAppIOConsole() As Boolean
    IsAppIOConsole = ((app_iotype And AppIO_Console) = AppIO_Console)
End Function
Friend Function IsAppIOStdError() As Boolean
    IsAppIOStdError = ((app_iotype And AppIO_StdError) = AppIO_StdError)
End Function
Friend Function IsAppIOStdInnerException() As Boolean
    IsAppIOStdInnerException = ((app_iotype And AppIO_StdInnerException) = AppIO_StdInnerException)
End Function
Friend Function IsAppIOStdOSException() As Boolean
    IsAppIOStdOSException = ((app_iotype And AppIO_StdOSException) = AppIO_StdOSException)
End Function

Public Property Get AppIOType() As AppIOType
    AppIOType = app_iotype
End Property
Public Property Let AppIOType(Value As AppIOType)
    app_iotype = Value
End Property

Private Function ITargetStream_Available(direction As StreamDirection) As Long
    ITargetStream_Available = IIf(direction = sdInStream, inState And IsAppIOStdInput, outState And (Not IsAppIOStdInput))
End Function

Private Sub ITargetStream_CloseStream(Optional direction As StreamDirection = 3&)
    If (direction And sdInStream) = sdInStream Then
        inState = False
    End If
    If (direction And sdOutStream) = sdOutStream Then
        outState = False
    End If
    If direction = sdBoth Then isOpened = False
End Sub

Private Function ITargetStream_getState(Dir As StreamDirection) As Boolean
    ITargetStream_getState = IIf(Dir = sdInStream, inState, outState)
End Function

Private Sub ITargetStream_OpenStream(Optional direction As StreamDirection = 3&)
    isOpened = True
    If (direction And sdInStream) = sdInStream Then
        If IsAppIOStdInput Then
            inState = True
            isOpened = True
        Else
            throw InvalidStatusException("Unable To Open Input Stream While Not AppIOStdInput.")
        End If
    End If
    If (direction And sdOutStream) = sdOutStream Then
        If IsAppIOStdOutput Then
            If Not output_File Is Nothing Then
                Call output_File.TryOpen
            End If
            outState = True
            isOpened = True
        Else
            throw InvalidStatusException("Unable To Open Output Stream While Not AppIOStdInput.")
        End If
    End If
End Sub

Private Sub ITargetStream_inStream(Data, Optional Length As Long = -1&)
    If Not isOpened Then throw TargetNotOpenedException
    If Not inState Then throw InvalidStatusException("At ITargetStream<AppIO>::inStream()")
    Dim str_read_buffer As String
    If IsAppIOStdInput Then
        If IsAppIOConsole Then
            str_read_buffer = ReadFromConsole(AppConsole)
        Else
            str_read_buffer = ReadFromStdInput
        End If
    Else
        throw InvalidStatusException("At ITargetStream<AppIO>::inStream()")
    End If
    Call mint_put_byte_array_to(Data, str_read_buffer)
End Sub
Private Function ITargetStream_readBytes(Length As Long, Optional StartAt As Long = 0) As Byte()
    
End Function
Private Sub ITargetStream_outStream(Data, Optional Length As Long = -1&)
    If Not isOpened Then throw TargetNotOpenedException
    If Not outState Then throw InvalidStatusException("At ITargetStream<AppIO>::inStream()")
    Dim output_buffer As String
    output_buffer = ByteArray(Data).toString()

    If Not output_File Is Nothing Then _
        Call output_File.WriteLine(output_buffer)

    If IsAppIOStdOSException Then
        Call PrintToStdError(output_buffer)
        'this will raise a operating system exception and end application.
        throw SystemException(output_buffer)
    ElseIf IsAppIOStdInnerException Then
        Call PrintToStdError(output_buffer)
        throw Exception(output_buffer)
    ElseIf IsAppIOStdOutput Then
        Call PrintToStdOutput(output_buffer)
    ElseIf IsAppIOStdError Then
        Call PrintToStdError(output_buffer)
    ElseIf IsAppIOConsole Then
        Call PrintToConsole(AppConsole, output_buffer)
    Else
        Call PrintToStdOutput(output_buffer)
    End If
End Sub

Private Sub ITargetStream_setInState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, inState, inState_LOCK, Reserved)
End Sub
Private Sub ITargetStream_setOutState(Optional ByVal State As Boolean = True, Optional Reserved As Variant)
    Call mint_setstream_state(State, outState, outState_LOCK, Reserved)
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
