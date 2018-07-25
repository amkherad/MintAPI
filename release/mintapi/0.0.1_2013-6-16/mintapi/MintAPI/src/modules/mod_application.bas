Attribute VB_Name = "mod_application"
Option Explicit

Private Type CrossBoundedObject
    hWnd As Long
    VBObject As Object

    isAPIProperty As Boolean 'it has first priority
    isVBProperty As Boolean

    VBPropertyName As String
End Type
Private Type BoundedObject
    tr As String 'Text
    obj As CrossBoundedObject
End Type
Private Type BoundedObjectsBuffer
    objs() As BoundedObject
    objsCount As Long
End Type


Public application_SignalEmitter As SignalEmitter

Public application_ap
Public application_Disposed As Boolean

Public application_developer_mode As Boolean

Public application_ln() As Language
Public application_currentLanguageID As Long
Public application_lnCount As Long
Public application_InitCommonControls As Boolean

Public application_notifyArg As Variant

Dim application_default_IOStream As ITargetStream
Dim application_default_InputStream As ITargetStream
Dim application_default_OutputStream As ITargetStream

Public tApplication As IApplication

Dim bob As BoundedObjectsBuffer
Dim theOnlyConsoleInstance As Console

Dim tmCLanguageChanged_CallBack As Timer

Dim application_operations As New Collection


Public Function InitializeAppConsole() As Console
    If theOnlyConsoleInstance Is Nothing Then _
        Set theOnlyConsoleInstance = New Console
    Set InitializeAppConsole = theOnlyConsoleInstance
End Function
Public Sub UnloadAppConsole()
    Set theOnlyConsoleInstance = Nothing
End Sub
Public Function IsAppConsoleInitialized() As Boolean
    IsAppConsoleInitialized = (theOnlyConsoleInstance Is Nothing)
End Function

Public Function CreateStdInputStream() As ITargetStream
    Dim app_io As New AppIO
    Dim app_Sio As ITargetStream
    Set app_Sio = app_io
    Call app_io.Initialize(AppIO_StdInput)
    Call app_Sio.setInState
    Call app_Sio.setOutState(False, "1@26application_input_lock_key$01")
    Set CreateStdInputStream = app_io
End Function
Public Function CreateStdOutputStream() As ITargetStream
    Dim app_io As New AppIO
    Dim app_Sio As ITargetStream
    Set app_Sio = app_io
    Call app_io.Initialize(AppIO_StdOutput)
    Call app_Sio.setOutState
    Call app_Sio.setInState(False, "1@26application_output_lock_key$02")
    Set CreateStdOutputStream = app_io
End Function
Public Function CreateStdStream() As ITargetStream
    Dim app_io As New AppIO
    Dim app_Sio As ITargetStream
    Set app_Sio = app_io
    Call app_io.Initialize(AppIO_StdStream)
    Call app_Sio.setInState
    Call app_Sio.setOutState
    Set CreateStdStream = app_io
End Function

Public Function InitializeInputStream() As ITargetStream
    If application_default_InputStream Is Nothing Then _
        Set application_default_InputStream = CreateStdInputStream
    Set InitializeInputStream = application_default_InputStream
End Function
Public Function InitializeOutputStream() As ITargetStream
    If application_default_OutputStream Is Nothing Then _
        Set application_default_OutputStream = CreateStdOutputStream
    Set InitializeOutputStream = application_default_OutputStream
End Function
Public Function InitializeStream() As ITargetStream
    If application_default_IOStream Is Nothing Then _
        Set application_default_IOStream = CreateStdStream
    Set application_default_InputStream = application_default_IOStream
    Set application_default_OutputStream = application_default_IOStream
    Set InitializeStream = application_default_IOStream
End Function

Public Sub trigStream(tStream As ITargetStream, Optional Dir As StreamDirection = StreamDirection.sdBoth)
    If (Dir And sdInStream) = sdInStream Then
        Set application_default_InputStream = tStream
    End If
    If (Dir And sdOutStream) = sdOutStream Then
        Set application_default_OutputStream = tStream
    End If
    If Dir = sdBoth Then Set application_default_IOStream = tStream
End Sub
Public Sub trigStdIO(Optional Dir As StreamDirection = StreamDirection.sdBoth)
    
End Sub
Public Sub trigConsole(Optional Dir As StreamDirection = StreamDirection.sdOutStream)
    Dim Con_1 As New Console
    Call trigStream(Con_1, Dir)
End Sub
Public Sub trigFile(File As File, Optional Dir As StreamDirection = StreamDirection.sdBoth)
    If File Is Nothing Then throw ArgumentNullException
    If Not File.TryOpen Then throw InvalidFileException("Unable To Open File.")
    Call trigStream(File)
End Sub
Public Sub trigProcess(p As Process, Optional Dir As StreamDirection = StreamDirection.sdBoth)
    
End Sub

Public Function appStream() As ITargetStream
    Set appStream = InitializeStream
End Function
Public Function appInput(Optional ForceLockStream As Boolean = False) As ITargetStream
    Set appInput = InitializeInputStream
    Call appInput.OpenStream(sdInStream)
End Function
Public Function appOutput(Optional ForceLockStream As Boolean = False) As ITargetStream
    Set appOutput = InitializeOutputStream
    Call appOutput.OpenStream(sdOutStream)
End Function

Public Function appError(Optional ForceLockStream As Boolean = False) As ITargetStream
    Dim AppIO As AppIO, tStream As ITargetStream
    Set AppIO = CreateStdOutputStream
    Set tStream = AppIO
    AppIO.AppIOType = AppIO_Error
    Call tStream.OpenStream(sdOutStream)
    Set appError = AppIO
End Function
Public Function appDebug(Optional ForceLockStream As Boolean = False) As ITargetStream
    Dim AppIO As AppIO, tStream As ITargetStream
    Set AppIO = CreateStdOutputStream
    Set tStream = AppIO
    AppIO.AppIOType = AppIO_Debug
    Call tStream.OpenStream(sdOutStream)
    Set appDebug = AppIO
End Function

Public Function appWarning(Optional ForceLockStream As Boolean = False) As ITargetStream
    Dim AppIO As AppIO, tStream As ITargetStream
    Set AppIO = CreateStdOutputStream
    Set tStream = AppIO
    AppIO.AppIOType = AppIO_Warning
    Call tStream.OpenStream(sdOutStream)
    Set appWarning = AppIO
End Function
Public Function appCritical(Optional ForceLockStream As Boolean = False) As ITargetStream
    Dim AppIO As AppIO, tStream As ITargetStream
    Set AppIO = CreateStdOutputStream
    Set tStream = AppIO
    AppIO.AppIOType = AppIO_Critical
    Call tStream.OpenStream(sdOutStream)
    Set appCritical = AppIO
End Function
Public Function appFatal(Optional ForceLockStream As Boolean = False) As ITargetStream
    Dim AppIO As AppIO, tStream As ITargetStream
    Set AppIO = CreateStdOutputStream
    Set tStream = AppIO
    AppIO.AppIOType = AppIO_Fatal
    Call tStream.OpenStream(sdOutStream)
    Set appFatal = AppIO
End Function
Public Function staticAppConsole(Optional ForceLockStream As Boolean = False) As ITargetStream
    Dim AppIO As AppIO, tStream As ITargetStream
    Set AppIO = CreateStdStream
    Set tStream = AppIO
    AppIO.AppIOType = AppIO.AppIOType Or AppIO_Console
    Call tStream.OpenStream(sdOutStream)
    Set staticAppConsole = AppIO
End Function


Public Sub out(ByVal Stream As ITargetStream, Data, Optional ByVal Length As Long = -1)
    If Stream Is Nothing Then Set Stream = appOutput
    If Not Stream.getState(sdOutStream) Then throw InvalidStatusException
    Call Stream.outStream(Data, Length)
End Sub
Public Sub inp(ByVal Stream As ITargetStream, Data, Optional ByVal Length As Long = -1)
    If Stream Is Nothing Then Set Stream = appInput
    If Not Stream.getState(sdInStream) Then throw InvalidStatusException
    Call Stream.inStream(Data, Length)
End Sub

Public Sub stdin(Data)
    inp appInput, Data
End Sub
Public Sub stdout(Data)
    out appOutput, Data
End Sub
Public Sub stderr(Data)
    out appError, Data
End Sub
Public Sub stdWarning(Data)
    out appWarning, Data
End Sub
Public Sub stdDebug(Data)
    out appDebug, Data
End Sub
Public Sub stdError(Data)
    out appError, Data
End Sub
Public Sub stdCritical(Data)
    out appCritical, Data
End Sub
Public Sub conout(Data)
    Dim sConsole As ITargetStream
    Set sConsole = New Console
    Call sConsole.OpenStream(sdOutStream)
    out sConsole, Data
End Sub
Public Sub conin(Data)
    Dim sConsole As ITargetStream
    Set sConsole = New Console
    Call sConsole.OpenStream(sdInStream)
    inp sConsole, Data
End Sub

Public Function CurrentLanguage() As Language
On Error GoTo try_retSysKey
    If application_lnCount > 0 Then
        If application_currentLanguageID = -1 Then GoTo try_retSysKey
        If application_lnCount > application_currentLanguageID Then
            Set CurrentLanguage = application_ln(application_currentLanguageID)
            Exit Function
        Else
            application_currentLanguageID = -1
            GoTo try_retSysKey
        End If
    End If
try_retSysKey:
    Dim l As Language
    Dim Key As String
    Key = LocaleLanguageName 'English
    Dim i As Long
    For i = 0 To application_lnCount - 1
        If application_ln(i).Name = Key Then
            application_currentLanguageID = i
            Set CurrentLanguage = application_ln(i)
            Call CLanguageChanged(False)
            Exit Function
        End If
    Next
    Set l = New Language
    Set CurrentLanguage = l
End Function
Public Sub SetCurrentLanguage(Name As String)
    Dim i As Long
    For i = 0 To application_lnCount - 1
        If application_ln(i).Name = Name Then
            Set CurrentLanguage = application_ln(i)
            Exit Sub
        End If
    Next
    throw ItemNotExistsException, Nothing, "SetCurrentLanguage", "Language With Specified Name Does Not Exists."
End Sub

Private Sub CLanguageChanged_CallBack()
    If Not tmCLanguageChanged_CallBack Is Nothing Then _
        tmCLanguageChanged_CallBack.EnsureUnload
    Set tmCLanguageChanged_CallBack = Nothing
    Call CLanguageChanged(True)
End Sub
Public Sub CLanguageChanged(Optional ByVal NotifyNow As Boolean = True)
    If Not NotifyNow Then
        On Error GoTo TimerErr
        If tmCLanguageChanged_CallBack Is Nothing Then _
            Set tmCLanguageChanged_CallBack = New Timer
        Dim CMethod As New Method
        Call CMethod.Initialize("CLanguageChanged_CallBack", _
            AddressOf CLanguageChanged_CallBack)
            
        Call tmCLanguageChanged_CallBack.Initialize(CMethod)
        Call tmCLanguageChanged_CallBack.Start
        
        Exit Sub
TimerErr: tmCLanguageChanged_CallBack.EnsureUnload: Exit Sub
    End If
    
    If Not tApplication Is Nothing Then
        On Error GoTo tApplicationErr
        Call tApplication.ApplicationLanguageChanged(EventArgs(CurrentLanguage))
tApplicationErr:
    End If
End Sub

Public Sub CreateManifestFor(Path As String, App As IApplication)
    
End Sub

Public Sub boundtranslationto(Key As String, Object As Object, targetProperty As String, SetNow As Boolean)
    bob.objsCount = 0
End Sub
Public Sub unboundtranslationto(Object As Object, Optional targetProperty As String)
    
End Sub

'=============================================
' Operation Segment
'=============================================

Public Function IndexOfOperation(Key As String) As Long
    Dim iC As Long
    Dim OPV As Variant
    Dim Opr As Operation
    For Each OPV In application_operations
        Set Opr = OPV
        If Opr.Key = Key Then
            IndexOfOperation = iC
            Exit Function
        End If
        iC = iC + 1
    Next
    IndexOfOperation = -1
End Function

Public Sub AddOperation(Key As String, State As Boolean, Optional Args As ArgumentList = Nothing, Optional ReadOnly As Boolean = False)
    If IndexOfOperation(Key) <> -1 Then throw ItemExistsException("Operation already exists.")
    Dim Opr As New Operation
    Call Opr.Initialize(Key, State, Args, ReadOnly)
    Call application_operations.Add(Opr)
End Sub
Public Sub RemoveOperation(Key As String)
    Dim Index As Long
    Index = IndexOfOperation(Key)
    If Index = -1 Then throw ItemExistsException("Operation does not exists.")
    Call application_operations.Remove(Index + 1)
End Sub

Public Function Operation(Key As String) As Operation
    Dim OPV As Variant
    Dim Opr As Operation
    For Each OPV In application_operations
        Set Opr = OPV
        If Opr.Key = Key Then
            Set Operation = Opr
            Exit Function
        End If
    Next
    
    Set Operation = New Operation
    Call Operation.Initialize(Key, False)
    Call Operation.MakeAbstractOperation
End Function

Public Function Operations() As List
    Dim lst As New List
    Dim OPV As Variant
    For Each OPV In application_operations
        Call lst.Add(OPV)
    Next
    Set Operations = lst
End Function

Public Sub mint_garbagec_CleanOperations()
    Set application_operations = New Collection
End Sub
Public Sub mint_garbagec_CleanBoundTR()
    
End Sub
