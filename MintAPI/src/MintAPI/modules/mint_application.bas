Attribute VB_Name = "mint_application"
Option Explicit

Private Declare Function API_EbMode Lib "vba6" Alias "EbMode" () As Long

Public Type AppStreams
    AppOutput As IClassStream
    AppInput As IClassStream
    AppError As IClassStream
    AppGeneric As IClassStream 'Obsolete
End Type
Public Type ApplicationInformation
    SignalEmitter As SignalEmitter
    Streams As AppStreams
    UserApp As App
    Terminated As Boolean
    Languages As List
    CurrentTranslation As Translation
    NotifyArg As Variant
    ApplicationStartType As ApplicationStartConstants
    TargetApplication As IApplication
    DevelopmentEnvironmentState As AppDevelopEnvState
    Operations As Collection
End Type

Public Type CrossBoundedObject
    hWnd As Long
    VBObject As Object

    IsAPIProperty As Boolean 'it has first priority
    IsVBProperty As Boolean

    VBPropertyName As String
End Type

Public Enum AppDevelopEnvState
    adesIDE = &H1000
    adesRuntime = 1
    adesBreakMode = adesIDE Or 2
    adesDesignTime = adesIDE Or adesBreakMode Or adesRuntime
End Enum

Public AppInfo As ApplicationInformation

Public Sub Construct()
    '1.In IDE if VB6.EXE & Title Microsoft Visual Basic
    Debug.Assert SetInIDE
    Call SetIfInDesignTime
End Sub

Private Function SetInIDE() As Boolean 'by Kelly Ethridge added on 7/30/2013
    AppInfo.DevelopmentEnvironmentState = adesIDE Or adesRuntime
    SetInIDE = True
End Function
Private Function IsVB6EXE() As Boolean
    Dim RetVal As String
    RetVal = String$(XLARGELPSTR, 0)
    
    Call API_GetModuleFileName(vbNullPtr, RetVal, XLARGELPSTR)
    
    RetVal = GetLPSTR(RetVal)
    
    If (UCase$(Right$(RetVal, 8)) = "\VB6.EXE") Then IsVB6EXE = True
End Function
Private Sub SetIfInDesignTime()
    If IsVB6EXE Then
        If (API_GetModuleHandle("vba6") <> vbNullPtr) Then
            Dim EbVal As Long
            On Error GoTo CatchErr
            EbVal = API_EbMode
            If EbVal = 2 Then
                AppInfo.DevelopmentEnvironmentState = AppInfo.DevelopmentEnvironmentState Or adesBreakMode
            Else
CatchErr:
                AppInfo.DevelopmentEnvironmentState = AppInfo.DevelopmentEnvironmentState Or adesDesignTime
            End If
        Else
            AppInfo.DevelopmentEnvironmentState = AppInfo.DevelopmentEnvironmentState Or adesDesignTime
        End If
    Else
        AppInfo.DevelopmentEnvironmentState = AppInfo.DevelopmentEnvironmentState Or adesRuntime
    End If
End Sub

''<summary>Determines that MintAPI is running on IDE by F5.</summary>
''<retval>A boolean to determine that application is running on IDE.</retval>
Public Property Get IsInIDE() As Boolean
    IsInIDE = ((AppInfo.DevelopmentEnvironmentState And adesIDE) = adesIDE)
End Property
''<summary>Determines that application is executing at design time in VB instance.</summary>
''<retval>A boolean to determine that application is executing in VB debugger.</retval>
Public Property Get IsDesignMode() As Boolean
    IsDesignMode = AppInfo.DevelopmentEnvironmentState = adesDesignTime
End Property
''<summary>Determines that application is executing as a stand-alone application (not in debugger).</summary>
''<retval>A boolean to determine that application is executing as a stand-alone application.</retval>
Public Property Get IsRuntime() As Boolean
    'IsRuntime = ((AppInfo.DevelopmentEnvironmentState And adesRuntime) = adesRuntime)
    IsRuntime = (AppInfo.DevelopmentEnvironmentState = adesRuntime)
End Property
Public Sub VirtualizeDesignMode()
    AppInfo.DevelopmentEnvironmentState = AppInfo.DevelopmentEnvironmentState Or adesDesignTime
End Sub
Public Sub EndDebugging()
    Stop
End Sub

'==============================================

Private Function InitializeStdInput() As IClassStream
    Dim AppStream As New ApplicationStream
    Call AppStream.Constructor0(API_GetStdHandle(STD_INPUT_HANDLE), sdInStream, "Standard Input")
    Set InitializeStdInput = AppStream
    Call InitializeStdInput.OpenStream(sdInStream)
End Function
Private Function InitializeStdOutput() As IClassStream
    Dim AppStream As New ApplicationStream
    Call AppStream.Constructor0(API_GetStdHandle(STD_OUTPUT_HANDLE), sdOutStream, "Standard Output")
    Set InitializeStdOutput = AppStream
    Call InitializeStdOutput.OpenStream(sdOutStream)
End Function
Private Function InitializeStdError() As IClassStream
    Dim AppStream As New ApplicationStream
    Call AppStream.Constructor0(API_GetStdHandle(STD_ERROR_HANDLE), sdOutStream, "Standard Error")
    Set InitializeStdError = AppStream
    Call InitializeStdError.OpenStream(sdOutStream)
End Function

Private Function InitializeInputStream() As IClassStream
    If AppInfo.Streams.AppInput Is Nothing Then
        Set AppInfo.Streams.AppInput = InitializeStdError
    End If
    Set InitializeInputStream = AppInfo.Streams.AppInput
End Function
Private Function InitializeOutputStream() As IClassStream
    If AppInfo.Streams.AppOutput Is Nothing Then
        Set AppInfo.Streams.AppOutput = InitializeStdOutput
    End If
    Set InitializeOutputStream = AppInfo.Streams.AppOutput
End Function
Private Function InitializeErrorStream() As IClassStream
    If AppInfo.Streams.AppError Is Nothing Then
        Set AppInfo.Streams.AppError = InitializeStdError
    End If
    Set InitializeErrorStream = AppInfo.Streams.AppError
End Function

Public Sub trigStream(ByVal tStream As IClassStream, Optional ByVal Direction As ApplicationStreams = ApplicationStreams.asInputOutput)
    Dim IsSet As Boolean
    If (Direction And asInput) = asInput Then
        Set AppInfo.Streams.AppInput = tStream
        IsSet = True
    End If
    If (Direction And asOutput) = asOutput Then
        Set AppInfo.Streams.AppOutput = tStream
        IsSet = True
    End If
    If (Direction And asError) = asError Then
        Set AppInfo.Streams.AppError = tStream
        IsSet = True
    End If
    If Direction = asInputOutput Then
        Set AppInfo.Streams.AppGeneric = tStream
        IsSet = True
    End If
    If Not IsSet Then _
        throw Exps.InvalidArgumentException(Mtr("Invalid direction value."))
End Sub
Public Sub trigStdIO(Optional ByVal Direction As ApplicationStreams = ApplicationStreams.asInputOutput)
    Select Case Direction
        Case ApplicationStreams.asInput
            Call trigStream(InitializeStdInput, asInput)
        Case ApplicationStreams.asOutput
            Call trigStream(InitializeStdOutput, asOutput)
        Case ApplicationStreams.asInputOutput
            Call trigStream(InitializeStdInput, asInput)
            Call trigStream(InitializeStdOutput, asOutput)
        Case ApplicationStreams.asError
            Call trigStream(InitializeStdError, asError)
        Case Else
            throw Exps.InvalidArgumentException("Direction")
    End Select
End Sub
Public Sub trigConsole(Optional ByVal Direction As ApplicationStreams = ApplicationStreams.asInputOutput)
    Call API_AllocConsole
    Dim ConsoleStream As ApplicationStream
    Select Case Direction
        Case ApplicationStreams.asInput
            Call trigStream(InitializeStdInput, asInput)
        Case ApplicationStreams.asOutput
            Call trigStream(InitializeStdOutput, asOutput)
        Case ApplicationStreams.asInputOutput
            Call trigStream(InitializeStdInput, asInput)
            Call trigStream(InitializeStdOutput, asOutput)
        Case ApplicationStreams.asError
            Call trigStream(InitializeStdError, asError)
        Case Else
            throw Exps.InvalidArgumentException("Direction")
    End Select
End Sub
Public Sub trigFile(ByVal File As File, Optional ByVal Dir As ApplicationStreams = ApplicationStreams.asInputOutput)
    
End Sub
Public Sub trigProcess(ByVal Proc As Process, Optional ByVal Dir As ApplicationStreams = ApplicationStreams.asInputOutput)
    
End Sub

Public Function AppInput() As IClassStream
    Set AppInput = InitializeInputStream
    Call AppInput.OpenStream(sdInStream)
End Function
Public Function AppOutput() As IClassStream
    Set AppOutput = InitializeOutputStream
    Call AppOutput.OpenStream(sdOutStream)
End Function
Public Function AppError() As IClassStream
    Set AppError = InitializeErrorStream
    Call AppError.OpenStream(sdOutStream)
End Function
'Public Function AppDebug() As IClassStream
'    Set AppDebug = AppError
'End Function
'Public Function AppWarning() As IClassStream
'    Set AppWarning = AppError
'End Function
'Public Function AppCritical() As IClassStream
'    Set AppCritical = AppError
'End Function
'Public Function AppFatal() As IClassStream
'    Set AppFatal = AppError
'End Function

Public Sub out(ByVal Stream As IClassStream, ByRef Data As Variant, Optional ByVal Length As Long = -1)
    If Stream Is Nothing Then Set Stream = AppOutput
    If Not Stream.GetState(sdOutStream) Then throw Exps.InvalidOperationException
    Call Stream.OutStream(Data, Length)
End Sub
Public Sub inp(ByVal Stream As IClassStream, ByRef Data As Variant, Optional ByVal Length As Long = -1)
    If Stream Is Nothing Then Set Stream = AppInput
    If Not Stream.GetState(sdInStream) Then throw Exps.InvalidOperationException
    Call Stream.InStream(Data, Length)
End Sub

Public Function EmitForApp(Signal, ParamArray Args() As Variant)
    
End Function

'Public Sub stdin(ByRef Data As Variant)
'    inp AppInput, Data
'End Sub
'Public Sub stdout(ByRef Data As Variant)
'    out AppOutput, Data
'End Sub
'Public Sub stderr(ByRef Data As Variant)
'    out AppError, Data
'End Sub
'Public Sub stdWarning(ByRef Data As Variant)
'    out AppWarning, Data
'End Sub
'Public Sub stdDebug(ByRef Data As Variant)
'    out AppDebug, Data
'End Sub
'Public Sub stdError(ByRef Data As Variant)
'    out AppError, Data
'End Sub
'Public Sub stdCritical(ByRef Data As Variant)
'    out AppCritical, Data
'End Sub
'Public Sub conout(ByRef Data As Variant)
'    Dim sConsole As IClassStream
'    Set sConsole = New Console
'    Call sConsole.OpenStream(sdOutStream)
'    out sConsole, Data
'End Sub
'Public Sub conin(ByRef Data As Variant)
'    Dim sConsole As IClassStream
'    Set sConsole = New Console
'    Call sConsole.OpenStream(sdInStream)
'    inp sConsole, Data
'End Sub

Public Function CurrentTranslation() As Translation
'On Error GoTo try_retSysKey
'    If application_lnCount > 0 Then
'        If application_currentLanguageID = -1 Then GoTo try_retSysKey
'        If application_lnCount > application_currentLanguageID Then
'            Set CurrentLanguage = application_ln(application_currentLanguageID)
'            Exit Function
'        Else
'            application_currentLanguageID = -1
'            GoTo try_retSysKey
'        End If
'    End If
'try_retSysKey:
    Dim tr As Translation
'    Dim Key As String
'    Key = LocaleLanguageName 'English
'    Dim i As Long
'    For i = 0 To application_lnCount - 1
'        If application_ln(i).Name = Key Then
'            application_currentLanguageID = i
'            Set CurrentLanguage = application_ln(i)
'            Call CLanguageChanged(False)
'            Exit Function
'        End If
'    Next
    Set tr = New Translation
    Set CurrentTranslation = tr
End Function
Public Sub SetCurrentTranslation(ByVal Name As String)
'    Dim i As Long
'    For i = 0 To application_lnCount - 1
'        If application_ln(i).Name = Name Then
'            Set CurrentLanguage = application_ln(i)
'            Exit Sub
'        End If
'    Next
'    throw Exps.ItemNotExistsException, Nothing, "SetCurrentLanguage", "Language With Specified Name Does Not Exists."
End Sub

Private Sub CLanguageChanged_CallBack()
'    If Not tmCLanguageChanged_CallBack Is Nothing Then _
'        tmCLanguageChanged_CallBack.EnsureUnload
'    Set tmCLanguageChanged_CallBack = Nothing
'    Call CLanguageChanged(True)
End Sub
Public Sub CLanguageChanged(Optional ByVal NotifyNow As Boolean = True)
'    If Not NotifyNow Then
'        On Error GoTo TimerErr
'        If tmCLanguageChanged_CallBack Is Nothing Then _
'            Set tmCLanguageChanged_CallBack = New Timer
'        Dim CMethod As New Method
'        Call CMethod.Constructor0("CLanguageChanged_CallBack", _
'            AddressOf CLanguageChanged_CallBack)
'
'        Call tmCLanguageChanged_CallBack.Initialize(CMethod)
'        Call tmCLanguageChanged_CallBack.Start
'
'        Exit Sub
'TimerErr: tmCLanguageChanged_CallBack.EnsureUnload: Exit Sub
'    End If
'
'    If Not tApplication Is Nothing Then
'        On Error GoTo tApplicationErr
'        Call tApplication.ApplicationLanguageChanged(EventArgs(CurrentLanguage))
'tApplicationErr:
'    End If
End Sub

Public Sub bindtranslationto(Key As String, Object As Object, targetProperty As String, SetNow As Boolean)
    'BOB.objsCount = 0
End Sub
Public Sub unboundtranslation(Object As Object, Optional targetProperty As String)
    
End Sub

'=============================================
' Operation Segment
'=============================================

'Public Function IndexOfOperation(Key As String) As Long
''    Dim iC As Long
''    Dim OPV As Variant
''    Dim Opr As Operation
''    For Each OPV In application_operations
''        Set Opr = OPV
''        If Opr.Key = Key Then
''            IndexOfOperation = iC
''            Exit Function
''        End If
''        iC = iC + 1
''    Next
''    IndexOfOperation = -1
'End Function

'Public Sub AddOperation(Key As String, State As Boolean, Optional Args As ArgumentList = Nothing, Optional ReadOnly As Boolean = False)
''    If IndexOfOperation(Key) <> -1 Then throw Exps.ItemExistsException("Operation already exists.")
''    Dim Opr As New Operation
''    Call Opr.Initialize(Key, State, Args, ReadOnly)
''    Call application_operations.Add(Opr)
'End Sub
'Public Sub RemoveOperation(Key As String)
''    Dim Index As Long
''    Index = IndexOfOperation(Key)
''    If Index = -1 Then throw Exps.ItemExistsException("Operation does not exists.")
''    Call application_operations.Remove(Index + 1)
'End Sub
'
'Public Function Operation(Key As String) As Operation
''    Dim OPV As Variant
''    Dim Opr As Operation
''    For Each OPV In application_operations
''        Set Opr = OPV
''        If Opr.Key = Key Then
''            Set Operation = Opr
''            Exit Function
''        End If
''    Next
''
''    Set Operation = New Operation
''    Call Operation.Initialize(Key, False)
''    Call Operation.MakeAbstractOperation
'End Function

'Public Function Operations() As List
''    Dim Lst As New List
''    Dim OPV As Variant
''    For Each OPV In application_operations
''        Call Lst.Append(OPV)
''    Next
''    Set Operations = Lst
'End Function
