VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Contains basic methods for application including StartApplication and TerminateApplication methods.\r\nStartApplication start an IApplication inherited class as an application ,this provides some MintAPI specified controls such as log and debug engine ,application globalization ,and error controllers.\r\nIf you use an IApplication class to create your application ,Application class provides you some features such as rasing event on ApplicationNetworkStateChange,ApplicationInputMethodChaged and more usefull attributes."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Operation"
Attribute VB_Ext_KEY = "Member1" ,"GarbageCollector"
'@PROJECT_LICENSE


'Ideas from Me , Qt , .net Framework

Option Base 0
Option Explicit
Const CLASSID As String = "Application"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Const SIGNAL_START As String = "start"
Const SIGNAL_TERMINATING As String = "terminating"
Const SIGNAL_TERMINATED As String = "terminated"
Const SIGNAL_DEVELOPERMODEON As String = "developermodeon"
Const SIGNAL_DEVELOPERMODEOFF As String = "developermodeoff"
Const SIGNAL_CONSOLETRIGGED As String = "consoletrigged"
Const SIGNAL_FILETRIGGED As String = "filetrigged"
Const SIGNAL_PROCESSTRIGGED As String = "processtrigged"
Const SIGNAL_STREAMTRIGGED As String = "streamtrigged"
Const SIGNAL_STDIOTRIGGED As String = "stdiotrigged"
Const SIGNAL_NOTHINGTRIGGED As String = "nothingtrigged"
Const SIGNAL_PROCESSEVENTS As String = "processevents"
Const SIGNAL_LANGUAGEREGISTERED As String = "languageregistered"
Const SIGNAL_LANGUAGEPATHREGISTERED As String = "languagepathregistered"
Const SIGNAL_LANGUAGEUNREGISTERED As String = "languageunregistered"
Const SIGNAL_VISUALSTYLESENABLED As String = "visualstylesenabled"
Const SIGNAL_GLOBALNOTIFICATION As String = "globalnotification"

Const SIGNAL_EVENTRECIEVED As String = "event"

Implements IAliasable
Implements ISupportSignal
Implements ISupportSlot

Const SLOT_TERMINATE As String = "terminate"
Const SLOT_EXTERMINATE As String = "exterminate"

Public Enum ApplicationType
    ActiveXHost = 1
    OwnConsole = 2
    Service = 4
    StaticLinkLibrary = ActiveXHost Or &H10000000
    DynamicLinkLibrary = ActiveXHost Or &H20000000
    ApplicationExecutable = &H40000000
    ApplicationExecutableConsole = ApplicationExecutable Or OwnConsole
    ApplicationExecutable_ActiveXHost = ActiveXHost Or ApplicationExecutable
    ApplicationExecutable_ActiveXHostConsole = ApplicationExecutable_ActiveXHost Or OwnConsole
End Enum


Dim WithEvents MintAPI_Application_Impl As ImplementedApplication
Attribute MintAPI_Application_Impl.VB_VarHelpID = -1

Dim mySignalEmitter As SignalEmitter

Private Sub Class_Initialize()
    If application_SignalEmitter Is Nothing Then _
        Set application_SignalEmitter = New SignalEmitter
    Set mySignalEmitter = application_SignalEmitter
End Sub

Friend Sub Initialize(Optional Arguments)
    
End Sub


Public Sub Run(ByVal Application As IApplication, ByVal TargetApplicationType As ApplicationType, App, Args)
    If Application Is Nothing Then throw Exps.ArgumentNullException("Application")
    If Not TypeOf App Is VB.App Then throw Exps.InvalidArgumentTypeException
    Set tApplication = Application
    If VarType(App) = VBObject Then
        Set application_ap = App
    Else
        application_ap = App
    End If

    If (TargetApplicationType And OwnConsole) = OwnConsole Then _
        Call trigConsole(sdBoth)

    On Error GoTo errStarting
    Call Application.Starting(EventArgs(Me))
    GoTo continueStarting
errStarting:
    throw Exps.Exception(Err.Description)
continueStarting:

    application_Disposed = False
    Call Application.Main(Args)
    Call Application.Started(EventArgs(Me))
    Call mySignalEmitter.Emit(SIGNAL_START)
End Sub
Public Property Get InitedApplication() As IApplication
    Set InitedApplication = tApplication
End Property
Public Function Terminate(Optional ByVal IgnoreCancel As Boolean = False) As Boolean
    Call mySignalEmitter.Emit(SIGNAL_TERMINATING)
    If application_Disposed Then
        Terminate = True
        Exit Function
    End If
    If tApplication Is Nothing Then
        Terminate = False
        Exit Function
    End If
    Dim ContinueValue As Boolean
    ContinueValue = True
    Dim d As CancelEventArgs
    Set d = CancelEventArgs(tApplication, ContinueValue)
    Call tApplication.Terminating(d)
    If Not IgnoreCancel Then
        If d.Cancel Then
            Terminate = False
            Exit Function
        End If
    End If
    On Error GoTo terminatedError
    Call tApplication.Terminated(EventArgs(tApplication))
terminatedError:
    application_Disposed = True
    Terminate = True

    If mint_application.IsDesignMode Then
        Debug.Print "Terminating application..."
        Call EndDebugging
    Else
        Call Process.EndProcess(0)
    End If
    Call mySignalEmitter.Emit(SIGNAL_TERMINATED)
End Function
Public Sub Exterminate()
    If Not Terminate(True) Then
        throw Exps.Exception("Unable to dispose the application.")
    End If
End Sub

'Public Sub EnterDeveloperMode()
'    application_developer_mode = True
'    Call mySignalEmitter.Emit(SIGNAL_DEVELOPERMODEON)
'End Sub
'Public Sub LeaveDeveloperMode()
'    application_developer_mode = False
'    Call mySignalEmitter.Emit(SIGNAL_DEVELOPERMODEOFF)
'End Sub
'Public Property Get IsInIDE() As Boolean
'    IsInIDE = mint_application.IsInIDE
'End Property
Public Property Get IsDesignMode() As Boolean
    IsDesignMode = mint_application.IsDesignMode
End Property
Public Property Get IsRuntime() As Boolean
    IsRuntime = mint_application.IsRuntime
End Property

Public Sub EnableLibraryCalls()

End Sub
Public Sub DisableLibraryCalls()

End Sub

Public Function CreateTestingInstance() As Object
    Dim API As New API
    Set CreateTestingInstance = Objects.CreateObjectBuffer(OBJECTBUFFER_CONTROLLERINSTANCE, _
            API, Me)
End Function

Public Function InitializeNewApplication(Optional Process As Process, Optional AppDomain As String, _
        Optional AppName As String, Optional ProductName As String, _
        Optional AppPath As String, Optional BinaryPath As String, _
        Optional ExecutableName As String, Optional AppRegistryConfigPath As String, _
        Optional AppRegistryPath As String, Optional AppConfigPath As String, _
        Optional AppDataPath As String, Optional AppLogPath As String, _
        Optional AppPluginsPath As String, Optional AppTempPath As String, _
        Optional CompanyName As String, _
        Optional MajorVersion As Long, Optional MinorVersion As Long, _
        Optional Revision As Long, Optional VersionTags As String, _
        Optional Copyright As String, Optional Author As String, _
        Optional ProductCode As String, Optional strGUID As String, _
        Optional Args, Optional App, Optional AppWebsite As String, _
        Optional AppHelpLink As String, Optional AppUpdateLink As String, _
        Optional AppSupportLink As String, Optional AppServiceWebsite As String) _
            As ImplementedApplication
    '-------------------------------------------

    Dim ImplApp As New ImplementedApplication
    '-------------------------------------------ImplApp.Initialize
    Call ImplApp.Initialize(Process, AppDomain, _
        AppName, ProductName, AppPath, BinaryPath, ExecutableName, _
        AppRegistryPath, AppRegistryConfigPath, AppConfigPath, _
        AppDataPath, AppLogPath, AppPluginsPath, AppTempPath, _
        CompanyName, MajorVersion, MinorVersion, Revision, _
        VersionTags, Copyright, Author, ProductCode, strGUID, _
        Args, App, AppWebsite, AppHelpLink, AppUpdateLink, _
        AppSupportLink, AppServiceWebsite)
    '-------------------------------------------ImplApp.Initialize
    Set InitializeNewApplication = ImplApp
End Function

Public Sub trigConsole(Optional Dir As StreamDirection = StreamDirection.sdOutStream)
    Call mint_application.trigConsole(Dir)
    Call mySignalEmitter.Emit(SIGNAL_CONSOLETRIGGED)
End Sub
Public Sub trigFile(File As File, Optional Dir As StreamDirection = StreamDirection.sdOutStream)
    Call mint_application.trigFile(File, Dir)
    Call mySignalEmitter.Emit(SIGNAL_FILETRIGGED)
End Sub
Public Sub trigProcess(p As Process, Optional Dir As StreamDirection = StreamDirection.sdOutStream)
    Call mint_application.trigProcess(p, Dir)
    Call mySignalEmitter.Emit(SIGNAL_PROCESSTRIGGED)
End Sub
Public Sub trigStream(tStream As IClassStream, Optional Dir As StreamDirection = StreamDirection.sdOutStream)
    Call mint_application.trigStream(tStream, Dir)
    If tStream Is Nothing Then
        Call mySignalEmitter.Emit(SIGNAL_NOTHINGTRIGGED)
    Else
        Call mySignalEmitter.Emit(SIGNAL_STREAMTRIGGED)
    End If
End Sub
Public Sub trigStdIO(Optional Dir As StreamDirection = StreamDirection.sdOutStream)
    Call mint_application.trigStdIO(Dir)
    Call mySignalEmitter.Emit(SIGNAL_STDIOTRIGGED)
End Sub


Public Function appInput(Optional ForceLockStream As Boolean = False) As IClassStream
    Set appInput = mint_application.appInput(ForceLockStream)
End Function
Public Function appOutput(Optional ForceLockStream As Boolean = False) As IClassStream
    Set appOutput = mint_application.appOutput(ForceLockStream)
End Function
Public Function appError(Optional ForceLockStream As Boolean = False) As IClassStream
    Set appError = mint_application.appError(ForceLockStream)
End Function
Public Function appDebug(Optional ForceLockStream As Boolean = False) As IClassStream
    Set appDebug = mint_application.appDebug(ForceLockStream)
End Function
Public Function appWarning(Optional ForceLockStream As Boolean = False) As IClassStream
    Set appWarning = mint_application.appWarning(ForceLockStream)
End Function
Public Function appCritical(Optional ForceLockStream As Boolean = False) As IClassStream
    Set appCritical = mint_application.appCritical(ForceLockStream)
End Function
Public Function appFatal(Optional ForceLockStream As Boolean = False) As IClassStream
    Set appFatal = mint_application.appFatal(ForceLockStream)
End Function

Public Property Get IsApplicationDisposed() As Boolean
    IsApplicationDisposed = application_Disposed
End Property

Public Sub ProccessEvents(Optional TreatAsIdleTime As Boolean = False)
    DoEvents
    Call mySignalEmitter.Emit(SIGNAL_PROCESSEVENTS)
End Sub


Public Sub Notify(Arg)
    If IsObject(Arg) Then
        Set application_notifyArg = Arg
    Else
            application_notifyArg = Arg
    End If
    Call mySignalEmitter.Emit(SIGNAL_GLOBALNOTIFICATION)
End Sub

Public Property Get NotifierArgumentType() As VbVarType
    NotifierArgumentType = VarType(application_notifyArg)
End Property
Public Property Get NotifierArgument() As Variant
    If IsObject(application_notifyArg) Then
        Set NotifierArgument = application_notifyArg
    Else
            NotifierArgument = application_notifyArg
    End If
End Property
Public Property Let NotifierArgument(Value As Variant)
    If IsObject(Value) Then
        Set application_notifyArg = Value
    Else
            application_notifyArg = Value
    End If
End Property



Public Function LastException() As Exception
    Set LastException = mint_exceptions.LastException
End Function

Private Sub CheckIfNull()
    If tApplication Is Nothing Then _
        throw Exps.InvalidStatusException("No Application Is Started.")
End Sub

Public Property Get Name() As String
CheckIfNull
    Name = tApplication.ApplicationName
End Property
Public Property Get MintAPI() As IApplication
    If MintAPI_Application_Impl Is Nothing Then
        Set MintAPI_Application_Impl = New ImplementedApplication
MintAPI_Initializer:
        Call MintAPI_Application_Impl.Initialize( _
        Process.Current, APPLICATIONDOMAIN, APPLICATIONID, APPLICATIONID, _
        App.Path, ConcatPath(App.Path, App.EXEName), App.EXEName, _
        APP_REGISTRYPATH, APP_REGISTRYPATH, mint_config.ConfigPath, _
        mint_config.MintAPIDataPath, mint_config.MintAPILogPath, _
        mint_config.MintAPIPluginsPath, mint_config.MintAPITempPath, _
        App.CompanyName, App.Major, App.Minor, App.Revision, _
        APP_VERSIONTAGS, App.LegalCopyright, App.LegalTrademarks, _
        APP_PRODUCTCODE, APP_GUID, Command, App, APP_WEBSITE, _
        APP_HELPLINK, APP_UPDATELINK, APP_SUPPORTLINK, _
        APP_SERVICEWEBSITE, False)
    End If
    If Not MintAPI_Application_Impl.Initialized Then _
        GoTo MintAPI_Initializer
    Set MintAPI = MintAPI_Application_Impl
End Property
Public Property Get Version() As Long
CheckIfNull
    Dim maj_Version As Long
    Dim min_Version As Long
    Dim rev_Version As Long
    maj_Version = tApplication.VersionMajor
    min_Version = tApplication.VersionMinor
    rev_Version = tApplication.VersionRevision
    If maj_Version > 255 Then GoTo outOfRangeExp_Err
    If min_Version > 255 Then GoTo outOfRangeExp_Err
    If rev_Version > 65535 Then GoTo outOfRangeExp_Err
    maj_Version = ShiftLeft(maj_Version And &HFF, 24)
    min_Version = ShiftLeft(min_Version And &HFF, 16)
    rev_Version = rev_Version And &HFFFF
    Version = (maj_Version + min_Version + rev_Version)
    Exit Property
outOfRangeExp_Err:
    Version = 0
End Property
Public Property Get VersionString() As String
CheckIfNull
    Dim maj_Version As Long
    Dim min_Version As Long
    Dim rev_Version As Long
    maj_Version = tApplication.VersionMajor
    min_Version = tApplication.VersionMinor
    rev_Version = tApplication.VersionRevision
    VersionString = (maj_Version & "." & min_Version & "." & rev_Version)
End Property
Public Property Get MajorVersion() As Long
CheckIfNull
    MajorVersion = tApplication.VersionMajor
End Property
Public Property Get MinorVersion() As Long
CheckIfNull
    MinorVersion = tApplication.VersionMinor
End Property
Public Property Get Revision() As Long
CheckIfNull
    Revision = tApplication.VersionRevision
End Property
Public Property Get Author() As String
CheckIfNull
    Author = tApplication.Author
End Property
Public Property Get LegalCopyright() As String
CheckIfNull
    LegalCopyright = tApplication.LegalCopyright
End Property
Public Property Get ApplicationPath() As String
CheckIfNull
    ApplicationPath = tApplication.ApplicationPath
End Property
Public Function CheckForUpdates() As Boolean
CheckIfNull
    'check here....
End Function
Public Sub InstallUpdates()
CheckIfNull

End Sub
Public Sub Install()
CheckIfNull

End Sub
Public Sub Uninstall()
CheckIfNull

End Sub




Public Property Get StandardOutput() As Stream
    Set StandardOutput = OpenProcessStream(API_GetStdHandle(STD_OUTPUT_HANDLE), sdOutStream, "Standard Output")
End Property
Public Property Get StandardInput() As Stream
    Set StandardInput = OpenProcessStream(API_GetStdHandle(STD_INPUT_HANDLE), sdInStream, "Standard Input")
End Property
Public Property Get StandardError() As Stream
    Set StandardError = OpenProcessStream(API_GetStdHandle(STD_ERROR_HANDLE), sdOutStream, "Standard Error Output")
End Property
Private Function OpenProcessStream(ByVal Handle As Long, ByVal Direction As StreamDirection, ByVal DirName As String) As ProcessStream
    Select Case Direction
        Case sdInStream, sdOutStream, sdBoth
        Case Else
            throw Exps.InvalidArgumentValueException
    End Select
    
    Set OpenProcessStream = New ProcessStream
    Call OpenProcessStream.Constructor0(Handle, Direction, DirName)
End Function



Public Sub AddOperation(Key As String, State As Boolean, Optional Args As ArgumentList = Nothing, Optional ReadOnly As Boolean = False)
    Call mint_application.AddOperation(Key, State, Args, ReadOnly)
End Sub
Public Sub RemoveOperation(Key As String)
    Call mint_application.RemoveOperation(Key)
End Sub
Public Function Operation(Key As String) As Operation
    Set Operation = mint_application.Operation(Key)
End Function
Public Function Operations() As List
    Set Operations = mint_application.Operations
End Function
Public Function IsOperation(Key As String) As Boolean
    IsOperation = (mint_application.IndexOfOperation(Key) >= 0)
End Function


Private Function Language_IndexOf(Key As String) As Long
    Dim i As Long
    For i = 0 To application_lnCount - 1
        If application_ln(i).ShortName = Key Then
            Language_IndexOf = i
            Exit Function
        End If
    Next
    Language_IndexOf = -1
End Function
Public Function tr(Key As String) As String
    tr = CurrentLanguage.translate(Key)
End Function
Public Sub boundtr(Key As String, Object As Object, targetPropertyName As String, Optional SetNow As Boolean = True)
    Call boundtranslationto(Key, Object, targetPropertyName, SetNow)
End Sub
Public Function translate(Key As String) As String
    translate = CurrentLanguage.translate(Key)
End Function
Public Function CurrentLanguage() As Language
    Set CurrentLanguage = mint_application.CurrentLanguage
End Function
Public Sub SetCurrentLanguage(LangName As String)
    Call mint_application.SetCurrentLanguage(LangName)
End Sub
Public Function RegisterLanguagePath(Path As String)
    On Error GoTo Err
    Dim f() As String
    f = Directory(Path).FilteredFileNames(FileFilters(StringArray("*.lang", _
        "*.lng", "*.mlang", "*.mlng", "*.minttranslation"), StringArray("")), vbNormal, True)
    Dim Size As Long
    Size = ArraySize(f)
    RegisterLanguagePath = Size
    Dim i As Long
    Call mySignalEmitter.BlockSignal(SIGNAL_LANGUAGEREGISTERED)
    For i = 0 To Size - 1
        Call RegisterLanguage(Language(f(i)).LoadFromFile)
    Next
    Call mySignalEmitter.UnblockSignal(SIGNAL_LANGUAGEREGISTERED)
    Call mySignalEmitter.Emit(SIGNAL_LANGUAGEPATHREGISTERED)
    Exit Function
Err:
    throw Exps.Exception("Some of language files were load but others fail to load.")
End Function
Public Sub RegisterLanguage(ByVal Language As Language, Optional ByVal SetAsCurrentLanguage As Boolean = False)
    If Language Is Nothing Then throw Exps.InvalidArgumentValueException
    If Not Language.IsLoaded Then Call Language.Load
    ReDim Preserve application_ln(application_lnCount)
    Set application_ln(application_lnCount) = Language
    If application_lnCount = 0 Then SetAsCurrentLanguage = True
    If SetAsCurrentLanguage Then
        application_currentLanguageID = application_lnCount
        Call mint_application.CLanguageChanged
    End If
    application_lnCount = application_lnCount + 1
    Call mySignalEmitter.Emit(SIGNAL_LANGUAGEREGISTERED)
End Sub
Public Sub UnRegisterLanguage(ShortName As String)
    Dim Index As Long
    Dim i As Long
    Index = Language_IndexOf(ShortName)
    If Index = -1 Then throw Exps.ItemNotExistsException
    If Index = application_lnCount - 1 Then
        Call application_ln(Index).Dispose
    Else
        Call application_ln(Index).Dispose
        For i = Index To application_lnCount - 2
            Set application_ln(i) = application_ln(i + 1)
        Next
    End If
    application_lnCount = application_lnCount - 1
    If application_lnCount = 0 Then
        Erase application_ln
    Else
        ReDim Preserve application_ln(application_lnCount - 1)
    End If
    Call mySignalEmitter.Emit(SIGNAL_LANGUAGEUNREGISTERED)
End Sub
Public Property Get LanguageName() As String
    LanguageName = CurrentLanguage.Name
End Property
Public Property Get LanguageShortName() As String
    LanguageShortName = CurrentLanguage.ShortName
End Property
Public Property Get LanguageRegion() As String
    LanguageRegion = CurrentLanguage.Region
End Property
Public Property Get LanguageDescription() As String
    LanguageDescription = CurrentLanguage.Description
End Property
Public Property Get LanguageRightToLeft() As String
    LanguageRightToLeft = CurrentLanguage.RightToLeft
End Property

Public Sub EnableVisualStyles(Optional ByVal alsoCreateManifest As Boolean = False)
    
    Call mySignalEmitter.Emit(SIGNAL_VISUALSTYLESENABLED)
End Sub

Public Function GetAppDataPath(Optional ByVal System As Boolean = False) As String
    'GetAppDataPath = baseMethods2.GetAppDataPath(System)
End Function
Public Function GetDesktopPath(Optional ByVal System As Boolean = False) As String
    'GetDesktopPath = baseMethods2.GetDesktopPath(System)
End Function
Public Function GetLocalUserPath(Optional ByVal System As Boolean = False) As String
    'GetLocalUserPath = baseMethods2.GetLocalUserPath(System)
End Function
Public Function GetApplicationSpecifiedTempPath(Optional ByVal UseRevision As Boolean = True) As String
    If tApplication Is Nothing Then throw Exps.InvalidStatusException("Not any Application running.")
    'GetApplicationSpecifiedTempPath = baseMethods2.GetApplicationSpecifiedTempPath_specified(tApplication.CompanyName, tApplication.ProductName, tApplication.VersionMajor, tApplication.VersionMinor, tApplication.VersionRevision, UseRevision)
End Function
Public Function GetApplicationDataPath(Optional ByVal System As Boolean = False, Optional ByVal UseRevision As Boolean = True) As String
    If tApplication Is Nothing Then throw Exps.InvalidStatusException("Not any Application running.")
    'GetApplicationDataPath = baseMethods2.GetApplicationDataPath_specified(tApplication.CompanyName, tApplication.ProductName, tApplication.VersionMajor, tApplication.VersionMinor, tApplication.VersionRevision, System, UseRevision)
End Function

Private Sub MintAPI_Application_Impl_AppMain(Args As Variant)
    Call Instance.About
End Sub
Private Sub MintAPI_Application_Impl_AppRestart(E As EventArgs)
    Call Instance.About(True)
End Sub

Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter.ExternEmitter
End Function

Private Property Get ISupportSlot_AutoManage() As Boolean
    ISupportSlot_AutoManage = False
End Property
Private Sub ISupportSlot_ManualManageCall(SlotName As String)
    Select Case SlotName
        Case SLOT_TERMINATE
            Call Terminate
        Case SLOT_EXTERMINATE
            Call Exterminate
    End Select
End Sub
Private Function ISupportSlot_Methods() As String
    ISupportSlot_Methods = _
        "0" & SLOT_TERMINATE & _
        " 0" & SLOT_EXTERMINATE
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
