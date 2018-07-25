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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Member1" ,"Version"
'@PROJECT_LICENSE

''<class name="Application" tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

Implements IObject

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Application" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"

Const SIGNAL_START As String = "start"
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


Const SLOT_TERMINATE As String = "terminate"
Const SLOT_EXTERMINATE As String = "exterminate"
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Public Enum ApplicationType
    ActiveXHost = 1
    OwnConsole = 2
    Service = 4
    StaticLinkLibrary = ActiveXHost Or &H10000000
    DynamicLinkLibrary = ActiveXHost Or &H20000000
    ApplicationExecutable = &H40000000
    ApplicationExecutable_Console = ApplicationExecutable Or OwnConsole
    ApplicationExecutable_ActiveXHost = ActiveXHost Or ApplicationExecutable
    ApplicationExecutable_ActiveXHost_Console = ApplicationExecutable_ActiveXHost Or OwnConsole
End Enum
Public Enum ApplicationStreams
    asNull = 0
    asInput = 1
    asOutput = 2
    asError = 4
    asInputOutput = asInput Or asOutput
End Enum
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables/Events>
'
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Dim WithEvents MintAPI_Application_Impl As ImplementedApplication
Attribute MintAPI_Application_Impl.VB_VarHelpID = -1
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Constructors>
'
''<summary>
''  Basic constructor to use by MintAPI class provider,
''  When creating new instance of this class.
''</summary>
''<params>
''  <param name="Arguments">
''      A list of construction arguments.
''  </param>
''</params>
''<cargs order="0" count="1">
''  <carg type="long" name="ID"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)

    This_Constructed = True
End Sub
'Friend Sub Constructor0(ByVal ID As Long)
'
'    This_Constructed = True
'End Sub
Public Function InitializeNewApplication(Optional ByVal Process As Process, Optional ByVal AppDomain As String, _
        Optional ByVal AppName As String, Optional ByVal ProductName As String, _
        Optional ByVal AppPath As String, Optional ByVal BinaryPath As String, _
        Optional ByVal ExecutableName As String, Optional ByVal AppRegistryConfigPath As String, _
        Optional ByVal AppRegistryPath As String, Optional ByVal AppConfigPath As String, _
        Optional ByVal AppDataPath As String, Optional ByVal AppLogPath As String, _
        Optional ByVal AppPluginsPath As String, Optional ByVal AppTempPath As String, _
        Optional ByVal CompanyName As String, Optional ByVal Version As Version = Nothing, _
        Optional ByVal Copyright As String, Optional ByVal Author As String, _
        Optional ByVal ProductCode As String, Optional StrGuid As String, _
        Optional Args, Optional App, Optional ByVal AppWebsite As String, _
        Optional ByVal AppHelpLink As String, Optional ByVal AppUpdateLink As String, _
        Optional ByVal AppSupportLink As String, Optional ByVal AppServiceWebsite As String) _
            As ImplementedApplication
    '-------------------------------------------

    Dim ImplApp As New ImplementedApplication
    '-------------------------------------------ImplApp.Initialize
    'Call ImplApp.Initialize(Process, AppDomain, _
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

Public Sub Run(ByVal Application As IApplication, ByVal TargetApplicationType As ApplicationType, App, Args)
    If Application Is Nothing Then throw Exps.ArgumentNullException("Application")
    If Not TypeOf App Is VB.App Then throw Exps.InvalidArgumentException
    Set AppInfo.TargetApplication = Application
    Call API_VariantCopyInd(AppInfo.UserApp, App)

    If (TargetApplicationType And OwnConsole) = OwnConsole Then _
        Call trigConsole(sdBoth)

    On Error GoTo ErrStarting
    Call Application.Starting(EventArgs(Me))
    GoTo ContinueStarting
ErrStarting:
    throw Exps.Exception(Err.Description)
ContinueStarting:

    AppInfo.Terminated = False
    Call Application.Main(Args)
    Call Application.Started(EventArgs(Me))
    Call myEmit(SIGNAL_START)
End Sub
Public Function Terminate(Optional ByVal IgnoreCancel As Boolean = False) As Boolean
    If AppInfo.Terminated Then
        Terminate = True
        Exit Function
    End If
    If AppInfo.TargetApplication Is Nothing Then
        Terminate = False
        Exit Function
    End If
    Dim ContinueValue As Boolean
    ContinueValue = True
    Call myEmit(SIGNAL_TERMINATING, ContinueValue)
    Dim d As CancelEventArgs
    Set d = CancelEventArgs(AppInfo.TargetApplication, Not ContinueValue)
    Call AppInfo.TargetApplication.Terminating(d)
    If Not IgnoreCancel Then
        If d.Cancel Then
            Terminate = False
            Exit Function
        End If
    End If
    On Error GoTo terminatedError
    Call AppInfo.TargetApplication.Terminated(EventArgs(AppInfo.TargetApplication))
terminatedError:
    AppInfo.Terminated = True
    Terminate = True

    If mint_application.IsDesignMode Then
        Debug.Print "Terminating application..."
        Call EndDebugging
    Else
        Call Process.EndProcess(0)
    End If
    Call myEmit(SIGNAL_TERMINATED)
End Function
Public Sub Exterminate()
    If Not Terminate(True) Then
        throw Exps.Exception("Unable to dispose the application.")
    End If
End Sub

Public Sub UsingPackage(ByVal strPackages As String)
    'core query text runtime script data
End Sub

Public Property Get IsApplicationTerminated() As Boolean
    IsApplicationTerminated = AppInfo.Terminated
End Property

Public Property Get IsDesignMode() As Boolean
    IsDesignMode = mint_application.IsDesignMode
End Property
Public Property Get IsRuntime() As Boolean
    IsRuntime = mint_application.IsRuntime
End Property

Public Sub trigStream(ByVal tStream As IClassStream, Optional ByVal Dir As ApplicationStreams = ApplicationStreams.asOutput)
    Call mint_application.trigStream(tStream, Dir)
    If tStream Is Nothing Then
        Call myEmit(SIGNAL_NOTHINGTRIGGED)
    Else
        Call myEmit(SIGNAL_STREAMTRIGGED)
    End If
End Sub
Public Sub trigConsole(Optional ByVal Dir As ApplicationStreams = ApplicationStreams.asOutput)
    Call mint_application.trigConsole(Dir)
    Call myEmit(SIGNAL_CONSOLETRIGGED)
End Sub
Public Sub trigFile(ByVal File As File, Optional ByVal Dir As ApplicationStreams = ApplicationStreams.asOutput)
    Call mint_application.trigFile(File, Dir)
    Call myEmit(SIGNAL_FILETRIGGED)
End Sub
Public Sub trigProcess(ByVal P As Process, Optional ByVal Dir As ApplicationStreams = ApplicationStreams.asOutput)
    Call mint_application.trigProcess(P, Dir)
    Call myEmit(SIGNAL_PROCESSTRIGGED)
End Sub
Public Sub trigStdIO(Optional ByVal Dir As ApplicationStreams = ApplicationStreams.asOutput)
    Call mint_application.trigStdIO(Dir)
    Call myEmit(SIGNAL_STDIOTRIGGED)
End Sub


Public Function AppInput() As IClassStream
    Set AppInput = mint_application.AppInput
End Function
Public Function AppOutput() As IClassStream
    Set AppOutput = mint_application.AppOutput
End Function
Public Function AppError() As IClassStream
    Set AppError = mint_application.AppError
End Function
Public Function AppDebug() As IClassStream
    'Set AppDebug = mint_application.AppDebug
End Function
Public Function AppWarning() As IClassStream
    'Set AppWarning = mint_application.AppWarning
End Function
Public Function AppCritical() As IClassStream
    'Set AppCritical = mint_application.AppCritical
End Function
Public Function AppFatal() As IClassStream
    'Set AppFatal = mint_application.AppFatal
End Function

Public Property Get StandardOutput() As Stream
    Set StandardOutput = OpenApplicationStream(API_GetStdHandle(STD_OUTPUT_HANDLE), sdOutStream, "Standard Output")
End Property
Public Property Get StandardInput() As Stream
    Set StandardInput = OpenApplicationStream(API_GetStdHandle(STD_INPUT_HANDLE), sdInStream, "Standard Input")
End Property
Public Property Get StandardError() As Stream
    Set StandardError = OpenApplicationStream(API_GetStdHandle(STD_ERROR_HANDLE), sdOutStream, "Standard Error Output")
End Property
Private Function OpenApplicationStream(ByVal Handle As Long, ByVal Direction As StreamDirection, ByVal DirName As String) As ApplicationStream
    Select Case Direction
        Case sdInStream, sdOutStream, sdBoth
        Case Else
            throw Exps.InvalidArgumentException
    End Select
    
    Set OpenApplicationStream = New ApplicationStream
    Call OpenApplicationStream.Constructor0(Handle, Direction, DirName)
End Function

Public Sub ProccessEvents(Optional TreatAsIdleTime As Boolean = False)
    DoEvents
    Call myEmit(SIGNAL_PROCESSEVENTS)
End Sub


Public Function LoadPlugin(ByVal Path As String, Optional ByVal EntryClassName As String = "") As IPlugin
    
End Function
Public Sub UnloadPlugin(ByVal IPlugin As IPlugin)
    
End Sub


Public Property Get Name() As String: THROW_IfNull
    Name = AppInfo.TargetApplication.Name
End Property
'Public Property Get MintAPI() As IApplication
'    If MintAPI_Application_Impl Is Nothing Then
'        Set MintAPI_Application_Impl = New ImplementedApplication
'MintAPI_Initializer:
'        Call MintAPI_Application_Impl.Constructor0( _
'        Process.Current, APPLICATIONDOMAIN, APPLICATIONID, APPLICATIONID, _
'        App.Path, ConcatPath(App.Path, App.EXEName), App.EXEName, _
'        APP_REGISTRYPATH, APP_REGISTRYPATH, mint_config.ConfigPath, _
'        mint_config.MintAPIDataPath, mint_config.MintAPILogPath, _
'        mint_config.MintAPIPluginsPath, mint_config.MintAPITempPath, _
'        App.CompanyName, App.Major, App.Minor, App.Revision, _
'        APP_VERSIONTAGS, App.LegalCopyright, App.LegalTrademarks, _
'        APP_PRODUCTCODE, APP_GUID, Command, App, APP_WEBSITE, _
'        APP_HELPLINK, APP_UPDATELINK, APP_SUPPORTLINK, _
'        APP_SERVICEWEBSITE, False)
'    End If
'    If Not MintAPI_Application_Impl.Initialized Then _
'        GoTo MintAPI_Initializer
'    Set MintAPI = MintAPI_Application_Impl
'End Property
Public Property Get Version() As Version: THROW_IfNull
    Set Version = AppInfo.TargetApplication.Version
End Property
Public Property Get LegalCopyright() As String: THROW_IfNull
    LegalCopyright = AppInfo.TargetApplication.LegalCopyright
End Property
Public Property Get ApplicationPath() As String: THROW_IfNull
    ApplicationPath = Process.Current.Executable
End Property
Public Function CheckForUpdates() As Boolean: THROW_IfNull
    'check here....
End Function
Public Sub InstallUpdates(): THROW_IfNull

End Sub
Public Sub Install(): THROW_IfNull

End Sub
Public Sub Uninstall(): THROW_IfNull

End Sub


Public Property Get CommandLine() As String
    
End Property
Public Function GetCommandLineArgs() As String()
    
End Function


Public Function Translate(ByVal Key As String) As String
    Translate = mint_application.CurrentTranslation.Translate(Key)
End Function
Public Function tr(ByVal Key As String) As String
    tr = mint_application.CurrentTranslation.Translate(Key)
End Function
Public Sub boundtr(ByVal Key As String, ByVal Object As Object, ByVal TargetPropertyName As String, Optional ByVal SetNow As Boolean = True)
    Call bindtranslationto(Key, Object, TargetPropertyName, SetNow)
End Sub
Public Sub SetCurrentTranslation(ByVal LangName As String)
    Call mint_application.SetCurrentTranslation(LangName)
End Sub
Public Function RegisterTranslationPath(ByVal Path As String)
    
End Function
Public Sub RegisterTranslation(ByVal Translation As Translation, Optional ByVal SetAsCurrentLanguage As Boolean = False)
    
End Sub
Public Sub UnRegisterTranslation(ByVal ShortName As String)
    
End Sub


Public Sub EnableVisualStyles(Optional ByVal alsoCreateManifest As Boolean = False)
    
    Call myEmit(SIGNAL_VISUALSTYLESENABLED)
End Sub


Public Property Get AppSharedPath() As String
    
End Property

'Public Function GetAppDataPath(Optional ByVal System As Boolean = False) As String
'    'GetAppDataPath = baseMethods2.GetAppDataPath(System)
'End Function
'Public Function GetDesktopPath(Optional ByVal System As Boolean = False) As String
'    'GetDesktopPath = baseMethods2.GetDesktopPath(System)
'End Function
'Public Function GetLocalUserPath(Optional ByVal System As Boolean = False) As String
'    'GetLocalUserPath = baseMethods2.GetLocalUserPath(System)
'End Function
'Public Function GetApplicationSpecifiedTempPath(Optional ByVal UseRevision As Boolean = True) As String
'    If tApplication Is Nothing Then throw Exps.InvalidStatusException("Not any Application running.")
'    'GetApplicationSpecifiedTempPath = baseMethods2.GetApplicationSpecifiedTempPath_specified(tApplication.CompanyName, tApplication.ProductName, tApplication.VersionMajor, tApplication.VersionMinor, tApplication.VersionRevision, UseRevision)
'End Function
'Public Function GetApplicationDataPath(Optional ByVal System As Boolean = False, Optional ByVal UseRevision As Boolean = True) As String
'    If tApplication Is Nothing Then throw Exps.InvalidStatusException("Not any Application running.")
'    'GetApplicationDataPath = baseMethods2.GetApplicationDataPath_specified(tApplication.CompanyName, tApplication.ProductName, tApplication.VersionMajor, tApplication.VersionMinor, tApplication.VersionRevision, System, UseRevision)
'End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Basic Members>
'
Public Property Get Alias() As String
    Alias = CLASSALIASSTRING
End Property
Public Function Equals(Value As Variant) As Boolean
    Equals = Objects.Equals(Me, Value)
End Function
Public Function ToString() As String
    ToString = CLASSALIAS
End Function
Public Function GetHashCode() As Long
    GetHashCode = Objects.GetHashCode(Me)
End Function
Public Function MetaObject() As MetaObject
    If myMetaObject Is Nothing Then
        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    End If
    Set MetaObject = myMetaObject
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
Private Sub Class_Terminate()
    Call myEmit(SIGNAL_TERMINATING)
    If Not myMetaObject Is Nothing Then
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Private Function myEmit(Signal, ParamArray Args() As Variant)
    
End Function

Private Sub THROW_IfNull()
    If AppInfo.TargetApplication Is Nothing Then _
        throw Exps.InvalidOperationException("Not any application is started.")
End Sub

Private Sub MintAPI_Application_Impl_AppMain(Args As Variant)
    Call Instance.About
End Sub
Private Sub MintAPI_Application_Impl_AppRestart(ByVal E As EventArgs)
    Call Instance.About(True)
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo): Call Constructor(CInfo): End Sub
Private Function IObject_Equals(Value As Variant) As Boolean: IObject_Equals = Equals(Value): End Function
Private Function IObject_GetHashCode() As Long: IObject_GetHashCode = GetHashCode: End Function
Private Function IObject_MetaObject() As MetaObject: Set IObject_MetaObject = MetaObject: End Function
Private Function IObject_ToString() As String: IObject_ToString = ToString: End Function
Private Function IObject_Stub(ByVal ProxyInfo As ObjectProxyInfo) As Boolean
    
    IObject_Stub = False
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
'Private Function ICloneable_Clone() As Object
'    Set ICloneable_Clone = Clone
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
