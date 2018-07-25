VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImplementedApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "ImplementedApplication"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Const SIGNAL_START As String = "start"
Const SIGNAL_STARTING As String = "starting"
Const SIGNAL_STARTED As String = "started"
Const SIGNAL_MAIN As String = "main"
Const SIGNAL_MESSAGE As String = "message"
Const SIGNAL_DEBUG As String = "debug"
Const SIGNAL_DISPOSING As String = "disposing"
Const SIGNAL_RESTART As String = "restart"
Const SIGNAL_CONFIGURATIONCHANGED As String = "configurationchanged"
Const SIGNAL_INPUTMETHODCHANGED As String = "inputmethodchanged"
Const SIGNAL_LANGUAGECHANGED As String = "languagechanged"
Const SIGNAL_NETWORKSTATECHANGED As String = "networkstatechanged"
Const SIGNAL_ERROR As String = "error"
Const SIGNAL_TERMINATING As String = "terminating"
Const SIGNAL_TERMINATED As String = "terminated"

Const SLOT_RESTART As String = "restart"
Const SLOT_STARTING As String = "starting"
Const SLOT_STARTED As String = "started"
Const SLOT_MAIN As String = "main"
Const SLOT_MESSAGE As String = "message"
Const SLOT_TERMINATED As String = "terminated"
Const SLOT_TERMINATING As String = "terminating"

Implements IApplication
Implements ISupportSignal
Implements ISupportSlot
Implements IAliasable

Public Event AppStart(E As EventArgs)
Public Event AppStarting(E As EventArgs)
Public Event AppStarted(E As EventArgs)
Public Event appDebug(E As EventArgs)
Public Event AppDisposing(E As DisposingEventArgs)
Public Event AppMain(Args As Variant)
Public Event AppMessage(E As ApplicationMessageEventArgs)
Public Event AppRestart(E As EventArgs)
Public Event AppConfigurationChanged(E As EventArgs)
Public Event AppInputMethodChanged(E As EventArgs)
Public Event AppLanguageChanged(E As EventArgs)
Public Event AppNetworkStateChanged(E As EventArgs)
Public Event AppDebugger(E As ExceptionOccuredEventArgs)
Public Event appError(E As ExceptionOccuredEventArgs)
Public Event AppTerminating(E As DisposingEventArgs)
Public Event AppTerminated(E As EventArgs)


Dim AppApp
Dim Appargs
Public AppName As String
Public AppProductName As String
Public AppDomain As String
Public AppBinPath As String
Public AppExeName As String
Public AppPath As String
Public AppRegConfigPath As String
Public AppRegistryPath As String
Public AppConfigPath As String
Public AppDataPath As String
Public AppLogPath As String
Public AppPluginsPath As String
Public AppTempPath As String
Public AppMajorVersion As Long
Public AppMinorVersion As Long
Public AppRevision As Long
Public AppVersionTags As String
Public AppCompanyName As String
Public AppCopyright As String
Public AppProductCode As String
Public AppAuthor As String
Public AppGUID As String
Public AppProcess As Process
Public AppWebsite As String
Public AppHelpLink As String
Public AppUpdateLink As String
Public AppSupportLink As String
Public AppServiceWebsite As String

Dim IsInited As Boolean
Dim IsEditable As Boolean

Dim mySignalEmitter As New SignalEmitter

Private Sub Class_Initialize()
    Call mySignalEmitter.Initialize(Me)
End Sub

Friend Sub Initialize(Optional Process As Process, Optional AppDomain_ As String, _
        Optional AppName_ As String, Optional ProductName_ As String, _
        Optional AppPath_ As String, Optional BinaryPath_ As String, _
        Optional ExecutableName As String, Optional AppRegistryPath_ As String, _
        Optional AppRegistryConfigPath_ As String, Optional AppConfigPath_ As String, _
        Optional AppDataPath_ As String, Optional AppLogPath_ As String, _
        Optional AppPluginsPath_ As String, Optional AppTempPath_ As String, _
        Optional CompanyName_ As String, _
        Optional MajorVersion_ As Long, Optional MinorVersion_ As Long, _
        Optional Revision_ As Long, Optional VersionTags_ As String, _
        Optional Copyright_ As String, Optional Author_ As String, _
        Optional ProductCode_ As String, Optional strGUID As String, _
        Optional Args_, Optional App_, Optional AppWebsite_ As String, _
        Optional AppHelpLink_ As String, Optional AppUpdateLink_ As String, _
        Optional AppSupportLink_ As String, Optional AppServiceWebsite_ As String, _
        Optional Editable As Boolean = True)

    Set AppProcess = Process
    If Not IsMissing(App_) Then
        Set AppApp = App_
    End If
    AppDomain = AppDomain_
    AppName = AppName_
    AppBinPath = BinaryPath_
    AppExeName = ExecutableName
    AppRegistryPath = AppRegistryPath_
    AppRegConfigPath = AppRegistryConfigPath_
    AppLogPath = AppLogPath_
    AppTempPath = AppTempPath_
    AppPluginsPath = AppPluginsPath_
    AppConfigPath = AppConfigPath_
    AppProductName = ProductName_
    AppCompanyName = CompanyName_
    AppVersionTags = VersionTags_
    AppMajorVersion = MajorVersion_
    AppMinorVersion = MinorVersion_
    AppRevision = Revision_
    If Not IsMissing(Args_) Then
        If VarType(Args_) = VBObject Then
            Set Appargs = Args_
        Else
                Appargs = Args_
        End If
    End If
    AppGUID = strGUID
    AppCopyright = Copyright_
    AppAuthor = Author_
    
    AppWebsite = AppWebsite_
    AppHelpLink = AppHelpLink_
    AppUpdateLink = AppUpdateLink_
    AppSupportLink = AppSupportLink_
    AppServiceWebsite = AppServiceWebsite_
    
    IsInited = True
    IsEditable = Editable
End Sub
Public Property Get Initialized() As Boolean
    Initialized = IsInited
End Property
Public Property Get Editable() As Boolean
    Editable = IsEditable
End Property

Private Sub CheckIfNotEditable()
    
End Sub

Public Property Get ApplicationVBAAppProperty() As Object
CheckIfNotEditable
    Set ApplicationVBAAppProperty = AppApp
End Property
Public Property Let ApplicationVBAAppProperty(Value As Object)
    Set AppApp = Value
End Property
Public Property Set ApplicationVBAAppProperty(Value As Object)
    Set AppApp = Value
End Property
Public Property Get ApplicationArguments() As Variant
    If VarType(Appargs) = VBObject Then
        Set ApplicationArguments = Appargs
    Else
            ApplicationArguments = Appargs
    End If
End Property
Public Property Let ApplicationArguments(Value As Variant)
    If VarType(Appargs) = VBObject Then
        Set Appargs = Value
    Else
            Appargs = Value
    End If
End Property
Public Property Set ApplicationArguments(Value As Variant)
    Set Appargs = Value
End Property

Private Sub IApplication_Starting(ByVal E As EventArgs)
    RaiseEvent AppStarting(E)
    Call mySignalEmitter.Emit(SIGNAL_STARTING)
End Sub
Private Sub IApplication_Started(ByVal E As EventArgs)
    RaiseEvent AppStarted(E)
    Call mySignalEmitter.Emit(SIGNAL_STARTED)
End Sub
Private Sub IApplication_Main(Args As Variant)
    RaiseEvent AppMain(Args)
    Call mySignalEmitter.Emit(SIGNAL_MAIN)
End Sub
Private Sub IApplication_Message(ByVal E As ApplicationMessageEventArgs)
    RaiseEvent AppMessage(E)
    Call mySignalEmitter.Emit(SIGNAL_MESSAGE)
End Sub
Private Sub IApplication_Restart(ByVal E As EventArgs)
    RaiseEvent AppRestart(E)
    Call mySignalEmitter.Emit(SIGNAL_RESTART)
End Sub
Private Sub IApplication_ApplicationConfigurationChanged(ByVal E As EventArgs)
    RaiseEvent AppConfigurationChanged(E)
    Call mySignalEmitter.Emit(SIGNAL_CONFIGURATIONCHANGED)
End Sub
Private Sub IApplication_ApplicationInputMethodChanged(ByVal E As EventArgs)
    RaiseEvent AppInputMethodChanged(E)
    Call mySignalEmitter.Emit(SIGNAL_INPUTMETHODCHANGED)
End Sub
Private Sub IApplication_ApplicationLanguageChanged(ByVal E As EventArgs)
    RaiseEvent AppLanguageChanged(E)
    Call mySignalEmitter.Emit(SIGNAL_LANGUAGECHANGED)
End Sub
Private Sub IApplication_ApplicationNetworkStateChanged(ByVal E As EventArgs)
    RaiseEvent AppNetworkStateChanged(E)
    Call mySignalEmitter.Emit(SIGNAL_NETWORKSTATECHANGED)
End Sub
Private Sub IApplication_Debugger(ByVal E As ExceptionOccuredEventArgs)
    RaiseEvent AppDebugger(E)
    Call mySignalEmitter.Emit(SIGNAL_DEBUG)
End Sub
Private Sub IApplication_Error(ByVal E As ExceptionOccuredEventArgs)
    RaiseEvent appError(E)
    Call mySignalEmitter.Emit(SIGNAL_ERROR)
End Sub
Private Sub IApplication_Terminating(ByVal E As DisposingEventArgs)
    RaiseEvent AppTerminating(E)
    Call mySignalEmitter.Emit(SIGNAL_TERMINATING)
End Sub
Private Sub IApplication_Terminated(ByVal E As EventArgs)
    RaiseEvent AppTerminated(E)
    Call mySignalEmitter.Emit(SIGNAL_TERMINATED)
End Sub

Private Function IApplication_GUID() As String
    If AppGUID <> "" Then
        IApplication_GUID = AppGUID
    Else
        IApplication_GUID = "" 'generate one
    End If
End Function

Private Function IApplication_ApplicationBinPath() As String
    If AppBinPath <> "" Then
        IApplication_ApplicationBinPath = AppBinPath
    Else
        IApplication_ApplicationBinPath = AppPath
    End If
End Function
Private Function IApplication_ApplicationPath() As String
    If AppPath <> "" Then
        IApplication_ApplicationPath = AppPath
    Else
        IApplication_ApplicationPath = AppBinPath
    End If
End Function
Private Function IApplication_ApplicationConfigurationPath() As String
    If AppConfigPath <> "" Then
        IApplication_ApplicationConfigurationPath = AppConfigPath
    Else
        IApplication_ApplicationConfigurationPath = "" 'generate one
    End If
End Function
Private Function IApplication_ApplicationConfigurationRegistryPath(Optional ByVal System As Boolean = True) As String
    If AppRegConfigPath <> "" Then
        IApplication_ApplicationConfigurationRegistryPath = _
            AppRegConfigPath
    Else
        IApplication_ApplicationConfigurationRegistryPath = "" 'generate one
    End If
End Function
Private Function IApplication_ApplicationDataPath() As String
    If AppDataPath <> "" Then
        IApplication_ApplicationDataPath = AppDataPath
    Else
        IApplication_ApplicationDataPath = GetApplicationDataPath_specified(AppCompanyName, AppProductName, AppMajorVersion, AppMinorVersion, AppRevision, False, False)
    End If
End Function
Private Function IApplication_ApplicationLogFile() As String
    If AppLogPath <> "" Then
        IApplication_ApplicationLogFile = AppLogPath
    Else
        IApplication_ApplicationLogFile = "" 'generate one
    End If
End Function
Private Function IApplication_ApplicationPluginsPath() As String
    If AppPluginsPath <> "" Then
        IApplication_ApplicationPluginsPath = AppPluginsPath
    Else
        IApplication_ApplicationPluginsPath = "" 'generate one
    End If
End Function
Private Function IApplication_ApplicationRegistryPath(Optional ByVal System As Boolean = True) As String
    If AppRegistryPath <> "" Then
        IApplication_ApplicationRegistryPath = AppRegistryPath
    Else
        IApplication_ApplicationRegistryPath = "" 'generate one
    End If
End Function
Private Function IApplication_ApplicationTempPath() As String
    If AppTempPath <> "" Then
        IApplication_ApplicationTempPath = AppTempPath
    Else
        IApplication_ApplicationTempPath = GetApplicationSpecifiedTempPath_specified(AppCompanyName, AppProductName, AppMajorVersion, AppMinorVersion, AppRevision, True)
    End If
End Function
Private Function IApplication_ExecutionPath() As String
    If AppExeName <> "" Then
        IApplication_ExecutionPath = AppExeName
    Else
        IApplication_ExecutionPath = "" 'generate one
    End If
End Function

Private Function IApplication_AppDomain() As String
    If AppDomain <> "" Then
        IApplication_AppDomain = AppDomain
    Else
        IApplication_AppDomain = "" 'generate one
    End If
End Function
Private Function IApplication_ApplicationName() As String
    If AppName <> "" Then
        IApplication_ApplicationName = AppName
    Else
        IApplication_ApplicationName = "" 'generate one
    End If
End Function
Private Function IApplication_ProductName() As String
    If AppProductName <> "" Then
        IApplication_ProductName = AppProductName
    Else
        IApplication_ProductName = ""  'generate one
    End If
End Function
Private Function IApplication_VersionString() As String
    Dim strVersionString As String
    strVersionString = AppMajorVersion & "." & AppMinorVersion & "." & AppRevision
    IApplication_VersionString = strVersionString & IIf(AppVersionTags = "", "", " " & AppVersionTags)
End Function
Private Function IApplication_VersionTags() As String
    IApplication_VersionTags = AppVersionTags
End Function
Private Function IApplication_VersionMajor() As Integer
    IApplication_VersionMajor = AppMajorVersion
End Function
Private Function IApplication_VersionMinor() As Long
    IApplication_VersionMinor = AppMinorVersion
End Function
Private Function IApplication_VersionRevision() As Long
    IApplication_VersionRevision = AppRevision
End Function
Private Function IApplication_Author() As String
    If AppAuthor <> "" Then
        IApplication_Author = AppAuthor
    Else
        IApplication_Author = "" 'generate one
    End If
End Function
Private Function IApplication_CompanyName() As String
    If AppCompanyName <> "" Then
        IApplication_CompanyName = AppCompanyName
    Else
        IApplication_CompanyName = "" 'generate one
    End If
End Function
Private Function IApplication_LegalCopyright() As String
    If AppCopyright <> "" Then
        IApplication_LegalCopyright = AppCopyright
    Else
        IApplication_LegalCopyright = "" 'generate one
    End If
End Function
Private Function IApplication_ProductCode() As String
    If AppProductCode <> "" Then
        IApplication_ProductCode = AppProductCode
    Else
        IApplication_ProductCode = "" 'generate one
    End If
End Function
Private Function IApplication_ProductCode50chars() As String
    Dim bufProductCode As String
    If AppProductCode <> "" Then
        bufProductCode = AppProductCode
    Else
        bufProductCode = "" 'generate one
    End If
    'process bufProductCode here....
    IApplication_ProductCode50chars = bufProductCode
End Function
Private Function IApplication_Website() As String
    If AppWebsite <> "" Then
        IApplication_Website = AppWebsite
    Else
        IApplication_Website = "" 'generate one
    End If
End Function
Private Function IApplication_ServiceWebsite() As String
    If AppServiceWebsite <> "" Then
        IApplication_ServiceWebsite = AppServiceWebsite
    Else
        IApplication_ServiceWebsite = "" 'generate one
    End If
End Function
Private Function IApplication_SupportLink() As String
    If AppSupportLink <> "" Then
        IApplication_SupportLink = AppSupportLink
    Else
        IApplication_SupportLink = "" 'generate one
    End If
End Function
Private Function IApplication_UpdateLink() As String
    If AppUpdateLink <> "" Then
        IApplication_UpdateLink = AppUpdateLink
    Else
        IApplication_UpdateLink = "" 'generate one
    End If
End Function
Private Function IApplication_HelpLink() As String
    If AppHelpLink <> "" Then
        IApplication_HelpLink = AppHelpLink
    Else
        IApplication_HelpLink = "" 'generate one
    End If
End Function



Private Function ISupportSignal_Emitter() As SignalEmitter
    Set ISupportSignal_Emitter = mySignalEmitter
End Function

Private Property Get ISupportSlot_AutoManage() As Boolean
    ISupportSlot_AutoManage = False
End Property
Private Sub ISupportSlot_ManualManageCall(SlotName As String)
    Select Case SlotName
        Case SLOT_MESSAGE
            Call IApplication_Message(EventArgs(Nothing))
        Case SLOT_RESTART
            Call IApplication_Restart(EventArgs(Nothing))
        Case SLOT_MAIN
            Call IApplication_Main(0)
        Case SLOT_STARTED
            Call IApplication_Started(EventArgs(Nothing))
        Case SLOT_STARTING
            Call IApplication_Starting(EventArgs(Nothing))
        Case SLOT_TERMINATED
            Call IApplication_Terminated(EventArgs(Nothing))
        Case SLOT_TERMINATING
            Call IApplication_Terminating(EventArgs(Nothing))
        Case Else
            throw InvalidCallException
    End Select
End Sub
Private Function ISupportSlot_Methods() As String
    ISupportSlot_Methods = _
        "0" & SLOT_MAIN & _
        " 0" & SLOT_MESSAGE & _
        " 0" & SLOT_STARTING & _
        " 0" & SLOT_STARTED & _
        " 0" & SLOT_RESTART & _
        " 0" & SLOT_TERMINATING & _
        " 0" & SLOT_TERMINATED
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
