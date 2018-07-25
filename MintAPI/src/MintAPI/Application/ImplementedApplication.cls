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
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''<class name="ImplementedApplication" tags="">
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
Implements IApplication

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ImplementedApplications" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"


Const SIGNAL_START As String = "start"
Const SIGNAL_STARTING As String = "starting"
Const SIGNAL_STARTED As String = "started"
Const SIGNAL_MAIN As String = "main"
Const SIGNAL_EVENTMESSAGE As String = "eventmessage"
Const SIGNAL_DEBUG As String = "debug"
Const SIGNAL_DISPOSING As String = "disposing"
Const SIGNAL_RESTART As String = "restart"
Const SIGNAL_CONFIGURATIONCHANGED As String = "configurationchanged"
Const SIGNAL_INPUTMETHODCHANGED As String = "inputmethodchanged"
Const SIGNAL_LANGUAGECHANGED As String = "languagechanged"
Const SIGNAL_NETWORKSTATECHANGED As String = "networkstatechanged"
Const SIGNAL_ERROR As String = "error"
Const SIGNAL_TERMINATED As String = "terminated"

Const SLOT_RESTART As String = "restart"
Const SLOT_STARTING As String = "starting"
Const SLOT_STARTED As String = "started"
Const SLOT_MAIN As String = "main"
Const SIGNAL_EVENTMESSAGE As String = "eventmessage"
Const SLOT_TERMINATED As String = "terminated"
Const SLOT_TERMINATING As String = "terminating"
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
'<section Events>
'
Public Event AppStart(ByVal E As EventArgs)
Public Event AppStarting(ByVal E As EventArgs)
Public Event AppStarted(ByVal E As EventArgs)
Public Event AppDebug(ByVal E As EventArgs)
Public Event AppDisposing(ByVal E As CancelEventArgs)
Public Event AppMain(Args As Variant)
Public Event AppEventMessage(ByVal E As ApplicationMessageEventArgs)
Public Event AppRestart(ByVal E As EventArgs)
Public Event AppConfigurationChanged(ByVal E As EventArgs)
Public Event AppDebugger(ByVal E As DebuggerNotificationEventArgs)
Public Event AppError(ByVal E As ExceptionOccuredEventArgs)
Public Event AppTerminating(ByVal E As CancelEventArgs)
Public Event AppTerminated(ByVal E As EventArgs)
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

Dim IsEditable As Boolean

Dim Process                                     As Process
Dim AppInfos                                    As ApplicationInfo
Dim Args
Dim App

'Dim AppAlias                                    As String
'Dim Name                                        As String
'Dim AppBinPath                                  As String
'Dim AppConfigurationPath                        As String
'Dim AppConfigurationRegistryPathSystem          As String
'Dim AppConfigurationRegistryPath                As String
'Dim AppDataPath                                 As String
'Dim AppLogFile                                  As String
'Dim AppPath                                     As String
'Dim AppPluginsPath                              As String
'Dim AppRegistryPathSystem                       As String
'Dim AppRegistryPath                             As String
'Dim AppTempPath                                 As String
'Dim Author                                      As String
'Dim CompanyName                                 As String
'Dim ExecutionPath                               As String
'Dim Guid                                        As Guid
'Dim LegalCopyright                              As String
'Dim ProductCode                                 As String
'Dim ProductCode50chars                          As String
'Dim ProductName                                 As String
'Dim Version                                     As Version
'Dim ExecutableName                              As String
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
Friend Sub Constructor0(ByVal Process_ As Process, _
    ByRef ApplicationInfo As ApplicationInfo, Args_, App_, _
    AppLinks_ As ApplicationLinks, _
    ByVal Editable As Boolean)
    
    
    Set Process = Process_
    AppInfos = ApplicationInfo
    Call API_VariantCopy(Args, Args_)
    Call API_VariantCopy(App, App_)
    IsEditable = Editable
    
    This_Constructed = True
End Sub
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

Public Property Get Initialized() As Boolean
    Initialized = This_Constructed
End Property
Public Property Get Editable() As Boolean
    Editable = IsEditable
End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IApplication Implementation>
'

Private Function IApplication_GetApplicationInfo() As ApplicationInfo
    IApplication_GetApplicationInfo = AppInfo
End Function


Private Sub IApplication_Starting(ByVal E As EventArgs)
    RaiseEvent AppStarting(E)
    Call Signal.Emit(Me, SIGNAL_STARTING)
End Sub
Private Sub IApplication_Started(ByVal E As EventArgs)
    RaiseEvent AppStarted(E)
    Call Signal.Emit(Me, SIGNAL_STARTED)
End Sub
Private Sub IApplication_Main(Args As Variant)
    RaiseEvent AppMain(Args)
    Call Signal.Emit(Me, SIGNAL_MAIN)
End Sub
Private Sub IApplication_Message(ByVal E As ApplicationMessageEventArgs)
    RaiseEvent AppMessage(E)
    Call Signal.Emit(Me, SIGNAL_MESSAGE)
End Sub
Private Sub IApplication_Restart(ByVal E As EventArgs)
    RaiseEvent AppRestart(E)
    Call Signal.Emit(Me, SIGNAL_RESTART)
End Sub
Private Sub IApplication_Debugger(ByVal E As DebuggerNotificationEventArgs)
    RaiseEvent AppDebugger(E)
    Call Signal.Emit(Me, SIGNAL_DEBUG)
End Sub
Private Sub IApplication_Error(ByVal E As ExceptionOccuredEventArgs)
    RaiseEvent AppError(E)
    Call Signal.Emit(Me, SIGNAL_ERROR)
End Sub
Private Sub IApplication_Terminating(ByVal E As CancelEventArgs)
    RaiseEvent AppTerminating(E)
    Call Signal.Emit(Me, SIGNAL_TERMINATING)
End Sub
Private Sub IApplication_Terminated(ByVal E As EventArgs)
    RaiseEvent AppTerminated(E)
    Call Signal.Emit(Me, SIGNAL_TERMINATED)
End Sub


Private Sub IApplication_EventMessage(ByVal E As ApplicationMessageEventArgs)
    RaiseEvent AppEventMessage(E)
    Call Signal.Emit(Me, SIGNAL_EVENTMESSAGE)
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Basic Members>
'
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
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub

Private Sub CheckIfNotEditable()
    If Not This_Constructed Then throw Exps.InvalidOperationException
    If Not IsEditable Then throw Exps.InvalidOperationException
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
