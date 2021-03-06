VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Application generic configuration provider.\r\nThis class provides auto configuration saving and retrieving from diffrent kind of configuration saving places such as registry,application specified .ini file,.xml configuration,... ."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "Configuration"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Public Enum ConfigurationTypes
    ctRegistry
    ctFile_Ini
    ctFile_XML
End Enum

Private Type Column
    Name As String
    Value As Variant
End Type
Private Type Columns
    Col As Column
    colsCount As Long
End Type
Private Type Settings__
    Name As String
    C As Columns
End Type

Dim AppName As String
Dim AppPath As String
Dim plgpath As String
Dim cfgpath As String
Dim regcfgpath As String
Dim regcfgpathp As String

Dim sys As Boolean
Dim ct As ConfigurationTypes

Dim tApp As IApplication

Dim defaultSettings As Settings__
Dim currentSettings As Settings__
Dim pushedSettings() As Settings__
Dim pushedCount As Long


Private Sub Class_Initialize()
    sys = True
    ct = ctFile_Ini
    If Not tApplication Is Nothing Then
        On Error GoTo Err
        AppName = tApplication.ApplicationName
        AppPath = tApplication.ApplicationPath
        plgpath = tApplication.ApplicationPluginsPath
        cfgpath = tApplication.ApplicationConfigurationPath
        regcfgpath = tApplication.ApplicationConfigurationRegistryPath(True)
        regcfgpathp = tApplication.ApplicationConfigurationRegistryPath(False)
        Set tApp = tApplication
    End If
    Exit Sub
Err:
    throw InvalidStatusException
End Sub


Friend Sub Initialize(Optional Arguments) 'is default
    
End Sub
Friend Sub InitializeW(s As Settings__)

End Sub


Public Sub SetApplication(Application As IApplication)
    If Application Is Nothing Then throw ArgumentNullException
    On Error GoTo Err
    AppName = Application.ApplicationName
    AppPath = Application.ApplicationPath
    plgpath = Application.ApplicationPluginsPath
    cfgpath = Application.ApplicationConfigurationPath
    regcfgpath = Application.ApplicationConfigurationRegistryPath(True)
    regcfgpathp = Application.ApplicationConfigurationRegistryPath(False)
    Set tApp = Application
    Exit Sub
Err:
    throw InvalidStatusException("Invalid IApplication class methods call.")
End Sub
Public Sub Set_Values(ApplicationName As String, ApplicationPath As String, PluginPath As String, ConfigurationPath As String, _
                     RegistryConfigurationPath As String, RegistryUserConfigurationPath As String)
    AppName = ApplicationName
    AppPath = ApplicationPath
    plgpath = PluginPath
    cfgpath = ConfigurationPath
    regcfgpath = RegistryConfigurationPath
    regcfgpathp = RegistryUserConfigurationPath
End Sub

Public Sub Set_ApplicationName(ApplicationName As String): AppName = ApplicationName: End Sub
Public Sub Set_ApplicationPath(ApplicationPath As String): AppPath = ApplicationPath: End Sub
Public Sub Set_PluginPath(PluginPath As String): plgpath = PluginPath: End Sub
Public Sub Set_ConfigurationPath(ConfigurationPath As String): cfgpath = ConfigurationPath: End Sub
Public Sub Set_RegistryConfigurationPath(RegistryConfigurationPath As String): regcfgpath = RegistryConfigurationPath: End Sub
Public Sub Set_RegistryUserConfigurationPath(RegistryUserConfigurationPath As String): regcfgpathp = RegistryUserConfigurationPath: End Sub


Public Property Get ConfigurationType() As ConfigurationTypes
    ConfigurationType = ct
End Property
Public Property Let ConfigurationType(Value As ConfigurationTypes)
    ct = Value
End Property


Public Sub LoadSettings(Optional Path)
    
End Sub
Public Sub ResetSettings()
    
End Sub
Public Sub SaveSettings(Optional Path_or_Key)
    
End Sub

Public Sub DeleteBuffer(Name As String)

End Sub
Public Sub BufferSettings(Name As String)

End Sub
Public Sub RestoreSettings(Name As String)

End Sub
Public Sub ClearBufferedSettings()

End Sub

Public Function Default() As Configuration
    Dim C As New Configuration
    Call C.InitializeW(defaultSettings)
    Set Default = C
End Function


Public Sub AddSetting(Name As String, ValueType As VbVarType)

End Sub
Public Sub SetSetting(Name As String, Value)

End Sub
Public Function GetSetting(Name As String, Optional DefaultValue)

End Function


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
