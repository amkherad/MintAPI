VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MintAPI_IMPLAPPLICATION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Implements IApplication
'
'Private Function IApplication_AppDomain() As String
'    IApplication_AppDomain = "net.Sourceforge.MintAPI"
'End Function
'
'Private Function IApplication_ApplicationBinPath() As String
'    IApplication_ApplicationBinPath = MintAPIRegistry.GetValue("bin", App.Path)
'End Function
'
'Private Function IApplication_ApplicationConfigurationPath() As String
'    IApplication_ApplicationConfigurationPath = MintAPIRegistry.GetValue("config_file", Directory.ConcatPath(App.Path, "config.bin"))
'End Function
'
'Private Function IApplication_ApplicationDataPath() As String
'    IApplication_ApplicationDataPath = MintAPIRegistry.GetValue("data", Directory.ConcatPath(App.Path, "data"))
'End Function
'
'Private Function IApplication_ApplicationLogFile() As String
'    IApplication_ApplicationLogFile = Directory.ConcatPath(App.Path, "log.html")
'End Function
'
'Private Function IApplication_ApplicationName() As String
'    IApplication_ApplicationName = "MintAPI"
'End Function
'
'Private Function IApplication_ApplicationPath() As String
'    IApplication_ApplicationPath = App.Path
'End Function
'
'Private Function IApplication_ApplicationPluginsPath() As String
'    IApplication_ApplicationPluginsPath = MintAPIRegistry.GetValue("plugins", Directory.ConcatPath(App.Path, "plugins"))
'End Function
'
'Private Function IApplication_ApplicationRegistryPath(Optional ByVal System As Boolean = True) As String
'    IApplication_ApplicationRegistryPath = IIf(System, APP_REGISTRYPATH, APP_REGISTRYPATH_USER)
'End Function
'
'Private Function IApplication_ApplicationTempPath() As String
'    IApplication_ApplicationTempPath = baseMethods2.GetApplicationSpecifiedTempPath_specified(IApplication_CompanyName, IApplication_ProductName, IApplication_VersionMajor, IApplication_VersionMinor, IApplication_VersionRevision, False)
'End Function
'
'Private Function IApplication_Author() As String
'    IApplication_Author = "Ali Mousavi Kherad"
'End Function
'
'Private Function IApplication_CompanyName() As String
'    IApplication_CompanyName = "MintAPI / Ali Mousavi Kherad"
'End Function
'
'Private Function IApplication_ExecutionPath() As String
'    IApplication_ExecutionPath = Directory.ConcatPath(App.Path, App.EXEName)
'End Function
'
'Private Function IApplication_HelpLink() As String
'
'End Function
'
'Private Function IApplication_LegalCopyright() As String
'    IApplication_LegalCopyright = "Ali Mousavi Kherad"
'End Function
'
'Private Function IApplication_ProductName() As String
'    IApplication_ProductName = "MintAPI"
'End Function
'
'Private Sub IApplication_Restart(ByVal e As EventArgs)
'
'End Sub
'
'Private Function IApplication_ServiceWebsite() As String
'    IApplication_ServiceWebsite = ""
'End Function
'
'Private Function IApplication_VersionMajor() As Integer
'    IApplication_VersionMajor = App.Major
'End Function
'
'Private Function IApplication_VersionMinor() As Long
'    IApplication_VersionMinor = App.Minor
'End Function
'
'Private Function IApplication_VersionRevision() As Long
'    IApplication_VersionRevision = App.Revision
'End Function
'
'Private Function IApplication_VersionString() As String
'    IApplication_VersionString = APP_VERSIONSTRING
'End Function
'
'Private Function IApplication_VersionTags() As String
'    IApplication_VersionTags = APP_VERSIONTAGS
'End Function
'
'Private Function IApplication_ApplicationConfigurationRegistryPath(Optional ByVal System As Boolean = True) As String
'    IApplication_ApplicationConfigurationRegistryPath = IIf(System, APP_REGISTRYPATH, APP_REGISTRYPATH_USER)
'End Function
'
'Private Function IApplication_ProductCode() As String
'    IApplication_ProductCode = ""
'End Function
'
'Private Function IApplication_ProductCode50chars() As String
'
'End Function
'
'Private Function IApplication_GUID() As String
'    IApplication_GUID = ""
'End Function
'
'Private Function IApplication_UpdateLink() As String
'    IApplication_UpdateLink = ""
'End Function
'
'Private Function IApplication_SupportLink() As String
'
'End Function
'
'Private Function IApplication_Website() As String
'    IApplication_Website = "sourceforge.net/projects/mintapi"
'End Function
'
'Private Sub IApplication_Main(args As Variant)
'
'End Sub
'
'Private Sub IApplication_Started(ByVal e As EventArgs)
'
'End Sub
'
'Private Sub IApplication_Starting(ByVal e As EventArgs)
'
'End Sub
'
'Private Sub IApplication_Terminated(ByVal e As EventArgs)
'
'End Sub
'
'Private Sub IApplication_Terminating(ByVal e As DisposingEventArgs)
'
'End Sub
'
'Private Sub IApplication_ApplicationConfigurationChanged(ByVal e As EventArgs)
'
'End Sub
'
'Private Sub IApplication_ApplicationInputMethodChanged(ByVal e As EventArgs)
'
'End Sub
'
'Private Sub IApplication_ApplicationLanguageChanged(ByVal e As EventArgs)
'
'End Sub
'
'Private Sub IApplication_ApplicationNetworkStateChanged(ByVal e As EventArgs)
'
'End Sub
'
'Private Sub IApplication_Debugger(ByVal e As ExceptionOccuredEventArgs)
'
'End Sub
'
'Private Sub IApplication_Error(ByVal e As ExceptionOccuredEventArgs)
'
'End Sub
'
