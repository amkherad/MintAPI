VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------
'  MintAPI by Ali Mousavi Kherad
'  - alimousavikherad@gmail.com
'----------------------------------------------

Option Explicit
Implements IApplication
Implements IPlugin

Const VERSIONSTRING As String = "0.0.1.2013 twoflowers - advanced"

Const COMPILEYEAR As String = "2013"
Const COMPILEDATE As String = "0/0/2013"
Const COMPILERINFO As String = "VB6 Advanced"
Const COMPILEBY As String = "Compiled by | support@test.com"


Private Sub IApplication_Main(Args As Variant)
    'code here.
    'Main Application Entry...
    
    
End Sub
Private Sub IApplication_Message(ByVal e As MintAPI.ApplicationMessageEventArgs)
    'code here.
End Sub
Private Sub IApplication_Restart(ByVal e As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_Starting(ByVal e As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_Started(ByVal e As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_ApplicationInputMethodChanged(ByVal e As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_ApplicationNetworkStateChanged(ByVal e As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_ApplicationConfigurationChanged(ByVal e As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_ApplicationLanguageChanged(ByVal e As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_Debugger(ByVal e As MintAPI.ExceptionOccuredEventArgs)
    'code here.
End Sub
Private Sub IApplication_Error(ByVal e As MintAPI.ExceptionOccuredEventArgs)
    'code here.
End Sub
Private Sub IApplication_Terminating(ByVal e As MintAPI.CancelEventArgs)
    'code here.
End Sub
Private Sub IApplication_Terminated(ByVal e As MintAPI.EventArgs)
    'code here.
End Sub

Private Function IApplication_AppDomain() As String
    IApplication_AppDomain = "com.test.testerapp"
End Function
Private Function IApplication_ExecutionPath() As String
    IApplication_ExecutionPath = Directory.ConcatPath(App.Path, App.EXEName)
End Function
Private Function IApplication_ApplicationPath() As String
    IApplication_ApplicationPath = App.Path
End Function
Private Function IApplication_ApplicationPluginsPath() As String
    IApplication_ApplicationPluginsPath = Directory.ConcatPath(App.Path, "plugins")
End Function
Private Function IApplication_ApplicationRegistryPath(Optional ByVal System As Boolean = True) As String
    IApplication_ApplicationRegistryPath = "HKEY_LOCAL_MACHINE\Software\MintAPI\tester\"
End Function
Private Function IApplication_ApplicationTempPath() As String
    IApplication_ApplicationTempPath = Application.GetApplicationSpecifiedTempPath(False)
End Function
Private Function IApplication_Author() As String
    IApplication_Author = "Ali Mousavi Kherad"
End Function
Private Function IApplication_CompanyName() As String
    IApplication_CompanyName = "MintAPI Project"
End Function
Private Function IApplication_ApplicationLogFile() As String
    IApplication_ApplicationLogFile = Directory.ConcatPath(Application.GetAppDataPath(False), "log.txt")
End Function
Private Function IApplication_ApplicationName() As String
    IApplication_ApplicationName = "MintAPI_Project"
End Function
Private Function IApplication_ApplicationBinPath() As String
    IApplication_ApplicationBinPath = App.Path
End Function
Private Function IApplication_ApplicationConfigurationPath() As String
    IApplication_ApplicationConfigurationPath = Directory.ConcatPath(App.Path, "config.ini")
End Function
Private Function IApplication_ApplicationConfigurationRegistryPath(Optional ByVal System As Boolean = True) As String
    IApplication_ApplicationConfigurationRegistryPath = "HKEY_LOCAL_MACHINE\Software\Test\"
End Function
Private Function IApplication_ApplicationDataPath() As String
    IApplication_ApplicationDataPath = Application.GetAppDataPath(True)
End Function
Private Function IApplication_GUID() As String
    IApplication_GUID = "{xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}"
End Function

Private Function IApplication_HelpLink() As String
    IApplication_HelpLink = "test.com/help"
End Function
Private Function IApplication_LegalCopyright() As String
    IApplication_LegalCopyright = "Ali Mousavi Kherad"
End Function
Private Function IApplication_ProductCode() As String
    IApplication_ProductCode = ""
End Function
Private Function IApplication_ProductCode50chars() As String
    IApplication_ProductCode50chars = ""
End Function
Private Function IApplication_ProductName() As String
    IApplication_ProductName = App.ProductName
End Function
Private Function IApplication_ServiceWebsite() As String
    IApplication_ServiceWebsite = "support.test.com"
End Function
Private Function IApplication_SupportLink() As String
    IApplication_SupportLink = "support.test.com"
End Function
Private Function IApplication_UpdateLink() As String
    IApplication_UpdateLink = "update.test.com"
End Function
Private Function IApplication_VersionMajor() As Integer
    IApplication_VersionMajor = App.Major
End Function
Private Function IApplication_VersionMinor() As Long
    IApplication_VersionMinor = App.Minor
End Function
Private Function IApplication_VersionRevision() As Long
    IApplication_VersionRevision = App.Revision
End Function
Private Function IApplication_VersionString() As String
    IApplication_VersionString = App.Major & "." & App.Minor & "." & App.Revision & "." & COMPILEYEAR
End Function
Private Function IApplication_VersionTags() As String
    IApplication_VersionTags = "versiontag - advanced"
End Function
Private Function IApplication_Website() As String
    IApplication_Website = "test.com"
End Function

Private Function IPlugin_Application() As MintAPI.IApplication
    Set IPlugin_Application = Me
End Function
Private Function IPlugin_Execute(strExecuteCommand As String, Args() As Variant) As Variant
    
End Function
Private Function IPlugin_GeneralProvider(ProviderEngine As String) As MintAPI.IProvider
    
End Function
Private Function IPlugin_GetObject(Name As String, Args As MintAPI.ArgumentList) As Object
    
End Function
Private Function IPlugin_Main(Args As MintAPI.ArgumentList) As Long
    
End Function
Private Sub IPlugin_StartupEntry(ByVal Provider As MintAPI.Provider, ByVal StandAloneType As MintAPI.StandAloneType, ByVal Args As MintAPI.ArgumentList)
    
End Sub
Private Function IPlugin_SupportStandAlone() As Boolean
    IPlugin_SupportStandAlone = False
End Function