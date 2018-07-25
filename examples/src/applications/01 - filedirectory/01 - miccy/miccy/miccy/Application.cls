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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements MintAPI.IApplication

Private Sub IApplication_Message(ByVal e As MintAPI.ApplicationMessageEventArgs)
    
End Sub

Private Sub IApplication_Starting(ByVal e As EventArgs)
    '
End Sub
Private Sub IApplication_Started(ByVal e As EventArgs)
    '
End Sub
Private Sub IApplication_Main(args As Variant)
    
End Sub
Private Sub IApplication_Restart(ByVal e As EventArgs)
    '
End Sub
Private Sub IApplication_ApplicationConfigurationChanged(ByVal e As EventArgs)
    '
End Sub
Private Sub IApplication_ApplicationInputMethodChanged(ByVal e As EventArgs)
    '
End Sub
Private Sub IApplication_ApplicationLanguageChanged(ByVal e As EventArgs)
    '
End Sub
Private Sub IApplication_ApplicationNetworkStateChanged(ByVal e As EventArgs)
    '
End Sub
Private Sub IApplication_Debugger(ByVal e As ExceptionOccuredEventArgs)
    '
End Sub
Private Sub IApplication_Error(ByVal e As ExceptionOccuredEventArgs)
    '
End Sub
Private Sub IApplication_Terminating(ByVal e As DisposingEventArgs)
    '
End Sub
Private Sub IApplication_Terminated(ByVal e As EventArgs)
    '
End Sub

Private Function IApplication_GUID() As String
    IApplication_GUID = "{blabla-blabla-blabla-blabla}"
End Function

Private Function IApplication_ApplicationBinPath() As String
    IApplication_ApplicationBinPath = App.Path
End Function
Private Function IApplication_ApplicationConfigurationPath() As String
    IApplication_ApplicationConfigurationPath = Directory.ConcatPath(App.Path, "config.ini")
End Function
Private Function IApplication_ApplicationConfigurationRegistryPath(Optional ByVal System As Boolean = True) As String
    IApplication_ApplicationConfigurationRegistryPath = ""
End Function
Private Function IApplication_ApplicationDataPath() As String
    IApplication_ApplicationDataPath = Application.GetApplicationDataPath 'GetApplicationDataPath_specified(CompanyName, ProductName, MajorVersion, MinorVersion, Revision, False, False)
End Function
Private Function IApplication_ApplicationLogFile() As String
    IApplication_ApplicationLogFile = Directory.ConcatPath(App.Path, "log.txt")
End Function
Private Function IApplication_ApplicationPath() As String
    IApplication_ApplicationPath = App.Path
End Function
Private Function IApplication_ApplicationPluginsPath() As String
    IApplication_ApplicationPluginsPath = Directory.ConcatPath(App.Path, "plugins")
End Function
Private Function IApplication_ApplicationRegistryPath(Optional ByVal System As Boolean = True) As String
    IApplication_ApplicationRegistryPath = "HKEY_LOCALE_MACHINE/Software/miccy"
End Function
Private Function IApplication_ApplicationTempPath() As String
    IApplication_ApplicationTempPath = Application.GetApplicationSpecifiedTempPath(True)
End Function
Private Function IApplication_ExecutionPath() As String
    IApplication_ExecutionPath = Directory.ConcatPath(App.Path, App.EXEName)
End Function

Private Function IApplication_AppDomain() As String
    IApplication_AppDomain = "net.miccyultimatetools"
End Function
Private Function IApplication_ApplicationName() As String
    IApplication_ApplicationName = App.Title
End Function
Private Function IApplication_ProductName() As String
    IApplication_ProductName = "miccyUltimateTools"
End Function
Private Function IApplication_VersionString() As String
    IApplication_VersionString = ""
End Function
Private Function IApplication_VersionTags() As String
    IApplication_VersionTags = ""
End Function
Private Function IApplication_VersionMajor() As Integer
    IApplication_VersionMajor = App.Major
End Function
Private Function IApplication_VersionMinor() As Long
    IApplication_VersionMinor = App.Minor
End Function
Private Function IApplication_VersionRevision() As Long
    IApplication_VersionRevision = ""
End Function
Private Function IApplication_Author() As String
    IApplication_Author = "Ali Mousavi Kherad"
End Function
Private Function IApplication_CompanyName() As String
    IApplication_CompanyName = "Ali Mousavi Kherad"
End Function
Private Function IApplication_LegalCopyright() As String
    IApplication_LegalCopyright = "by Ali Mousavi Kherad"
End Function
Private Function IApplication_ProductCode() As String
    IApplication_ProductCode = IApplication_ProductCode50chars
End Function
Private Function IApplication_ProductCode50chars() As String
    IApplication_ProductCode50chars = "345sfd455ydfu65u659hk"
End Function
Private Function IApplication_Website() As String
    IApplication_Website = ""
End Function
Private Function IApplication_ServiceWebsite() As String
    IApplication_ServiceWebsite = ""
End Function
Private Function IApplication_SupportLink() As String
    IApplication_SupportLink = ""
End Function
Private Function IApplication_UpdateLink() As String
    IApplication_UpdateLink = ""
End Function
Private Function IApplication_HelpLink() As String
    IApplication_HelpLink = "[No Help!]"
End Function


