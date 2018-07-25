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
Option Explicit
Implements IApplication

Const VERSIONSTRING As String = "0.0.1.2013 twoflowers - advanced"

Const COMPILEYEAR As String = "2013"
Const COMPILEDATE As String = "0/0/2013"
Const COMPILERINFO As String = "VB6 Advanced"
Const COMPILEBY As String = "Ali Mousavi Kherad | alimousavikherad@gmail.com"


Private Sub IApplication_Main(args As Variant)
    'code here.
    'Main Application Entry...
    Dim C As Console
    Set C = Application.appOutput
    C.Title = "Mouse Position Capture..."
    '=======================================
looper:
    DoEvents
    stdout sclr, "fps:1.5", endl, Environment.MousePosition.Left, "x", Environment.MousePosition.Top
    Thread.Sleep 333
    GoTo looper
    '=======================================
    stdout "Press Any Key to Continue...", endl
    C.ReadKey
    
    stdout "Main()", endl
    
    Call Application.DisposeApplication
End Sub

Private Sub IApplication_Message(ByVal E As MintAPI.ApplicationMessageEventArgs)

End Sub

Private Sub IApplication_Restart(ByVal E As MintAPI.EventArgs)
    stdout "Restart...", endl
    'code here.
End Sub
Private Sub IApplication_Starting(ByVal E As MintAPI.EventArgs)
    stdout "Starting...", endl
    'code here.
End Sub
Private Sub IApplication_Started(ByVal E As MintAPI.EventArgs)
    stdout "Started...", endl
    'code here.
End Sub
Private Sub IApplication_ApplicationInputMethodChanged(ByVal E As MintAPI.EventArgs)
    'code here.
End Sub
Private Sub IApplication_ApplicationNetworkStateChanged(ByVal E As MintAPI.EventArgs)
    stdout "Internet Connection State:", NetAPI.InternetConnectionStatus, endl
    'code here.
End Sub
Private Sub IApplication_ApplicationConfigurationChanged(ByVal E As MintAPI.EventArgs)
    stdout "Current Configuration:", Configuration.ConfigurationType, endl
    'code here.
End Sub
Private Sub IApplication_ApplicationLanguageChanged(ByVal E As MintAPI.EventArgs)
    stdout "Current Language:", Application.CurrentLanguage.Name, endl
    'code here.
End Sub
Private Sub IApplication_Debugger(ByVal E As MintAPI.ExceptionOccuredEventArgs)
    Call MintAPI.Log(E.Exception.Message, Colors.cOrange, Me)
    'code here.
End Sub
Private Sub IApplication_Error(ByVal E As MintAPI.ExceptionOccuredEventArgs)
    Call MintAPI.Log(E.Exception.Message, Colors.cRed, Me)
    'code here.
End Sub
Private Sub IApplication_Terminating(ByVal E As MintAPI.DisposingEventArgs)
    stdout "Terminating ...", endl
    'code here.
End Sub
Private Sub IApplication_Terminated(ByVal E As MintAPI.EventArgs)
    stdout "Terminated.", endl
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

